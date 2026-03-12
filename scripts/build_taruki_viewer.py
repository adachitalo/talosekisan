#!/usr/bin/env python3
"""垂木拾い3Dビューア: 屋根勾配方向に垂木を自動配置し積算
Usage:
  python build_taruki_viewer.py <input.ifc> <output.html>

垂木配置ルール:
  - 基本ピッチ: 455mm
  - ログ壁の両脇には必ず配置
  - 天窓・煙突の両脇はダブル配置
"""
import sys
import os
import ifcopenshell
import ifcopenshell.geom
import ifcopenshell.util.element
import json
import numpy as np
from collections import defaultdict

WALL_SUB = {"log": "ログ壁", "majikiri": "間仕切壁"}
SKIP = {"I", "kiso", "dodai"}
FULL_WALL_SUB = {"log": "ログ壁", "majikiri": "間仕切壁", "kiso": "基礎", "dodai": "土台", "I": "断熱壁"}
TYPE_NAMES = {
    "IfcWall": "壁", "IfcWallStandardCase": "壁",
    "IfcSlab": "スラブ", "IfcColumn": "柱", "IfcBeam": "梁",
    "IfcDoor": "ドア", "IfcWindow": "窓", "IfcStair": "階段",
    "IfcRailing": "手摺", "IfcRoof": "屋根", "IfcMember": "部材",
    "IfcPlate": "板", "IfcCurtainWall": "カーテンウォール", "IfcFooting": "基礎",
}
CATEGORY_COLORS = {
    "ログ壁": "#D2691E", "間仕切壁": "#CD853F", "基礎": "#808080",
    "土台": "#8B7355", "1F床": "#DEB887", "2F床": "#D2B48C",
    "テラス": "#BC8F8F", "バルコニー": "#C4A882", "屋根": "#8B0000",
    "柱": "#A0522D", "梁": "#B8860B", "ドア": "#4682B4", "窓": "#87CEEB",
    "手摺": "#708090", "階段": "#696969",
}

RAFTER_PITCHES = [0.455, 0.303]  # 455mm, 303mm


def classify_wall_full(name):
    if not name:
        return "壁（その他）"
    for k, v in FULL_WALL_SUB.items():
        if name.startswith(k):
            return v
    return "壁（その他）"


def classify_slab(en, ln):
    n, l = (en or "").lower(), (ln or "").lower()
    if "yane" in n or "屋根" in l:
        return "屋根"
    if "terrace" in n:
        return "テラス"
    if "balcony" in n or "balconi" in n:
        return "バルコニー"
    if "2-yuka" in n or "2f" in n:
        return "2F床"
    if "yuka" in n or "1-yuka" in n:
        return "1F床"
    return "スラブ（その他）"


###############################################################################
# 屋根パラメータ自動検出
###############################################################################

def detect_roof_params(ifc, settings):
    """IFC屋根メッシュから勾配天井パラメータを自動検出。
    面法線で屋根上面を特定し、棟の方向・位置・勾配比を算出。
    軒先Y座標は面法線から算出した勾配比で計算（鼻隠し下端ではなく屋根面の延長）。
    Returns dict with ridge_axis, slope_axis, ridge_pos, ridge_height,
    eave_pos_min, eave_pos_max, eave_height_min, eave_height_max,
    ridge_range_min, ridge_range_max, slope_ratio
    """
    top_verts = []  # 上向き法線の面の頂点（屋根上面）
    # 法線と面積のペアを前面/背面別に保持
    front_normals = []  # slope_axis負方向の面
    back_normals = []   # slope_axis正方向の面

    for elem in list(ifc.by_type("IfcSlab")) + list(ifc.by_type("IfcRoof")):
        ename = (elem.Name or "").lower()
        if elem.is_a("IfcSlab") and "yane" not in ename and "roof" not in ename:
            continue
        try:
            shape = ifcopenshell.geom.create_shape(settings, elem)
            vf = shape.geometry.verts
            ff = shape.geometry.faces
            pts = []
            for i in range(0, len(vf), 3):
                pts.append([vf[i], vf[i+2], -vf[i+1]])  # IFC→Three.js
            pts = np.array(pts)
            for i in range(0, len(ff), 3):
                i0, i1, i2 = ff[i], ff[i+1], ff[i+2]
                p0, p1, p2 = pts[i0], pts[i1], pts[i2]
                normal = np.cross(p1 - p0, p2 - p0)
                norm_len = np.linalg.norm(normal)
                if norm_len < 1e-10:
                    continue
                normal = normal / norm_len
                if normal[1] > 0.3:  # 上向き法線 = 屋根上面
                    top_verts.extend([pts[i0], pts[i1], pts[i2]])
                    area = norm_len / 2
                    front_normals.append((normal, area))  # 後でaxis判定後に振り分け
        except Exception:
            pass

    if not top_verts:
        print("WARNING: No roof top faces found!")
        return None

    tv = np.array(top_verts)
    tv = np.unique(np.round(tv, 3), axis=0)

    print(f"  Roof top vertices: {len(tv)} unique points")
    print(f"    X range: [{tv[:,0].min():.3f}, {tv[:,0].max():.3f}]")
    print(f"    Y range: [{tv[:,1].min():.3f}, {tv[:,1].max():.3f}]")
    print(f"    Z range: [{tv[:,2].min():.3f}, {tv[:,2].max():.3f}]")

    # 棟: Y最高の頂点群
    ridge_height = tv[:, 1].max()
    ridge_tol = 0.15
    ridge_mask = tv[:, 1] > ridge_height - ridge_tol
    ridge_pts = tv[ridge_mask]

    ridge_x_span = ridge_pts[:, 0].max() - ridge_pts[:, 0].min()
    ridge_z_span = ridge_pts[:, 2].max() - ridge_pts[:, 2].min()

    if ridge_x_span > ridge_z_span:
        ridge_axis = "x"
        slope_axis = "z"
        slope_idx = 2  # Z
        ridge_pos = float(np.median(ridge_pts[:, 2]))
        eave_pos_min = float(tv[:, 2].min())
        eave_pos_max = float(tv[:, 2].max())
        ridge_range_min = float(tv[:, 0].min())
        ridge_range_max = float(tv[:, 0].max())
    else:
        ridge_axis = "z"
        slope_axis = "x"
        slope_idx = 0  # X
        ridge_pos = float(np.median(ridge_pts[:, 0]))
        eave_pos_min = float(tv[:, 0].min())
        eave_pos_max = float(tv[:, 0].max())
        ridge_range_min = float(tv[:, 2].min())
        ridge_range_max = float(tv[:, 2].max())

    # 頂点ベースで勾配比を算出（面法線は鼻隠し/破風の面に引きずられるため）
    # 各頂点の slope = (ridge_height - Y) / |ridge_pos - slope_coord| を計算し
    # 最頻値（mode）を採用することで鼻隠し/破風の外れ値を除外
    vertex_slopes = []
    min_dist = 0.3  # 棟から近すぎる頂点は除外（計算誤差が大きい）
    for v in tv:
        sc = v[slope_idx]
        dist = abs(sc - ridge_pos)
        if dist < min_dist:
            continue
        rise = ridge_height - v[1]
        if rise < 0:
            continue
        s = rise / dist
        if 0.1 < s < 3.0:  # 妥当な範囲のみ
            vertex_slopes.append(round(s * 20) / 20)  # 0.05単位に丸め

    slope_ratio = 0.6  # デフォルト（6寸勾配）
    if vertex_slopes:
        # 最頻値を採用（鼻隠し/破風の頂点は異なる値になるので自然に除外される）
        from collections import Counter
        slope_counts = Counter(vertex_slopes)
        mode_slope = slope_counts.most_common(1)[0][0]
        slope_ratio = mode_slope
        slope_angle = np.degrees(np.arctan(slope_ratio))
        print(f"\n  Vertex-based slope computation:")
        print(f"    Total slope samples: {len(vertex_slopes)}")
        print(f"    Top 5 slopes: {slope_counts.most_common(5)}")
        print(f"    Mode slope: {mode_slope:.2f} (angle={slope_angle:.1f}deg)")

    # 軒先Y座標: 勾配比から算出（鼻隠し下端ではなく屋根面の実際の高さ）
    eave_height_min = ridge_height - abs(ridge_pos - eave_pos_min) * slope_ratio
    eave_height_max = ridge_height - abs(eave_pos_max - ridge_pos) * slope_ratio

    result = {
        "ridge_axis": ridge_axis,
        "slope_axis": slope_axis,
        "ridge_pos": ridge_pos,
        "ridge_height": float(ridge_height),
        "eave_pos_min": eave_pos_min,
        "eave_pos_max": eave_pos_max,
        "eave_height_min": float(eave_height_min),
        "eave_height_max": float(eave_height_max),
        "ridge_range_min": ridge_range_min,
        "ridge_range_max": ridge_range_max,
        "slope_ratio": float(slope_ratio),
    }

    print(f"\n  Detected roof parameters:")
    print(f"    Ridge axis: {ridge_axis} (ridge runs along {ridge_axis.upper()})")
    print(f"    Slope axis: {slope_axis} (slope varies along {slope_axis.upper()})")
    print(f"    Ridge position ({slope_axis.upper()}): {ridge_pos:.3f}")
    print(f"    Ridge height (Y): {ridge_height:.3f}")
    print(f"    Eave positions ({slope_axis.upper()}): [{eave_pos_min:.3f}, {eave_pos_max:.3f}]")
    print(f"    Eave heights (Y): min={eave_height_min:.3f}, max={eave_height_max:.3f}")
    print(f"    Slope ratio (rise/run): {slope_ratio:.4f}")
    print(f"    Ridge range ({ridge_axis.upper()}): [{ridge_range_min:.3f}, {ridge_range_max:.3f}]")

    return result


def roof_surface_y(slope_coord, rp):
    """屋根上面のY座標を返す（切妻屋根・勾配比ベース）"""
    return rp["ridge_height"] - abs(slope_coord - rp["ridge_pos"]) * rp["slope_ratio"]


###############################################################################
# ログ壁エッジ検出（垂木配置の基準位置）
###############################################################################

def detect_log_wall_edges_along_ridge(ifc, settings, rp):
    """棟軸方向に沿ったログ壁のエッジ座標（ridge_axis座標）を返す。
    垂木はログ壁の両脇に必ず配置される。
    Returns: list of (ridge_coord_min, ridge_coord_max) for each log wall
    """
    ridge_axis = rp["ridge_axis"]
    slope_axis = rp["slope_axis"]
    wall_edges = []

    for wall in ifc.by_type("IfcWall"):
        ename = wall.Name or ""
        if not ename.startswith("log"):
            continue
        try:
            shape = ifcopenshell.geom.create_shape(settings, wall)
            vf = shape.geometry.verts
            pts = []
            for i in range(0, len(vf), 3):
                pts.append([vf[i], vf[i+2], -vf[i+1]])
            pts = np.array(pts)

            dx = pts[:, 0].max() - pts[:, 0].min()
            dz = pts[:, 2].max() - pts[:, 2].min()

            # 壁の主方向を判定
            if ridge_axis == "x":
                wall_dir = "x" if dx > dz else "z"
            else:
                wall_dir = "z" if dz > dx else "x"

            # 棟軸と垂直な壁（妻壁: slope_axis方向に走る壁）の位置を取得
            if wall_dir == slope_axis:
                # 妻壁: ridge_axis上の位置
                if ridge_axis == "x":
                    rc = float(np.mean([pts[:, 0].min(), pts[:, 0].max()]))
                    wall_edges.append({
                        "type": "gable",
                        "ridge_coord": rc,
                        "width": float(dx),
                    })
                else:
                    rc = float(np.mean([pts[:, 2].min(), pts[:, 2].max()]))
                    wall_edges.append({
                        "type": "gable",
                        "ridge_coord": rc,
                        "width": float(dz),
                    })
            else:
                # eave壁: ridge_axis方向に走る壁 → ridge_axis上のmin/max
                if ridge_axis == "x":
                    r_min = float(pts[:, 0].min())
                    r_max = float(pts[:, 0].max())
                else:
                    r_min = float(pts[:, 2].min())
                    r_max = float(pts[:, 2].max())
                wall_edges.append({
                    "type": "eave",
                    "ridge_min": r_min,
                    "ridge_max": r_max,
                })
        except Exception:
            pass

    return wall_edges


def detect_roof_openings(ifc, settings, rp):
    """天窓（IfcWindow on roof）と煙突を検出。
    Returns: list of {type, ridge_coord_min, ridge_coord_max, slope_coord_min, slope_coord_max}
    """
    ridge_axis = rp["ridge_axis"]
    openings = []

    # 天窓: 屋根面に配置されたIfcWindow（Y座標がeave_height以上）
    for win in ifc.by_type("IfcWindow"):
        try:
            shape = ifcopenshell.geom.create_shape(settings, win)
            vf = shape.geometry.verts
            pts = []
            for i in range(0, len(vf), 3):
                pts.append([vf[i], vf[i+2], -vf[i+1]])
            pts = np.array(pts)

            center_y = (pts[:, 1].min() + pts[:, 1].max()) / 2

            # 屋根上の窓かどうか: Y座標がeave_height付近以上
            eave_h = min(rp["eave_height_min"], rp["eave_height_max"])
            if center_y > eave_h - 0.3:
                if ridge_axis == "x":
                    r_min = float(pts[:, 0].min())
                    r_max = float(pts[:, 0].max())
                    s_min = float(pts[:, 2].min())
                    s_max = float(pts[:, 2].max())
                else:
                    r_min = float(pts[:, 2].min())
                    r_max = float(pts[:, 2].max())
                    s_min = float(pts[:, 0].min())
                    s_max = float(pts[:, 0].max())

                # 窓が屋根面に寝ている（高さが幅/奥行きより小さい）= 天窓
                height_range = pts[:, 1].max() - pts[:, 1].min()
                horiz_range = max(r_max - r_min, s_max - s_min)
                if height_range < horiz_range * 1.5:
                    openings.append({
                        "type": "skylight",
                        "name": win.Name or "天窓",
                        "ridge_min": r_min,
                        "ridge_max": r_max,
                        "slope_min": s_min,
                        "slope_max": s_max,
                    })
        except Exception:
            pass

    # 煙突: IfcBuildingElementProxy or IfcColumn on roof
    for elem_type in ["IfcBuildingElementProxy", "IfcColumn"]:
        for elem in ifc.by_type(elem_type):
            ename = (elem.Name or "").lower()
            if "chimney" in ename or "entotsu" in ename or "煙突" in ename:
                try:
                    shape = ifcopenshell.geom.create_shape(settings, elem)
                    vf = shape.geometry.verts
                    pts = []
                    for i in range(0, len(vf), 3):
                        pts.append([vf[i], vf[i+2], -vf[i+1]])
                    pts = np.array(pts)

                    if ridge_axis == "x":
                        r_min = float(pts[:, 0].min())
                        r_max = float(pts[:, 0].max())
                        s_min = float(pts[:, 2].min())
                        s_max = float(pts[:, 2].max())
                    else:
                        r_min = float(pts[:, 2].min())
                        r_max = float(pts[:, 2].max())
                        s_min = float(pts[:, 0].min())
                        s_max = float(pts[:, 0].max())

                    openings.append({
                        "type": "chimney",
                        "name": elem.Name or "煙突",
                        "ridge_min": r_min,
                        "ridge_max": r_max,
                        "slope_min": s_min,
                        "slope_max": s_max,
                    })
                except Exception:
                    pass

    return openings


###############################################################################
# 垂木配置
###############################################################################

def place_rafters(rp, wall_edges, openings, pitch=0.455):
    """垂木の配置位置（ridge_axis座標）を計算。
    pitch: 垂木ピッチ（m）。デフォルト0.455(455mm)。
    Returns: list of {ridge_coord, reason, double}
    """
    ridge_axis = rp["ridge_axis"]
    ridge_min = rp["ridge_range_min"]
    ridge_max = rp["ridge_range_max"]

    # 0. 屋根両端に必ず配置
    rafter_positions = {}  # ridge_coord → {reason, double}
    rafter_positions[round(ridge_min, 4)] = {"reason": "屋根端部", "double": False}
    rafter_positions[round(ridge_max, 4)] = {"reason": "屋根端部", "double": False}

    # 1. 基本ピッチ配置
    pos = ridge_min
    while pos <= ridge_max + 0.001:
        key = round(pos, 4)
        if key not in rafter_positions:
            rafter_positions[key] = {"reason": "基本ピッチ", "double": False}
        pos += pitch

    # 2. ログ壁エッジに配置（垂木と平行な壁＝勾配方向に走る壁のみ）
    wall_ridge_edges = set()
    for we in wall_edges:
        if we["type"] == "gable":
            # 妻壁（勾配方向に走る壁）: 壁の中心位置の両脇
            rc = we["ridge_coord"]
            half_w = we.get("width", 0.12) / 2
            wall_ridge_edges.add(round(rc - half_w, 4))
            wall_ridge_edges.add(round(rc + half_w, 4))
        # eave壁（棟軸方向に走る壁）は垂木と直交するため、脇配置不要

    for edge in wall_ridge_edges:
        if ridge_min - 0.01 <= edge <= ridge_max + 0.01:
            key = round(edge, 4)
            if key not in rafter_positions:
                rafter_positions[key] = {"reason": "ログ壁脇", "double": False}
            else:
                rafter_positions[key]["reason"] = "ログ壁脇"

    # 3. 天窓・煙突の両脇にダブル配置
    DOUBLE_OFFSET = 0.045  # 45mm offset for double rafter
    for op in openings:
        edge_min = round(op["ridge_min"], 4)
        edge_max = round(op["ridge_max"], 4)

        # 開口部の両端に垂木（シングル→ダブルに昇格 or 新規ダブル）
        for edge in [edge_min, edge_max]:
            if ridge_min - 0.01 <= edge <= ridge_max + 0.01:
                key = round(edge, 4)
                label = "天窓脇" if op["type"] == "skylight" else "煙突脇"
                rafter_positions[key] = {"reason": label, "double": True}

    # ソートして返す
    result = []
    for coord in sorted(rafter_positions.keys()):
        info = rafter_positions[coord]
        result.append({
            "ridge_coord": coord,
            "reason": info["reason"],
            "double": info["double"],
        })

    return result


def generate_rafter_lines(rafters, rp):
    """各垂木のridge_coord位置から、勾配方向に走る3Dラインセグメントを生成。
    切妻屋根の両面にそれぞれ生成。
    Returns: list of {seg: [[x,y,z],[x,y,z]], length_m, reason, double, side}
    """
    ridge_axis = rp["ridge_axis"]
    slope_axis = rp["slope_axis"]
    lines = []

    for r in rafters:
        rc = r["ridge_coord"]
        reason = r["reason"]
        double = r["double"]

        # 各垂木位置で、棟から軒先への2本のライン（左右勾配面）
        for side_label, eave_pos, eave_y in [
            ("前面", rp["eave_pos_min"], rp["eave_height_min"]),
            ("背面", rp["eave_pos_max"], rp["eave_height_max"]),
        ]:
            # 棟での座標
            ridge_y = rp["ridge_height"]

            if ridge_axis == "x":
                p_ridge = [rc, ridge_y, rp["ridge_pos"]]
                p_eave = [rc, eave_y, eave_pos]
            else:
                p_ridge = [rp["ridge_pos"], ridge_y, rc]
                p_eave = [eave_pos, eave_y, rc]

            dx = p_eave[0] - p_ridge[0]
            dy = p_eave[1] - p_ridge[1]
            dz = p_eave[2] - p_ridge[2]
            length = np.sqrt(dx*dx + dy*dy + dz*dz)

            seg = [
                [round(p_ridge[0], 4), round(p_ridge[1], 4), round(p_ridge[2], 4)],
                [round(p_eave[0], 4), round(p_eave[1], 4), round(p_eave[2], 4)],
            ]

            lines.append({
                "seg": seg,
                "length_m": round(length, 4),
                "reason": reason,
                "double": double,
                "side": side_label,
            })

            # ダブル配置: 45mmオフセットした追加垂木（同じ理由名で集計統合）
            if double:
                offset = 0.045  # 45mm
                if ridge_axis == "x":
                    p_ridge2 = [rc + offset, ridge_y, rp["ridge_pos"]]
                    p_eave2 = [rc + offset, eave_y, eave_pos]
                else:
                    p_ridge2 = [rp["ridge_pos"], ridge_y, rc + offset]
                    p_eave2 = [eave_pos, eave_y, rc + offset]

                seg2 = [
                    [round(p_ridge2[0], 4), round(p_ridge2[1], 4), round(p_ridge2[2], 4)],
                    [round(p_eave2[0], 4), round(p_eave2[1], 4), round(p_eave2[2], 4)],
                ]

                lines.append({
                    "seg": seg2,
                    "length_m": round(length, 4),
                    "reason": reason,
                    "double": True,
                    "side": side_label,
                })

    return lines


###############################################################################
# メイン
###############################################################################

def main():
    if len(sys.argv) >= 3:
        IFC_PATH = sys.argv[1]
        OUTPUT_HTML = sys.argv[2]
    else:
        print("Usage: python build_taruki_viewer.py <input.ifc> <output.html>")
        sys.exit(1)

    model_name = os.path.splitext(os.path.basename(IFC_PATH))[0]
    print(f"IFC読み込み中... {IFC_PATH}")
    ifc = ifcopenshell.open(IFC_PATH)
    settings = ifcopenshell.geom.settings()
    settings.set(settings.USE_WORLD_COORDS, True)

    skip_types = {"IfcBuildingElementProxy", "IfcCovering"}
    seen_ids = set()
    meshes = []

    # === Part 1: 全建物ジオメトリ ===
    for ifc_type in TYPE_NAMES:
        if ifc_type in skip_types:
            continue
        for elem in ifc.by_type(ifc_type):
            gid = elem.GlobalId
            if gid in seen_ids:
                continue
            seen_ids.add(gid)
            cat = TYPE_NAMES.get(ifc_type, ifc_type)
            ename = elem.Name or ""
            if cat == "壁":
                if ename.startswith("I"):
                    continue
                cat = classify_wall_full(ename)
            elif cat == "スラブ":
                psets = ifcopenshell.util.element.get_psets(elem)
                layer = psets.get('ArchiCADProperties', {}).get('レイヤー', '')
                cat = classify_slab(ename, layer)
            try:
                shape = ifcopenshell.geom.create_shape(settings, elem)
                vf = shape.geometry.verts
                verts = []
                for i in range(0, len(vf), 3):
                    verts.append(round(vf[i], 4))
                    verts.append(round(vf[i + 2], 4))
                    verts.append(round(-vf[i + 1], 4))
                faces = list(shape.geometry.faces)
                meshes.append({"cat": cat, "name": ename, "gid": gid,
                               "verts": verts, "faces": faces})
            except Exception:
                pass
    print(f"  建物ジオメトリ: {len(meshes)}個")

    # === Part 2: 屋根パラメータ検出 ===
    print("\n屋根パラメータ検出中...")
    rp = detect_roof_params(ifc, settings)
    if rp is None:
        print("ERROR: 屋根が検出できませんでした")
        sys.exit(1)

    # === Part 3: ログ壁エッジ検出 ===
    print("\nログ壁エッジ検出中...")
    wall_edges = detect_log_wall_edges_along_ridge(ifc, settings, rp)
    print(f"  ログ壁エッジ: {len(wall_edges)}個")
    for we in wall_edges:
        if we["type"] == "gable":
            print(f"    妻壁 ridge_coord={we['ridge_coord']:.3f} width={we['width']:.3f}")
        else:
            print(f"    eave壁 range=[{we['ridge_min']:.3f}, {we['ridge_max']:.3f}]")

    # === Part 4: 天窓・煙突検出 ===
    print("\n天窓・煙突検出中...")
    openings = detect_roof_openings(ifc, settings, rp)
    print(f"  開口部: {len(openings)}個")
    for op in openings:
        print(f"    {op['type']}: {op['name']} "
              f"ridge=[{op['ridge_min']:.3f},{op['ridge_max']:.3f}] "
              f"slope=[{op['slope_min']:.3f},{op['slope_max']:.3f}]")

    # === Part 5-6: ピッチ別に垂木配置＆ライン生成 ===
    all_pitch_data = {}  # pitch_mm → {rafters, rafter_lines, summary}
    first_pitch = True

    for pitch in RAFTER_PITCHES:
        pitch_mm = int(pitch * 1000)
        print(f"\n垂木配置計算中... (ピッチ: {pitch_mm}mm)")
        rafters = place_rafters(rp, wall_edges, openings, pitch=pitch)
        print(f"  配置位置: {len(rafters)}箇所")

        n_double = sum(1 for r in rafters if r["double"])
        n_edge = sum(1 for r in rafters if "屋根端部" in r["reason"])
        n_wall = sum(1 for r in rafters if "ログ壁" in r["reason"])
        n_skylight = sum(1 for r in rafters if "天窓" in r["reason"])
        n_chimney = sum(1 for r in rafters if "煙突" in r["reason"])
        n_basic = len(rafters) - n_edge - n_wall - n_skylight - n_chimney
        if first_pitch:
            print(f"    基本ピッチ: {n_basic}箇所")
            print(f"    屋根端部: {n_edge}箇所")
            print(f"    ログ壁脇: {n_wall}箇所")
            print(f"    天窓脇: {n_skylight}箇所")
            print(f"    煙突脇: {n_chimney}箇所")
            print(f"    ダブル配置: {n_double}箇所")

        rafter_lines = generate_rafter_lines(rafters, rp)
        print(f"  垂木ライン: {len(rafter_lines)}本")

        total_count = len(rafter_lines)
        total_length = sum(r["length_m"] for r in rafter_lines)
        single_length = rafter_lines[0]["length_m"] if rafter_lines else 0

        count_by_reason = defaultdict(int)
        length_by_reason = defaultdict(float)
        for r in rafter_lines:
            count_by_reason[r["reason"]] += 1
            length_by_reason[r["reason"]] += r["length_m"]

        count_front = sum(1 for r in rafter_lines if r["side"] == "前面")
        count_back = sum(1 for r in rafter_lines if r["side"] == "背面")

        if first_pitch:
            print(f"\n=== 垂木積算 ({pitch_mm}mm) ===")
            print(f"  垂木1本の長さ（勾配長）: {single_length:.3f}m")
            print(f"  総本数: {total_count}本")
            print(f"  総長さ: {total_length:.2f}m")
            print(f"\n  理由別:")
            for reason in sorted(count_by_reason.keys()):
                print(f"    {reason}: {count_by_reason[reason]}本 / {length_by_reason[reason]:.2f}m")
            print(f"\n  前面: {count_front}本")
            print(f"  背面: {count_back}本")

        summary_info = {
            "total_count": total_count,
            "total_length": round(total_length, 2),
            "single_length": round(single_length, 3),
            "count_by_reason": {k: v for k, v in count_by_reason.items()},
            "length_by_reason": {k: round(v, 2) for k, v in length_by_reason.items()},
            "count_front": count_front,
            "count_back": count_back,
            "pitch": pitch_mm,
        }

        all_pitch_data[pitch_mm] = {
            "rafters": rafters,
            "rafter_lines": rafter_lines,
            "summary": summary_info,
        }
        first_pitch = False

    # === HTML出力 ===
    default_pitch = int(RAFTER_PITCHES[0] * 1000)
    meshes_json = json.dumps(meshes, ensure_ascii=False)
    colors_json = json.dumps(CATEGORY_COLORS, ensure_ascii=False)

    # ピッチ別データをまとめたオブジェクト
    pitch_data_for_html = {}
    for pitch_mm, pdata in all_pitch_data.items():
        pitch_data_for_html[pitch_mm] = {
            "rafters": pdata["rafter_lines"],
            "summary": pdata["summary"],
        }
    pitch_data_json = json.dumps(pitch_data_for_html, ensure_ascii=False)
    pitches_json = json.dumps([int(p * 1000) for p in RAFTER_PITCHES])

    html = generate_html(meshes_json, pitch_data_json, pitches_json,
                         default_pitch, colors_json, model_name)

    os.makedirs(os.path.dirname(OUTPUT_HTML) or ".", exist_ok=True)
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"\n出力: {OUTPUT_HTML}")

    # === JSON集計出力（デフォルトピッチ） ===
    json_path = os.path.splitext(OUTPUT_HTML)[0] + "_summary.json"
    dp = all_pitch_data[default_pitch]
    summary_out = {
        "tool": "垂木拾い",
        "model": model_name,
        "pitch_mm": default_pitch,
        "single_length_m": dp["summary"]["single_length"],
        "total_count": dp["summary"]["total_count"],
        "total_length_m": dp["summary"]["total_length"],
        "count_front": dp["summary"]["count_front"],
        "count_back": dp["summary"]["count_back"],
        "count_by_reason": dp["summary"]["count_by_reason"],
        "length_by_reason": dp["summary"]["length_by_reason"],
        "rafters": [
            {
                "ridge_coord": r["ridge_coord"],
                "reason": r["reason"],
                "double": r["double"],
            }
            for r in dp["rafters"]
        ],
    }
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(summary_out, f, ensure_ascii=False, indent=2)
    print(f"JSON集計: {json_path}")


def generate_html(meshes_json, pitch_data_json, pitches_json,
                   default_pitch, colors_json, model_name=""):
    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>垂木拾い 3Dビューア - {model_name}</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ background:#1a1a2e; overflow:hidden; font-family:Arial,sans-serif; }}
#info {{ position:fixed; top:10px; left:10px; color:#fff; background:rgba(0,0,0,0.75);
  padding:12px 16px; border-radius:8px; font-size:13px; z-index:100; max-width:360px; }}
#info h3 {{ margin-bottom:6px; color:#ffa500; font-size:15px; }}
#legend {{ position:fixed; top:10px; right:10px; color:#fff; background:rgba(0,0,0,0.75);
  padding:12px; border-radius:8px; font-size:12px; z-index:100; max-height:80vh; overflow-y:auto; }}
#legend div {{ cursor:pointer; padding:3px 6px; border-radius:3px; margin:2px 0; white-space:nowrap; }}
#legend div:hover {{ background:rgba(255,255,255,0.15); }}
.cb {{ display:inline-block; width:14px; height:14px; border-radius:3px; margin-right:6px; vertical-align:middle; }}
#taruki-info {{ position:fixed; bottom:10px; left:10px; color:#fff; background:rgba(0,0,0,0.85);
  padding:14px 18px; border-radius:8px; font-size:13px; z-index:100; line-height:1.6; max-width:400px; }}
#controls {{ position:fixed; bottom:10px; right:10px; z-index:100; display:flex; flex-wrap:wrap; gap:4px; align-items:center; }}
#controls button {{ background:rgba(255,255,255,0.15); color:#fff; border:1px solid rgba(255,255,255,0.3);
  padding:8px 14px; border-radius:6px; cursor:pointer; font-size:12px; }}
#controls button:hover {{ background:rgba(255,255,255,0.3); }}
#controls button.active {{ background:rgba(255,160,0,0.5); border-color:#ffa500; }}
#pitch-group {{ display:flex; gap:2px; margin-right:8px; }}
#pitch-group button {{ padding:8px 12px; }}
</style>
</head>
<body>
<div id="info">
  <h3>垂木拾い 3Dビューア</h3>
  <div>左ドラッグ: 回転 / 右ドラッグ: 移動 / ホイール: ズーム</div>
  <div id="sel-info" style="margin-top:6px;color:#aaa;">クリックで部材選択</div>
</div>
<div id="taruki-info"></div>
<div id="legend"></div>
<div id="controls">
  <div id="pitch-group"></div>
  <button id="btn-taruki" class="active" onclick="toggleTaruki()">垂木表示</button>
  <button id="btn-building" class="active" onclick="toggleBuilding()">建物表示</button>
  <button onclick="resetCam()">リセット</button>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
<script>
const MESHES={meshes_json};
const PITCH_DATA={pitch_data_json};
const PITCHES={pitches_json};
const DEFAULT_PITCH={default_pitch};
const COLORS={colors_json};

let currentPitch=DEFAULT_PITCH;

const REASON_COLORS = {{
  "基本ピッチ": 0xffa500,
  "屋根端部": 0x00ccff,
  "ログ壁脇": 0x00ff88,
  "天窓脇": 0xff4444,
  "煙突脇": 0xff00ff,
}};

const W=innerWidth, H=innerHeight;
const scene=new THREE.Scene();
scene.background=new THREE.Color(0x1a1a2e);
const camera=new THREE.PerspectiveCamera(50,W/H,0.01,1000);
const renderer=new THREE.WebGLRenderer({{antialias:true}});
renderer.setSize(W,H); renderer.setPixelRatio(devicePixelRatio);
document.body.appendChild(renderer.domElement);

scene.add(new THREE.AmbientLight(0xffffff,0.5));
const dl=new THREE.DirectionalLight(0xffffff,0.7); dl.position.set(5,10,5); scene.add(dl);
const dl2=new THREE.DirectionalLight(0xffffff,0.3); dl2.position.set(-5,5,-5); scene.add(dl2);

class Orbit{{
constructor(c,e){{this.c=c;this.e=e;this.t=new THREE.Vector3();this.s=new THREE.Spherical();
this._d=0;this._st=0;this._sv=new THREE.Vector2();
e.addEventListener('mousedown',ev=>{{this._st=ev.button===0?1:ev.button===2?2:0;this._sv.set(ev.clientX,ev.clientY);
const mm=ev2=>{{const dx=ev2.clientX-this._sv.x,dy=ev2.clientY-this._sv.y;this._sv.set(ev2.clientX,ev2.clientY);
if(this._st===1)this._rot(dx,dy);else if(this._st===2)this._pan(dx,dy);}};
const mu=()=>{{this._st=0;document.removeEventListener('mousemove',mm);document.removeEventListener('mouseup',mu);}};
document.addEventListener('mousemove',mm);document.addEventListener('mouseup',mu);}});
e.addEventListener('wheel',ev=>{{ev.preventDefault();this._zm(ev.deltaY>0?1.1:0.9);}},{{passive:false}});
e.addEventListener('contextmenu',ev=>ev.preventDefault());
let td=0;e.addEventListener('touchstart',ev=>{{ev.preventDefault();if(ev.touches.length===1){{this._st=1;this._sv.set(ev.touches[0].clientX,ev.touches[0].clientY);}}
else if(ev.touches.length===2){{this._st=3;td=Math.hypot(ev.touches[1].clientX-ev.touches[0].clientX,ev.touches[1].clientY-ev.touches[0].clientY);}}}},{{passive:false}});
e.addEventListener('touchmove',ev=>{{ev.preventDefault();if(this._st===1&&ev.touches.length===1){{const dx=ev.touches[0].clientX-this._sv.x,dy=ev.touches[0].clientY-this._sv.y;
this._sv.set(ev.touches[0].clientX,ev.touches[0].clientY);this._rot(dx,dy);}}else if(this._st===3&&ev.touches.length===2){{
const d=Math.hypot(ev.touches[1].clientX-ev.touches[0].clientX,ev.touches[1].clientY-ev.touches[0].clientY);this._zm(td/d);td=d;}}}},{{passive:false}});
e.addEventListener('touchend',()=>{{this._st=0;}});this.initSM();}}
_rot(dx,dy){{const o=this.c.position.clone().sub(this.t);this.s.setFromVector3(o);this.s.theta-=dx*0.008;this.s.phi-=dy*0.008;
this.s.phi=Math.max(0.01,Math.min(Math.PI-0.01,this.s.phi));o.setFromSpherical(this.s);this.c.position.copy(this.t).add(o);this.c.lookAt(this.t);}}
_pan(dx,dy){{const d=this.c.position.distanceTo(this.t)*0.001;const r=new THREE.Vector3().setFromMatrixColumn(this.c.matrix,0);
const u=new THREE.Vector3().setFromMatrixColumn(this.c.matrix,1);const p=r.multiplyScalar(-dx*d).add(u.multiplyScalar(dy*d));
this.c.position.add(p);this.t.add(p);this.c.lookAt(this.t);}}
_zm(f){{const o=this.c.position.clone().sub(this.t);o.multiplyScalar(f);this.c.position.copy(this.t).add(o);this.c.lookAt(this.t);}}
initSM(){{this._smId=null;const cn=e=>{{if(e.gamepad&&(e.gamepad.id.includes('SpaceMouse')||e.gamepad.id.includes('3Dconnexion')||e.gamepad.axes.length>=6))this._smId=e.gamepad.index;}};
addEventListener('gamepadconnected',cn);addEventListener('gamepaddisconnected',e=>{{if(this._smId===e.gamepad.index)this._smId=null;}});
for(const g of navigator.getGamepads())if(g&&(g.id.includes('SpaceMouse')||g.id.includes('3Dconnexion')||g.axes.length>=6)){{this._smId=g.index;break;}}}}
pollSM(){{if(this._smId===null)return;const g=navigator.getGamepads()[this._smId];if(!g)return;const a=g.axes,dz=0.05;
const tx=Math.abs(a[0])>dz?a[0]:0,ty=Math.abs(a[1])>dz?a[1]:0,tz=Math.abs(a[2])>dz?a[2]:0;
const rx=Math.abs(a[3])>dz?a[3]:0,ry=Math.abs(a[4])>dz?a[4]:0;
if(tx||tz)this._pan(tx*-8,tz*8);if(ty)this._zm(1+ty*0.02);if(rx||ry)this._rot(ry*-6,rx*-6);}}
}}

const buildingGroup=new THREE.Group();
const catGroups={{}};const allMeshes=[];const bbox=new THREE.Box3();
MESHES.forEach(m=>{{
const g=new THREE.BufferGeometry();g.setAttribute('position',new THREE.Float32BufferAttribute(m.verts,3));
g.setIndex(m.faces);g.computeVertexNormals();
const color=COLORS[m.cat]||'#999';
const mat=new THREE.MeshLambertMaterial({{color:new THREE.Color(color),transparent:true,opacity:0.2,side:THREE.DoubleSide,depthWrite:false}});
const mesh=new THREE.Mesh(g,mat);mesh.userData={{cat:m.cat,name:m.name,gid:m.gid}};
if(!catGroups[m.cat]){{catGroups[m.cat]=new THREE.Group();buildingGroup.add(catGroups[m.cat]);}}
catGroups[m.cat].add(mesh);allMeshes.push(mesh);bbox.expandByObject(mesh);
}});
scene.add(buildingGroup);

// 垂木ライン描画（ピッチ切り替え対応）
const tarukiGroup=new THREE.Group();
scene.add(tarukiGroup);

function buildTarukiLines(pitchMm){{
  // 既存ラインをクリア
  while(tarukiGroup.children.length>0) tarukiGroup.remove(tarukiGroup.children[0]);
  const pd=PITCH_DATA[pitchMm];
  if(!pd) return;
  pd.rafters.forEach(r=>{{
    const color=REASON_COLORS[r.reason]||0xffa500;
    const pts=[new THREE.Vector3(...r.seg[0]),new THREE.Vector3(...r.seg[1])];
    const g=new THREE.BufferGeometry().setFromPoints(pts);
    const mat=new THREE.LineBasicMaterial({{color,linewidth:2}});
    const line=new THREE.Line(g,mat);
    line.userData={{reason:r.reason,length:r.length_m,side:r.side,double:r.double}};
    tarukiGroup.add(line);
  }});
  // 集計表示更新
  const S=pd.summary;
  let h='<b style="font-size:14px;">垂木積算</b><br><br>';
  h+=`ピッチ: ${{S.pitch}}mm<br>`;
  h+=`垂木長さ（1本）: ${{S.single_length}}m<br><br>`;
  h+=`<b>総本数: ${{S.total_count}}本</b><br>`;
  h+=`総長さ: ${{S.total_length}}m<br><br>`;
  h+=`前面: ${{S.count_front}}本<br>`;
  h+=`背面: ${{S.count_back}}本<br><br>`;
  h+='<b>理由別:</b><br>';
  for(const [reason, count] of Object.entries(S.count_by_reason)){{
    const color=REASON_COLORS[reason]||0xffa500;
    const hex='#'+color.toString(16).padStart(6,'0');
    const len=S.length_by_reason[reason]||0;
    h+=`<span style="color:${{hex}}">━━</span> ${{reason}}: ${{count}}本 / ${{len}}m<br>`;
  }}
  document.getElementById('taruki-info').innerHTML=h;
}}

function switchPitch(pitchMm){{
  currentPitch=pitchMm;
  buildTarukiLines(pitchMm);
  document.querySelectorAll('#pitch-group button').forEach(b=>{{
    b.classList.toggle('active', parseInt(b.dataset.pitch)===pitchMm);
  }});
}}

// ピッチ切り替えボタン生成
const pg=document.getElementById('pitch-group');
PITCHES.forEach(p=>{{
  const b=document.createElement('button');
  b.textContent=p+'mm';
  b.dataset.pitch=p;
  b.onclick=()=>switchPitch(p);
  if(p===DEFAULT_PITCH) b.classList.add('active');
  pg.appendChild(b);
}});

// 初期描画
buildTarukiLines(DEFAULT_PITCH);

const center=new THREE.Vector3();bbox.getCenter(center);
const size=bbox.getSize(new THREE.Vector3());const maxDim=Math.max(size.x,size.y,size.z);
camera.position.set(center.x+maxDim*0.8,center.y+maxDim*0.6,center.z+maxDim*0.8);
const controls=new Orbit(camera,renderer.domElement);
controls.t.copy(center);camera.lookAt(center);

function resetCam(){{
camera.position.set(center.x+maxDim*0.8,center.y+maxDim*0.6,center.z+maxDim*0.8);
controls.t.copy(center);camera.lookAt(center);}}

const leg=document.getElementById('legend');
[...new Set(MESHES.map(m=>m.cat))].forEach(cat=>{{
const d=document.createElement('div');d.innerHTML='<span class="cb" style="background:'+(COLORS[cat]||'#999')+'"></span>'+cat;
d.onclick=()=>{{if(catGroups[cat]){{catGroups[cat].visible=!catGroups[cat].visible;d.style.opacity=catGroups[cat].visible?1:0.3;}}}};
leg.appendChild(d);
}});

const rc=new THREE.Raycaster(),mouse=new THREE.Vector2();let selMesh=null;
renderer.domElement.addEventListener('click',e=>{{
mouse.x=(e.clientX/W)*2-1;mouse.y=-(e.clientY/H)*2+1;rc.setFromCamera(mouse,camera);
const hits=rc.intersectObjects(allMeshes);
if(selMesh){{selMesh.material.emissive.setHex(0);selMesh=null;}}
if(hits.length>0){{selMesh=hits[0].object;selMesh.material.emissive.setHex(0x333333);
document.getElementById('sel-info').textContent=selMesh.userData.cat+' / '+selMesh.userData.name;}}
else{{document.getElementById('sel-info').textContent='クリックで部材選択';}}}});

let showT=true,showB=true;
function toggleTaruki(){{showT=!showT;tarukiGroup.visible=showT;document.getElementById('btn-taruki').classList.toggle('active',showT);}}
function toggleBuilding(){{showB=!showB;buildingGroup.visible=showB;document.getElementById('btn-building').classList.toggle('active',showB);}}

const grid=new THREE.GridHelper(20,40,0x444444,0x333333);grid.position.copy(center);grid.position.y=0;scene.add(grid);
(function anim(){{requestAnimationFrame(anim);controls.pollSM();renderer.render(scene,camera);}})();
addEventListener('resize',()=>{{camera.aspect=innerWidth/innerHeight;camera.updateProjectionMatrix();renderer.setSize(innerWidth,innerHeight);}});
</script>
</body>
</html>"""


if __name__ == "__main__":
    main()
