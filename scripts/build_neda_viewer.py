#!/usr/bin/env python3
"""根太拾い3Dビューア: 基礎区画ごとに根太を自動配置し積算
Usage:
  python build_neda_viewer.py <input.ifc> <output.html>

根太配置ルール:
  - 方向: 垂木と平行（slope_axis方向）
  - ピッチ: 455mm / 303mm（切り替え可能）
  - 基礎区画ごとに配置
  - 区画の両端には必ず根太を配置
  - ユニットバスのスラブ切れは無視
"""
import sys
import os
import ifcopenshell
import ifcopenshell.geom
import ifcopenshell.util.element
import json
import numpy as np
from collections import defaultdict, Counter

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

JOIST_PITCHES = [0.455, 0.303]  # 455mm, 303mm


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
# 屋根パラメータ自動検出（slope_axis / ridge_axis を得るため）
###############################################################################

def detect_roof_params(ifc, settings):
    """IFC屋根メッシュから棟方向・勾配方向を自動検出。
    根太の方向（= slope_axis）を決定するために使用。
    """
    top_verts = []

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
                if normal[1] > 0.3:
                    top_verts.extend([pts[i0], pts[i1], pts[i2]])
        except Exception:
            pass

    if not top_verts:
        print("WARNING: No roof top faces found!")
        return None

    tv = np.array(top_verts)
    tv = np.unique(np.round(tv, 3), axis=0)

    ridge_height = tv[:, 1].max()
    ridge_tol = 0.15
    ridge_mask = tv[:, 1] > ridge_height - ridge_tol
    ridge_pts = tv[ridge_mask]

    ridge_x_span = ridge_pts[:, 0].max() - ridge_pts[:, 0].min()
    ridge_z_span = ridge_pts[:, 2].max() - ridge_pts[:, 2].min()

    if ridge_x_span > ridge_z_span:
        ridge_axis = "x"
        slope_axis = "z"
    else:
        ridge_axis = "z"
        slope_axis = "x"

    print(f"  Ridge axis: {ridge_axis}, Slope axis: {slope_axis}")
    return {"ridge_axis": ridge_axis, "slope_axis": slope_axis}


###############################################################################
# 基礎区画検出
###############################################################################

def detect_kiso_compartments(ifc, settings, rp):
    """基礎壁から区画を検出する。
    IFC座標系（X, Y が水平面、Z が高さ）で壁を解析。
    Three.js座標系へは最後に変換する。

    Returns: list of dicts with x_min, x_max, y_min, y_max (IFC座標)
    """
    slope_axis = rp["slope_axis"]
    ridge_axis = rp["ridge_axis"]

    walls = ifc.by_type("IfcWall") + ifc.by_type("IfcWallStandardCase")
    kiso_walls = [w for w in walls if w.Name and "kiso" in w.Name.lower()]

    # IFC座標で壁のパラメータを取得（重複排除）
    seen = set()
    wall_segments = []
    for w in kiso_walls:
        try:
            shape = ifcopenshell.geom.create_shape(settings, w)
            vf = list(shape.geometry.verts)
            xs = [vf[i] for i in range(0, len(vf), 3)]
            ys = [vf[i+1] for i in range(0, len(vf), 3)]  # IFC Y = 水平方向
            xspan = max(xs) - min(xs)
            yspan = max(ys) - min(ys)
            direction = "x" if xspan > yspan else "y"
            cx = round((min(xs) + max(xs)) / 2, 2)
            cy = round((min(ys) + max(ys)) / 2, 2)
            key = (direction, cx, cy)
            if key not in seen:
                seen.add(key)
                wall_segments.append({
                    "dir": direction,
                    "x_min": min(xs), "x_max": max(xs),
                    "y_min": min(ys), "y_max": max(ys),
                    "cx": cx, "cy": cy,
                })
        except Exception:
            pass

    x_walls = sorted([w for w in wall_segments if w["dir"] == "x"], key=lambda w: w["cy"])
    y_walls = sorted([w for w in wall_segments if w["dir"] == "y"], key=lambda w: w["cx"])

    print(f"  基礎壁（重複排除後）: {len(wall_segments)}本")
    print(f"    X方向壁: {len(x_walls)}本")
    for w in x_walls:
        print(f"      Y={w['cy']:.3f}, X=[{w['x_min']:.3f}, {w['x_max']:.3f}]")
    print(f"    Y方向壁: {len(y_walls)}本")
    for w in y_walls:
        print(f"      X={w['cx']:.3f}, Y=[{w['y_min']:.3f}, {w['y_max']:.3f}]")

    # グリッド位置を収集
    x_positions = sorted(set(w["cx"] for w in y_walls))
    y_positions = sorted(set(w["cy"] for w in x_walls))

    if len(x_positions) < 2 or len(y_positions) < 2:
        print("WARNING: 基礎区画を構成するのに十分な壁がありません")
        return []

    tol = 0.2  # 壁のカバレッジ判定用許容差

    def x_wall_covers(y_pos, x1, x2):
        """Y位置y_posにあるX方向壁が[x1,x2]をカバーしているか"""
        return any(
            abs(w["cy"] - y_pos) < tol and w["x_min"] <= x1 + tol and w["x_max"] >= x2 - tol
            for w in x_walls
        )

    def y_wall_covers(x_pos, y1, y2):
        """X位置x_posにあるY方向壁が[y1,y2]をカバーしているか"""
        return any(
            abs(w["cx"] - x_pos) < tol and w["y_min"] <= y1 + tol and w["y_max"] >= y2 - tol
            for w in y_walls
        )

    # 全ペアのY方向壁を調べ、各Xストリップ内で有効なX方向壁を検出
    compartments = []
    for xi in range(len(x_positions)):
        for xj in range(xi + 1, len(x_positions)):
            x1, x2 = x_positions[xi], x_positions[xj]

            # この X 範囲で両側のY壁がカバーしているか確認
            # ただし、間に中間Y壁がある場合はスキップ（細かい区画を優先）
            # まず中間のY壁をチェック
            has_intermediate_y_wall = False
            for xk in range(xi + 1, xj):
                xm = x_positions[xk]
                # この中間Y壁が y全範囲をカバーしているか
                # → 簡易チェック: どのY範囲でもカバーしていれば中間壁あり
                for yk in range(len(y_positions)):
                    for yl in range(yk + 1, len(y_positions)):
                        y1t, y2t = y_positions[yk], y_positions[yl]
                        if y_wall_covers(xm, y1t, y2t):
                            has_intermediate_y_wall = True
                            break
                    if has_intermediate_y_wall:
                        break
                if has_intermediate_y_wall:
                    break
            if has_intermediate_y_wall:
                continue

            if not y_wall_covers(x1, y_positions[0], y_positions[-1]):
                # 左辺の壁が全くない場合はスキップ（部分カバーは後で確認）
                pass
            if not y_wall_covers(x2, y_positions[0], y_positions[-1]):
                pass

            # このXストリップでカバーされているY位置を収集
            valid_y_positions = []
            for yp in y_positions:
                if x_wall_covers(yp, x1, x2):
                    valid_y_positions.append(yp)

            # 隣接するY位置間で区画を作成
            for yi in range(len(valid_y_positions) - 1):
                y1, y2 = valid_y_positions[yi], valid_y_positions[yi + 1]

                # 中間のX壁チェック（このX範囲をカバーする中間X壁があれば分割済み）
                has_intermediate_x_wall = False
                for yp in y_positions:
                    if y1 < yp < y2 and x_wall_covers(yp, x1, x2):
                        has_intermediate_x_wall = True
                        break
                if has_intermediate_x_wall:
                    continue

                # 左右のY壁チェック
                has_left = y_wall_covers(x1, y1, y2)
                has_right = y_wall_covers(x2, y1, y2)

                if has_left and has_right:
                    compartments.append({
                        "x_min": x1, "x_max": x2,
                        "y_min": y1, "y_max": y2,
                        "width_x": x2 - x1,
                        "width_y": y2 - y1,
                    })

    print(f"\n  検出された区画: {len(compartments)}個")
    for idx, c in enumerate(compartments):
        print(f"    区画{idx+1}: X=[{c['x_min']:.3f},{c['x_max']:.3f}] Y=[{c['y_min']:.3f},{c['y_max']:.3f}]"
              f"  ({c['width_x']:.3f} x {c['width_y']:.3f})")

    return compartments


###############################################################################
# 2Fスラブ開口検出
###############################################################################

def detect_2f_openings(ifc, settings):
    """2Fスラブの開口（階段の吹き抜け等）を検出する。

    スラブ上面のポリゴン頂点からBBの欠け部分（開口）を矩形で検出する。
    スラブ形状がBB通りの単純矩形なら開口なし。
    内部頂点（BB辺上でない頂点）が存在すれば、グリッド分割して
    スラブが存在しないセルを開口として返す。

    Returns: list of dicts with x_min, x_max, y_min, y_max (IFC座標)
    """
    slabs = ifc.by_type("IfcSlab")
    openings = []

    for slab in slabs:
        name = (slab.Name or "").lower()
        if "2-yuka" not in name and "2f" not in name:
            continue
        try:
            shape = ifcopenshell.geom.create_shape(settings, slab)
            vf = shape.geometry.verts
            ff = shape.geometry.faces
            va = np.array(vf).reshape(-1, 3)
            fa = np.array(ff).reshape(-1, 3)

            z_max = va[:, 2].max()

            # 上面頂点を取得
            top_mask = va[:, 2] > z_max - 0.01
            top_pts = va[top_mask]
            unique_xy = np.unique(np.round(top_pts[:, :2], 3), axis=0)

            bb_xmin, bb_xmax = va[:, 0].min(), va[:, 0].max()
            bb_ymin, bb_ymax = va[:, 1].min(), va[:, 1].max()

            # 単純矩形（4頂点）なら開口なし
            if len(unique_xy) <= 4:
                continue

            # 全てのX, Y座標を収集してグリッドを作成
            all_x = sorted(set(round(p[0], 3) for p in unique_xy))
            all_y = sorted(set(round(p[1], 3) for p in unique_xy))

            # 上面の三角形を収集
            top_tris = []
            for f in fa:
                pts = va[f]
                if all(abs(p[2] - z_max) < 0.01 for p in pts):
                    top_tris.append(pts[:, :2])  # XY only

            # グリッドセルごとにスラブ存在判定（中心点がスラブ三角形内にあるか）
            def point_in_tri(p, t):
                """点pが三角形t内にあるか（2D）"""
                v0 = t[2] - t[0]
                v1 = t[1] - t[0]
                v2 = p - t[0]
                d00 = np.dot(v0, v0)
                d01 = np.dot(v0, v1)
                d02 = np.dot(v0, v2)
                d11 = np.dot(v1, v1)
                d12 = np.dot(v1, v2)
                inv = d00 * d11 - d01 * d01
                if abs(inv) < 1e-10:
                    return False
                u = (d11 * d02 - d01 * d12) / inv
                v = (d00 * d12 - d01 * d02) / inv
                return u >= -0.001 and v >= -0.001 and u + v <= 1.001

            for xi in range(len(all_x) - 1):
                for yi in range(len(all_y) - 1):
                    cx = (all_x[xi] + all_x[xi + 1]) / 2
                    cy = (all_y[yi] + all_y[yi + 1]) / 2
                    center = np.array([cx, cy])

                    in_slab = False
                    for tri in top_tris:
                        if point_in_tri(center, tri):
                            in_slab = True
                            break

                    if not in_slab:
                        openings.append({
                            "x_min": all_x[xi],
                            "x_max": all_x[xi + 1],
                            "y_min": all_y[yi],
                            "y_max": all_y[yi + 1],
                        })

        except Exception as e:
            print(f"  WARNING: 開口検出失敗 ({slab.Name}): {e}")

    # 隣接する開口セルを統合（同じX範囲のY連続セル / 同じY範囲のX連続セル）
    if len(openings) > 1:
        merged = [openings[0]]
        for o in openings[1:]:
            prev = merged[-1]
            # Y方向に隣接（同じX範囲）
            if (abs(o["x_min"] - prev["x_min"]) < 0.01 and
                abs(o["x_max"] - prev["x_max"]) < 0.01 and
                abs(o["y_min"] - prev["y_max"]) < 0.01):
                prev["y_max"] = o["y_max"]
            # X方向に隣接（同じY範囲）
            elif (abs(o["y_min"] - prev["y_min"]) < 0.01 and
                  abs(o["y_max"] - prev["y_max"]) < 0.01 and
                  abs(o["x_min"] - prev["x_max"]) < 0.01):
                prev["x_max"] = o["x_max"]
            else:
                merged.append(o)
        openings = merged

    if openings:
        print(f"\n  2Fスラブ開口: {len(openings)}箇所")
        for idx, o in enumerate(openings):
            w = o["x_max"] - o["x_min"]
            h = o["y_max"] - o["y_min"]
            print(f"    開口{idx+1}: X=[{o['x_min']:.3f},{o['x_max']:.3f}] "
                  f"Y=[{o['y_min']:.3f},{o['y_max']:.3f}] ({w:.3f} x {h:.3f}m)")
    else:
        print("\n  2Fスラブ開口: なし")

    return openings


###############################################################################
# 2F区画検出（集成梁方向判定付き）
###############################################################################

def detect_2f_compartments(ifc, settings):
    """2Fスラブと集成梁から2F根太の区画・方向を検出する。

    ロジック:
    1. "2-yuka" or "2f" を含む IfcSlab を検出 → 2F区画の矩形BB
    2. 集成梁（IfcBeam, ObjectType に "集成梁"）のうち、2Fスラブと高さが重なるものを検出
    3. 集成梁の主方向（長手方向）に直交する方向が根太方向
    4. 短い梁（<1.5m）はコネクタとして除外
    5. 集成梁がない場合は短手方向
    6. 全根太の方向を統一（集成梁の主方向の多数決）
    7. 4m超のスパンは区画を分割

    Returns: list of dicts (2F区画), str (joist_dir), list of dicts (openings)
    """
    slabs = ifc.by_type("IfcSlab")
    slab_2f_list = []

    for slab in slabs:
        name = (slab.Name or "").lower()
        if "2-yuka" not in name and "2f" not in name:
            continue
        try:
            shape = ifcopenshell.geom.create_shape(settings, slab)
            vf = shape.geometry.verts
            xs = [vf[i] for i in range(0, len(vf), 3)]
            ys = [vf[i+1] for i in range(0, len(vf), 3)]
            zs = [vf[i+2] for i in range(0, len(vf), 3)]
            slab_2f_list.append({
                "name": slab.Name,
                "x_min": min(xs), "x_max": max(xs),
                "y_min": min(ys), "y_max": max(ys),
                "z_min": min(zs), "z_max": max(zs),
                "slab_top_z": max(zs),
            })
        except Exception as e:
            print(f"  WARNING: 2Fスラブ解析失敗 ({slab.Name}): {e}")

    if not slab_2f_list:
        print("  2Fスラブなし（平屋）")
        return [], ""

    # 全2Fスラブの統合BB（通常は1つだが複数ある場合も対応）
    all_x_min = min(s["x_min"] for s in slab_2f_list)
    all_x_max = max(s["x_max"] for s in slab_2f_list)
    all_y_min = min(s["y_min"] for s in slab_2f_list)
    all_y_max = max(s["y_max"] for s in slab_2f_list)
    all_z_min = min(s["z_min"] for s in slab_2f_list)
    all_z_max = max(s["z_max"] for s in slab_2f_list)
    slab_top_z = max(s["slab_top_z"] for s in slab_2f_list)

    print(f"\n  2Fスラブ: {len(slab_2f_list)}枚")
    for s in slab_2f_list:
        print(f"    {s['name']}: X=[{s['x_min']:.3f},{s['x_max']:.3f}] "
              f"Y=[{s['y_min']:.3f},{s['y_max']:.3f}] Z={s['z_min']:.3f}-{s['z_max']:.3f}")

    # 集成梁の検出（2Fスラブ高さと重なるもの）
    beams = ifc.by_type("IfcBeam")
    beam_2f = []
    z_overlap_margin = 0.5  # スラブとの高さオーバーラップ許容

    for b in beams:
        obj_type = getattr(b, "ObjectType", "") or ""
        if "集成梁" not in obj_type:
            continue
        try:
            shape = ifcopenshell.geom.create_shape(settings, b)
            vf = shape.geometry.verts
            bxs = [vf[i] for i in range(0, len(vf), 3)]
            bys = [vf[i+1] for i in range(0, len(vf), 3)]
            bzs = [vf[i+2] for i in range(0, len(vf), 3)]
            b_z_min, b_z_max = min(bzs), max(bzs)

            # 2Fスラブとの高さオーバーラップチェック
            if b_z_max < all_z_min - z_overlap_margin or b_z_min > all_z_max + z_overlap_margin:
                continue  # 高さが合わない（屋根梁など）

            x_span = max(bxs) - min(bxs)
            y_span = max(bys) - min(bys)
            length = max(x_span, y_span)

            # 短い梁（< 1.5m）はコネクタとして除外
            if length < 1.5:
                print(f"    スキップ（短い梁）: {b.Name} len={length:.3f}m")
                continue

            beam_dir = "x" if x_span > y_span else "y"
            beam_2f.append({
                "name": b.Name,
                "type": obj_type,
                "dir": beam_dir,
                "length": length,
                "x_min": min(bxs), "x_max": max(bxs),
                "y_min": min(bys), "y_max": max(bys),
                "z_min": b_z_min, "z_max": b_z_max,
            })
        except Exception as e:
            print(f"  WARNING: 梁解析失敗 ({b.Name}): {e}")

    print(f"\n  2F集成梁: {len(beam_2f)}本")
    for b in beam_2f:
        print(f"    {b['name']} ({b['type']}): dir={b['dir']} len={b['length']:.3f}m "
              f"Z={b['z_min']:.3f}-{b['z_max']:.3f}")

    # 根太方向の決定
    if beam_2f:
        # 集成梁の主方向の多数決 → それに直交する方向が根太方向
        dir_votes = Counter(b["dir"] for b in beam_2f)
        beam_main_dir = dir_votes.most_common(1)[0][0]
        joist_dir = "y" if beam_main_dir == "x" else "x"  # 直交
        print(f"\n  集成梁主方向: {beam_main_dir} → 根太方向: {joist_dir}")
    else:
        # 集成梁なし → 短手方向
        x_span = all_x_max - all_x_min
        y_span = all_y_max - all_y_min
        joist_dir = "x" if x_span <= y_span else "y"
        print(f"\n  集成梁なし → 短手方向で根太配置: {joist_dir}")

    # =====================================================================
    # ログ壁検出 → 区画境界として使用（間仕切壁は無視）
    # =====================================================================
    walls = ifc.by_type("IfcWall") + ifc.by_type("IfcWallStandardCase")
    wall_x_boundaries = set()  # Y方向ログ壁 → X方向の境界
    wall_y_boundaries = set()  # X方向ログ壁 → Y方向の境界
    wall_seen = set()
    edge_tol = 0.15  # スラブ端との近接判定

    for w in walls:
        wname = (w.Name or "").lower()
        if not wname.startswith("log"):
            continue
        try:
            shape = ifcopenshell.geom.create_shape(settings, w)
            vf = shape.geometry.verts
            wxs = [vf[i] for i in range(0, len(vf), 3)]
            wys = [vf[i+1] for i in range(0, len(vf), 3)]
            wzs = [vf[i+2] for i in range(0, len(vf), 3)]

            # 2Fスラブ高さに到達する壁のみ
            if max(wzs) < all_z_min - 0.2:
                continue

            xspan = max(wxs) - min(wxs)
            yspan = max(wys) - min(wys)
            wdir = "x" if xspan > yspan else "y"
            cx = round((min(wxs) + max(wxs)) / 2, 2)
            cy = round((min(wys) + max(wys)) / 2, 2)
            key = (wdir, cx, cy)
            if key in wall_seen:
                continue
            wall_seen.add(key)

            if wdir == "y":
                # Y方向壁 → X境界（スラブY範囲を横断していること）
                if max(wys) > all_y_min + 0.1 and min(wys) < all_y_max - 0.1:
                    if all_x_min - 0.2 <= cx <= all_x_max + 0.2:
                        # スラブ端でない内部壁のみ
                        if cx > all_x_min + edge_tol and cx < all_x_max - edge_tol:
                            wall_x_boundaries.add(cx)
            else:
                # X方向壁 → Y境界（スラブX範囲を横断していること）
                if max(wxs) > all_x_min + 0.1 and min(wxs) < all_x_max - 0.1:
                    if all_y_min - 0.2 <= cy <= all_y_max + 0.2:
                        if cy > all_y_min + edge_tol and cy < all_y_max - edge_tol:
                            wall_y_boundaries.add(cy)
        except Exception:
            pass

    print(f"\n  ログ壁による内部境界:")
    print(f"    X方向境界（Y方向ログ壁）: {sorted(wall_x_boundaries)}")
    print(f"    Y方向境界（X方向ログ壁）: {sorted(wall_y_boundaries)}")

    # =====================================================================
    # 区画グリッド生成（ログ壁ベース → 梁はセル交差時のみ分割）
    # =====================================================================
    # Step 1: ログ壁のみでベースグリッドを作成
    x_bounds = sorted(set([all_x_min] + list(wall_x_boundaries) + [all_x_max]))
    y_bounds = sorted(set([all_y_min] + list(wall_y_boundaries) + [all_y_max]))

    print(f"\n  ログ壁ベースグリッド:")
    print(f"    X境界: {[round(x,3) for x in x_bounds]}")
    print(f"    Y境界: {[round(y,3) for y in y_bounds]}")

    # Step 2: ベースグリッドのセルを生成
    base_cells = []
    for xi in range(len(x_bounds) - 1):
        for yi in range(len(y_bounds) - 1):
            x1, x2 = x_bounds[xi], x_bounds[xi + 1]
            y1, y2 = y_bounds[yi], y_bounds[yi + 1]
            if x2 - x1 < 0.1 or y2 - y1 < 0.1:
                continue
            base_cells.append((x1, x2, y1, y2))

    # Step 3: 各梁について、実際に交差するセルのみ分割
    # 梁の実際のXY範囲を考慮して交差判定
    beam_margin = 0.3  # 梁の存在範囲の余裕
    for b in beam_2f:
        bcx = round((b["x_min"] + b["x_max"]) / 2, 3)
        bcy = round((b["y_min"] + b["y_max"]) / 2, 3)
        new_cells = []
        for (x1, x2, y1, y2) in base_cells:
            split_done = False
            if b["dir"] == "y":
                # Y方向梁 → X位置で分割（梁のY範囲がセルと重なる場合のみ）
                if (all_x_min + edge_tol < bcx < all_x_max - edge_tol and
                        x1 + edge_tol < bcx < x2 - edge_tol):
                    # 梁のY範囲がこのセルのY範囲と十分重なるか
                    beam_y_lo = b["y_min"] - beam_margin
                    beam_y_hi = b["y_max"] + beam_margin
                    overlap = min(y2, beam_y_hi) - max(y1, beam_y_lo)
                    cell_y_span = y2 - y1
                    if overlap > cell_y_span * 0.5:
                        # セルを梁位置で2分割
                        new_cells.append((x1, bcx, y1, y2))
                        new_cells.append((bcx, x2, y1, y2))
                        split_done = True
            else:
                # X方向梁 → Y位置で分割（梁のX範囲がセルと重なる場合のみ）
                if (all_y_min + edge_tol < bcy < all_y_max - edge_tol and
                        y1 + edge_tol < bcy < y2 - edge_tol):
                    beam_x_lo = b["x_min"] - beam_margin
                    beam_x_hi = b["x_max"] + beam_margin
                    overlap = min(x2, beam_x_hi) - max(x1, beam_x_lo)
                    cell_x_span = x2 - x1
                    if overlap > cell_x_span * 0.5:
                        new_cells.append((x1, x2, y1, bcy))
                        new_cells.append((x1, x2, bcy, y2))
                        split_done = True
            if not split_done:
                new_cells.append((x1, x2, y1, y2))
        base_cells = new_cells

    print(f"\n  梁分割後セル（統合前）: {len(base_cells)}個")

    # Step 3b: 梁分割で生じた狭いセルを隣接セルに統合
    # 根太方向の幅が min_cell_span 未満のセルは隣に吸収
    min_cell_span = 0.8  # 最小セル幅（m）
    merged = True
    while merged:
        merged = False
        new_cells = []
        skip = set()
        for i, (x1, x2, y1, y2) in enumerate(base_cells):
            if i in skip:
                continue
            wx = x2 - x1
            wy = y2 - y1
            narrow = (joist_dir == "x" and wx < min_cell_span) or \
                     (joist_dir == "y" and wy < min_cell_span)
            if not narrow:
                new_cells.append((x1, x2, y1, y2))
                continue
            # 隣接セルを探して統合
            best_j = -1
            best_area = 0
            for j, (ax1, ax2, ay1, ay2) in enumerate(base_cells):
                if j == i or j in skip:
                    continue
                # Y範囲が一致し、X方向で隣接（またはその逆）
                if abs(y1 - ay1) < 0.01 and abs(y2 - ay2) < 0.01:
                    if abs(x2 - ax1) < 0.01 or abs(ax2 - x1) < 0.01:
                        area = (ax2 - ax1) * (ay2 - ay1)
                        if area > best_area:
                            best_area = area
                            best_j = j
                if abs(x1 - ax1) < 0.01 and abs(x2 - ax2) < 0.01:
                    if abs(y2 - ay1) < 0.01 or abs(ay2 - y1) < 0.01:
                        area = (ax2 - ax1) * (ay2 - ay1)
                        if area > best_area:
                            best_area = area
                            best_j = j
            if best_j >= 0:
                ax1, ax2, ay1, ay2 = base_cells[best_j]
                mx1 = min(x1, ax1)
                mx2 = max(x2, ax2)
                my1 = min(y1, ay1)
                my2 = max(y2, ay2)
                new_cells.append((mx1, mx2, my1, my2))
                skip.add(best_j)
                merged = True
            else:
                new_cells.append((x1, x2, y1, y2))
        base_cells = new_cells

    print(f"  梁分割後セル（統合後）: {len(base_cells)}個")

    # Step 4: セルから区画を生成（根太方向4m超なら分割）
    compartments = []
    for (x1, x2, y1, y2) in base_cells:
        wx = x2 - x1
        wy = y2 - y1
        if wx < 0.1 or wy < 0.1:
            continue

        # 根太方向のスパンが4m超なら分割
        if joist_dir == "x":
            joist_span = wx
        else:
            joist_span = wy

        if joist_span > 4.0:
            n_div = int(np.ceil(joist_span / 4.0))
            div_span = joist_span / n_div
            for d in range(n_div):
                if joist_dir == "x":
                    compartments.append({
                        "x_min": x1 + d * div_span,
                        "x_max": x1 + (d + 1) * div_span,
                        "y_min": y1, "y_max": y2,
                        "width_x": div_span, "width_y": wy,
                        "slab_top_z": slab_top_z,
                        "joist_dir": joist_dir,
                    })
                else:
                    compartments.append({
                        "x_min": x1, "x_max": x2,
                        "y_min": y1 + d * div_span,
                        "y_max": y1 + (d + 1) * div_span,
                        "width_x": wx, "width_y": div_span,
                        "slab_top_z": slab_top_z,
                        "joist_dir": joist_dir,
                    })
        else:
            compartments.append({
                "x_min": x1, "x_max": x2,
                "y_min": y1, "y_max": y2,
                "width_x": wx, "width_y": wy,
                "slab_top_z": slab_top_z,
                "joist_dir": joist_dir,
            })

    print(f"\n  2F区画: {len(compartments)}個 (根太方向: {joist_dir})")
    for idx, c in enumerate(compartments):
        print(f"    区画{idx+1}: X=[{c['x_min']:.3f},{c['x_max']:.3f}] "
              f"Y=[{c['y_min']:.3f},{c['y_max']:.3f}] "
              f"({c['width_x']:.3f} x {c['width_y']:.3f})")

    return compartments, joist_dir


###############################################################################
# 2F根太配置（開口対応）
###############################################################################

def place_joists_2f(compartments, joist_dir, openings=None, pitch=0.455):
    """2F区画ごとに根太を配置。開口部分は根太をクリッピングまたはスキップ。

    開口処理:
    - 根太が開口と交差する場合、開口の手前と奥の2本に分割
    - 根太が開口内に完全に収まる場合はスキップ
    """
    joists = []
    if openings is None:
        openings = []

    for comp_idx, comp in enumerate(compartments):
        if joist_dir == "x":
            joist_start = comp["x_min"]
            joist_end = comp["x_max"]
            pitch_start = comp["y_min"]
            pitch_end = comp["y_max"]
        else:
            joist_start = comp["y_min"]
            joist_end = comp["y_max"]
            pitch_start = comp["x_min"]
            pitch_end = comp["x_max"]

        positions = [
            {"pos": pitch_start, "reason": "区画端部"},
            {"pos": pitch_end, "reason": "区画端部"},
        ]

        pos = pitch_start + pitch
        while pos < pitch_end - 0.01:
            positions.append({"pos": pos, "reason": "基本ピッチ"})
            pos += pitch

        positions.sort(key=lambda p: p["pos"])
        merged = []
        for p in positions:
            if merged and abs(p["pos"] - merged[-1]["pos"]) < 0.03:
                if p["reason"] == "区画端部":
                    merged[-1] = p
                continue
            merged.append(p)

        for p in merged:
            # 開口チェック: 根太が開口と交差するか
            j_start = joist_start
            j_end = joist_end
            pos_val = p["pos"]

            # この根太に影響する開口を収集
            # 根太のpitch方向位置(pos_val)が開口のpitch範囲内にあるか
            clipped_segments = [(j_start, j_end)]

            for opening in openings:
                if joist_dir == "x":
                    # 根太はX方向、pos=IFC Y座標
                    op_pitch_min = opening["y_min"]
                    op_pitch_max = opening["y_max"]
                    op_joist_min = opening["x_min"]
                    op_joist_max = opening["x_max"]
                else:
                    # 根太はY方向、pos=IFC X座標
                    op_pitch_min = opening["x_min"]
                    op_pitch_max = opening["x_max"]
                    op_joist_min = opening["y_min"]
                    op_joist_max = opening["y_max"]

                # この根太のpitch位置が開口のpitch範囲内か
                if pos_val < op_pitch_min - 0.01 or pos_val > op_pitch_max + 0.01:
                    continue  # この開口は影響しない

                # 根太の走行方向で開口と交差 → セグメントをクリップ
                new_segments = []
                for seg_start, seg_end in clipped_segments:
                    if seg_end <= op_joist_min + 0.01 or seg_start >= op_joist_max - 0.01:
                        # 開口と重ならない
                        new_segments.append((seg_start, seg_end))
                    else:
                        # 開口の手前部分
                        if seg_start < op_joist_min - 0.01:
                            new_segments.append((seg_start, op_joist_min))
                        # 開口の奥部分
                        if seg_end > op_joist_max + 0.01:
                            new_segments.append((op_joist_max, seg_end))
                        # 完全に開口内 → スキップ
                clipped_segments = new_segments

            # クリップ後のセグメントを根太として追加
            for seg_start, seg_end in clipped_segments:
                seg_len = abs(seg_end - seg_start)
                if seg_len < 0.30:  # 30cm未満はスキップ（開口端と梁の隙間等）
                    continue
                joists.append({
                    "comp_idx": comp_idx,
                    "pos": pos_val,
                    "reason": p["reason"],
                    "joist_start": seg_start,
                    "joist_end": seg_end,
                    "length": seg_len,
                    "joist_dir": joist_dir,
                    "slab_top_z": comp["slab_top_z"],
                })

    return joists


def generate_joist_lines_2f(joists):
    """2F根太の3Dラインセグメント(Three.js座標)を生成。
    IFC→Three.js: (X, Y, Z) → (X, Z, -Y)
    """
    lines = []

    for j in joists:
        slab_y = j["slab_top_z"]

        if j["joist_dir"] == "x":
            x1 = j["joist_start"]
            x2 = j["joist_end"]
            z = -j["pos"]
            seg = [[x1, slab_y, z], [x2, slab_y, z]]
        else:
            x = j["pos"]
            z1 = -j["joist_start"]
            z2 = -j["joist_end"]
            seg = [[x, slab_y, z1], [x, slab_y, z2]]

        lines.append({
            "seg": seg,
            "length_m": round(j["length"], 3),
            "reason": j["reason"],
            "comp_idx": j["comp_idx"],
        })

    return lines


###############################################################################
# テラス・バルコニーの区画検出
###############################################################################

def detect_terrace_balcony_compartments(ifc, settings):
    """IFCのテラス・バルコニースラブから区画を検出する。
    各スラブのBBから矩形区画を生成し、短手方向（根太の走る方向）を判定する。

    Returns: list of dicts with:
      x_min, x_max, y_min, y_max (IFC座標),
      slab_top_z (IFC Z座標 = スラブ上面),
      joist_dir ("x" or "y"): 根太が走る方向（短手方向）,
      category: "テラス" or "バルコニー"
    """
    slabs = ifc.by_type("IfcSlab")
    compartments = []

    for slab in slabs:
        name = (slab.Name or "").lower()
        if "terrace" in name:
            category = "テラス"
        elif "balcony" in name or "balconi" in name:
            category = "バルコニー"
        else:
            continue

        try:
            shape = ifcopenshell.geom.create_shape(settings, slab)
            vf = shape.geometry.verts
            xs = [vf[i] for i in range(0, len(vf), 3)]
            ys = [vf[i+1] for i in range(0, len(vf), 3)]
            zs = [vf[i+2] for i in range(0, len(vf), 3)]

            x_min, x_max = min(xs), max(xs)
            y_min, y_max = min(ys), max(ys)
            x_span = x_max - x_min
            y_span = y_max - y_min

            # スラブ上面 = IFC Z座標の最大値
            slab_top_z = max(zs)

            # 短手方向に根太を走らせる
            # IFC X方向が短い → 根太はIFC X方向に走る → joist_dir="x"
            # IFC Y方向が短い → 根太はIFC Y方向に走る → joist_dir="y"
            if x_span <= y_span:
                joist_dir = "x"
            else:
                joist_dir = "y"

            compartments.append({
                "x_min": x_min, "x_max": x_max,
                "y_min": y_min, "y_max": y_max,
                "width_x": x_span,
                "width_y": y_span,
                "slab_top_z": slab_top_z,
                "joist_dir": joist_dir,
                "category": category,
                "name": slab.Name,
            })
        except Exception as e:
            print(f"  WARNING: スラブ解析失敗 ({slab.Name}): {e}")

    print(f"\n  テラス・バルコニー区画: {len(compartments)}個")
    for idx, c in enumerate(compartments):
        print(f"    {c['category']}{idx+1}: X=[{c['x_min']:.3f},{c['x_max']:.3f}] "
              f"Y=[{c['y_min']:.3f},{c['y_max']:.3f}] "
              f"({c['width_x']:.3f} x {c['width_y']:.3f}) "
              f"根太方向={c['joist_dir']} 上面Z={c['slab_top_z']:.3f}")

    return compartments


###############################################################################
# テラス・バルコニーの根太配置
###############################################################################

def place_joists_tb(compartments, pitch=0.455):
    """テラス・バルコニー区画ごとに根太を配置。
    根太はjoist_dir方向（短手方向）に走り、長手方向にpitch間隔で並ぶ。
    """
    joists = []

    for comp_idx, comp in enumerate(compartments):
        joist_dir = comp["joist_dir"]

        if joist_dir == "x":
            # 根太はIFC X方向に走る、IFC Y方向にピッチで並ぶ
            joist_start = comp["x_min"]
            joist_end = comp["x_max"]
            pitch_start = comp["y_min"]
            pitch_end = comp["y_max"]
        else:
            # 根太はIFC Y方向に走る、IFC X方向にピッチで並ぶ
            joist_start = comp["y_min"]
            joist_end = comp["y_max"]
            pitch_start = comp["x_min"]
            pitch_end = comp["x_max"]

        joist_length = abs(joist_end - joist_start)

        # 両端に必ず配置
        positions = [
            {"pos": pitch_start, "reason": "区画端部"},
            {"pos": pitch_end, "reason": "区画端部"},
        ]

        # ピッチ間隔で配置
        pos = pitch_start + pitch
        while pos < pitch_end - 0.01:
            positions.append({"pos": pos, "reason": "基本ピッチ"})
            pos += pitch

        # 重複マージ（30mm以内は区画端部を優先）
        positions.sort(key=lambda p: p["pos"])
        merged = []
        for p in positions:
            if merged and abs(p["pos"] - merged[-1]["pos"]) < 0.03:
                if p["reason"] == "区画端部":
                    merged[-1] = p
                continue
            merged.append(p)

        for p in merged:
            joists.append({
                "comp_idx": comp_idx,
                "pos": p["pos"],
                "reason": p["reason"],
                "joist_start": joist_start,
                "joist_end": joist_end,
                "length": joist_length,
                "joist_dir": joist_dir,
                "slab_top_z": comp["slab_top_z"],
            })

    return joists


def generate_joist_lines_tb(joists):
    """テラス・バルコニー根太の3Dラインセグメント(Three.js座標)を生成。
    IFC→Three.js: (X, Y, Z) → (X, Z, -Y)
    """
    lines = []

    for j in joists:
        # スラブ上面のThree.js Y座標 = IFC Z
        slab_y = j["slab_top_z"]

        if j["joist_dir"] == "x":
            # 根太はIFC X方向に走る
            x1 = j["joist_start"]
            x2 = j["joist_end"]
            z = -j["pos"]  # IFC Y → Three.js -Z
            seg = [[x1, slab_y, z], [x2, slab_y, z]]
        else:
            # 根太はIFC Y方向に走る
            x = j["pos"]   # IFC X → Three.js X
            z1 = -j["joist_start"]  # IFC Y → Three.js -Z
            z2 = -j["joist_end"]
            seg = [[x, slab_y, z1], [x, slab_y, z2]]

        lines.append({
            "seg": seg,
            "length_m": round(j["length"], 3),
            "reason": j["reason"],
            "comp_idx": j["comp_idx"],
        })

    return lines


###############################################################################
# 根太配置（1F基礎区画）
###############################################################################

def place_joists(compartments, rp, pitch=0.455):
    """基礎区画ごとに根太を配置。

    根太はslope_axis方向に走る。
    ridge_axis方向にpitchで並ぶ。
    区画の両端には必ず配置。

    IFC座標系:
      slope_axis="x" → 根太はIFC X方向に走る、IFC Yの方向にピッチで並ぶ
      slope_axis="z" → 根太はIFC Y方向に走る(Three.jsのZ)、IFC Xの方向にピッチで並ぶ

    Three.js座標系: IFC(X,Y,Z) → Three.js(X,Z,-Y)
      slope_axis="x" → 根太はThree.js X方向に走る、Three.js Zの方向(-IFC_Y方向)にピッチで並ぶ
      slope_axis="z" → 根太はThree.js Z方向(-IFC_Y)に走る、Three.js X方向にピッチで並ぶ
    """
    slope_axis = rp["slope_axis"]
    joists = []

    for comp_idx, comp in enumerate(compartments):
        # 根太の方向（slope_axis方向）と並ぶ方向（ridge_axis方向）を特定
        if slope_axis == "x":
            # 根太はIFC X方向に走る
            joist_start = comp["x_min"]  # 根太の始点（IFC X）
            joist_end = comp["x_max"]    # 根太の終点（IFC X）
            pitch_start = comp["y_min"]  # ピッチ方向の始点（IFC Y）
            pitch_end = comp["y_max"]    # ピッチ方向の終点（IFC Y）
        else:
            # slope_axis == "z" → 根太はIFC Y方向に走る
            joist_start = comp["y_min"]  # 根太の始点（IFC Y）
            joist_end = comp["y_max"]    # 根太の終点（IFC Y）
            pitch_start = comp["x_min"]  # ピッチ方向の始点（IFC X）
            pitch_end = comp["x_max"]    # ピッチ方向の終点（IFC X）

        joist_length = abs(joist_end - joist_start)
        pitch_span = abs(pitch_end - pitch_start)

        # ピッチ方向に根太を配置
        positions = []

        # 両端に必ず配置
        positions.append({"pos": pitch_start, "reason": "区画端部"})
        positions.append({"pos": pitch_end, "reason": "区画端部"})

        # ピッチ間隔で配置（start側から）
        pos = pitch_start + pitch
        while pos < pitch_end - 0.01:  # 端部と重複しないよう少しマージン
            positions.append({"pos": pos, "reason": "基本ピッチ"})
            pos += pitch

        # 重複除去（近すぎるものはマージ）
        positions.sort(key=lambda p: p["pos"])
        merged = []
        for p in positions:
            if merged and abs(p["pos"] - merged[-1]["pos"]) < 0.03:
                # 近い場合は端部優先
                if p["reason"] == "区画端部":
                    merged[-1] = p
                continue
            merged.append(p)

        for p in merged:
            joist = {
                "comp_idx": comp_idx,
                "pos": p["pos"],
                "reason": p["reason"],
                "joist_start": joist_start,
                "joist_end": joist_end,
                "length": joist_length,
            }
            joists.append(joist)

    return joists


###############################################################################
# 根太3Dライン生成
###############################################################################

def generate_joist_lines(joists, rp):
    """根太の3Dラインセグメント(Three.js座標)を生成。
    根太はスラブ上面（Y≈0.172m）に配置。
    IFC→Three.js: (X, Y, Z) → (X, Z, -Y)
    """
    slope_axis = rp["slope_axis"]
    slab_top_y = 0.172  # 1Fスラブ上面高さ（Three.js Y座標）
    lines = []

    for j in joists:
        if slope_axis == "x":
            # 根太はThree.js X方向に走る
            # IFC (X, Y, Z) → Three.js (X, Z, -Y)
            # joist_start/end = IFC X, pos = IFC Y
            x1 = j["joist_start"]
            x2 = j["joist_end"]
            z = -j["pos"]  # IFC Y → Three.js -Y = Three.js Z
            seg = [[x1, slab_top_y, z], [x2, slab_top_y, z]]
        else:
            # slope_axis == "z" → 根太はIFC Y方向に走る
            # joist_start/end = IFC Y, pos = IFC X
            x = j["pos"]  # IFC X → Three.js X
            z1 = -j["joist_start"]  # IFC Y → Three.js -Y
            z2 = -j["joist_end"]
            seg = [[x, slab_top_y, z1], [x, slab_top_y, z2]]

        lines.append({
            "seg": seg,
            "length_m": round(j["length"], 3),
            "reason": j["reason"],
            "comp_idx": j["comp_idx"],
        })

    return lines


###############################################################################
# メイン関数
###############################################################################

def main():
    if len(sys.argv) < 3:
        print("Usage: python build_neda_viewer.py <input.ifc> <output.html>")
        sys.exit(1)

    IFC_PATH = sys.argv[1]
    OUTPUT_HTML = sys.argv[2]

    model_name = os.path.splitext(os.path.basename(IFC_PATH))[0]
    print(f"=== 根太拾い: {model_name} ===")
    print(f"  IFC: {IFC_PATH}")

    ifc = ifcopenshell.open(IFC_PATH)
    settings = ifcopenshell.geom.settings()
    settings.set(settings.USE_WORLD_COORDS, True)

    # === Part 1: 建物ジオメトリ抽出 ===
    print("\n建物ジオメトリ抽出中...")
    meshes = []
    for elem in ifc.by_type("IfcProduct"):
        if elem.is_a("IfcOpeningElement"):
            continue
        ename = elem.Name or ""
        etype = elem.is_a()
        if etype in ("IfcWall", "IfcWallStandardCase"):
            cat = classify_wall_full(ename)
        elif etype == "IfcSlab":
            ln = ""
            try:
                t = ifcopenshell.util.element.get_type(elem)
                ln = t.Name if t else ""
            except Exception:
                pass
            cat = classify_slab(ename, ln)
        else:
            cat = TYPE_NAMES.get(etype)
        if cat is None:
            continue
        try:
            shape = ifcopenshell.geom.create_shape(settings, elem)
            vf = shape.geometry.verts
            ff = shape.geometry.faces
            gid = shape.geometry.id
            verts = []
            for i in range(0, len(vf), 3):
                verts.extend([vf[i], vf[i+2], -vf[i+1]])
            faces = list(ff)
            meshes.append({"cat": cat, "name": ename, "gid": gid,
                           "verts": verts, "faces": faces})
        except Exception:
            pass
    print(f"  建物ジオメトリ: {len(meshes)}個")

    # === Part 2: 屋根パラメータ検出（根太の方向を決定） ===
    print("\n屋根パラメータ検出中...")
    rp = detect_roof_params(ifc, settings)
    if rp is None:
        print("ERROR: 屋根が検出できませんでした")
        sys.exit(1)

    # === Part 3: 基礎区画検出 ===
    print("\n基礎区画検出中...")
    compartments = detect_kiso_compartments(ifc, settings, rp)
    if not compartments:
        print("ERROR: 基礎区画が検出できませんでした")
        sys.exit(1)

    # === Part 4: テラス・バルコニー区画検出 ===
    print("\nテラス・バルコニー区画検出中...")
    tb_compartments = detect_terrace_balcony_compartments(ifc, settings)

    # === Part 4b: 2F区画検出 ===
    print("\n2F区画検出中...")
    f2_compartments, f2_joist_dir = detect_2f_compartments(ifc, settings)

    # === Part 4c: 2Fスラブ開口検出 ===
    f2_openings = []
    if f2_compartments:
        print("\n2Fスラブ開口検出中...")
        f2_openings = detect_2f_openings(ifc, settings)

    # === Part 5-6: ピッチ別に根太配置＆ライン生成（1F + テラス・バルコニー + 2F） ===
    all_pitch_data = {}
    all_tb_pitch_data = {}
    all_f2_pitch_data = {}
    first_pitch = True

    for pitch in JOIST_PITCHES:
        pitch_mm = int(pitch * 1000)

        # --- 1F根太 ---
        print(f"\n1F根太配置計算中... (ピッチ: {pitch_mm}mm)")
        joists = place_joists(compartments, rp, pitch=pitch)
        joist_lines = generate_joist_lines(joists, rp)
        print(f"  1F根太ライン: {len(joist_lines)}本")

        total_count = len(joist_lines)
        total_length = sum(j["length_m"] for j in joist_lines)

        count_by_reason = defaultdict(int)
        length_by_reason = defaultdict(float)
        count_by_comp = defaultdict(int)
        length_by_comp = defaultdict(float)
        for j in joist_lines:
            count_by_reason[j["reason"]] += 1
            length_by_reason[j["reason"]] += j["length_m"]
            count_by_comp[j["comp_idx"]] += 1
            length_by_comp[j["comp_idx"]] += j["length_m"]

        summary_info = {
            "total_count": total_count,
            "total_length": round(total_length, 2),
            "count_by_reason": {k: v for k, v in count_by_reason.items()},
            "length_by_reason": {k: round(v, 2) for k, v in length_by_reason.items()},
            "count_by_comp": {str(k): v for k, v in count_by_comp.items()},
            "length_by_comp": {str(k): round(v, 2) for k, v in length_by_comp.items()},
            "pitch": pitch_mm,
            "compartment_count": len(compartments),
        }

        all_pitch_data[pitch_mm] = {
            "joists": joists,
            "joist_lines": joist_lines,
            "summary": summary_info,
        }

        # --- テラス・バルコニー根太 ---
        if tb_compartments:
            print(f"  テラス・バルコニー根太配置計算中... (ピッチ: {pitch_mm}mm)")
            tb_joists = place_joists_tb(tb_compartments, pitch=pitch)
            tb_joist_lines = generate_joist_lines_tb(tb_joists)
            print(f"  テラス・バルコニー根太ライン: {len(tb_joist_lines)}本")

            tb_total_count = len(tb_joist_lines)
            tb_total_length = sum(j["length_m"] for j in tb_joist_lines)

            tb_count_by_comp = defaultdict(int)
            tb_length_by_comp = defaultdict(float)
            for j in tb_joist_lines:
                tb_count_by_comp[j["comp_idx"]] += 1
                tb_length_by_comp[j["comp_idx"]] += j["length_m"]

            tb_summary = {
                "total_count": tb_total_count,
                "total_length": round(tb_total_length, 2),
                "count_by_comp": {str(k): v for k, v in tb_count_by_comp.items()},
                "length_by_comp": {str(k): round(v, 2) for k, v in tb_length_by_comp.items()},
                "pitch": pitch_mm,
                "compartment_count": len(tb_compartments),
            }

            all_tb_pitch_data[pitch_mm] = {
                "joist_lines": tb_joist_lines,
                "summary": tb_summary,
            }

        # --- 2F根太 ---
        if f2_compartments:
            print(f"  2F根太配置計算中... (ピッチ: {pitch_mm}mm)")
            f2_joists = place_joists_2f(f2_compartments, f2_joist_dir, openings=f2_openings, pitch=pitch)
            f2_joist_lines = generate_joist_lines_2f(f2_joists)
            print(f"  2F根太ライン: {len(f2_joist_lines)}本")

            f2_total_count = len(f2_joist_lines)
            f2_total_length = sum(j["length_m"] for j in f2_joist_lines)

            f2_count_by_comp = defaultdict(int)
            f2_length_by_comp = defaultdict(float)
            for j in f2_joist_lines:
                f2_count_by_comp[j["comp_idx"]] += 1
                f2_length_by_comp[j["comp_idx"]] += j["length_m"]

            f2_summary = {
                "total_count": f2_total_count,
                "total_length": round(f2_total_length, 2),
                "count_by_comp": {str(k): v for k, v in f2_count_by_comp.items()},
                "length_by_comp": {str(k): round(v, 2) for k, v in f2_length_by_comp.items()},
                "pitch": pitch_mm,
                "compartment_count": len(f2_compartments),
            }

            all_f2_pitch_data[pitch_mm] = {
                "joist_lines": f2_joist_lines,
                "summary": f2_summary,
            }

        if first_pitch:
            print(f"\n=== 1F根太積算 ({pitch_mm}mm) ===")
            print(f"  総本数: {total_count}本")
            print(f"  総長さ: {total_length:.2f}m")
            print(f"\n  理由別:")
            for reason in sorted(count_by_reason.keys()):
                print(f"    {reason}: {count_by_reason[reason]}本 / {length_by_reason[reason]:.2f}m")
            print(f"\n  区画別:")
            for ci in sorted(count_by_comp.keys()):
                c = compartments[ci]
                print(f"    区画{ci+1} ({c['width_x']:.2f}x{c['width_y']:.2f}m): "
                      f"{count_by_comp[ci]}本 / {length_by_comp[ci]:.2f}m")
            if tb_compartments:
                print(f"\n=== テラス・バルコニー根太積算 ({pitch_mm}mm) ===")
                print(f"  総本数: {tb_total_count}本")
                print(f"  総長さ: {tb_total_length:.2f}m")
                for ci, c in enumerate(tb_compartments):
                    cnt = tb_count_by_comp.get(ci, 0)
                    ln = tb_length_by_comp.get(ci, 0)
                    print(f"    {c['category']} ({c['width_x']:.2f}x{c['width_y']:.2f}m): "
                          f"{cnt}本 / {ln:.2f}m")
            if f2_compartments:
                print(f"\n=== 2F根太積算 ({pitch_mm}mm) ===")
                print(f"  総本数: {f2_total_count}本")
                print(f"  総長さ: {f2_total_length:.2f}m")
                for ci, c in enumerate(f2_compartments):
                    cnt = f2_count_by_comp.get(ci, 0)
                    ln = f2_length_by_comp.get(ci, 0)
                    print(f"    2F区画{ci+1} ({c['width_x']:.2f}x{c['width_y']:.2f}m): "
                          f"{cnt}本 / {ln:.2f}m")

        first_pitch = False

    # === HTML出力 ===
    default_pitch = int(JOIST_PITCHES[0] * 1000)
    meshes_json = json.dumps(meshes, ensure_ascii=False)
    colors_json = json.dumps(CATEGORY_COLORS, ensure_ascii=False)

    # 1F区画情報をJSON化（Three.js座標系）
    comp_data = []
    for c in compartments:
        comp_data.append({
            "x_min": c["x_min"], "x_max": c["x_max"],
            "z_min": -c["y_max"], "z_max": -c["y_min"],
            "width_x": c["width_x"], "width_y": c["width_y"],
        })
    compartments_json = json.dumps(comp_data, ensure_ascii=False)

    # テラス・バルコニー区画情報をJSON化（Three.js座標系）
    tb_comp_data = []
    for c in tb_compartments:
        tb_comp_data.append({
            "x_min": c["x_min"], "x_max": c["x_max"],
            "z_min": -c["y_max"], "z_max": -c["y_min"],
            "width_x": c["width_x"], "width_y": c["width_y"],
            "slab_y": c["slab_top_z"],  # Three.js Y = IFC Z
            "category": c["category"],
        })
    tb_compartments_json = json.dumps(tb_comp_data, ensure_ascii=False)

    # 2F区画情報をJSON化（Three.js座標系）
    f2_comp_data = []
    for c in f2_compartments:
        f2_comp_data.append({
            "x_min": c["x_min"], "x_max": c["x_max"],
            "z_min": -c["y_max"], "z_max": -c["y_min"],
            "width_x": c["width_x"], "width_y": c["width_y"],
            "slab_y": c["slab_top_z"],
        })
    f2_compartments_json = json.dumps(f2_comp_data, ensure_ascii=False)

    # 2F開口情報をJSON化（Three.js座標系）
    f2_opening_data = []
    slab_top_z_for_opening = f2_compartments[0]["slab_top_z"] if f2_compartments else 0
    for o in f2_openings:
        f2_opening_data.append({
            "x_min": o["x_min"], "x_max": o["x_max"],
            "z_min": -o["y_max"], "z_max": -o["y_min"],
            "slab_y": slab_top_z_for_opening,
        })
    f2_openings_json = json.dumps(f2_opening_data, ensure_ascii=False)

    pitch_data_for_html = {}
    for pitch_mm, pdata in all_pitch_data.items():
        pitch_data_for_html[pitch_mm] = {
            "joists": pdata["joist_lines"],
            "summary": pdata["summary"],
        }
    pitch_data_json = json.dumps(pitch_data_for_html, ensure_ascii=False)

    tb_pitch_data_json = json.dumps(
        {pm: {"joists": d["joist_lines"], "summary": d["summary"]}
         for pm, d in all_tb_pitch_data.items()},
        ensure_ascii=False
    ) if all_tb_pitch_data else "{}"

    f2_pitch_data_json = json.dumps(
        {pm: {"joists": d["joist_lines"], "summary": d["summary"]}
         for pm, d in all_f2_pitch_data.items()},
        ensure_ascii=False
    ) if all_f2_pitch_data else "{}"

    pitches_json = json.dumps([int(p * 1000) for p in JOIST_PITCHES])

    html = generate_html(meshes_json, pitch_data_json, pitches_json,
                         default_pitch, colors_json, compartments_json,
                         tb_compartments_json, tb_pitch_data_json,
                         f2_compartments_json, f2_pitch_data_json,
                         f2_openings_json, model_name)

    os.makedirs(os.path.dirname(OUTPUT_HTML) or ".", exist_ok=True)
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"\n出力: {OUTPUT_HTML}")

    # === JSON集計出力 ===
    json_path = os.path.splitext(OUTPUT_HTML)[0] + "_summary.json"
    dp = all_pitch_data[default_pitch]
    summary_out = {
        "tool": "根太拾い",
        "model": model_name,
        "pitch_mm": default_pitch,
        "total_count": dp["summary"]["total_count"],
        "total_length_m": dp["summary"]["total_length"],
        "count_by_reason": dp["summary"]["count_by_reason"],
        "length_by_reason": dp["summary"]["length_by_reason"],
        "compartment_count": len(compartments),
    }
    if default_pitch in all_tb_pitch_data:
        tb_dp = all_tb_pitch_data[default_pitch]
        summary_out["tb_total_count"] = tb_dp["summary"]["total_count"]
        summary_out["tb_total_length_m"] = tb_dp["summary"]["total_length"]
        summary_out["tb_compartment_count"] = len(tb_compartments)
    if default_pitch in all_f2_pitch_data:
        f2_dp = all_f2_pitch_data[default_pitch]
        summary_out["f2_total_count"] = f2_dp["summary"]["total_count"]
        summary_out["f2_total_length_m"] = f2_dp["summary"]["total_length"]
        summary_out["f2_compartment_count"] = len(f2_compartments)
        summary_out["f2_joist_dir"] = f2_joist_dir
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(summary_out, f, ensure_ascii=False, indent=2)
    print(f"JSON集計: {json_path}")


###############################################################################
# HTML生成
###############################################################################

def generate_html(meshes_json, pitch_data_json, pitches_json,
                  default_pitch, colors_json, compartments_json,
                  tb_compartments_json="{}", tb_pitch_data_json="{}",
                  f2_compartments_json="{}", f2_pitch_data_json="{}",
                  f2_openings_json="[]", model_name=""):
    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>根太拾い 3Dビューア - {model_name}</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ background:#1a1a2e; overflow:hidden; font-family:Arial,sans-serif; }}
#info {{ position:fixed; top:10px; left:10px; color:#fff; background:rgba(0,0,0,0.75);
  padding:12px 16px; border-radius:8px; font-size:13px; z-index:100; max-width:360px; }}
#info h3 {{ margin-bottom:6px; color:#4fc3f7; font-size:15px; }}
#legend {{ position:fixed; top:10px; right:10px; color:#fff; background:rgba(0,0,0,0.75);
  padding:12px; border-radius:8px; font-size:12px; z-index:100; max-height:80vh; overflow-y:auto; }}
#legend div {{ cursor:pointer; padding:3px 6px; border-radius:3px; margin:2px 0; white-space:nowrap; }}
#legend div:hover {{ background:rgba(255,255,255,0.15); }}
.cb {{ display:inline-block; width:14px; height:14px; border-radius:3px; margin-right:6px; vertical-align:middle; }}
#neda-info {{ position:fixed; bottom:10px; left:10px; color:#fff; background:rgba(0,0,0,0.85);
  padding:14px 18px; border-radius:8px; font-size:13px; z-index:100; line-height:1.6; max-width:450px;
  max-height:50vh; overflow-y:auto; }}
#controls {{ position:fixed; bottom:10px; right:10px; z-index:100; display:flex; flex-wrap:wrap; gap:4px; align-items:center; }}
#controls button {{ background:rgba(255,255,255,0.15); color:#fff; border:1px solid rgba(255,255,255,0.3);
  padding:8px 14px; border-radius:6px; cursor:pointer; font-size:12px; }}
#controls button:hover {{ background:rgba(255,255,255,0.3); }}
#controls button.active {{ background:rgba(79,195,247,0.5); border-color:#4fc3f7; }}
#pitch-group {{ display:flex; gap:2px; margin-right:8px; }}
#pitch-group button {{ padding:8px 12px; }}
.comp-label {{
  color:#fff; font-size:12px; font-weight:bold; padding:3px 8px;
  border-radius:4px; white-space:nowrap; pointer-events:none;
  text-shadow:0 0 3px rgba(0,0,0,0.8);
  font-family:Arial,sans-serif; line-height:1.3; text-align:center;
}}
</style>
</head>
<body>
<div id="info">
  <h3>根太拾い 3Dビューア</h3>
  <div>左ドラッグ: 回転 / 右ドラッグ: 移動 / ホイール: ズーム</div>
  <div id="sel-info" style="margin-top:6px;color:#aaa;">クリックで部材選択</div>
</div>
<div id="neda-info"></div>
<div id="legend"></div>
<div id="controls">
  <div id="pitch-group"></div>
  <button id="btn-neda" class="active" onclick="toggleNeda()">1F根太</button>
  <button id="btn-tb" class="active" onclick="toggleTB()">テラス・バルコニー根太</button>
  <button id="btn-f2" class="active" onclick="toggleF2()">2F根太</button>
  <button id="btn-comp" class="active" onclick="toggleComp()">区画表示</button>
  <button id="btn-building" class="active" onclick="toggleBuilding()">建物表示</button>
  <button onclick="resetCam()">リセット</button>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/lines/LineSegmentsGeometry.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/lines/LineGeometry.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/lines/LineMaterial.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/lines/LineSegments2.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/lines/Line2.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/renderers/CSS2DRenderer.js"></script>
<script>
const MESHES={meshes_json};
const PITCH_DATA={pitch_data_json};
const TB_PITCH_DATA={tb_pitch_data_json};
const TB_COMPARTMENTS={tb_compartments_json};
const F2_PITCH_DATA={f2_pitch_data_json};
const F2_COMPARTMENTS={f2_compartments_json};
const F2_OPENINGS={f2_openings_json};
const PITCHES={pitches_json};
const DEFAULT_PITCH={default_pitch};
const COLORS={colors_json};
const COMPARTMENTS={compartments_json};

let currentPitch=DEFAULT_PITCH;

const REASON_COLORS = {{
  "基本ピッチ": 0x4fc3f7,
  "区画端部": 0x00e676,
}};

// 区画ごとの色
const COMP_COLORS = [0xff6b6b, 0x51cf66, 0x339af0, 0xfcc419, 0xcc5de8,
                      0x20c997, 0xff922b, 0x845ef7, 0xe64980];

const W=innerWidth, H=innerHeight;
const scene=new THREE.Scene();
scene.background=new THREE.Color(0x1a1a2e);
const camera=new THREE.PerspectiveCamera(50,W/H,0.01,1000);
const renderer=new THREE.WebGLRenderer({{antialias:true}});
renderer.setSize(W,H); renderer.setPixelRatio(devicePixelRatio);
document.body.appendChild(renderer.domElement);

const labelRenderer=new THREE.CSS2DRenderer();
labelRenderer.setSize(W,H);
labelRenderer.domElement.style.position='absolute';
labelRenderer.domElement.style.top='0px';
labelRenderer.domElement.style.pointerEvents='none';
document.body.appendChild(labelRenderer.domElement);

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
e.addEventListener('touchend',()=>{{this._st=0;}});}}
_rot(dx,dy){{const o=this.c.position.clone().sub(this.t);this.s.setFromVector3(o);this.s.theta-=dx*0.008;this.s.phi-=dy*0.008;
this.s.phi=Math.max(0.01,Math.min(Math.PI-0.01,this.s.phi));o.setFromSpherical(this.s);this.c.position.copy(this.t).add(o);this.c.lookAt(this.t);}}
_pan(dx,dy){{const d=this.c.position.distanceTo(this.t)*0.001;const r=new THREE.Vector3().setFromMatrixColumn(this.c.matrix,0);
const u=new THREE.Vector3().setFromMatrixColumn(this.c.matrix,1);const p=r.multiplyScalar(-dx*d).add(u.multiplyScalar(dy*d));
this.c.position.add(p);this.t.add(p);this.c.lookAt(this.t);}}
_zm(f){{const o=this.c.position.clone().sub(this.t);o.multiplyScalar(f);this.c.position.copy(this.t).add(o);this.c.lookAt(this.t);}}
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

// 区画表示
const compGroup=new THREE.Group();
scene.add(compGroup);
const slabY=0.172;
COMPARTMENTS.forEach((c,i)=>{{
  const w=c.x_max-c.x_min, d=c.z_max-c.z_min;
  const g=new THREE.PlaneGeometry(w,Math.abs(d));
  const color=COMP_COLORS[i%COMP_COLORS.length];
  const mat=new THREE.MeshBasicMaterial({{color,transparent:true,opacity:0.15,side:THREE.DoubleSide}});
  const mesh=new THREE.Mesh(g,mat);
  mesh.rotation.x=-Math.PI/2;
  mesh.position.set((c.x_min+c.x_max)/2, slabY+0.005, (c.z_min+c.z_max)/2);
  compGroup.add(mesh);
  // ラベル
  const hex='#'+color.toString(16).padStart(6,'0');
  const lDiv=document.createElement('div');lDiv.className='comp-label';
  lDiv.style.background='rgba(0,0,0,0.6)';lDiv.style.borderLeft='3px solid '+hex;
  lDiv.innerHTML='1F-区画'+(i+1)+'<br>'+w.toFixed(2)+'×'+Math.abs(d).toFixed(2)+'m';
  const lObj=new THREE.CSS2DObject(lDiv);
  lObj.position.set((c.x_min+c.x_max)/2, slabY+0.05, (c.z_min+c.z_max)/2);
  compGroup.add(lObj);
}});

// テラス・バルコニー区画表示
const TB_COLORS = {{ "テラス": 0x66bb6a, "バルコニー": 0x42a5f5 }};
TB_COMPARTMENTS.forEach((c,i)=>{{
  const w=c.x_max-c.x_min, d=c.z_max-c.z_min;
  const g=new THREE.PlaneGeometry(w,Math.abs(d));
  const color=TB_COLORS[c.category]||0x888888;
  const mat=new THREE.MeshBasicMaterial({{color,transparent:true,opacity:0.15,side:THREE.DoubleSide}});
  const mesh=new THREE.Mesh(g,mat);
  mesh.rotation.x=-Math.PI/2;
  mesh.position.set((c.x_min+c.x_max)/2, c.slab_y+0.005, (c.z_min+c.z_max)/2);
  compGroup.add(mesh);
  // ラベル
  const hex='#'+color.toString(16).padStart(6,'0');
  const lDiv=document.createElement('div');lDiv.className='comp-label';
  lDiv.style.background='rgba(0,0,0,0.6)';lDiv.style.borderLeft='3px solid '+hex;
  lDiv.innerHTML=c.category+'<br>'+w.toFixed(2)+'×'+Math.abs(d).toFixed(2)+'m';
  const lObj=new THREE.CSS2DObject(lDiv);
  lObj.position.set((c.x_min+c.x_max)/2, c.slab_y+0.05, (c.z_min+c.z_max)/2);
  compGroup.add(lObj);
}});

// 2F区画表示（区画ごとに色分け）
const F2_COMP_COLORS = [0xce93d8, 0xba68c8, 0xab47bc, 0x9c27b0, 0x8e24aa, 0x7b1fa2];
F2_COMPARTMENTS.forEach((c,i)=>{{
  const w=c.x_max-c.x_min, d=c.z_max-c.z_min;
  const g=new THREE.PlaneGeometry(w,Math.abs(d));
  const color=F2_COMP_COLORS[i%F2_COMP_COLORS.length];
  const mat=new THREE.MeshBasicMaterial({{color,transparent:true,opacity:0.15,side:THREE.DoubleSide}});
  const mesh=new THREE.Mesh(g,mat);
  mesh.rotation.x=-Math.PI/2;
  mesh.position.set((c.x_min+c.x_max)/2, c.slab_y+0.005, (c.z_min+c.z_max)/2);
  compGroup.add(mesh);
  // ラベル
  const hex='#'+color.toString(16).padStart(6,'0');
  const lDiv=document.createElement('div');lDiv.className='comp-label';
  lDiv.style.background='rgba(0,0,0,0.6)';lDiv.style.borderLeft='3px solid '+hex;
  lDiv.innerHTML='2F-区画'+(i+1)+'<br>'+w.toFixed(2)+'×'+Math.abs(d).toFixed(2)+'m';
  const lObj=new THREE.CSS2DObject(lDiv);
  lObj.position.set((c.x_min+c.x_max)/2, c.slab_y+0.05, (c.z_min+c.z_max)/2);
  compGroup.add(lObj);
}});

// 2F開口表示（赤い半透明）
F2_OPENINGS.forEach((o,i)=>{{
  const w=o.x_max-o.x_min, d=o.z_max-o.z_min;
  const g=new THREE.PlaneGeometry(w,Math.abs(d));
  const mat=new THREE.MeshBasicMaterial({{color:0xff1744,transparent:true,opacity:0.25,side:THREE.DoubleSide}});
  const mesh=new THREE.Mesh(g,mat);
  mesh.rotation.x=-Math.PI/2;
  mesh.position.set((o.x_min+o.x_max)/2, o.slab_y+0.01, (o.z_min+o.z_max)/2);
  compGroup.add(mesh);
  // ラベル
  const lDiv=document.createElement('div');lDiv.className='comp-label';
  lDiv.style.background='rgba(255,23,68,0.6)';lDiv.style.borderLeft='3px solid #ff1744';
  lDiv.innerHTML='開口'+(i+1)+'<br>'+w.toFixed(2)+'×'+Math.abs(d).toFixed(2)+'m';
  const lObj=new THREE.CSS2DObject(lDiv);
  lObj.position.set((o.x_min+o.x_max)/2, o.slab_y+0.06, (o.z_min+o.z_max)/2);
  compGroup.add(lObj);
}});

// 1F根太ライン描画
const nedaGroup=new THREE.Group();
scene.add(nedaGroup);

// テラス・バルコニー根太ライン描画
const tbNedaGroup=new THREE.Group();
scene.add(tbNedaGroup);

// 2F根太ライン描画
const f2NedaGroup=new THREE.Group();
scene.add(f2NedaGroup);

function buildNedaLines(pitchMm){{
  while(nedaGroup.children.length>0) nedaGroup.remove(nedaGroup.children[0]);
  const pd=PITCH_DATA[pitchMm];
  if(!pd) return;
  pd.joists.forEach(j=>{{
    const compColor=COMP_COLORS[j.comp_idx%COMP_COLORS.length];
    const color=j.reason==="区画端部"?0x00e676:compColor;
    const s=j.seg;
    const g=new THREE.LineGeometry();
    g.setPositions([s[0][0],s[0][1],s[0][2],s[1][0],s[1][1],s[1][2]]);
    const mat=new THREE.LineMaterial({{color,linewidth:4,resolution:new THREE.Vector2(innerWidth,innerHeight)}});
    const line=new THREE.Line2(g,mat);
    line.userData={{reason:j.reason,length:j.length_m,comp_idx:j.comp_idx}};
    nedaGroup.add(line);
  }});
  // 集計表示更新
  const S=pd.summary;
  let h='<b style="font-size:14px;">根太積算</b><br><br>';
  h+=`ピッチ: ${{S.pitch}}mm<br>`;
  h+=`<b>総本数: ${{S.total_count}}本</b><br>`;
  h+=`総長さ: ${{S.total_length}}m<br><br>`;
  h+='<b>理由別:</b><br>';
  for(const [reason, count] of Object.entries(S.count_by_reason)){{
    const color=REASON_COLORS[reason]||0x4fc3f7;
    const hex='#'+color.toString(16).padStart(6,'0');
    const len=S.length_by_reason[reason]||0;
    h+=`<span style="color:${{hex}}">━━</span> ${{reason}}: ${{count}}本 / ${{len}}m<br>`;
  }}
  h+='<br><b>区画別:</b><br>';
  for(let i=0;i<S.compartment_count;i++){{
    const cnt=S.count_by_comp[String(i)]||0;
    const len=S.length_by_comp[String(i)]||0;
    const color=COMP_COLORS[i%COMP_COLORS.length];
    const hex='#'+color.toString(16).padStart(6,'0');
    const c=COMPARTMENTS[i];
    h+=`<span style="color:${{hex}}">■</span> 区画${{i+1}} (${{c.width_x.toFixed(2)}}x${{c.width_y.toFixed(2)}}m): ${{cnt}}本 / ${{len}}m<br>`;
  }}
  // テラス・バルコニー根太の集計
  const tbPd=TB_PITCH_DATA[pitchMm];
  if(tbPd && tbPd.summary.total_count>0){{
    h+='<br><b style="font-size:14px;">テラス・バルコニー根太</b><br><br>';
    h+=`ピッチ: ${{tbPd.summary.pitch}}mm<br>`;
    h+=`<b>総本数: ${{tbPd.summary.total_count}}本</b><br>`;
    h+=`総長さ: ${{tbPd.summary.total_length}}m<br><br>`;
    for(let i=0;i<tbPd.summary.compartment_count;i++){{
      const cnt=tbPd.summary.count_by_comp[String(i)]||0;
      const len=tbPd.summary.length_by_comp[String(i)]||0;
      const c=TB_COMPARTMENTS[i];
      const color=TB_COLORS[c.category]||0x888888;
      const hex='#'+color.toString(16).padStart(6,'0');
      h+=`<span style="color:${{hex}}">■</span> ${{c.category}} (${{c.width_x.toFixed(2)}}x${{c.width_y.toFixed(2)}}m): ${{cnt}}本 / ${{len}}m<br>`;
    }}
  }}
  // 2F根太の集計
  const f2Pd=F2_PITCH_DATA[pitchMm];
  if(f2Pd && f2Pd.summary.total_count>0){{
    h+='<br><b style="font-size:14px;">2F根太</b><br><br>';
    h+=`ピッチ: ${{f2Pd.summary.pitch}}mm<br>`;
    h+=`<b>総本数: ${{f2Pd.summary.total_count}}本</b><br>`;
    h+=`総長さ: ${{f2Pd.summary.total_length}}m<br><br>`;
    for(let i=0;i<f2Pd.summary.compartment_count;i++){{
      const cnt=f2Pd.summary.count_by_comp[String(i)]||0;
      const len=f2Pd.summary.length_by_comp[String(i)]||0;
      const c=F2_COMPARTMENTS[i];
      const ccolor=F2_COMP_COLORS[i%F2_COMP_COLORS.length];
      const hex='#'+ccolor.toString(16).padStart(6,'0');
      h+=`<span style="color:${{hex}}">■</span> 2F区画${{i+1}} (${{c.width_x.toFixed(2)}}x${{c.width_y.toFixed(2)}}m): ${{cnt}}本 / ${{len}}m<br>`;
    }}
  }}
  document.getElementById('neda-info').innerHTML=h;
}}

function buildTBNedaLines(pitchMm){{
  while(tbNedaGroup.children.length>0) tbNedaGroup.remove(tbNedaGroup.children[0]);
  const pd=TB_PITCH_DATA[pitchMm];
  if(!pd) return;
  pd.joists.forEach(j=>{{
    const c=TB_COMPARTMENTS[j.comp_idx];
    const color=TB_COLORS[c.category]||0x888888;
    const s=j.seg;
    const g=new THREE.LineGeometry();
    g.setPositions([s[0][0],s[0][1],s[0][2],s[1][0],s[1][1],s[1][2]]);
    const mat=new THREE.LineMaterial({{color,linewidth:4,resolution:new THREE.Vector2(innerWidth,innerHeight)}});
    const line=new THREE.Line2(g,mat);
    line.userData={{reason:j.reason,length:j.length_m,comp_idx:j.comp_idx}};
    tbNedaGroup.add(line);
  }});
}}

function buildF2NedaLines(pitchMm){{
  while(f2NedaGroup.children.length>0) f2NedaGroup.remove(f2NedaGroup.children[0]);
  const pd=F2_PITCH_DATA[pitchMm];
  if(!pd) return;
  pd.joists.forEach(j=>{{
    const compColor=F2_COMP_COLORS[j.comp_idx%F2_COMP_COLORS.length];
    const color=j.reason==="区画端部"?0xea80fc:compColor;
    const s=j.seg;
    const g=new THREE.LineGeometry();
    g.setPositions([s[0][0],s[0][1],s[0][2],s[1][0],s[1][1],s[1][2]]);
    const mat=new THREE.LineMaterial({{color,linewidth:4,resolution:new THREE.Vector2(innerWidth,innerHeight)}});
    const line=new THREE.Line2(g,mat);
    line.userData={{reason:j.reason,length:j.length_m,comp_idx:j.comp_idx}};
    f2NedaGroup.add(line);
  }});
}}

function switchPitch(pitchMm){{
  currentPitch=pitchMm;
  buildNedaLines(pitchMm);
  buildTBNedaLines(pitchMm);
  buildF2NedaLines(pitchMm);
  document.querySelectorAll('#pitch-group button').forEach(b=>{{
    b.classList.toggle('active', parseInt(b.dataset.pitch)===pitchMm);
  }});
}}

const pg=document.getElementById('pitch-group');
PITCHES.forEach(p=>{{
  const b=document.createElement('button');
  b.textContent=p+'mm';
  b.dataset.pitch=p;
  b.onclick=()=>switchPitch(p);
  if(p===DEFAULT_PITCH) b.classList.add('active');
  pg.appendChild(b);
}});

buildNedaLines(DEFAULT_PITCH);
buildTBNedaLines(DEFAULT_PITCH);
buildF2NedaLines(DEFAULT_PITCH);

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

let showN=true,showTB=true,showF2=true,showC=true,showB=true;
function toggleNeda(){{showN=!showN;nedaGroup.visible=showN;document.getElementById('btn-neda').classList.toggle('active',showN);}}
function toggleTB(){{showTB=!showTB;tbNedaGroup.visible=showTB;document.getElementById('btn-tb').classList.toggle('active',showTB);}}
function toggleF2(){{showF2=!showF2;f2NedaGroup.visible=showF2;document.getElementById('btn-f2').classList.toggle('active',showF2);}}
function toggleComp(){{showC=!showC;compGroup.visible=showC;document.getElementById('btn-comp').classList.toggle('active',showC);}}
function toggleBuilding(){{showB=!showB;buildingGroup.visible=showB;document.getElementById('btn-building').classList.toggle('active',showB);}}

const grid=new THREE.GridHelper(20,40,0x444444,0x333333);grid.position.copy(center);grid.position.y=0;scene.add(grid);
(function anim(){{requestAnimationFrame(anim);renderer.render(scene,camera);labelRenderer.render(scene,camera);}})();
addEventListener('resize',()=>{{camera.aspect=innerWidth/innerHeight;camera.updateProjectionMatrix();renderer.setSize(innerWidth,innerHeight);labelRenderer.setSize(innerWidth,innerHeight);scene.traverse(c=>{{if(c.material&&c.material.resolution)c.material.resolution.set(innerWidth,innerHeight);}});}});
</script>
</body>
</html>"""


if __name__ == "__main__":
    main()

