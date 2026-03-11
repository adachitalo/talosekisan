#!/usr/bin/env python3
"""廻り縁拾い3Dビューア: 天井-壁取り合い部の廻り縁長さを自動算出
Usage:
  python build_mawari_buchi_viewer.py <input.ifc> <output.html>
  python build_mawari_buchi_viewer.py   # デフォルトパス使用
"""
import sys
import os
import ifcopenshell
import ifcopenshell.geom
import ifcopenshell.util.element
import json
import numpy as np
from collections import defaultdict

# デフォルトパス（CLI引数で上書き可能）
IFC_PATH = "input.ifc"
OUTPUT_HTML = "output/index.html"

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


###############################################################################
# 屋根勾配パラメータ自動検出
###############################################################################
# グローバル変数（main()内で自動検出結果で上書き）
RIDGE_AXIS = "z"      # 棟が走る軸 ("x" or "z")
SLOPE_AXIS = "x"      # 勾配が変わる軸 ("x" or "z")
RIDGE_POS = 3.500     # 棟の slope_axis 座標
RIDGE_HEIGHT = 5.880  # 棟での天井高さ（Y）
EAVE_POS_MIN = -0.778 # 軒先1の slope_axis 座標
EAVE_POS_MAX = 7.778  # 軒先2の slope_axis 座標
EAVE_HEIGHT = 3.313   # 軒先での天井高さ（Y）


def detect_roof_params(ifc, settings):
    """IFC屋根メッシュから勾配天井パラメータを自動検出。
    面法線で屋根底面（天井面）を特定し、棟の方向・位置を算出。
    Returns dict with ridge_axis, slope_axis, ridge_pos, ridge_height,
    eave_pos_min, eave_pos_max, eave_height (all in Three.js coords).
    """
    bottom_verts = []  # 下向き法線の面の頂点（屋根底面＝天井面）

    for slab in ifc.by_type("IfcSlab"):
        ename = (slab.Name or "").lower()
        if "yane" not in ename and "roof" not in ename:
            continue
        try:
            shape = ifcopenshell.geom.create_shape(settings, slab)
            vf = shape.geometry.verts
            ff = shape.geometry.faces
            pts = []
            for i in range(0, len(vf), 3):
                pts.append([vf[i], vf[i+2], -vf[i+1]])
            pts = np.array(pts)
            for i in range(0, len(ff), 3):
                i0, i1, i2 = ff[i], ff[i+1], ff[i+2]
                p0, p1, p2 = pts[i0], pts[i1], pts[i2]
                normal = np.cross(p1 - p0, p2 - p0)
                norm_len = np.linalg.norm(normal)
                if norm_len < 1e-10:
                    continue
                normal = normal / norm_len
                if normal[1] < -0.3:
                    bottom_verts.extend([pts[i0], pts[i1], pts[i2]])
        except Exception as e:
            print(f"  Error processing slab: {e}")

    for roof in ifc.by_type("IfcRoof"):
        try:
            shape = ifcopenshell.geom.create_shape(settings, roof)
            vf = shape.geometry.verts
            ff = shape.geometry.faces
            pts = []
            for i in range(0, len(vf), 3):
                pts.append([vf[i], vf[i+2], -vf[i+1]])
            pts = np.array(pts)
            for i in range(0, len(ff), 3):
                i0, i1, i2 = ff[i], ff[i+1], ff[i+2]
                p0, p1, p2 = pts[i0], pts[i1], pts[i2]
                normal = np.cross(p1 - p0, p2 - p0)
                norm_len = np.linalg.norm(normal)
                if norm_len < 1e-10:
                    continue
                normal = normal / norm_len
                if normal[1] < -0.3:
                    bottom_verts.extend([pts[i0], pts[i1], pts[i2]])
        except Exception:
            pass

    if not bottom_verts:
        print("WARNING: No roof underside faces found!")
        return None

    bv = np.array(bottom_verts)
    bv = np.unique(np.round(bv, 3), axis=0)

    print(f"  Roof underside vertices: {len(bv)} unique points")
    print(f"    X range: [{bv[:,0].min():.3f}, {bv[:,0].max():.3f}]")
    print(f"    Y range: [{bv[:,1].min():.3f}, {bv[:,1].max():.3f}]")
    print(f"    Z range: [{bv[:,2].min():.3f}, {bv[:,2].max():.3f}]")

    # 棟: Y最高の底面頂点群
    ridge_height = bv[:, 1].max()
    ridge_tol = 0.15
    ridge_mask = bv[:, 1] > ridge_height - ridge_tol
    ridge_pts = bv[ridge_mask]

    # 棟の方向: X方向スパン vs Z方向スパン
    ridge_x_span = ridge_pts[:, 0].max() - ridge_pts[:, 0].min()
    ridge_z_span = ridge_pts[:, 2].max() - ridge_pts[:, 2].min()

    if ridge_x_span > ridge_z_span:
        ridge_axis = "x"
        slope_axis = "z"
        ridge_pos = float(np.median(ridge_pts[:, 2]))
        eave_pos_min = float(bv[:, 2].min())
        eave_pos_max = float(bv[:, 2].max())
    else:
        ridge_axis = "z"
        slope_axis = "x"
        ridge_pos = float(np.median(ridge_pts[:, 0]))
        eave_pos_min = float(bv[:, 0].min())
        eave_pos_max = float(bv[:, 0].max())

    # 軒先高さ: slope_axis端の頂点Y
    if slope_axis == "x":
        eave_mask = (np.abs(bv[:, 0] - eave_pos_min) < 0.3) | \
                    (np.abs(bv[:, 0] - eave_pos_max) < 0.3)
    else:
        eave_mask = (np.abs(bv[:, 2] - eave_pos_min) < 0.3) | \
                    (np.abs(bv[:, 2] - eave_pos_max) < 0.3)
    eave_pts = bv[eave_mask]
    eave_height = float(eave_pts[:, 1].min()) if len(eave_pts) > 0 else float(bv[:, 1].min())

    result = {
        "ridge_axis": ridge_axis,
        "slope_axis": slope_axis,
        "ridge_pos": ridge_pos,
        "ridge_height": float(ridge_height),
        "eave_pos_min": eave_pos_min,
        "eave_pos_max": eave_pos_max,
        "eave_height": eave_height,
    }

    print(f"\n  Detected roof parameters:")
    print(f"    Ridge axis: {ridge_axis} (ridge runs along {ridge_axis.upper()})")
    print(f"    Slope axis: {slope_axis} (slope varies along {slope_axis.upper()})")
    print(f"    Ridge position ({slope_axis.upper()}): {ridge_pos:.3f}")
    print(f"    Ridge height (Y): {ridge_height:.3f}")
    print(f"    Eave positions ({slope_axis.upper()}): [{eave_pos_min:.3f}, {eave_pos_max:.3f}]")
    print(f"    Eave height (Y): {eave_height:.3f}")

    return result


def detect_ceiling_levels(ifc, settings, roof_params):
    """IFCジオメトリから天井レベルを自動検出。
    1F天井: 2-yuka スラブ下面（なければNone = 平屋）
    2F天井: ログ壁top_y分布のギャップ解析
    """
    # --- 1F天井: 2-yuka スラブ下面 ---
    ceiling_1f = None
    for slab in ifc.by_type("IfcSlab"):
        ename = (slab.Name or "").lower()
        if ename == "2-yuka":
            try:
                shape = ifcopenshell.geom.create_shape(settings, slab)
                vf = shape.geometry.verts
                ys = [vf[i+2] for i in range(0, len(vf), 3)]  # IFC Z → Three.js Y
                ceiling_1f = min(ys)
                print(f"  ceiling_1f from 2-yuka slab bottom: {ceiling_1f:.3f}")
            except Exception:
                pass

    if ceiling_1f is None:
        # 平屋の可能性 → majikiri壁から推定試行
        majikiri_tops = []
        for wall in ifc.by_type("IfcWall"):
            ename = wall.Name or ""
            if ename.startswith("majikiri"):
                try:
                    shape = ifcopenshell.geom.create_shape(settings, wall)
                    vf = shape.geometry.verts
                    ys = [vf[i+2] for i in range(0, len(vf), 3)]
                    majikiri_tops.append(max(ys))
                except Exception:
                    pass
        if majikiri_tops:
            ceiling_1f = float(np.median(majikiri_tops))
            print(f"  ceiling_1f from majikiri wall tops (median): {ceiling_1f:.3f}")
        else:
            print(f"  ceiling_1f: None (平屋/no 2-yuka, no majikiri)")

    # --- 2F天井: ログ壁top_y分布のギャップ解析 ---
    ceiling_2f = None
    if roof_params:
        eave_height = roof_params["eave_height"]
        ridge_axis = roof_params["ridge_axis"]
        ridge_height = roof_params["ridge_height"]

        log_tops = []
        for wall in ifc.by_type("IfcWall"):
            ename = wall.Name or ""
            if not ename.startswith("log"):
                continue
            try:
                shape = ifcopenshell.geom.create_shape(settings, wall)
                vf = shape.geometry.verts
                pts_3js = []
                for i in range(0, len(vf), 3):
                    pts_3js.append([vf[i], vf[i+2], -vf[i+1]])
                pts_3js = np.array(pts_3js)
                top_y = float(pts_3js[:, 1].max())
                log_tops.append(top_y)
            except Exception:
                pass

        if log_tops:
            # ギャップ解析: 壁top_yの分布でeave_heightに最も近いギャップを探す
            sorted_tops = sorted(set(round(t, 2) for t in log_tops))
            print(f"  Log wall top_y values: {sorted_tops}")

            best_gap = None
            best_gap_dist = float('inf')
            for i in range(len(sorted_tops) - 1):
                gap = sorted_tops[i+1] - sorted_tops[i]
                gap_mid = (sorted_tops[i] + sorted_tops[i+1]) / 2
                dist_to_eave = abs(gap_mid - eave_height)
                if gap > 0.3 and dist_to_eave < best_gap_dist:
                    best_gap = i
                    best_gap_dist = dist_to_eave

            if best_gap is not None:
                ceiling_2f = sorted_tops[best_gap + 1]
                print(f"  ceiling_2f from gap analysis: {ceiling_2f:.3f}"
                      f" (gap {sorted_tops[best_gap]:.2f}→{sorted_tops[best_gap+1]:.2f})")
            else:
                # フォールバック: eave_height + 0.5
                ceiling_2f = eave_height + 0.5
                print(f"  ceiling_2f fallback (eave_height + 0.5): {ceiling_2f:.3f}")

    has_2yuka = any((slab.Name or "").lower() == "2-yuka" for slab in ifc.by_type("IfcSlab"))

    return {
        "ceiling_1f": ceiling_1f,
        "ceiling_2f": ceiling_2f,
        "has_2yuka": has_2yuka,
    }


def roof_ceiling_y(slope_coord):
    """屋根勾配天井のY座標を返す（切妻屋根、汎用軸対応）
    slope_coord: SLOPE_AXIS上の座標（ie4d1ならX、HI-4AならZ）
    """
    if slope_coord <= RIDGE_POS:
        if abs(RIDGE_POS - EAVE_POS_MIN) < 0.01:
            return RIDGE_HEIGHT
        return EAVE_HEIGHT + (RIDGE_HEIGHT - EAVE_HEIGHT) * \
            (slope_coord - EAVE_POS_MIN) / (RIDGE_POS - EAVE_POS_MIN)
    else:
        if abs(EAVE_POS_MAX - RIDGE_POS) < 0.01:
            return RIDGE_HEIGHT
        return EAVE_HEIGHT + (RIDGE_HEIGHT - EAVE_HEIGHT) * \
            (EAVE_POS_MAX - slope_coord) / (EAVE_POS_MAX - RIDGE_POS)


def get_slope_coord_from_face(face_coord, direction):
    """壁面のface_coordからslope_axis上の座標を取得。
    eave壁(direction==RIDGE_AXIS)のface_coordはslope_axis上の値。
    gable壁(direction==SLOPE_AXIS)のface_coordはridge_axis上の値。
    """
    if direction == RIDGE_AXIS:
        # eave壁: face_coord = slope_axis座標
        return face_coord
    else:
        # gable壁: face_coordはridge_axis座標、slope_coordは壁の範囲で変わる
        return None  # 個別に処理


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


def extract_wall_edges(verts_flat):
    """壁ジオメトリから底辺情報を抽出"""
    pts = np.array(verts_flat).reshape(-1, 3)
    if len(pts) == 0:
        return None
    min_y = pts[:, 1].min()
    max_y = pts[:, 1].max()
    tol = 0.005
    bottom = pts[pts[:, 1] < (min_y + tol)]
    if len(bottom) < 2:
        return None
    dx = bottom[:, 0].max() - bottom[:, 0].min()
    dz = bottom[:, 2].max() - bottom[:, 2].min()
    if dx > dz:
        direction = "x"
        x_min, x_max = float(bottom[:, 0].min()), float(bottom[:, 0].max())
        full_x_min, full_x_max = float(pts[:, 0].min()), float(pts[:, 0].max())
        z_vals = np.unique(np.round(bottom[:, 2], 2))
        if len(z_vals) < 2:
            z_c = float(z_vals[0])
            return {
                "edges": [[[x_min, min_y, z_c], [x_max, min_y, z_c]]],
                "floor_y": float(min_y), "top_y": float(max_y),
                "direction": direction, "length_m": float(dx),
                "face_coords": [z_c],
                "full_x_range": (full_x_min, full_x_max),
            }
        z_f1, z_f2 = float(z_vals.min()), float(z_vals.max())
        return {
            "edges": [
                [[x_min, min_y, z_f1], [x_max, min_y, z_f1]],
                [[x_min, min_y, z_f2], [x_max, min_y, z_f2]],
            ],
            "floor_y": float(min_y), "top_y": float(max_y),
            "direction": direction, "length_m": float(dx),
            "face_coords": [z_f1, z_f2],
            "full_x_range": (full_x_min, full_x_max),
        }
    else:
        direction = "z"
        z_min, z_max = float(bottom[:, 2].min()), float(bottom[:, 2].max())
        full_z_min, full_z_max = float(pts[:, 2].min()), float(pts[:, 2].max())
        x_vals = np.unique(np.round(bottom[:, 0], 2))
        if len(x_vals) < 2:
            x_c = float(x_vals[0])
            return {
                "edges": [[[x_c, min_y, z_min], [x_c, min_y, z_max]]],
                "floor_y": float(min_y), "top_y": float(max_y),
                "direction": direction, "length_m": float(dz),
                "face_coords": [x_c],
                "full_z_range": (full_z_min, full_z_max),
            }
        x_f1, x_f2 = float(x_vals.min()), float(x_vals.max())
        return {
            "edges": [
                [[x_f1, min_y, z_min], [x_f1, min_y, z_max]],
                [[x_f2, min_y, z_min], [x_f2, min_y, z_max]],
            ],
            "floor_y": float(min_y), "top_y": float(max_y),
            "direction": direction, "length_m": float(dz),
            "face_coords": [x_f1, x_f2],
            "full_z_range": (full_z_min, full_z_max),
        }


def wall_cross_section_at_y(vf_raw, ff_raw, y_level, tol=0.01):
    """壁メッシュをY=y_levelで切断し、XZ平面上の断面線分を返す"""
    pts = []
    for i in range(0, len(vf_raw), 3):
        pts.append([vf_raw[i], vf_raw[i + 2], -vf_raw[i + 1]])
    pts = np.array(pts)
    segments = []
    for i in range(0, len(ff_raw), 3):
        i0, i1, i2 = ff_raw[i], ff_raw[i + 1], ff_raw[i + 2]
        y0, y1, y2 = pts[i0, 1], pts[i1, 1], pts[i2, 1]
        above = [y0 > y_level + tol, y1 > y_level + tol, y2 > y_level + tol]
        below = [y0 < y_level - tol, y1 < y_level - tol, y2 < y_level - tol]
        if all(above) or all(below):
            continue
        on_plane = [abs(y0 - y_level) <= tol, abs(y1 - y_level) <= tol, abs(y2 - y_level) <= tol]
        tri_pts = [pts[i0], pts[i1], pts[i2]]
        tri_ys = [y0, y1, y2]
        cross_pts = []
        edges_list = [(0, 1), (1, 2), (2, 0)]
        for a, b in edges_list:
            ya, yb = tri_ys[a], tri_ys[b]
            if on_plane[a]:
                cross_pts.append(tri_pts[a][[0, 2]])
            if (ya - y_level) * (yb - y_level) < 0:
                t = (y_level - ya) / (yb - ya)
                ix = tri_pts[a][0] + t * (tri_pts[b][0] - tri_pts[a][0])
                iz = tri_pts[a][2] + t * (tri_pts[b][2] - tri_pts[a][2])
                cross_pts.append(np.array([ix, iz]))
        if len(cross_pts) >= 2:
            unique = [cross_pts[0]]
            for p in cross_pts[1:]:
                if all(np.linalg.norm(p - u) > 0.001 for u in unique):
                    unique.append(p)
            if len(unique) >= 2:
                segments.append((unique[0], unique[1]))
    return segments


def extract_molding_from_cross_section(vf_raw, ff_raw, y_level, direction):
    """壁メッシュのY=y_levelでの断面から廻り縁ラインセグメントを抽出。"""
    cs_segs = wall_cross_section_at_y(vf_raw, ff_raw, y_level)
    face_segments = defaultdict(list)
    for p1, p2 in cs_segs:
        if direction == "x":
            dx = abs(p2[0] - p1[0])
            dz = abs(p2[1] - p1[1])
            if dx > 0.02 and dz < 0.005:
                z_face = round((p1[1] + p2[1]) / 2, 3)
                x_min = min(p1[0], p2[0])
                x_max = max(p1[0], p2[0])
                face_segments[z_face].append((x_min, x_max))
        else:
            dx = abs(p2[0] - p1[0])
            dz = abs(p2[1] - p1[1])
            if dz > 0.02 and dx < 0.005:
                x_face = round((p1[0] + p2[0]) / 2, 3)
                z_min = min(p1[1], p2[1])
                z_max = max(p1[1], p2[1])
                face_segments[x_face].append((z_min, z_max))
    result = {}
    for face_coord, segs in face_segments.items():
        segs.sort()
        merged = []
        for s, e in segs:
            if merged and s <= merged[-1][1] + 0.02:
                merged[-1] = (merged[-1][0], max(merged[-1][1], e))
            else:
                merged.append((s, e))
        edges_3d = []
        for s, e in merged:
            seg_len = e - s
            if seg_len < 0.05:
                continue
            if direction == "x":
                edges_3d.append(([s, y_level, face_coord], [e, y_level, face_coord]))
            else:
                edges_3d.append(([face_coord, y_level, s], [face_coord, y_level, e]))
        if edges_3d:
            result[face_coord] = edges_3d
    return result


def classify_exterior_with_side(walls_list):
    """外周/室内を判定し、外周壁にはどちら側が外側かも返す"""
    log_walls = [(c, g, n, inf) for c, g, n, inf in walls_list if c == "ログ壁"]
    result = {}
    for direction in ["x", "z"]:
        dir_walls = [(c, g, n, inf) for c, g, n, inf in log_walls
                     if inf["direction"] == direction]
        if not dir_walls:
            continue
        perp_groups = defaultdict(list)
        for c, g, n, inf in dir_walls:
            fc = inf["face_coords"]
            perp_coord = round(np.mean(fc), 1)
            matched = False
            for existing in list(perp_groups.keys()):
                if abs(perp_coord - existing) < 0.3:
                    perp_groups[existing].append((g, inf, perp_coord))
                    matched = True
                    break
            if not matched:
                perp_groups[perp_coord].append((g, inf, perp_coord))
        sorted_coords = sorted(perp_groups.keys())
        if len(sorted_coords) >= 2:
            min_c, max_c = sorted_coords[0], sorted_coords[-1]
            for coord in sorted_coords:
                is_ext = (coord == min_c or coord == max_c)
                ext_side = None
                if is_ext:
                    ext_side = "min" if coord == min_c else "max"
                for g, inf, _ in perp_groups[coord]:
                    result[g] = {"is_ext": is_ext, "ext_side": ext_side}
        else:
            for coord in sorted_coords:
                for g, inf, _ in perp_groups[coord]:
                    result[g] = {"is_ext": True, "ext_side": "min"}
    for c, g, n, inf in walls_list:
        if g not in result:
            result[g] = {"is_ext": False, "ext_side": None}
    return result


def point_in_triangle_2d(px, pz, ax, az, bx, bz, cx, cz):
    """XZ平面上の点(px,pz)が三角形(a,b,c)内にあるか"""
    d = (bz - cz) * (ax - cx) + (cx - bx) * (az - cz)
    if abs(d) < 1e-10:
        return False
    a = ((bz - cz) * (px - cx) + (cx - bx) * (pz - cz)) / d
    b = ((cz - az) * (px - cx) + (ax - cx) * (pz - cz)) / d
    c = 1 - a - b
    tol = 0.02
    return a >= -tol and b >= -tol and c >= -tol


def point_on_slab(px, pz, slab_tris, margin=0):
    """点(px,pz)がスラブ三角形メッシュ上にあるか"""
    for offx, offz in [(0, 0), (margin, 0), (-margin, 0), (0, margin), (0, -margin)]:
        for tri in slab_tris:
            if point_in_triangle_2d(px + offx, pz + offz,
                                     tri[0][0], tri[0][1],
                                     tri[1][0], tri[1][1],
                                     tri[2][0], tri[2][1]):
                return True
    return False


def extract_slab_top_tris(slab, settings):
    """スラブ上面の三角形をXZ平面座標で返す"""
    shape = ifcopenshell.geom.create_shape(settings, slab)
    vf = shape.geometry.verts
    ff = shape.geometry.faces
    pts = []
    for i in range(0, len(vf), 3):
        pts.append([vf[i], vf[i + 2], -vf[i + 1]])
    pts = np.array(pts)
    top_y = pts[:, 1].max()
    tol = 0.01
    tris_xz = []
    for i in range(0, len(ff), 3):
        i0, i1, i2 = ff[i], ff[i + 1], ff[i + 2]
        if (abs(pts[i0, 1] - top_y) < tol and
                abs(pts[i1, 1] - top_y) < tol and
                abs(pts[i2, 1] - top_y) < tol):
            tris_xz.append([
                [float(pts[i0, 0]), float(pts[i0, 2])],
                [float(pts[i1, 0]), float(pts[i1, 2])],
                [float(pts[i2, 0]), float(pts[i2, 2])],
            ])
    return tris_xz


def extract_slab_bottom_tris(slab, settings):
    """スラブ下面の三角形をXZ平面座標で返す（天井面）"""
    shape = ifcopenshell.geom.create_shape(settings, slab)
    vf = shape.geometry.verts
    ff = shape.geometry.faces
    pts = []
    for i in range(0, len(vf), 3):
        pts.append([vf[i], vf[i + 2], -vf[i + 1]])
    pts = np.array(pts)
    bottom_y = pts[:, 1].min()
    tol = 0.01
    tris_xz = []
    for i in range(0, len(ff), 3):
        i0, i1, i2 = ff[i], ff[i + 1], ff[i + 2]
        if (abs(pts[i0, 1] - bottom_y) < tol and
                abs(pts[i1, 1] - bottom_y) < tol and
                abs(pts[i2, 1] - bottom_y) < tol):
            tris_xz.append([
                [float(pts[i0, 0]), float(pts[i0, 2])],
                [float(pts[i1, 0]), float(pts[i1, 2])],
                [float(pts[i2, 0]), float(pts[i2, 2])],
            ])
    return tris_xz


def clip_segment_to_slab(p1, p2, slab_tris, direction, n_samples=20):
    """セグメントをスラブ範囲でクリップ"""
    if not slab_tris:
        return [(p1, p2)]
    coord_idx = 0 if direction == "x" else 2
    c1, c2 = p1[coord_idx], p2[coord_idx]
    if c1 > c2:
        p1, p2 = p2, p1
        c1, c2 = c2, c1
    seg_len = c2 - c1
    if seg_len < 0.01:
        mx = (p1[0] + p2[0]) / 2
        mz = (p1[2] + p2[2]) / 2
        if point_on_slab(mx, mz, slab_tris, margin=0):
            return [(p1, p2)]
        return []
    on_slab = []
    for i in range(n_samples + 1):
        t = i / n_samples
        sx = p1[0] + t * (p2[0] - p1[0])
        sz = p1[2] + t * (p2[2] - p1[2])
        on_slab.append(point_on_slab(sx, sz, slab_tris, margin=0))
    results = []
    start = None
    for i, on in enumerate(on_slab):
        if on and start is None:
            start = i
        elif not on and start is not None:
            t0 = start / n_samples
            t1 = (i - 1) / n_samples
            sp1 = [p1[j] + t0 * (p2[j] - p1[j]) for j in range(3)]
            sp2 = [p1[j] + t1 * (p2[j] - p1[j]) for j in range(3)]
            results.append((sp1, sp2))
            start = None
    if start is not None:
        t0 = start / n_samples
        sp1 = [p1[j] + t0 * (p2[j] - p1[j]) for j in range(3)]
        results.append((sp1, list(p2)))
    return results


def get_wall_full_range(info, axis):
    """壁のfull_x_range or full_z_rangeを取得（フォールバック付き）"""
    if axis == "x":
        if "full_x_range" in info:
            return info["full_x_range"]
        return (min(e[0] for edge in info["edges"] for e in edge),
                max(e[0] for edge in info["edges"] for e in edge))
    else:
        if "full_z_range" in info:
            return info["full_z_range"]
        return (min(e[2] for edge in info["edges"] for e in edge),
                max(e[2] for edge in info["edges"] for e in edge))


def get_wall_range_along_axis(info, axis):
    """壁の指定軸方向の範囲を取得"""
    if axis == "x":
        return get_wall_full_range(info, "x")
    else:
        return get_wall_full_range(info, "z")


def make_segment_along_direction(direction, face_coord, range_min, range_max, y_val):
    """壁方向に沿ったセグメントを生成"""
    if direction == "x":
        return [range_min, y_val, face_coord], [range_max, y_val, face_coord]
    else:
        return [face_coord, y_val, range_min], [face_coord, y_val, range_max]


def main():
    global RIDGE_AXIS, SLOPE_AXIS, RIDGE_POS, RIDGE_HEIGHT
    global EAVE_POS_MIN, EAVE_POS_MAX, EAVE_HEIGHT

    print("IFC読み込み中...")
    ifc = ifcopenshell.open(IFC_PATH)
    settings = ifcopenshell.geom.settings()
    settings.set(settings.USE_WORLD_COORDS, True)

    # === 屋根パラメータ自動検出 ===
    print("\n--- 屋根パラメータ自動検出 ---")
    roof_params = detect_roof_params(ifc, settings)
    if roof_params:
        RIDGE_AXIS = roof_params["ridge_axis"]
        SLOPE_AXIS = roof_params["slope_axis"]
        RIDGE_POS = roof_params["ridge_pos"]
        RIDGE_HEIGHT = roof_params["ridge_height"]
        EAVE_POS_MIN = roof_params["eave_pos_min"]
        EAVE_POS_MAX = roof_params["eave_pos_max"]
        EAVE_HEIGHT = roof_params["eave_height"]
    else:
        print("  WARNING: 屋根パラメータ検出失敗、デフォルト値を使用")

    # === 天井レベル自動検出 ===
    print("\n--- 天井レベル自動検出 ---")
    ceiling_info = detect_ceiling_levels(ifc, settings, roof_params)
    ceiling_1f = ceiling_info["ceiling_1f"]
    ceiling_2f = ceiling_info["ceiling_2f"]
    has_2yuka = ceiling_info["has_2yuka"]

    # 平屋判定: 2-yukaスラブがない
    is_single_story = not has_2yuka
    if is_single_story:
        print(f"\n  *** 平屋建物検出（2-yukaスラブなし）***")
        print(f"  → 1F水平天井セクションをスキップ、全壁を2F/屋根として処理")

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

    # === Part 2: 壁ジオメトリ解析 ===
    print("\n壁ジオメトリ解析中...")
    seen2 = set()
    wall_data = []

    for wall in ifc.by_type("IfcWall"):
        gid = wall.GlobalId
        if gid in seen2:
            continue
        seen2.add(gid)
        ename = wall.Name or ""
        if any(ename.startswith(p) for p in SKIP):
            continue
        cat = None
        for k, v in WALL_SUB.items():
            if ename.startswith(k):
                cat = v
                break
        if not cat:
            continue
        try:
            shape = ifcopenshell.geom.create_shape(settings, wall)
            vf = shape.geometry.verts
            ff = shape.geometry.faces
            verts_3js = []
            for i in range(0, len(vf), 3):
                verts_3js.extend([vf[i], vf[i + 2], -vf[i + 1]])
            info = extract_wall_edges(verts_3js)
            if info:
                wall_data.append((cat, gid, ename, info, list(vf), list(ff)))
        except Exception:
            pass

    # === 廻り縁ライン生成 ===
    molding_lines = []
    total_by_floor = defaultdict(float)
    total_by_type = defaultdict(float)

    def slope_side_at(fc_slope, face_coords_all_slope):
        """eave壁（RIDGE_AXIS方向壁）の面が勾配の上側か下側かを判定。
        fc_slope: 面のslope_axis座標
        face_coords_all_slope: 壁全面のslope_axis座標リスト
        """
        if len(face_coords_all_slope) < 2:
            return "上側"
        avg = sum(face_coords_all_slope) / len(face_coords_all_slope)
        dist = abs(fc_slope - RIDGE_POS)
        avg_dist = abs(avg - RIDGE_POS)
        if avg_dist < 0.3:
            return "上側"  # 棟上の壁
        if avg < RIDGE_POS:
            return "上側" if fc_slope >= avg else "下側"
        else:
            return "上側" if fc_slope <= avg else "下側"

    # =================================================================
    # 1F天井: 壁フルエクステント + 2Fスラブクリッピング
    # =================================================================
    if not is_single_story and ceiling_1f is not None:
        # 1F対象壁: 天井高さまで達する壁
        walls_1f_raw = [(c, g, n, inf, vf, ff) for c, g, n, inf, vf, ff in wall_data
                        if inf["top_y"] > ceiling_1f - 0.1]
        walls_1f = [(c, g, n, inf) for c, g, n, inf, vf, ff in walls_1f_raw]

        # スラブメッシュ取得（1F天井クリッピング用）
        print("  スラブメッシュを取得中...")
        slab_tris_1f_ceiling = []
        for slab in ifc.by_type("IfcSlab"):
            ename = slab.Name or ""
            try:
                if ename == "2-yuka":
                    tris_bottom = extract_slab_bottom_tris(slab, settings)
                    slab_tris_1f_ceiling.extend(tris_bottom)
                    print(f"    {ename} 下面: {len(tris_bottom)}三角形 → 1F天井クリッピング")
            except Exception:
                pass

        cs_y_1f = ceiling_1f - 0.02
        print(f"\n  1F天井 (Y={ceiling_1f:.3f}m): {len(walls_1f_raw)}壁 [全面・フルエクステント]")

        for cat, gid, ename, info, vf_raw, ff_raw in walls_1f_raw:
            direction = info["direction"]
            face_coords = info["face_coords"]
            wall_seg_total = 0

            for fc in face_coords:
                # セグメント生成: 壁方向に沿ったフルエクステント
                wall_range = get_wall_full_range(info, direction)
                p1, p2 = make_segment_along_direction(
                    direction, fc, wall_range[0], wall_range[1], cs_y_1f)

                clipped = clip_segment_to_slab(p1, p2, slab_tris_1f_ceiling, direction)

                for sp1, sp2 in clipped:
                    seg_len = np.sqrt(sum((a - b) ** 2 for a, b in zip(sp1, sp2)))
                    if seg_len < 0.05:
                        continue
                    mt = "廻り縁１" if cat == "ログ壁" else "廻り縁２"
                    wall_seg_total += seg_len
                    total_by_type[mt] += seg_len
                    molding_lines.append({
                        "floor": "1F", "type": "wall", "cat": cat,
                        "molding_type": mt,
                        "seg": [[round(sp1[0], 4), round(sp1[1], 4), round(sp1[2], 4)],
                                [round(sp2[0], 4), round(sp2[1], 4), round(sp2[2], 4)]],
                        "length_m": round(seg_len, 4),
                    })

            total_by_floor["1F"] += wall_seg_total
            print(f"    {cat:8s} {ename:15s} {len(face_coords)}面 → {wall_seg_total:.2f}m")
    else:
        walls_1f_raw = []
        slab_tris_1f_ceiling = []
        if is_single_story:
            print(f"\n  1F天井: スキップ（平屋建物）")
        else:
            print(f"\n  1F天井: スキップ（ceiling_1f未検出）")

    # =================================================================
    # 2F天井（勾配天井）: 全面・スラブクリップなし
    # =================================================================
    if ceiling_2f is not None:
        # 2F対象壁: 桁レベルまで達するログ壁
        walls_2f_raw = []
        seen_2f = set()
        for wall in ifc.by_type("IfcWall"):
            gid = wall.GlobalId
            if gid in seen_2f:
                continue
            seen_2f.add(gid)
            ename = wall.Name or ""
            if any(ename.startswith(p) for p in SKIP):
                continue
            cat = None
            for k, v in WALL_SUB.items():
                if ename.startswith(k):
                    cat = v
                    break
            if cat != "ログ壁":
                continue
            try:
                shape = ifcopenshell.geom.create_shape(settings, wall)
                vf = shape.geometry.verts
                ff = shape.geometry.faces
                verts_3js = []
                for i in range(0, len(vf), 3):
                    verts_3js.extend([vf[i], vf[i + 2], -vf[i + 1]])
                info = extract_wall_edges(verts_3js)
                if info and info["top_y"] > ceiling_2f - 0.1:
                    walls_2f_raw.append((cat, gid, ename, info, list(vf), list(ff)))
            except Exception:
                pass

        walls_2f = [(c, g, n, inf) for c, g, n, inf, vf, ff in walls_2f_raw]

        print(f"\n  2F天井（勾配天井）: {len(walls_2f_raw)}壁 [全面・スラブクリップなし]")
        print(f"    棟軸={RIDGE_AXIS.upper()} 勾配軸={SLOPE_AXIS.upper()}"
              f" 棟位置({SLOPE_AXIS.upper()})={RIDGE_POS:.3f}"
              f" 棟高さ={RIDGE_HEIGHT:.3f} 軒高さ={EAVE_HEIGHT:.3f}")

        for cat, gid, ename, info, vf_raw, ff_raw in walls_2f_raw:
            direction = info["direction"]
            face_coords = info["face_coords"]
            wall_seg_total = 0

            if direction == RIDGE_AXIS:
                # === eave壁（棟と平行）: 水平ライン at roof_ceiling_y(slope_coord) ===
                for fc in face_coords:
                    # fc = slope_axis座標（eave壁のface_coordはslope_axis上の値）
                    ceil_y = roof_ceiling_y(fc)
                    cut_y = min(ceil_y, info["top_y"]) - 0.02
                    if cut_y < EAVE_HEIGHT - 0.2:
                        continue

                    # 壁の長さ: ridge_axis方向のフル範囲
                    wall_range = get_wall_full_range(info, direction)
                    seg_len = abs(wall_range[1] - wall_range[0])
                    if seg_len < 0.05:
                        continue

                    side = slope_side_at(fc, face_coords)
                    mt = "廻り縁２" if side == "上側" else "廻り縁３"
                    wall_seg_total += seg_len
                    total_by_type[mt] += seg_len

                    p1, p2 = make_segment_along_direction(
                        direction, fc, wall_range[0], wall_range[1], cut_y)
                    molding_lines.append({
                        "floor": "2F", "type": "wall", "cat": cat,
                        "molding_type": mt, "slope_side": side,
                        "seg": [[round(p1[0], 4), round(p1[1], 4), round(p1[2], 4)],
                                [round(p2[0], 4), round(p2[1], 4), round(p2[2], 4)]],
                        "length_m": round(seg_len, 4),
                    })

                dir_label = f"{direction.upper()}方向(水平/eave)"
                print(f"    {cat:8s} {ename:15s} {dir_label} {len(face_coords)}面 → {wall_seg_total:.2f}m")

            elif direction == SLOPE_AXIS:
                # === gable壁（棟と垂直）: 勾配天井に沿った斜めライン ===
                # 壁のslope_axis方向範囲
                wall_range = get_wall_full_range(info, direction)
                slope_min, slope_max = wall_range

                for fc in face_coords:
                    # fc = ridge_axis座標（gable壁のface_coordはridge_axis上の値）
                    # slope_axis方向に勾配ラインを生成
                    slope_segments = []
                    if slope_min < RIDGE_POS and slope_max > RIDGE_POS:
                        slope_segments.append((slope_min, RIDGE_POS))
                        slope_segments.append((RIDGE_POS, slope_max))
                    else:
                        slope_segments.append((slope_min, slope_max))

                    for ss, se in slope_segments:
                        y_s = min(roof_ceiling_y(ss), info["top_y"]) - 0.02
                        y_e = min(roof_ceiling_y(se), info["top_y"]) - 0.02
                        if y_s < EAVE_HEIGHT - 0.2 and y_e < EAVE_HEIGHT - 0.2:
                            continue
                        seg_len = np.sqrt((se - ss)**2 + (y_e - y_s)**2)
                        if seg_len < 0.05:
                            continue
                        mt = "廻り縁１"
                        wall_seg_total += seg_len
                        total_by_type[mt] += seg_len

                        # セグメント: slope_axis方向 + Y変化
                        if SLOPE_AXIS == "x":
                            seg = [[round(ss, 4), round(y_s, 4), round(fc, 4)],
                                   [round(se, 4), round(y_e, 4), round(fc, 4)]]
                        else:
                            seg = [[round(fc, 4), round(y_s, 4), round(ss, 4)],
                                   [round(fc, 4), round(y_e, 4), round(se, 4)]]

                        molding_lines.append({
                            "floor": "2F", "type": "wall", "cat": cat,
                            "molding_type": mt,
                            "seg": seg,
                            "length_m": round(seg_len, 4),
                        })

                dir_label = f"{direction.upper()}方向(勾配/gable)"
                print(f"    {cat:8s} {ename:15s} {dir_label} {len(face_coords)}面 → {wall_seg_total:.2f}m")

            total_by_floor["2F"] += wall_seg_total

        # --- 2F桁レベル壁の廻り縁（軒側外壁: 桁で止まるeave方向壁） ---
        keta_walls_raw = []
        for c, g, n, inf, vf, ff in wall_data:
            if c != "ログ壁":
                continue
            if inf["direction"] != RIDGE_AXIS:
                continue
            if inf["top_y"] < ceiling_2f - 0.1 and inf["top_y"] > EAVE_HEIGHT + 0.1:
                keta_walls_raw.append((c, g, n, inf, vf, ff))

        if keta_walls_raw:
            print(f"\n  2F桁レベル壁（軒側）: {len(keta_walls_raw)}壁")
            for c, g, n, inf, vf, ff in keta_walls_raw:
                face_coords = inf["face_coords"]
                wall_seg_total = 0

                for fc in face_coords:
                    # fc = slope_axis座標
                    ceil_y = roof_ceiling_y(fc)
                    if ceil_y <= EAVE_HEIGHT + 0.01:
                        continue
                    cut_y = min(ceil_y, inf["top_y"]) - 0.02
                    if cut_y < 2.0:
                        continue
                    wall_range = get_wall_full_range(inf, RIDGE_AXIS)
                    seg_len = abs(wall_range[1] - wall_range[0])
                    if seg_len < 0.05:
                        continue
                    side = slope_side_at(fc, face_coords)
                    mt = "廻り縁２" if side == "上側" else "廻り縁３"
                    wall_seg_total += seg_len
                    total_by_type[mt] += seg_len

                    p1, p2 = make_segment_along_direction(
                        RIDGE_AXIS, fc, wall_range[0], wall_range[1], cut_y)
                    molding_lines.append({
                        "floor": "2F", "type": "wall", "cat": c,
                        "molding_type": mt, "slope_side": side,
                        "seg": [[round(p1[0], 4), round(p1[1], 4), round(p1[2], 4)],
                                [round(p2[0], 4), round(p2[1], 4), round(p2[2], 4)]],
                        "length_m": round(seg_len, 4),
                    })

                if wall_seg_total > 0:
                    total_by_floor["2F"] += wall_seg_total
                    print(f"    {c:8s} {n:15s} {RIDGE_AXIS.upper()}方向(桁)"
                          f" {len(face_coords)}面 → {wall_seg_total:.2f}m")
    else:
        print(f"\n  2F天井: スキップ（ceiling_2f未検出）")

    # --- 集成梁の廻り縁 ---
    print(f"\n  集成梁の廻り縁:")
    for beam in ifc.by_type("IfcBeam"):
        bname = beam.Name or ""
        try:
            shape = ifcopenshell.geom.create_shape(settings, beam)
            vf = shape.geometry.verts
            pts = []
            for i in range(0, len(vf), 3):
                pts.append([vf[i], vf[i + 2], -vf[i + 1]])
            pts = np.array(pts)

            min_y = pts[:, 1].min()
            max_y = pts[:, 1].max()
            dx = pts[:, 0].max() - pts[:, 0].min()
            dz = pts[:, 2].max() - pts[:, 2].min()
            beam_dir = "x" if dx > dz else "z"
            beam_seg_total = 0

            # (A) 屋根レベルの梁（棟木 L300 など）
            roof_beam_threshold = EAVE_HEIGHT + 0.2 if EAVE_HEIGHT else 3.5
            if max_y > roof_beam_threshold:
                if beam_dir == RIDGE_AXIS:
                    # 梁がridge_axis方向（棟木など）: 両側面に水平ライン
                    ridge_range = (float(pts[:, 0].min()), float(pts[:, 0].max())) \
                        if RIDGE_AXIS == "x" else (float(pts[:, 2].min()), float(pts[:, 2].max()))
                    # 両side面のslope_axis座標
                    if SLOPE_AXIS == "x":
                        face_min = float(np.round(pts[:, 0].min(), 3))
                        face_max = float(np.round(pts[:, 0].max(), 3))
                    else:
                        face_min = float(np.round(pts[:, 2].min(), 3))
                        face_max = float(np.round(pts[:, 2].max(), 3))

                    for fc_slope in [face_min, face_max]:
                        ceil_y = roof_ceiling_y(fc_slope) - 0.02
                        seg_len = abs(ridge_range[1] - ridge_range[0])
                        if seg_len < 0.05:
                            continue
                        mt = "廻り縁２"
                        beam_seg_total += seg_len
                        total_by_type[mt] += seg_len

                        p1, p2 = make_segment_along_direction(
                            RIDGE_AXIS, fc_slope, ridge_range[0], ridge_range[1], ceil_y)
                        molding_lines.append({
                            "floor": "2F", "type": "beam", "cat": "梁",
                            "molding_type": mt, "beam_name": bname,
                            "seg": [[round(p1[0], 4), round(p1[1], 4), round(p1[2], 4)],
                                    [round(p2[0], 4), round(p2[1], 4), round(p2[2], 4)]],
                            "length_m": round(seg_len, 4),
                        })

                    print(f"    [屋根] 梁 {bname:15s} {RIDGE_AXIS.upper()}方向"
                          f" 両面 → {beam_seg_total:.2f}m")

                else:
                    # 梁がslope_axis方向: 両側面に勾配ライン
                    slope_range = (float(pts[:, 0].min()), float(pts[:, 0].max())) \
                        if SLOPE_AXIS == "x" else (float(pts[:, 2].min()), float(pts[:, 2].max()))
                    # 両side面のridge_axis座標
                    if RIDGE_AXIS == "x":
                        face_min = float(np.round(pts[:, 0].min(), 3))
                        face_max = float(np.round(pts[:, 0].max(), 3))
                    else:
                        face_min = float(np.round(pts[:, 2].min(), 3))
                        face_max = float(np.round(pts[:, 2].max(), 3))

                    for fc_ridge in [face_min, face_max]:
                        slope_segs = []
                        if slope_range[0] < RIDGE_POS and slope_range[1] > RIDGE_POS:
                            slope_segs.append((slope_range[0], RIDGE_POS))
                            slope_segs.append((RIDGE_POS, slope_range[1]))
                        else:
                            slope_segs.append((slope_range[0], slope_range[1]))

                        for ss, se in slope_segs:
                            y_s = roof_ceiling_y(ss) - 0.02
                            y_e = roof_ceiling_y(se) - 0.02
                            seg_len = np.sqrt((se - ss)**2 + (y_e - y_s)**2)
                            if seg_len < 0.05:
                                continue
                            mt = "廻り縁２"
                            beam_seg_total += seg_len
                            total_by_type[mt] += seg_len

                            if SLOPE_AXIS == "x":
                                seg = [[round(ss, 4), round(y_s, 4), round(fc_ridge, 4)],
                                       [round(se, 4), round(y_e, 4), round(fc_ridge, 4)]]
                            else:
                                seg = [[round(fc_ridge, 4), round(y_s, 4), round(ss, 4)],
                                       [round(fc_ridge, 4), round(y_e, 4), round(se, 4)]]

                            molding_lines.append({
                                "floor": "2F", "type": "beam", "cat": "梁",
                                "molding_type": mt, "beam_name": bname,
                                "seg": seg,
                                "length_m": round(seg_len, 4),
                            })

                    print(f"    [屋根] 梁 {bname:15s} {SLOPE_AXIS.upper()}方向"
                          f" 両面 → {beam_seg_total:.2f}m")

                total_by_floor["2F"] += beam_seg_total

            # (B) 2F床梁（1F天井レベルを跨ぐ梁 = min_y < ceiling_1f < max_y）
            elif not is_single_story and ceiling_1f and min_y < ceiling_1f and max_y > ceiling_1f:
                beam_1f_total = 0
                cs_y_1f = ceiling_1f - 0.02

                if beam_dir == "z":
                    z_min_b = float(pts[:, 2].min())
                    z_max_b = float(pts[:, 2].max())
                    x_face_min = float(np.round(pts[:, 0].min(), 3))
                    x_face_max = float(np.round(pts[:, 0].max(), 3))
                    for x_face in [x_face_min, x_face_max]:
                        p1 = [x_face, cs_y_1f, z_min_b]
                        p2 = [x_face, cs_y_1f, z_max_b]
                        clipped = clip_segment_to_slab(p1, p2, slab_tris_1f_ceiling, "z")
                        for sp1, sp2 in clipped:
                            seg_len = np.sqrt(sum((a - b) ** 2 for a, b in zip(sp1, sp2)))
                            if seg_len < 0.05:
                                continue
                            mt = "廻り縁１"
                            beam_1f_total += seg_len
                            total_by_type[mt] += seg_len
                            molding_lines.append({
                                "floor": "1F", "type": "beam", "cat": "梁",
                                "molding_type": mt, "beam_name": bname,
                                "seg": [[round(sp1[0], 4), round(sp1[1], 4), round(sp1[2], 4)],
                                        [round(sp2[0], 4), round(sp2[1], 4), round(sp2[2], 4)]],
                                "length_m": round(seg_len, 4),
                            })
                    print(f"    [1F床] 梁 {bname:15s} Z方向 → {beam_1f_total:.2f}m")
                else:
                    x_min_b = float(pts[:, 0].min())
                    x_max_b = float(pts[:, 0].max())
                    z_face_min = float(np.round(pts[:, 2].min(), 3))
                    z_face_max = float(np.round(pts[:, 2].max(), 3))
                    for z_face in [z_face_min, z_face_max]:
                        p1 = [x_min_b, cs_y_1f, z_face]
                        p2 = [x_max_b, cs_y_1f, z_face]
                        clipped = clip_segment_to_slab(p1, p2, slab_tris_1f_ceiling, "x")
                        for sp1, sp2 in clipped:
                            seg_len = np.sqrt(sum((a - b) ** 2 for a, b in zip(sp1, sp2)))
                            if seg_len < 0.05:
                                continue
                            mt = "廻り縁１"
                            beam_1f_total += seg_len
                            total_by_type[mt] += seg_len
                            molding_lines.append({
                                "floor": "1F", "type": "beam", "cat": "梁",
                                "molding_type": mt, "beam_name": bname,
                                "seg": [[round(sp1[0], 4), round(sp1[1], 4), round(sp1[2], 4)],
                                        [round(sp2[0], 4), round(sp2[1], 4), round(sp2[2], 4)]],
                                "length_m": round(seg_len, 4),
                            })
                    print(f"    [1F床] 梁 {bname:15s} X方向 → {beam_1f_total:.2f}m")

                total_by_floor["1F"] += beam_1f_total

        except Exception:
            pass

    # 集計表示
    grand_total = 0
    print("\n=== 廻り縁集計（階別） ===")
    for fn in ["1F", "2F"]:
        if fn not in total_by_floor:
            continue
        floor_total = total_by_floor[fn]
        grand_total += floor_total
        print(f"  {fn}: {floor_total:.2f}m")
    print(f"  合計: {grand_total:.2f}m")

    print("\n=== 廻り縁集計（部材種別） ===")
    for mt in ["廻り縁１", "廻り縁２", "廻り縁３"]:
        if mt in total_by_type:
            print(f"  {mt}: {total_by_type[mt]:.2f}m")
    print(f"  合計: {sum(total_by_type.values()):.2f}m")

    # 詳細内訳
    print("\n=== 詳細内訳 ===")
    detail = defaultdict(float)
    for m in molding_lines:
        mt = m.get("molding_type", "?")
        floor = m["floor"]
        cat = m.get("cat", "?")
        mtype = m.get("type", "?")
        side = m.get("slope_side", "")
        if mtype == "beam":
            key = f"{mt} | {floor} {cat}({m.get('beam_name','')})"
        elif floor == "2F" and side:
            key = f"{mt} | {floor} {cat} {side}"
        else:
            key = f"{mt} | {floor} {cat}"
        detail[key] += m["length_m"]
    for key in sorted(detail.keys()):
        print(f"  {key}: {detail[key]:.2f}m")

    # === HTML出力 ===
    # IFCファイル名を取得してタイトルに使用
    ifc_basename = os.path.splitext(os.path.basename(IFC_PATH))[0]

    meshes_json = json.dumps(meshes, ensure_ascii=False)
    molding_json = json.dumps(molding_lines, ensure_ascii=False)
    colors_json = json.dumps(CATEGORY_COLORS, ensure_ascii=False)

    totals_data = {}
    for fn in ["1F", "2F"]:
        if fn in total_by_floor:
            totals_data[fn] = round(total_by_floor[fn], 2)
    type_totals = {}
    for mt in ["廻り縁１", "廻り縁２", "廻り縁３"]:
        if mt in total_by_type:
            type_totals[mt] = round(total_by_type[mt], 2)

    totals_json = json.dumps(totals_data, ensure_ascii=False)
    type_totals_json = json.dumps(type_totals, ensure_ascii=False)
    html = generate_html(meshes_json, molding_json, colors_json,
                         totals_json, type_totals_json, ifc_basename)

    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"\n出力: {OUTPUT_HTML}")


def generate_html(meshes_json, molding_json, colors_json,
                  totals_json, type_totals_json, title="廻り縁拾い"):
    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>廻り縁拾い 3Dビューア - {title}</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ background:#1a1a2e; overflow:hidden; font-family:Arial,sans-serif; }}
#info {{ position:fixed; top:10px; left:10px; color:#fff; background:rgba(0,0,0,0.75);
  padding:12px 16px; border-radius:8px; font-size:13px; z-index:100; max-width:360px; }}
#info h3 {{ margin-bottom:6px; color:#e066ff; font-size:15px; }}
#legend {{ position:fixed; top:10px; right:10px; color:#fff; background:rgba(0,0,0,0.75);
  padding:12px; border-radius:8px; font-size:12px; z-index:100; max-height:80vh; overflow-y:auto; }}
#legend div {{ cursor:pointer; padding:3px 6px; border-radius:3px; margin:2px 0; white-space:nowrap; }}
#legend div:hover {{ background:rgba(255,255,255,0.15); }}
.cb {{ display:inline-block; width:14px; height:14px; border-radius:3px; margin-right:6px; vertical-align:middle; }}
#molding-info {{ position:fixed; bottom:10px; left:10px; color:#fff; background:rgba(0,0,0,0.85);
  padding:14px 18px; border-radius:8px; font-size:13px; z-index:100; line-height:1.6; }}
#controls {{ position:fixed; bottom:10px; right:10px; z-index:100; }}
#controls button {{ background:rgba(255,255,255,0.15); color:#fff; border:1px solid rgba(255,255,255,0.3);
  padding:8px 14px; border-radius:6px; cursor:pointer; margin:2px; font-size:12px; }}
#controls button:hover {{ background:rgba(255,255,255,0.3); }}
#controls button.active {{ background:rgba(180,100,255,0.5); border-color:#e066ff; }}
</style>
</head>
<body>
<div id="info">
  <h3>廻り縁拾い 3Dビューア - {title}</h3>
  <div>左ドラッグ: 回転 / 右ドラッグ: 移動 / ホイール: ズーム</div>
  <div id="sel-info" style="margin-top:6px;color:#aaa;">クリックで部材選択</div>
</div>
<div id="molding-info"></div>
<div id="legend"></div>
<div id="controls">
  <button id="btn-m1" class="active" onclick="toggleType('廻り縁１')">廻り縁１</button>
  <button id="btn-m2" class="active" onclick="toggleType('廻り縁２')">廻り縁２</button>
  <button id="btn-m3" class="active" onclick="toggleType('廻り縁３')">廻り縁３</button>
  <button id="btn-building" class="active" onclick="toggleBuilding()">建物表示</button>
  <button onclick="resetCam()">リセット</button>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
<script>
const MESHES={meshes_json};
const MOLDING={molding_json};
const COLORS={colors_json};
const TOTALS={totals_json};
const TYPE_TOTALS={type_totals_json};

const MOLDING_TYPE_COLORS = {{
  "廻り縁１": 0xff6644,
  "廻り縁２": 0x44ccff,
  "廻り縁３": 0x88ff44
}};
const MOLDING_TYPE_CSS = {{
  "廻り縁１": "#ff6644",
  "廻り縁２": "#44ccff",
  "廻り縁３": "#88ff44"
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

const moldingGroups = {{}};
["廻り縁１","廻り縁２","廻り縁３"].forEach(mt=>{{moldingGroups[mt]=new THREE.Group();scene.add(moldingGroups[mt]);}});

MOLDING.forEach(h=>{{
const mt=h.molding_type||"廻り縁１";
const color=MOLDING_TYPE_COLORS[mt]||0xffffff;
const pts=[new THREE.Vector3(...h.seg[0]),new THREE.Vector3(...h.seg[1])];
const g=new THREE.BufferGeometry().setFromPoints(pts);
const mat=new THREE.LineBasicMaterial({{color,linewidth:3}});
const grp=moldingGroups[mt];
if(grp) grp.add(new THREE.Line(g,mat));
}});

let infoHtml='<b style="font-size:14px;">廻り縁集計（部材種別）</b><br><br>';
let grandTotal=0;
for(const mt of ["廻り縁１","廻り縁２","廻り縁３"]){{
  const t=TYPE_TOTALS[mt]||0;
  const cc=MOLDING_TYPE_CSS[mt]||"#fff";
  grandTotal+=t;
  infoHtml+=`<span style="color:${{cc}}">━━ ${{mt}}: ${{t.toFixed(2)}}m</span><br>`;
}}
infoHtml+=`<br>`;
for(const [fn,t] of Object.entries(TOTALS)){{
  infoHtml+=`<span style="color:#aaa">${{fn}}天井: ${{t.toFixed(2)}}m</span><br>`;
}}
infoHtml+=`<br><b style="font-size:15px;">合計: ${{grandTotal.toFixed(2)}}m</b>`;
document.getElementById('molding-info').innerHTML=infoHtml;

const center=new THREE.Vector3();bbox.getCenter(center);
const size=bbox.getSize(new THREE.Vector3());const maxDim=Math.max(size.x,size.y,size.z);
camera.position.set(center.x+maxDim*0.8,center.y+maxDim*0.6,center.z+maxDim*0.8);
const controls=new Orbit(camera,renderer.domElement);
controls.t.copy(center);camera.lookAt(center);

function resetCam(){{
camera.position.set(center.x+maxDim*0.8,center.y+maxDim*0.6,center.z+maxDim*0.8);
controls.t.copy(center);camera.lookAt(center);}}

const leg=document.getElementById('legend');
["廻り縁１","廻り縁２","廻り縁３"].forEach(mt=>{{
const cc=MOLDING_TYPE_CSS[mt]||"#fff";
const t=TYPE_TOTALS[mt]||0;
const d=document.createElement('div');
d.innerHTML='<span class="cb" style="background:'+cc+'"></span>'+mt+' ('+t.toFixed(1)+'m)';
d.style.fontWeight='bold';
leg.appendChild(d);
}});
const sep=document.createElement('div');sep.style.borderTop='1px solid rgba(255,255,255,0.3)';sep.style.margin='6px 0';leg.appendChild(sep);
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

const typeVisible = {{"廻り縁１":true,"廻り縁２":true,"廻り縁３":true}};
let showB=true;
function toggleType(mt){{
  typeVisible[mt]=!typeVisible[mt];
  if(moldingGroups[mt]) moldingGroups[mt].visible=typeVisible[mt];
  const btnMap = {{"廻り縁１":"btn-m1","廻り縁２":"btn-m2","廻り縁３":"btn-m3"}};
  const btn=document.getElementById(btnMap[mt]);
  if(btn) btn.classList.toggle('active',typeVisible[mt]);
}}
function toggleBuilding(){{showB=!showB;buildingGroup.visible=showB;document.getElementById('btn-building').classList.toggle('active',showB);}}

const grid=new THREE.GridHelper(20,40,0x444444,0x333333);grid.position.copy(center);grid.position.y=0;scene.add(grid);
(function anim(){{requestAnimationFrame(anim);renderer.render(scene,camera);}})();
addEventListener('resize',()=>{{camera.aspect=innerWidth/innerHeight;camera.updateProjectionMatrix();renderer.setSize(innerWidth,innerHeight);}});
</script>
</body>
</html>"""


if __name__ == "__main__":
    # CLI引数でIFCパスと出力パスを上書き
    if len(sys.argv) >= 2:
        IFC_PATH = sys.argv[1]
    if len(sys.argv) >= 3:
        OUTPUT_HTML = sys.argv[2]

    # 出力ディレクトリを自動作成
    out_dir = os.path.dirname(OUTPUT_HTML)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    print(f"入力IFC: {IFC_PATH}")
    print(f"出力HTML: {OUTPUT_HTML}")
    main()
