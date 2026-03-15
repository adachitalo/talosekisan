#!/usr/bin/env python3
"""巾木拾い3Dビューア v3: 外周壁室内側 + 2F対応 + スラブクリッピング"""
import sys
import os
import ifcopenshell
import ifcopenshell.geom
import ifcopenshell.util.element
import json
import numpy as np
from collections import defaultdict

# CLI: python build_habaki_viewer.py <input.ifc> <output.html>

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
    """
    壁ジオメトリから底辺情報を抽出。
    Returns: {edges, floor_y, top_y, direction, length_m,
              face_coords: [perp_coord_face0, perp_coord_face1]}
    """
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
        z_vals = np.unique(np.round(bottom[:, 2], 2))
        if len(z_vals) < 2:
            z_c = float(z_vals[0])
            return {
                "edges": [[[x_min, min_y, z_c], [x_max, min_y, z_c]]],
                "floor_y": float(min_y), "top_y": float(max_y),
                "direction": direction, "length_m": float(dx),
                "face_coords": [z_c],
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
        }
    else:
        direction = "z"
        z_min, z_max = float(bottom[:, 2].min()), float(bottom[:, 2].max())
        x_vals = np.unique(np.round(bottom[:, 0], 2))
        if len(x_vals) < 2:
            x_c = float(x_vals[0])
            return {
                "edges": [[[x_c, min_y, z_min], [x_c, min_y, z_max]]],
                "floor_y": float(min_y), "top_y": float(max_y),
                "direction": direction, "length_m": float(dz),
                "face_coords": [x_c],
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
        }


def edges_at_y(info, y_level):
    """壁のedges情報をY座標を差し替えた版として返す（ログ壁は直方体のためXZ断面は同一）"""
    new_edges = []
    for edge in info["edges"]:
        new_edge = [[edge[0][0], y_level, edge[0][2]],
                     [edge[1][0], y_level, edge[1][2]]]
        new_edges.append(new_edge)
    return new_edges


def wall_cross_section_at_y(vf_raw, ff_raw, y_level, tol=0.01):
    """
    壁メッシュをY=y_levelで切断し、XZ平面上の断面線分を返す。
    Returns: [(np.array([x,z]), np.array([x,z])), ...]
    """
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


def extract_baseboard_from_cross_section(vf_raw, ff_raw, y_level, direction):
    """
    壁メッシュのY=y_levelでの断面から巾木ラインセグメントを抽出。
    壁の主方向に沿った長い線分を各面ごとにグループ化して返す。
    Returns: {face_coord: [([x,y,z],[x,y,z]), ...], ...}
    """
    cs_segs = wall_cross_section_at_y(vf_raw, ff_raw, y_level)
    face_segments = defaultdict(list)

    for p1, p2 in cs_segs:
        if direction == "x":
            # X方向の壁: 巾木はX方向、面はZ座標で識別
            dx = abs(p2[0] - p1[0])
            dz = abs(p2[1] - p1[1])
            if dx > 0.02 and dz < 0.005:
                z_face = round((p1[1] + p2[1]) / 2, 3)
                x_min = min(p1[0], p2[0])
                x_max = max(p1[0], p2[0])
                face_segments[z_face].append((x_min, x_max))
        else:
            # Z方向の壁: 巾木はZ方向、面はX座標で識別
            dx = abs(p2[0] - p1[0])
            dz = abs(p2[1] - p1[1])
            if dz > 0.02 and dx < 0.005:
                x_face = round((p1[0] + p2[0]) / 2, 3)
                z_min = min(p1[1], p2[1])
                z_max = max(p1[1], p2[1])
                face_segments[x_face].append((z_min, z_max))

    # 各面のセグメントをマージ
    result = {}
    for face_coord, segs in face_segments.items():
        segs.sort()
        merged = []
        for s, e in segs:
            if merged and s <= merged[-1][1] + 0.02:
                merged[-1] = (merged[-1][0], max(merged[-1][1], e))
            else:
                merged.append((s, e))
        # 3D座標に変換
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
    """
    外周/室内を判定し、外周壁にはどちら側が外側かも返す。
    Returns: {gid: {"is_ext": bool, "ext_side": "min"|"max"|None}}
    ext_side: 壁の垂直方向座標で、外側がmin側かmax側か
    """
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
                    # この壁群が min 側なら外側は min、max 側なら外側は max
                    if coord == min_c:
                        ext_side = "min"
                    else:
                        ext_side = "max"
                for g, inf, _ in perp_groups[coord]:
                    result[g] = {"is_ext": is_ext, "ext_side": ext_side}
        else:
            for coord in sorted_coords:
                for g, inf, _ in perp_groups[coord]:
                    result[g] = {"is_ext": True, "ext_side": "min"}

    # 間仕切壁は全て室内
    for c, g, n, inf in walls_list:
        if g not in result:
            result[g] = {"is_ext": False, "ext_side": None}
    return result


def select_interior_face(info, ext_side):
    """
    外周壁の室内側面のインデックスを返す。
    face_coords[0] < face_coords[1] の順。
    ext_side="min" → 外側は face_coords[0]（小さい方） → 室内側は face_coords[1] = index 1
    ext_side="max" → 外側は face_coords[1]（大きい方） → 室内側は face_coords[0] = index 0
    """
    if len(info["edges"]) < 2:
        return 0
    if ext_side == "min":
        return 1  # 室内側 = 大きい座標側
    else:
        return 0  # 室内側 = 小さい座標側


def split_edge_by_doors(p1, p2, doors, direction):
    """エッジをドア開口で分割"""
    if not doors:
        return [(p1, p2)]

    coord_idx = 0 if direction == "x" else 2
    c1 = p1[coord_idx]
    c2 = p2[coord_idx]
    if c1 > c2:
        c1, c2 = c2, c1
        p1, p2 = p2, p1

    intervals = []
    for d in doors:
        ds = d["x_min"] if direction == "x" else d["z_min"]
        de = d["x_max"] if direction == "x" else d["z_max"]
        ds = max(ds, c1)
        de = min(de, c2)
        if ds < de:
            intervals.append((ds, de))

    intervals.sort()
    merged = []
    for s, e in intervals:
        if merged and s <= merged[-1][1] + 0.01:
            merged[-1] = (merged[-1][0], max(merged[-1][1], e))
        else:
            merged.append((s, e))

    segments = []
    pos = c1
    for ds, de in merged:
        if pos < ds - 0.01:
            segments.append((interp(p1, p2, pos, c1, c2), interp(p1, p2, ds, c1, c2)))
        pos = de
    if pos < c2 - 0.01:
        segments.append((interp(p1, p2, pos, c1, c2), list(p2)))

    return segments if segments else [(p1, p2)]


def interp(p1, p2, val, c1, c2):
    if abs(c2 - c1) < 0.001:
        return list(p1)
    t = (val - c1) / (c2 - c1)
    return [p1[i] + t * (p2[i] - p1[i]) for i in range(3)]


def point_in_triangle_2d(px, pz, ax, az, bx, bz, cx, cz):
    """XZ平面上の点(px,pz)が三角形(a,b,c)内にあるか"""
    d = (bz - cz) * (ax - cx) + (cx - bx) * (az - cz)
    if abs(d) < 1e-10:
        return False
    a = ((bz - cz) * (px - cx) + (cx - bx) * (pz - cz)) / d
    b = ((cz - az) * (px - cx) + (ax - cx) * (pz - cz)) / d
    c = 1 - a - b
    tol = 0.02  # 2cm tolerance
    return a >= -tol and b >= -tol and c >= -tol


def point_on_slab(px, pz, slab_tris, margin=0.06):
    """点(px,pz)がスラブ三角形メッシュ上にあるか（margin付き）"""
    # margetで壁の厚み分を考慮（壁面は壁芯から±0.06m）
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
        pts.append([vf[i], vf[i + 2], -vf[i + 1]])  # Three.js変換
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


def clip_segment_to_slab(p1, p2, slab_tris, direction, n_samples=20):
    """
    巾木セグメントをスラブ範囲でクリップ。
    セグメントを n_samples 個にサンプリングし、
    スラブ上にある連続区間のみ返す。
    """
    if not slab_tris:
        return [(p1, p2)]

    coord_idx = 0 if direction == "x" else 2
    c1, c2 = p1[coord_idx], p2[coord_idx]
    if c1 > c2:
        p1, p2 = p2, p1
        c1, c2 = c2, c1

    seg_len = c2 - c1
    if seg_len < 0.01:
        # very short → just test midpoint
        mx = (p1[0] + p2[0]) / 2
        mz = (p1[2] + p2[2]) / 2
        if point_on_slab(mx, mz, slab_tris, margin=0):
            return [(p1, p2)]
        return []

    # Sample points along segment
    on_slab = []
    for i in range(n_samples + 1):
        t = i / n_samples
        sx = p1[0] + t * (p2[0] - p1[0])
        sz = p1[2] + t * (p2[2] - p1[2])
        on_slab.append(point_on_slab(sx, sz, slab_tris, margin=0))

    # Extract contiguous "on" runs
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


def main():
    if len(sys.argv) >= 3:
        IFC_PATH = sys.argv[1]
        OUTPUT_HTML = sys.argv[2]
    else:
        print("Usage: python build_habaki_viewer.py <input.ifc> <output.html>")
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

    # === Part 2: 壁ジオメトリ解析 ===
    print("\n壁ジオメトリ解析中...")
    seen2 = set()
    wall_data = []  # (cat, gid, ename, info, vf_raw, ff_raw)

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

    # === 床レベル定義 ===
    # 1F: スラブ上面 Y=0.172m → 巾木高さ 0.172m（間仕切壁）/ 0.0m（ログ壁）
    # 2F: スラブ上面 Y=2.907m → 巾木高さ 2.907m
    FLOOR_LEVELS = {
        "1F": {"slab_top": 0.172, "threshold": 0.3},
        "2F": {"slab_top": 2.907, "threshold": 0.3},
    }

    # 1F巾木対象: 壁底辺が1Fスラブ付近 (bottom_Y < 0.3m)
    walls_1f_raw = [(c, g, n, inf, vf, ff) for c, g, n, inf, vf, ff in wall_data
                    if inf["floor_y"] < FLOOR_LEVELS["1F"]["threshold"]]
    walls_1f = [(c, g, n, inf) for c, g, n, inf, vf, ff in walls_1f_raw]

    # === 2Fスラブの存在チェック（平屋判定） ===
    # 2Fスラブ（2-yuka）が存在しない場合は平屋とみなし、2F巾木処理をスキップ
    has_2f_slab = any(
        (slab.Name or "") == "2-yuka"
        for slab in ifc.by_type("IfcSlab")
    )

    # 2F巾木対象: ログ壁で、2F床を跨ぐ壁（bottom < 2.907 かつ top > 2.907）
    # 間仕切壁は2Fには存在しない（IFCデータ上）
    # ※ 2Fスラブが存在しない場合（平屋など）はスキップ
    slab_2f = FLOOR_LEVELS["2F"]["slab_top"]
    walls_2f_raw = []  # (cat, gid, ename, info, vf_raw, ff_raw) — 断面計算用にメッシュも保持
    walls_2f = []

    if has_2f_slab:
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
                if info and info["floor_y"] < slab_2f and info["top_y"] > slab_2f + 0.1:
                    walls_2f_raw.append((cat, gid, ename, info, list(vf), list(ff)))
            except Exception:
                pass

        # 互換性のために walls_2f も保持（classify_exterior_with_side用）
        walls_2f = [(c, g, n, inf) for c, g, n, inf, vf, ff in walls_2f_raw]

    print(f"  1F巾木対象壁: {len(walls_1f)}本")
    if has_2f_slab:
        print(f"  2F巾木対象壁: {len(walls_2f)}本（2F床 Y={slab_2f}m を跨ぐログ壁）")
    else:
        print(f"  平屋（2Fスラブなし）→ 2F巾木処理をスキップ")

    # === スラブ上面メッシュ取得（巾木クリッピング用） ===
    print("  スラブ上面メッシュを取得中...")
    slab_tris_by_floor = {"1F": [], "2F": []}
    for slab in ifc.by_type("IfcSlab"):
        ename = slab.Name or ""
        try:
            if ename == "1-yuka" or ename == "1-yuka-tile":
                tris = extract_slab_top_tris(slab, settings)
                slab_tris_by_floor["1F"].extend(tris)
                print(f"    {ename}: {len(tris)}三角形")
            elif ename == "2-yuka":
                tris = extract_slab_top_tris(slab, settings)
                slab_tris_by_floor["2F"].extend(tris)
                print(f"    {ename}: {len(tris)}三角形")
        except Exception:
            pass

    # === ドア情報取得 ===
    print("  ドア開口位置を取得中...")
    door_bottoms = {}
    for door in ifc.by_type("IfcDoor"):
        try:
            shape = ifcopenshell.geom.create_shape(settings, door)
            vf = shape.geometry.verts
            verts_3js = []
            for i in range(0, len(vf), 3):
                verts_3js.extend([vf[i], vf[i + 2], -vf[i + 1]])
            pts = np.array(verts_3js).reshape(-1, 3)
            min_y = pts[:, 1].min()
            bottom = pts[pts[:, 1] < (min_y + 0.01)]
            dx = bottom[:, 0].max() - bottom[:, 0].min()
            dz = bottom[:, 2].max() - bottom[:, 2].min()
            door_bottoms[door.GlobalId] = {
                "x_min": float(bottom[:, 0].min()),
                "x_max": float(bottom[:, 0].max()),
                "z_min": float(bottom[:, 2].min()),
                "z_max": float(bottom[:, 2].max()),
                "floor_y": float(min_y),
                "direction": "x" if dx > dz else "z",
            }
        except Exception:
            pass

    door_to_wall = {}
    for rel in ifc.by_type("IfcRelVoidsElement"):
        host = rel.RelatingBuildingElement
        opening = rel.RelatedOpeningElement
        if not host or not opening:
            continue
        for fill_rel in ifc.by_type("IfcRelFillsElement"):
            if fill_rel.RelatingOpeningElement == opening:
                door = fill_rel.RelatedBuildingElement
                if door and door.is_a("IfcDoor"):
                    door_to_wall[door.GlobalId] = host.GlobalId

    wall_doors = defaultdict(list)
    for dgid, wgid in door_to_wall.items():
        if dgid in door_bottoms:
            wall_doors[wgid].append(door_bottoms[dgid])

    # === 巾木ライン生成 ===
    habaki_lines = []
    total_by_floor = defaultdict(lambda: {"ext": 0, "int": 0})

    # --- 1F: 壁メッシュ断面ベース + スラブクリッピング ---
    ext_map_1f = classify_exterior_with_side(walls_1f)
    slab_y_1f = FLOOR_LEVELS["1F"]["slab_top"]
    slab_tris_1f = slab_tris_by_floor.get("1F", [])

    print(f"\n  1F (巾木高さ Y={slab_y_1f:.3f}m): {len(walls_1f)}壁 [断面ベース]")

    for cat, gid, ename, info, vf_raw, ff_raw in walls_1f_raw:
        ext_info = ext_map_1f.get(gid, {"is_ext": False, "ext_side": None})
        is_ext = ext_info["is_ext"]
        ext_side = ext_info["ext_side"]
        direction = info["direction"]

        # 断面高さ: ログ壁はY=0.0(土台上)、間仕切壁はY=0.172(スラブ上)
        # → スラブ上面 0.172m + 少し上 で断面を取ることで、
        #   ドア開口部も正確に反映される
        cs_y = slab_y_1f + 0.05  # スラブ上面から5cm上で断面

        face_segs = extract_baseboard_from_cross_section(vf_raw, ff_raw, cs_y, direction)

        if not face_segs:
            print(f"    {cat:8s} {ename:15s} → 断面なし")
            continue

        face_coords_sorted = sorted(face_segs.keys())

        # 面の選択
        if len(face_coords_sorted) < 2:
            faces_to_use_coords = face_coords_sorted
        elif is_ext and cat == "ログ壁":
            if ext_side == "min":
                faces_to_use_coords = [face_coords_sorted[-1]]
            else:
                faces_to_use_coords = [face_coords_sorted[0]]
        elif not is_ext:
            faces_to_use_coords = face_coords_sorted
        else:
            if ext_side == "min":
                faces_to_use_coords = [face_coords_sorted[-1]]
            else:
                faces_to_use_coords = [face_coords_sorted[0]]

        wall_seg_total = 0
        for fc in faces_to_use_coords:
            for seg_3d in face_segs[fc]:
                p1, p2 = seg_3d

                # スラブ範囲でクリッピング
                clipped = clip_segment_to_slab(p1, p2, slab_tris_1f, direction)

                for sp1, sp2 in clipped:
                    seg_len = np.sqrt(sum((a - b) ** 2 for a, b in zip(sp1, sp2)))
                    if seg_len < 0.05:
                        continue
                    wall_seg_total += seg_len
                    habaki_lines.append({
                        "floor": "1F",
                        "type": "exterior" if is_ext else "interior",
                        "cat": cat,
                        "seg": [
                            [round(sp1[0], 4), round(sp1[1] + 0.02, 4), round(sp1[2], 4)],
                            [round(sp2[0], 4), round(sp2[1] + 0.02, 4), round(sp2[2], 4)],
                        ],
                        "length_m": round(seg_len, 4),
                    })

        if is_ext:
            total_by_floor["1F"]["ext"] += wall_seg_total
        else:
            total_by_floor["1F"]["int"] += wall_seg_total

        pos_label = "外周" if is_ext else "室内"
        n_faces = len(faces_to_use_coords)
        print(f"    {pos_label} {cat:8s} {ename:15s} "
              f"断面{n_faces}面 → {wall_seg_total:.2f}m")

    # --- 2F: 壁メッシュ断面ベース + スラブクリッピング ---
    # 平屋（2Fスラブなし）の場合はスキップ
    if has_2f_slab:
        ext_map_2f = classify_exterior_with_side(walls_2f)
        slab_y_2f = FLOOR_LEVELS["2F"]["slab_top"]
        slab_tris_2f = slab_tris_by_floor.get("2F", [])

        print(f"\n  2F (巾木高さ Y={slab_y_2f:.3f}m): {len(walls_2f)}壁 [断面ベース]")

        for cat, gid, ename, info, vf_raw, ff_raw in walls_2f_raw:
            ext_info = ext_map_2f.get(gid, {"is_ext": False, "ext_side": None})
            is_ext = ext_info["is_ext"]
            ext_side = ext_info["ext_side"]
            direction = info["direction"]

            # 壁メッシュの2F断面から巾木セグメントを抽出
            face_segs = extract_baseboard_from_cross_section(vf_raw, ff_raw, slab_y_2f, direction)

            if not face_segs:
                print(f"    {cat:8s} {ename:15s} → 2Fに断面なし")
                continue

            # 面座標を並べ替え
            face_coords_sorted = sorted(face_segs.keys())

            # 面の選択（exterior→室内側のみ、interior→両面）
            if len(face_coords_sorted) < 2:
                faces_to_use_coords = face_coords_sorted
            elif is_ext:
                # 外周壁: 室内側のみ
                if ext_side == "min":
                    faces_to_use_coords = [face_coords_sorted[-1]]  # max側 = 室内
                else:
                    faces_to_use_coords = [face_coords_sorted[0]]   # min側 = 室内
            else:
                faces_to_use_coords = face_coords_sorted  # 両面

            wall_seg_total = 0
            for fc in faces_to_use_coords:
                for seg_3d in face_segs[fc]:
                    p1, p2 = seg_3d

                    # スラブ範囲でクリッピング
                    clipped = clip_segment_to_slab(p1, p2, slab_tris_2f, direction)

                    for sp1, sp2 in clipped:
                        seg_len = np.sqrt(sum((a - b) ** 2 for a, b in zip(sp1, sp2)))
                        if seg_len < 0.05:
                            continue
                        wall_seg_total += seg_len
                        habaki_lines.append({
                            "floor": "2F",
                            "type": "exterior" if is_ext else "interior",
                            "cat": cat,
                            "seg": [
                                [round(sp1[0], 4), round(sp1[1] + 0.02, 4), round(sp1[2], 4)],
                                [round(sp2[0], 4), round(sp2[1] + 0.02, 4), round(sp2[2], 4)],
                            ],
                            "length_m": round(seg_len, 4),
                        })

            if is_ext:
                total_by_floor["2F"]["ext"] += wall_seg_total
            else:
                total_by_floor["2F"]["int"] += wall_seg_total

            pos_label = "外周" if is_ext else "室内"
            n_faces = len(faces_to_use_coords)
            print(f"    {pos_label} {cat:8s} {ename:15s} "
                  f"断面{n_faces}面 → {wall_seg_total:.2f}m")
    else:
        print(f"\n  2F: スキップ（平屋のため2Fスラブなし）")

    # 集計表示
    grand_total = 0
    print("\n=== 巾木集計 ===")
    for fn in ["1F", "2F"]:
        if fn not in total_by_floor:
            continue
        t = total_by_floor[fn]
        floor_total = t["ext"] + t["int"]
        grand_total += floor_total
        print(f"  {fn}: 外周(片面)={t['ext']:.2f}m  室内(両面)={t['int']:.2f}m  計={floor_total:.2f}m")
    print(f"  合計: {grand_total:.2f}m")

    # === HTML出力 ===
    meshes_json = json.dumps(meshes, ensure_ascii=False)
    habaki_json = json.dumps(habaki_lines, ensure_ascii=False)
    colors_json = json.dumps(CATEGORY_COLORS, ensure_ascii=False)

    totals_data = {}
    for fn in ["1F", "2F"]:
        if fn in total_by_floor:
            t = total_by_floor[fn]
            totals_data[fn] = {"ext": round(t["ext"], 2), "int": round(t["int"], 2)}

    totals_json = json.dumps(totals_data, ensure_ascii=False)
    html = generate_html(meshes_json, habaki_json, colors_json, totals_json, model_name)

    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"\n出力: {OUTPUT_HTML}")

    # === JSON集計出力（部材一覧Excel統合用） ===
    json_path = os.path.splitext(OUTPUT_HTML)[0] + "_summary.json"
    summary = {
        "tool": "巾木拾い",
        "model": model_name,
        "floor_totals": {},
        "grand_total": round(grand_total, 2),
        "lines": [
            {
                "floor": h["floor"],
                "type": h["type"],
                "cat": h["cat"],
                "length_m": h["length_m"],
            }
            for h in habaki_lines
        ],
    }
    for fn in ["1F", "2F"]:
        if fn in total_by_floor:
            t = total_by_floor[fn]
            summary["floor_totals"][fn] = {
                "ext": round(t["ext"], 2),
                "int": round(t["int"], 2),
                "total": round(t["ext"] + t["int"], 2),
            }
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"JSON集計: {json_path}")


def generate_html(meshes_json, habaki_json, colors_json, totals_json, model_name=""):
    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>巾木拾い 3Dビューア - {model_name}</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ background:#1a1a2e; overflow:hidden; font-family:Arial,sans-serif; }}
#info {{ position:fixed; top:10px; left:10px; color:#fff; background:rgba(0,0,0,0.75);
  padding:12px 16px; border-radius:8px; font-size:13px; z-index:100; max-width:360px; }}
#info h3 {{ margin-bottom:6px; color:#ff6b6b; font-size:15px; }}
#legend {{ position:fixed; top:10px; right:10px; color:#fff; background:rgba(0,0,0,0.75);
  padding:12px; border-radius:8px; font-size:12px; z-index:100; max-height:80vh; overflow-y:auto; }}
#legend div {{ cursor:pointer; padding:3px 6px; border-radius:3px; margin:2px 0; white-space:nowrap; }}
#legend div:hover {{ background:rgba(255,255,255,0.15); }}
.cb {{ display:inline-block; width:14px; height:14px; border-radius:3px; margin-right:6px; vertical-align:middle; }}
#habaki-info {{ position:fixed; bottom:10px; left:10px; color:#fff; background:rgba(0,0,0,0.85);
  padding:14px 18px; border-radius:8px; font-size:13px; z-index:100; line-height:1.6; }}
#controls {{ position:fixed; bottom:10px; right:10px; z-index:100; }}
#controls button {{ background:rgba(255,255,255,0.15); color:#fff; border:1px solid rgba(255,255,255,0.3);
  padding:8px 14px; border-radius:6px; cursor:pointer; margin:2px; font-size:12px; }}
#controls button:hover {{ background:rgba(255,255,255,0.3); }}
#controls button.active {{ background:rgba(255,100,100,0.5); border-color:#ff6b6b; }}
</style>
</head>
<body>
<div id="info">
  <h3>巾木拾い 3Dビューア</h3>
  <div>左ドラッグ: 回転 / 右ドラッグ: 移動 / ホイール: ズーム</div>
  <div id="sel-info" style="margin-top:6px;color:#aaa;">クリックで部材選択</div>
</div>
<div id="habaki-info"></div>
<div id="legend"></div>
<div id="controls">
  <button id="btn-habaki" class="active" onclick="toggleHabaki()">巾木表示</button>
  <button id="btn-building" class="active" onclick="toggleBuilding()">建物表示</button>
  <button onclick="resetCam()">リセット</button>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/lines/LineGeometry.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/lines/LineMaterial.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/lines/Line2.js"></script>
<script>
const MESHES={meshes_json};
const HABAKI={habaki_json};
const COLORS={colors_json};
const TOTALS={totals_json};

const FLOOR_COLORS = {{
  "1F": {{ ext: 0xff4444, int: 0x44ff44 }},
  "2F": {{ ext: 0xff8800, int: 0x44aaff }}
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

const habakiGroup=new THREE.Group();
HABAKI.forEach(h=>{{
const fc=FLOOR_COLORS[h.floor]||FLOOR_COLORS["1F"];
const color=h.type==="exterior"?fc.ext:fc.int;
const s=h.seg;
const g=new THREE.LineGeometry();
g.setPositions([s[0][0],s[0][1],s[0][2],s[1][0],s[1][1],s[1][2]]);
const mat=new THREE.LineMaterial({{color,linewidth:4,resolution:new THREE.Vector2(innerWidth,innerHeight)}});
habakiGroup.add(new THREE.Line2(g,mat));
}});
scene.add(habakiGroup);

let infoHtml='<b style="font-size:14px;">巾木集計</b><br><br>';
let grandTotal=0;
for(const [fn,t] of Object.entries(TOTALS)){{
  const fc=FLOOR_COLORS[fn]||FLOOR_COLORS["1F"];
  const floorTotal=t.ext+t.int;grandTotal+=floorTotal;
  infoHtml+=`<b>${{fn}}</b><br>`;
  infoHtml+=`<span style="color:#${{fc.ext.toString(16).padStart(6,'0')}}">━━ 外周(片面): ${{t.ext.toFixed(2)}}m</span><br>`;
  infoHtml+=`<span style="color:#${{fc.int.toString(16).padStart(6,'0')}}">━━ 室内(両面): ${{t.int.toFixed(2)}}m</span><br>`;
  infoHtml+=`小計: ${{floorTotal.toFixed(2)}}m<br><br>`;
}}
infoHtml+=`<b style="font-size:14px;">合計: ${{grandTotal.toFixed(2)}}m</b>`;
document.getElementById('habaki-info').innerHTML=infoHtml;

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

let showH=true,showB=true;
function toggleHabaki(){{showH=!showH;habakiGroup.visible=showH;document.getElementById('btn-habaki').classList.toggle('active',showH);}}
function toggleBuilding(){{showB=!showB;buildingGroup.visible=showB;document.getElementById('btn-building').classList.toggle('active',showB);}}

const grid=new THREE.GridHelper(20,40,0x444444,0x333333);grid.position.copy(center);grid.position.y=0;scene.add(grid);
(function anim(){{requestAnimationFrame(anim);renderer.render(scene,camera);}})();
addEventListener('resize',()=>{{camera.aspect=innerWidth/innerHeight;camera.updateProjectionMatrix();renderer.setSize(innerWidth,innerHeight);habakiGroup.traverse(c=>{{if(c.material&&c.material.resolution)c.material.resolution.set(innerWidth,innerHeight);}});}});
</script>
</body>
</html>"""


if __name__ == "__main__":
    main()
