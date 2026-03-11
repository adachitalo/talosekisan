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
# 根太配置
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

    # === Part 4-5: ピッチ別に根太配置＆ライン生成 ===
    all_pitch_data = {}
    first_pitch = True

    for pitch in JOIST_PITCHES:
        pitch_mm = int(pitch * 1000)
        print(f"\n根太配置計算中... (ピッチ: {pitch_mm}mm)")
        joists = place_joists(compartments, rp, pitch=pitch)
        print(f"  配置位置: {len(joists)}箇所")

        joist_lines = generate_joist_lines(joists, rp)
        print(f"  根太ライン: {len(joist_lines)}本")

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

        if first_pitch:
            print(f"\n=== 根太積算 ({pitch_mm}mm) ===")
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
        first_pitch = False

    # === HTML出力 ===
    default_pitch = int(JOIST_PITCHES[0] * 1000)
    meshes_json = json.dumps(meshes, ensure_ascii=False)
    colors_json = json.dumps(CATEGORY_COLORS, ensure_ascii=False)

    # 区画情報をJSON化（Three.js座標系）
    comp_data = []
    for c in compartments:
        # IFC座標 → Three.js座標: (X, Y, Z) → (X, Z, -Y)
        comp_data.append({
            "x_min": c["x_min"], "x_max": c["x_max"],
            "z_min": -c["y_max"], "z_max": -c["y_min"],  # IFC Y → Three.js -Z
            "width_x": c["width_x"], "width_y": c["width_y"],
        })
    compartments_json = json.dumps(comp_data, ensure_ascii=False)

    pitch_data_for_html = {}
    for pitch_mm, pdata in all_pitch_data.items():
        pitch_data_for_html[pitch_mm] = {
            "joists": pdata["joist_lines"],
            "summary": pdata["summary"],
        }
    pitch_data_json = json.dumps(pitch_data_for_html, ensure_ascii=False)
    pitches_json = json.dumps([int(p * 1000) for p in JOIST_PITCHES])

    html = generate_html(meshes_json, pitch_data_json, pitches_json,
                         default_pitch, colors_json, compartments_json, model_name)

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
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(summary_out, f, ensure_ascii=False, indent=2)
    print(f"JSON集計: {json_path}")


###############################################################################
# HTML生成
###############################################################################

def generate_html(meshes_json, pitch_data_json, pitches_json,
                  default_pitch, colors_json, compartments_json, model_name=""):
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
  <button id="btn-neda" class="active" onclick="toggleNeda()">根太表示</button>
  <button id="btn-comp" class="active" onclick="toggleComp()">区画表示</button>
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
  // 区画枠線
  const pts=[
    new THREE.Vector3(c.x_min, slabY+0.01, c.z_min),
    new THREE.Vector3(c.x_max, slabY+0.01, c.z_min),
    new THREE.Vector3(c.x_max, slabY+0.01, c.z_max),
    new THREE.Vector3(c.x_min, slabY+0.01, c.z_max),
    new THREE.Vector3(c.x_min, slabY+0.01, c.z_min),
  ];
  const lg=new THREE.BufferGeometry().setFromPoints(pts);
  const lmat=new THREE.LineBasicMaterial({{color,linewidth:2}});
  compGroup.add(new THREE.Line(lg,lmat));
}});

// 根太ライン描画
const nedaGroup=new THREE.Group();
scene.add(nedaGroup);

function buildNedaLines(pitchMm){{
  while(nedaGroup.children.length>0) nedaGroup.remove(nedaGroup.children[0]);
  const pd=PITCH_DATA[pitchMm];
  if(!pd) return;
  pd.joists.forEach(j=>{{
    const compColor=COMP_COLORS[j.comp_idx%COMP_COLORS.length];
    const color=j.reason==="区画端部"?0x00e676:compColor;
    const pts=[new THREE.Vector3(...j.seg[0]),new THREE.Vector3(...j.seg[1])];
    const g=new THREE.BufferGeometry().setFromPoints(pts);
    const mat=new THREE.LineBasicMaterial({{color,linewidth:2}});
    const line=new THREE.Line(g,mat);
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
  document.getElementById('neda-info').innerHTML=h;
}}

function switchPitch(pitchMm){{
  currentPitch=pitchMm;
  buildNedaLines(pitchMm);
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

let showN=true,showC=true,showB=true;
function toggleNeda(){{showN=!showN;nedaGroup.visible=showN;document.getElementById('btn-neda').classList.toggle('active',showN);}}
function toggleComp(){{showC=!showC;compGroup.visible=showC;document.getElementById('btn-comp').classList.toggle('active',showC);}}
function toggleBuilding(){{showB=!showB;buildingGroup.visible=showB;document.getElementById('btn-building').classList.toggle('active',showB);}}

const grid=new THREE.GridHelper(20,40,0x444444,0x333333);grid.position.copy(center);grid.position.y=0;scene.add(grid);
(function anim(){{requestAnimationFrame(anim);renderer.render(scene,camera);}})();
addEventListener('resize',()=>{{camera.aspect=innerWidth/innerHeight;camera.updateProjectionMatrix();renderer.setSize(innerWidth,innerHeight);}});
</script>
</body>
</html>"""


if __name__ == "__main__":
    main()
