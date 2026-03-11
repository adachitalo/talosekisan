#!/usr/bin/env python3
"""
build_gakubuchi_viewer.py – IFCから額縁拾い3Dビューアを生成

Usage: python build_gakubuchi_viewer.py input.ifc output.html

TALOログハウスIFCファイルから建具(IfcDoor/IfcWindow)を検出し、
額縁種別マスタに照合して額縁部材の必要数量を算出、
Three.js 3Dビューア付きHTMLとして出力する。

■ 額縁拾いルール:
  表面(front): 額縁 + 補助部材(額縁受け/T-bar/霧除け/木口)
  裏面(back):  額縁のみ(3辺=ドア, 4辺=窓)
  → T-bar/額縁受け/霧除け/木口は表面のみ(裏面なし)

■ 寸法オフセット:
  IFCのOverallWidth/Heightは建具のリーフ寸法。
  額縁計算にはログ壁開口寸法(=リーフ+枠)を使用する。
"""

import sys
import os
import json
import math
from collections import defaultdict
import numpy as np

import ifcopenshell
import ifcopenshell.geom

# ============================================================================
# 額縁種別ルール（建具タイプ→使用する額縁部材の種類）
# ============================================================================
# front_frame: "3" = 3辺(上左右), "4" = 4辺(上下左右)
# back_frame:  "3" = 3辺(上左右), "4" = 4辺(上下左右)
# front_extras: 表面の追加部材
FRAME_RULES = {
    "VMW":  {"front_frame": "4", "back_frame": "4",
             "front_extras": ["額縁受け", "T-bar", "霧除け"]},
    "VMSD": {"front_frame": "4", "back_frame": "3",
             "front_extras": ["額縁受け", "T-bar", "霧除け"]},
    "EURO": {"front_frame": "4", "back_frame": "3",
             "front_extras": ["額縁受け", "T-bar", "霧除け"]},
    "NV":   {"front_frame": "4", "back_frame": "3",
             "front_extras": ["額縁受け", "T-bar"]},
    "NVs":  {"front_frame": "3", "back_frame": "3",
             "front_extras": ["木口"]},
    "NW":   {"front_frame": "4", "back_frame": "3",
             "front_extras": ["T-bar", "木口"]},
    "JD":   {"front_frame": "3", "back_frame": "3",
             "front_extras": ["木口"]},
}

# ============================================================================
# IFC寸法 → ログ壁開口寸法 のオフセット(mm)
# ============================================================================
DIMENSION_OFFSETS = {
    "VMW":  (396, 438),
    "VMSD": (396, 268),
    "EURO": (396, 268),
    "NV":   (340, 120),
    "NVs":  (0, 0),
    "NW":   (340, 106),
    "JD":   (0, 0),
}

NW_SMALL_W_OFFSET = 280  # NW小型ドア(W<800)用

# 間仕切り壁の場合のオフセット差分 (NVのW offsetが変わる)
PARTITION_NV_W_OFFSET = 240  # ログ壁=340, 間仕切り=240

# 間仕切り壁の額縁ルール（ログ壁よりも簡素化）
PARTITION_FRAME_RULES = {
    "NV":  {"front_frame": "3", "back_frame": "3",
            "front_extras": []},
    "NW":  {"front_frame": "3", "back_frame": "3",
            "front_extras": ["木口"]},
}


# ============================================================================
# IFC壁タイプ判定
# ============================================================================
def get_host_wall_info(ifc_file):
    """各建具のホスト壁情報を収集 → {element_id: "log" or "partition"}"""
    # IfcRelVoidsElement: wall → opening
    opening_to_wall = {}
    for rel in ifc_file.by_type("IfcRelVoidsElement"):
        opening_to_wall[rel.RelatedOpeningElement.id()] = rel.RelatingBuildingElement

    # IfcRelFillsElement: opening → door/window
    element_wall_type = {}
    for rel in ifc_file.by_type("IfcRelFillsElement"):
        opening = rel.RelatingOpeningElement
        door = rel.RelatedBuildingElement
        wall = opening_to_wall.get(opening.id())
        if wall:
            wall_name = (wall.Name or "").lower()
            # TypeDefも確認
            wall_td = ""
            for r in ifc_file.by_type("IfcRelDefinesByType"):
                if wall in r.RelatedObjects:
                    wall_td = r.RelatingType.Name or ""
                    break
            if "ログ" in wall_td or "log" in wall_name:
                element_wall_type[door.id()] = "log"
            elif "間仕切" in wall_td or "majikiri" in wall_name:
                element_wall_type[door.id()] = "partition"
            else:
                element_wall_type[door.id()] = "log"  # デフォルトはログ壁

    return element_wall_type


# ============================================================================
# ArchiCAD TypeDef名 → 額縁タイプキー のマッピング
# ============================================================================
def classify_fixture_type(name: str, typedef: str) -> str:
    """ArchiCADのName/TypeDef名から額縁タイプキーを決定。
    額縁不要の建具はNoneを返す。
    """
    td = typedef.upper()
    nm = name.upper()

    # --- 額縁不要の建具を除外 ---
    # LogsOpeningDoor: ログ壁大開口（建具なし）
    if "LOGSOPENINGDOOR" in td or "LOGS_OPENING" in td:
        return None
    # open_ad: 開放部（建具リーフなし）
    if td == "OPEN_AD" or (nm.startswith("OPEN") and "DOOR" not in td):
        return None

    if "VMSD" in td or "PSD" in td:
        return "VMSD"
    if "VMW" in td or td.startswith("PW1") or td.startswith("PW2"):
        return "VMW"
    if "EURO" in td or "DOOR3" in td:
        return "EURO"
    if "NVS" in td or "ポケット" in typedef:
        return "NVs"
    if "NV" in td:
        return "NV"
    if "NW" in td:
        return "NW"
    if "スライド" in typedef or "JD" in td:
        return "JD"
    if "OPEN" in td:
        return "NV"

    if nm.startswith("AW"):
        return "VMW"
    if nm.startswith("SD"):
        return "JD"

    return "JD"


def make_fixture_label(name: str, typedef: str, ftype: str, w_mm: float) -> str:
    """表示用の建具ラベルを生成"""
    parts = [name]
    td = typedef.upper()
    for key in ["VMSD", "VMW", "EURO", "NVS", "NV", "NW"]:
        if key in td:
            parts.append(key.capitalize() if key != "NVS" else "NVs")
            break
    else:
        if "PSD" in td:
            parts.append("VMSD")
        elif "PW1" in td or "PW2" in td:
            parts.append("VMW")
        elif "ポケット" in typedef:
            parts.append("NVs")
        elif "スライド" in typedef:
            parts.append("JD")
        elif "DOOR3" in td:
            parts.append("EURO")
    return " ".join(parts)


def get_frame_dimensions(ftype: str, ifc_w: float, ifc_h: float, wall_type: str = "log"):
    """IFC寸法から額縁計算用の開口寸法(mm)を算出"""
    w_off, h_off = DIMENSION_OFFSETS.get(ftype, (0, 0))
    if ftype == "NW" and ifc_w < 800:
        w_off = NW_SMALL_W_OFFSET
    # 間仕切り壁のNVはW offsetが小さい
    if wall_type == "partition" and ftype == "NV":
        w_off = PARTITION_NV_W_OFFSET
    return ifc_w + w_off, ifc_h + h_off


# ============================================================================
# 額縁部材の長さ計算（mm単位）
# ============================================================================
def generate_frame_pieces(ftype: str, w_mm: float, h_mm: float, wall_type: str = "log"):
    """
    額縁タイプと開口寸法(mm)から、部材リスト(kind, piece_name, length_mm, side)を返す。
    side: "front" or "back"
    wall_type: "log" (ログ壁) or "partition" (間仕切り壁) — 間仕切りは簡素化
    """
    # 間仕切り壁で専用ルールがある場合はそちらを使用
    if wall_type == "partition" and ftype in PARTITION_FRAME_RULES:
        rule = PARTITION_FRAME_RULES[ftype]
    else:
        rule = FRAME_RULES.get(ftype, FRAME_RULES["JD"])
    pieces = []

    # --- 表面(front) ---
    if rule["front_frame"] == "4":
        pieces.append(("額縁", "上", w_mm + 300, "front"))
        pieces.append(("額縁", "左", h_mm + 170, "front"))
        pieces.append(("額縁", "右", h_mm + 170, "front"))
        pieces.append(("額縁", "下", w_mm, "front"))
    else:
        pieces.append(("額縁", "上", w_mm + 300, "front"))
        pieces.append(("額縁", "左", h_mm + 170, "front"))
        pieces.append(("額縁", "右", h_mm + 170, "front"))

    for extra in rule["front_extras"]:
        if extra == "額縁受け":
            pieces.append(("額縁受け", "横", w_mm, "front"))
        elif extra == "T-bar":
            pieces.append(("T-bar", "縦1", h_mm, "front"))
            pieces.append(("T-bar", "縦2", h_mm, "front"))
        elif extra == "霧除け":
            pieces.append(("霧除け", "横", w_mm + 300, "front"))
        elif extra == "木口":
            pieces.append(("木口", "上", w_mm, "front"))
            pieces.append(("木口", "左", h_mm, "front"))
            pieces.append(("木口", "右", h_mm, "front"))

    # --- 裏面(back) ---
    if rule["back_frame"] == "4":
        pieces.append(("額縁", "上", w_mm + 300, "back"))
        pieces.append(("額縁", "左", h_mm + 170, "back"))
        pieces.append(("額縁", "右", h_mm + 170, "back"))
        pieces.append(("額縁", "下", w_mm, "back"))
    else:
        pieces.append(("額縁", "上", w_mm + 300, "back"))
        pieces.append(("額縁", "左", h_mm + 170, "back"))
        pieces.append(("額縁", "右", h_mm + 170, "back"))

    return pieces


# ============================================================================
# IFC解析 – 建物メッシュ抽出
# ============================================================================
def extract_meshes(ifc_file):
    """IFCから建物要素のメッシュを抽出。
    天窓(IfcWindow on IfcRoof)は"天窓"カテゴリに分類。
    """
    settings = ifcopenshell.geom.settings()
    settings.set(settings.USE_WORLD_COORDS, True)

    meshes = []
    type_map = {
        "IfcWall": "壁", "IfcWallStandardCase": "壁",
        "IfcDoor": "ドア", "IfcWindow": "窓",
        "IfcBeam": "梁", "IfcColumn": "柱",
        "IfcSlab": "1F床", "IfcRailing": "手摺",
        "IfcRoof": "屋根", "IfcStair": "階段",
        "IfcMember": "部材",
    }

    # 天窓検出: IfcRoof に含まれる IfcWindow を特定
    skylight_ids = set()
    for rel in ifc_file.by_type("IfcRelVoidsElement"):
        host = rel.RelatingBuildingElement
        if host.is_a("IfcRoof"):
            opening = rel.RelatedOpeningElement
            for fill_rel in ifc_file.by_type("IfcRelFillsElement"):
                if fill_rel.RelatingOpeningElement.id() == opening.id():
                    skylight_ids.add(fill_rel.RelatedBuildingElement.id())
    # TypeDef名でも天窓を検出（スカイライト/skylight）
    for win in ifc_file.by_type("IfcWindow"):
        td = get_typedef_name(ifc_file, win).lower()
        nm = (win.Name or "").lower()
        if "スカイライト" in td or "skylight" in td or "スカイライト" in nm or "skylight" in nm or nm.startswith("sl"):
            skylight_ids.add(win.id())

    for element in ifc_file.by_type("IfcBuildingElement"):
        try:
            shape = ifcopenshell.geom.create_shape(settings, element)
            verts = list(shape.geometry.verts)
            faces = list(shape.geometry.faces)
            if not verts or not faces:
                continue
            ifc_type = element.is_a()
            cat = type_map.get(ifc_type, "")
            if not cat:
                continue
            # 天窓は専用カテゴリ
            if element.id() in skylight_ids:
                cat = "天窓"
            # IFC(Z-up) → Three.js(Y-up) 座標変換: (x,y,z)→(x,z,-y)
            verts_3js = []
            for vi in range(0, len(verts), 3):
                verts_3js.append(round(verts[vi], 4))
                verts_3js.append(round(verts[vi + 2], 4))
                verts_3js.append(round(-verts[vi + 1], 4))
            meshes.append({"cat": cat, "verts": verts_3js, "faces": faces})
        except Exception:
            pass

    return meshes


# ============================================================================
# IFC解析 – 建具検出と額縁ライン生成
# ============================================================================
def get_element_placement_and_center(element):
    """建具のワールド座標中心と方向ベクトルを取得。
    Returns: (center, width_dir, height_dir, depth_dir) すべてIFC座標系
    """
    settings_world = ifcopenshell.geom.settings()
    settings_world.set(settings_world.USE_WORLD_COORDS, True)
    settings_local = ifcopenshell.geom.settings()
    settings_local.set(settings_local.USE_WORLD_COORDS, False)

    # ワールド座標メッシュからバウンディングボックス中心を取得
    shape_w = ifcopenshell.geom.create_shape(settings_world, element)
    vf = shape_w.geometry.verts
    xs = [vf[i] for i in range(0, len(vf), 3)]
    ys = [vf[i + 1] for i in range(0, len(vf), 3)]
    zs = [vf[i + 2] for i in range(0, len(vf), 3)]
    center = np.array([
        (min(xs) + max(xs)) / 2,
        (min(ys) + max(ys)) / 2,
        (min(zs) + max(zs)) / 2
    ])

    # ローカル座標の変換行列から方向ベクトルを取得
    shape_l = ifcopenshell.geom.create_shape(settings_local, element)
    mat = list(shape_l.transformation.matrix)
    m44 = np.array(mat).reshape(4, 4).T
    # IFC建具配置: x_dir=幅方向, y_dir=奥行き方向, z_dir=高さ方向
    width_dir = m44[:3, 0]
    depth_dir = m44[:3, 1]
    height_dir = m44[:3, 2]

    width_dir = width_dir / (np.linalg.norm(width_dir) + 1e-12)
    height_dir = height_dir / (np.linalg.norm(height_dir) + 1e-12)
    depth_dir = depth_dir / (np.linalg.norm(depth_dir) + 1e-12)

    return center, width_dir, height_dir, depth_dir


def get_typedef_name(ifc_file, element):
    """IfcRelDefinesByType から型定義名を取得"""
    for rel in ifc_file.by_type("IfcRelDefinesByType"):
        if element in rel.RelatedObjects:
            return rel.RelatingType.Name or ""
    return ""


def detect_fixtures_and_frames(ifc_file):
    """建具を検出し、額縁ラインデータを生成"""
    fixtures_info = []
    all_frames = []

    # 壁タイプ情報を取得
    wall_info = get_host_wall_info(ifc_file)

    for el_type in ["IfcDoor", "IfcWindow"]:
        for element in ifc_file.by_type(el_type):
            name = element.Name or ""
            typedef = get_typedef_name(ifc_file, element)
            ifc_w = float(element.OverallWidth or 0)
            ifc_h = float(element.OverallHeight or 0)

            if ifc_w <= 0 or ifc_h <= 0:
                continue

            ftype = classify_fixture_type(name, typedef)
            if ftype is None:
                continue  # 額縁不要の建具（LogsOpeningDoor, open_ad等）
            label = make_fixture_label(name, typedef, ftype, ifc_w)

            wall_type = wall_info.get(element.id(), "log")
            frame_w, frame_h = get_frame_dimensions(ftype, ifc_w, ifc_h, wall_type)

            try:
                center, width_dir, height_dir, depth_dir = \
                    get_element_placement_and_center(element)
            except Exception:
                continue

            w_m = ifc_w / 1000.0
            h_m = ifc_h / 1000.0

            fixtures_info.append({
                "label": label,
                "ftype": ftype,
                "w_mm": ifc_w,
                "h_mm": ifc_h,
                "frame_w": frame_w,
                "frame_h": frame_h,
            })

            pieces = generate_frame_pieces(ftype, frame_w, frame_h, wall_type)

            # depth_dir = 壁面の法線方向（表/裏オフセット）
            front_offset = depth_dir * 0.015
            back_offset = -depth_dir * 0.015

            for kind, piece_name, length_mm, side in pieces:
                offset = front_offset if side == "front" else back_offset

                # center=メッシュ中心, width_dir=幅方向, height_dir=高さ方向
                p1, p2 = compute_frame_line_3d(
                    kind, piece_name,
                    center, width_dir, height_dir,
                    w_m, h_m, offset
                )

                # IFC(Z-up) → Three.js(Y-up): (x,y,z)→(x,z,-y)
                all_frames.append({
                    "kind": kind,
                    "points": [
                        [round(p1[0], 3), round(p1[2], 3), round(-p1[1], 3)],
                        [round(p2[0], 3), round(p2[2], 3), round(-p2[1], 3)]
                    ],
                    "length": round(length_mm),
                    "fixture": label,
                })

    return fixtures_info, all_frames


def compute_frame_line_3d(kind, piece_name, origin, x_dir, y_dir, w_m, h_m, offset):
    """額縁部材の3Dライン座標を計算"""
    hw = w_m / 2
    hh = h_m / 2
    base = origin + offset

    if piece_name in ("上", "横") and kind not in ("木口",):
        p1 = base + (-hw) * x_dir + hh * y_dir
        p2 = base + hw * x_dir + hh * y_dir
    elif piece_name == "下":
        p1 = base + (-hw) * x_dir + (-hh) * y_dir
        p2 = base + hw * x_dir + (-hh) * y_dir
    elif piece_name in ("左", "縦1"):
        p1 = base + (-hw) * x_dir + (-hh) * y_dir
        p2 = base + (-hw) * x_dir + hh * y_dir
    elif piece_name in ("右", "縦2"):
        p1 = base + hw * x_dir + (-hh) * y_dir
        p2 = base + hw * x_dir + hh * y_dir
    elif kind == "木口" and piece_name == "上":
        p1 = base + (-hw) * x_dir + hh * y_dir
        p2 = base + hw * x_dir + hh * y_dir
    elif kind == "木口" and piece_name == "左":
        p1 = base + (-hw) * x_dir + (-hh) * y_dir
        p2 = base + (-hw) * x_dir + hh * y_dir
    elif kind == "木口" and piece_name == "右":
        p1 = base + hw * x_dir + (-hh) * y_dir
        p2 = base + hw * x_dir + hh * y_dir
    elif kind == "額縁受け":
        p1 = base + (-hw) * x_dir + hh * y_dir
        p2 = base + hw * x_dir + hh * y_dir
    elif kind == "霧除け":
        p1 = base + (-hw) * x_dir + hh * y_dir
        p2 = base + hw * x_dir + hh * y_dir
    else:
        p1 = base
        p2 = base + x_dir * 0.1

    return p1, p2


# ============================================================================
# HTML生成
# ============================================================================
KIND_CSS = {
    "額縁": "#ff4444",
    "額縁受け": "#44ff44",
    "T-bar": "#4488ff",
    "霧除け": "#ffaa00",
    "木口": "#ff44ff",
}


def generate_html(model_name, meshes, frames, fixtures_info, type_totals):
    """Three.js 3Dビューア付きHTMLを生成"""

    fixture_count = len(fixtures_info)
    meshes_json = json.dumps(meshes, separators=(",", ":"))
    frames_json = json.dumps(frames, separators=(",", ":"))
    totals_json = json.dumps(type_totals, separators=(",", ":"), ensure_ascii=False)

    kind_order = ["額縁", "額縁受け", "T-bar", "霧除け", "木口"]
    kind_rows = ""
    total_all = 0
    for k in kind_order:
        v = type_totals.get(k, 0)
        if v == 0:
            continue
        total_all += v
        kind_rows += f'<div class="kind-row"><span class="kind-dot" style="background:{KIND_CSS[k]}"></span><span class="label">{k}</span><span class="val" style="margin-left:auto">{v:.2f}m</span></div>\n'

    return f'''<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>額縁拾い3D {model_name}</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{background:#1a1a2e;overflow:hidden;font-family:'Segoe UI',sans-serif}}
#cv{{display:block}}
#info{{position:absolute;top:10px;left:10px;color:#fff;background:rgba(0,0,0,0.75);
padding:12px 16px;border-radius:8px;font-size:13px;max-width:320px;z-index:10}}
#info h3{{margin-bottom:8px;font-size:15px;border-bottom:1px solid #555;padding-bottom:4px}}
.row{{display:flex;justify-content:space-between;padding:2px 0}}
.row .label{{color:#aaa}}.row .val{{font-weight:bold}}
.kind-row{{display:flex;align-items:center;gap:6px;padding:2px 0}}
.kind-dot{{width:10px;height:10px;border-radius:50%;display:inline-block}}
#controls{{position:absolute;top:10px;right:10px;display:flex;flex-direction:column;gap:4px;z-index:10}}
#controls button{{padding:6px 12px;border:none;border-radius:4px;cursor:pointer;
font-size:12px;color:#fff;opacity:0.85}}
#controls button:hover{{opacity:1}}
#controls button.active{{box-shadow:0 0 6px rgba(255,255,255,0.4)}}
#controls button.inactive{{opacity:0.4}}
#tooltip{{position:absolute;display:none;background:rgba(0,0,0,0.85);color:#fff;
padding:6px 10px;border-radius:4px;font-size:12px;pointer-events:none;z-index:20}}
</style></head><body>
<canvas id="cv"></canvas>
<div id="info">
<h3>額縁拾い3D {model_name}</h3>
<div class="row"><span class="label">対象建具:</span><span class="val">{fixture_count}箇所</span></div>
<div style="margin-top:6px;border-top:1px solid #444;padding-top:6px"></div>
{kind_rows}
<div class="row" style="margin-top:4px;border-top:1px solid #444;padding-top:4px">
<span class="label">合計</span><span class="val">{total_all:.2f}m</span></div>
</div>
<div id="controls"></div>
<div id="tooltip"></div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/controls/OrbitControls.js"></script>
<script>
const MESHES={meshes_json};
const FRAMES={frames_json};
const TYPE_TOTALS={totals_json};
const FRAME_COLORS={{"額縁":0xff4444,"額縁受け":0x44ff44,"T-bar":0x4488ff,"霧除け":0xffaa00,"木口":0xff44ff}};
const FRAME_CSS={{"額縁":"#ff4444","額縁受け":"#44ff44","T-bar":"#4488ff","霧除け":"#ffaa00","木口":"#ff44ff"}};

const MESH_COLORS={{
  "壁":0xcc7733,"ドア":0x5588cc,"窓":0x5588cc,"梁":0xddaa22,
  "柱":0xcc8844,"1F床":0xddbb77,"手摺":0x888888,"屋根":0xcc3333,
  "天窓":0x66aaff,"部材":0x999999,"階段":0xbbaa88
}};
const MESH_CSS={{
  "壁":"#cc7733","ドア":"#5588cc","窓":"#5588cc","梁":"#ddaa22",
  "柱":"#cc8844","1F床":"#ddbb77","手摺":"#888888","屋根":"#cc3333",
  "天窓":"#66aaff","部材":"#999999","階段":"#bbaa88"
}};
// 屋根はデフォルト非表示
const HIDDEN_CATS={{"屋根":true}};

const W=window.innerWidth,H=window.innerHeight;
const renderer=new THREE.WebGLRenderer({{canvas:document.getElementById('cv'),antialias:true}});
renderer.setSize(W,H);renderer.setPixelRatio(window.devicePixelRatio);
const scene=new THREE.Scene();scene.background=new THREE.Color(0x1a1a2e);
const camera=new THREE.PerspectiveCamera(50,W/H,0.1,500);
const controls=new THREE.OrbitControls(camera,renderer.domElement);
controls.enableDamping=true;controls.dampingFactor=0.08;
scene.add(new THREE.AmbientLight(0xffffff,0.5));
const dl=new THREE.DirectionalLight(0xffffff,0.7);dl.position.set(10,20,-10);scene.add(dl);

const buildingGroup=new THREE.Group();
const catGroups={{}};
const bbox=new THREE.Box3();
MESHES.forEach(m=>{{
  const g=new THREE.BufferGeometry();
  g.setAttribute('position',new THREE.Float32BufferAttribute(m.verts,3));
  g.setIndex(m.faces);g.computeVertexNormals();
  const color=MESH_COLORS[m.cat]||0x888888;
  const mat=new THREE.MeshLambertMaterial({{color,transparent:true,opacity:0.25,side:THREE.DoubleSide}});
  const mesh=new THREE.Mesh(g,mat);
  if(!catGroups[m.cat]){{catGroups[m.cat]=new THREE.Group();buildingGroup.add(catGroups[m.cat]);}}
  catGroups[m.cat].add(mesh);
  bbox.expandByObject(mesh);
}});
scene.add(buildingGroup);
// 屋根をデフォルト非表示
Object.keys(HIDDEN_CATS).forEach(c=>{{if(catGroups[c])catGroups[c].visible=false;}});
const center=new THREE.Vector3();bbox.getCenter(center);
const size=bbox.getSize(new THREE.Vector3());
const maxDim=Math.max(size.x,size.y,size.z);
camera.position.set(center.x+maxDim*0.8,center.y+maxDim*0.6,center.z+maxDim*0.8);
controls.target.copy(center);camera.lookAt(center);
const grid=new THREE.GridHelper(maxDim*1.5,20,0x444444,0x333333);
grid.position.copy(center);grid.position.y=bbox.min.y;scene.add(grid);

const frameGroups={{}};
const kindOrder=["額縁","額縁受け","T-bar","霧除け","木口"];
kindOrder.forEach(k=>{{frameGroups[k]=new THREE.Group();scene.add(frameGroups[k])}});

FRAMES.forEach(f=>{{
  const color=FRAME_COLORS[f.kind]||0xffffff;
  const g=new THREE.BufferGeometry();
  const pts=f.points;
  g.setAttribute('position',new THREE.Float32BufferAttribute([pts[0][0],pts[0][1],pts[0][2],pts[1][0],pts[1][1],pts[1][2]],3));
  const mat=new THREE.LineBasicMaterial({{color,linewidth:2}});
  const line=new THREE.LineSegments(g,mat);
  line.userData={{kind:f.kind,length:f.length,fixture:f.fixture}};
  if(frameGroups[f.kind])frameGroups[f.kind].add(line);
}});

let totalAll=0;
kindOrder.forEach(k=>{{
  if(!TYPE_TOTALS[k]) return;
  totalAll+=TYPE_TOTALS[k];
}});

const ctrlDiv=document.getElementById('controls');

// 建物表示トグル
let showB=true;
const btnB=document.createElement('button');
btnB.style.background='#666';btnB.textContent='建物表示';btnB.className='active';
btnB.onclick=()=>{{showB=!showB;buildingGroup.visible=showB;btnB.className=showB?'active':'inactive';}};
ctrlDiv.appendChild(btnB);

// 屋根トグル（デフォルト非表示）
if(catGroups['屋根']){{
  const btnR=document.createElement('button');
  btnR.style.background='#cc3333';btnR.textContent='屋根表示';btnR.className='inactive';
  btnR.onclick=()=>{{
    const g=catGroups['屋根'];g.visible=!g.visible;
    btnR.className=g.visible?'active':'inactive';
  }};
  ctrlDiv.appendChild(btnR);
}}

// 額縁種別トグル
kindOrder.forEach(k=>{{
  if(!TYPE_TOTALS[k]) return;
  const btn=document.createElement('button');
  btn.style.background=FRAME_CSS[k];btn.textContent=k;
  btn.className='active';
  btn.onclick=()=>{{
    const g=frameGroups[k];g.visible=!g.visible;
    btn.className=g.visible?'active':'inactive';
  }};
  ctrlDiv.appendChild(btn);
}});

// カテゴリ凡例（クリックで個別ON/OFF）
const cats=[...new Set(MESHES.map(m=>m.cat))];
if(cats.length>0){{
  const sep=document.createElement('div');sep.style.borderTop='1px solid rgba(255,255,255,0.3)';
  sep.style.margin='4px 0';sep.style.width='100%';ctrlDiv.appendChild(sep);
  cats.forEach(cat=>{{
    const d=document.createElement('button');
    d.style.background=MESH_CSS[cat]||'#888';d.textContent=cat;
    d.className=HIDDEN_CATS[cat]?'inactive':'active';d.style.fontSize='11px';
    d.onclick=()=>{{if(catGroups[cat]){{catGroups[cat].visible=!catGroups[cat].visible;d.className=catGroups[cat].visible?'active':'inactive';}}}};
    ctrlDiv.appendChild(d);
  }});
}}

const tooltip=document.getElementById('tooltip');
const raycaster=new THREE.Raycaster();raycaster.params.Line={{threshold:0.05}};
const mouse=new THREE.Vector2();
renderer.domElement.addEventListener('mousemove',e=>{{
  mouse.x=(e.clientX/W)*2-1;mouse.y=-(e.clientY/H)*2+1;
  raycaster.setFromCamera(mouse,camera);
  let all=[];kindOrder.forEach(k=>{{if(frameGroups[k])frameGroups[k].children.forEach(c=>all.push(c))}});
  const hits=raycaster.intersectObjects(all);
  if(hits.length>0){{
    const d=hits[0].object.userData;
    tooltip.style.display='block';
    tooltip.style.left=(e.clientX+12)+'px';tooltip.style.top=(e.clientY+12)+'px';
    tooltip.innerHTML=d.fixture+'<br>'+d.kind+': '+d.length+'mm';
  }}else{{tooltip.style.display='none'}}
}});

window.addEventListener('resize',()=>{{
  const w=window.innerWidth,h=window.innerHeight;
  camera.aspect=w/h;camera.updateProjectionMatrix();renderer.setSize(w,h);
}});

(function a(){{requestAnimationFrame(a);controls.update();renderer.render(scene,camera)}})();
</script></body></html>'''


# ============================================================================
# メイン処理
# ============================================================================
def main():
    if len(sys.argv) < 3:
        print(f"Usage: python {sys.argv[0]} input.ifc output.html")
        sys.exit(1)

    ifc_path = sys.argv[1]
    output_path = sys.argv[2]
    model_name = os.path.splitext(os.path.basename(ifc_path))[0]

    print(f"Processing: {ifc_path}")
    print(f"Output: {output_path}")

    ifc_file = ifcopenshell.open(ifc_path)

    print("Extracting meshes...")
    meshes = extract_meshes(ifc_file)
    print(f"  {len(meshes)} mesh elements")

    print("Detecting fixtures and calculating frames...")
    fixtures_info, frames = detect_fixtures_and_frames(ifc_file)
    print(f"  {len(fixtures_info)} fixtures → {len(frames)} frame lines")

    type_totals = defaultdict(float)
    for f in frames:
        type_totals[f["kind"]] += f["length"] / 1000.0
    type_totals = {k: round(v, 2) for k, v in type_totals.items()}

    print("Generating HTML viewer...")
    html = generate_html(model_name, meshes, frames, fixtures_info, type_totals)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\nHTML viewer saved to {output_path}")
    print("\nFrame totals:")
    total = 0
    for k in ["額縁", "額縁受け", "T-bar", "霧除け", "木口"]:
        v = type_totals.get(k, 0)
        if v > 0:
            print(f"  {k}: {v:.2f}m")
            total += v
    print(f"  合計: {total:.2f}m")
    print("\nSuccess!")


if __name__ == "__main__":
    main()
