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
             "front_extras": ["木口", "T-bar", "額縁受け"]},
    "SL":   {"front_frame": "4", "back_frame": None,
             "front_extras": []},
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
    "SL":   (0, 0),
}

# ============================================================================
# 規格寸法マスタ（CAD入力ミス検出用）
# マーカー名 → (W, H) のログ壁開口寸法 (mm)
# ============================================================================
SIZE_MASTER = {
    "VMW-1": (450, 597), "VMW-1(FIX)": (450, 597), "VMW-1(OBS)": (450, 597),
    "VMW-2": (889, 533), "VMW-3": (533, 889), "VMW-3(FIX)": (533, 889),
    "VMW-3(OBS)": (533, 889), "VMW-3+VMW-4": (1600, 889),
    "VMW-4": (1067, 889), "VMW-4a": (1422, 889),
    "VMW-5": (533, 1041), "VMW-6": (1067, 1041), "VMW-6a": (1422, 1041),
    "VMW-11": (533, 1194), "VMW-13": (1600, 1194), "VMW-13b": (1955, 1194),
    "VMSD": (1505, 2019), "VMSD-1": (1505, 2019),
    "VMSD 1505x2019": (1505, 2019), "VMSD 1810x2019": (1810, 2019),
    "VMSD 2419x2019": (2419, 2019),
}

# 規格寸法 → 品番 の逆引き
_SIZE_TO_CODE = {}
for _code, (_sw, _sh) in SIZE_MASTER.items():
    _key = f"{_sw}x{_sh}"
    if _key not in _SIZE_TO_CODE:
        _SIZE_TO_CODE[_key] = []
    if _code not in _SIZE_TO_CODE[_key]:
        _SIZE_TO_CODE[_key].append(_code)


def validate_dimensions(marker: str, frame_w: float, frame_h: float, tolerance: int = 30):
    """マーカー名と開口寸法を規格値と照合。不一致なら辞書を返す。"""
    # マーカーからベースコードを推定 (例: "AW3 Vmw" → VMW, label "NV Door1" → NV)
    base = marker.split()[0].upper() if marker else ""
    # SIZE_MASTERのキーで一致するものを探す
    spec = None
    matched_code = None
    for code in SIZE_MASTER:
        if code.upper() == base or base.startswith(code.upper()):
            spec = SIZE_MASTER[code]
            matched_code = code
            break
    # TypeDefベースのftype名でも検索 (例: VMW-3)
    if spec is None:
        for code in SIZE_MASTER:
            if code.upper() in marker.upper():
                spec = SIZE_MASTER[code]
                matched_code = code
                break
    if spec is None:
        return None  # 規格マスタにない品番

    spec_w, spec_h = spec
    fw, fh = round(frame_w), round(frame_h)
    if abs(fw - spec_w) <= tolerance and abs(fh - spec_h) <= tolerance:
        return None  # OK

    # 同じベースコードの別サイズバリエーションもチェック
    # 例: "VMSD" で最初に (1505,2019) にマッチしたが、実際は "VMSD 1810x2019" かもしれない
    base_upper = (matched_code or base).split()[0].upper()
    for code, (sw, sh) in SIZE_MASTER.items():
        if code.upper().startswith(base_upper) and code != matched_code:
            if abs(fw - sw) <= tolerance and abs(fh - sh) <= tolerance:
                return None  # 同品番の別サイズにマッチ → エラーではない

    # 不一致 → 実寸法から正しい品番を逆引き
    actual_key = f"{fw}x{fh}"
    suggestion = _SIZE_TO_CODE.get(actual_key, [])
    if not suggestion:
        # 近似検索
        for key, codes in _SIZE_TO_CODE.items():
            kw, kh = map(int, key.split("x"))
            if abs(kw - fw) <= tolerance and abs(kh - fh) <= tolerance:
                suggestion = codes
                break

    return {
        "marker": matched_code or base,
        "spec_w": spec_w, "spec_h": spec_h,
        "actual_w": fw, "actual_h": fh,
        "suggestion": [c for c in suggestion if c.upper() != (matched_code or base).upper()],
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
    "JD":  {"front_frame": "3", "back_frame": "3",
            "front_extras": ["木口"]},  # 間仕切りJDはT-bar/額縁受け不要
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
    # open_ad: 開放部（建具リーフなし）— ただしSD*（実ドア）は除外しない
    if td == "OPEN_AD" or (nm.startswith("OPEN") and "DOOR" not in td):
        if nm.startswith("SD"):
            return "JD"  # SD8等の実ドアはJD（木口あり）
        else:
            return None

    # --- 天窓(スカイライト) ---
    if "スカイライト" in typedef or "SKYLIGHT" in td or nm.startswith("SL"):
        return "SL"

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
def generate_frame_pieces(ftype: str, w_mm: float, h_mm: float, wall_type: str = "log",
                          ifc_w: float = None, ifc_h: float = None):
    """
    額縁タイプと開口寸法(mm)から、部材リスト(kind, piece_name, length_mm, side)を返す。
    side: "front" or "back"
    wall_type: "log" (ログ壁) or "partition" (間仕切り壁) — 間仕切りは簡素化
    ifc_w, ifc_h: IFC原寸法(mm)。T-bar等はオフセット前の開口寸法を使用。
    """
    # T-bar/木口はIFC開口寸法（オフセット前）を使用
    if ifc_w is None:
        ifc_w = w_mm
    if ifc_h is None:
        ifc_h = h_mm
    # 間仕切り壁で専用ルールがある場合はそちらを使用
    if wall_type == "partition" and ftype in PARTITION_FRAME_RULES:
        rule = PARTITION_FRAME_RULES[ftype]
    else:
        rule = FRAME_RULES.get(ftype, FRAME_RULES["JD"])
    pieces = []

    # --- 表面(front) ---
    # 3Dビューア用: すべてIFC寸法(ifc_w/ifc_h)ベースで統一
    if rule["front_frame"] == "4":
        pieces.append(("額縁", "上", ifc_w + 340, "front"))
        pieces.append(("額縁", "左", ifc_h + 170, "front"))
        pieces.append(("額縁", "右", ifc_h + 170, "front"))
        pieces.append(("額縁", "下", ifc_w, "front"))
    else:
        pieces.append(("額縁", "上", ifc_w + 340, "front"))
        pieces.append(("額縁", "左", ifc_h + 170, "front"))
        pieces.append(("額縁", "右", ifc_h + 170, "front"))

    for extra in rule["front_extras"]:
        if extra == "額縁受け":
            pieces.append(("額縁受け", "横", ifc_w, "front"))
        elif extra == "T-bar":
            pieces.append(("T-bar", "縦1", ifc_h, "front"))
            pieces.append(("T-bar", "縦2", ifc_h, "front"))
        elif extra == "霧除け":
            pieces.append(("霧除け", "横", ifc_w + 340, "front"))
        elif extra == "木口":
            pieces.append(("木口", "上", ifc_w, "front"))
            pieces.append(("木口", "左", ifc_h, "front"))
            pieces.append(("木口", "右", ifc_h, "front"))

    # --- 裏面(back) ---
    if rule["back_frame"] is None:
        pass  # 天窓等: 片面のみ
    elif rule["back_frame"] == "4":
        pieces.append(("額縁", "上", ifc_w + 340, "back"))
        pieces.append(("額縁", "左", ifc_h + 170, "back"))
        pieces.append(("額縁", "右", ifc_h + 170, "back"))
        pieces.append(("額縁", "下", ifc_w, "back"))
    else:
        pieces.append(("額縁", "上", ifc_w + 340, "back"))
        pieces.append(("額縁", "左", ifc_h + 170, "back"))
        pieces.append(("額縁", "右", ifc_h + 170, "back"))

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
        "IfcSlab": "床", "IfcRailing": "手摺",
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
            # IfcSlab: PredefinedType=ROOF → 屋根, それ以外 → 床
            if ifc_type == "IfcSlab":
                pt = getattr(element, "PredefinedType", None) or ""
                if pt == "ROOF":
                    cat = "屋根"
                else:
                    cat = "床"
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
def get_skylight_placement(element):
    """天窓の配置原点と方向ベクトルを取得。
    天窓はIfcRoof上に配置されるため、IFC配置行列の原点がフロアレベルにある。
    ワールド座標メッシュから実際の位置・勾配を算出する。
    Returns: (origin, width_dir, height_dir, depth_dir) すべてIFC座標系
    """
    # 配置行列からwidth_dirを取得（幅方向は棟と平行で正しい）
    settings_l = ifcopenshell.geom.settings()
    settings_l.set(settings_l.USE_WORLD_COORDS, False)
    shape_l = ifcopenshell.geom.create_shape(settings_l, element)
    mat = list(shape_l.transformation.matrix)
    m44 = np.array(mat).reshape(4, 4).T
    width_dir = m44[:3, 0]
    width_dir = width_dir / (np.linalg.norm(width_dir) + 1e-12)

    # ワールド座標メッシュから実際の位置を取得
    settings_w = ifcopenshell.geom.settings()
    settings_w.set(settings_w.USE_WORLD_COORDS, True)
    shape_w = ifcopenshell.geom.create_shape(settings_w, element)
    vf = shape_w.geometry.verts
    pts = np.array([[vf[i], vf[i+1], vf[i+2]]
                     for i in range(0, len(vf), 3)])

    bb_center = pts.mean(axis=0)

    # 勾配方向(height_dir): width方向を除去し、残りの最大分散方向
    rel = pts - bb_center
    w_proj = np.dot(rel, width_dir).reshape(-1, 1) * width_dir
    perp = rel - w_proj

    cov = np.cov(perp.T)
    eigvals, eigvecs = np.linalg.eigh(cov)
    height_dir = eigvecs[:, np.argmax(eigvals)]
    # Z成分が正（上向き勾配）になるよう統一
    if height_dir[2] < 0:
        height_dir = -height_dir
    height_dir = height_dir / (np.linalg.norm(height_dir) + 1e-12)

    # depth_dir: width × height の外積（屋根面の法線方向）
    depth_dir = np.cross(width_dir, height_dir)
    depth_dir = depth_dir / (np.linalg.norm(depth_dir) + 1e-12)

    # 基点 = BBの左下手前（width最小, height最小, depth最小）
    proj_w = np.dot(rel, width_dir)
    proj_h = np.dot(rel, height_dir)
    proj_d = np.dot(rel, depth_dir)

    origin = (bb_center
              + proj_w.min() * width_dir
              + proj_h.min() * height_dir
              + proj_d.min() * depth_dir)

    return origin, width_dir, height_dir, depth_dir


def get_element_placement_and_center(element):
    """建具のIFC配置原点と方向ベクトルを取得。
    基点 = IFC配置行列のtranslation（建具の配置原点）。
    Returns: (origin, width_dir, height_dir, depth_dir) すべてIFC座標系
    """
    settings_local = ifcopenshell.geom.settings()
    settings_local.set(settings_local.USE_WORLD_COORDS, False)

    # ローカル→ワールド変換行列から方向ベクトルと配置原点を取得
    shape_l = ifcopenshell.geom.create_shape(settings_local, element)
    mat = list(shape_l.transformation.matrix)
    m44 = np.array(mat).reshape(4, 4).T
    # IFC建具配置: col0=幅方向, col1=奥行き方向, col2=高さ方向, col3=配置原点
    width_dir = m44[:3, 0]
    depth_dir = m44[:3, 1]
    height_dir = m44[:3, 2]
    origin = m44[:3, 3]  # IFC配置原点（ワールド座標）

    width_dir = width_dir / (np.linalg.norm(width_dir) + 1e-12)
    height_dir = height_dir / (np.linalg.norm(height_dir) + 1e-12)
    depth_dir = depth_dir / (np.linalg.norm(depth_dir) + 1e-12)

    return origin, width_dir, height_dir, depth_dir


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
                if ftype == "SL":
                    center, width_dir, height_dir, depth_dir = \
                        get_skylight_placement(element)
                else:
                    center, width_dir, height_dir, depth_dir = \
                        get_element_placement_and_center(element)
            except Exception:
                continue

            w_m = ifc_w / 1000.0
            h_m = ifc_h / 1000.0

            # 規格寸法チェック（SIZE_MASTERはIFC寸法基準なのでオフセット前で照合）
            dim_error = validate_dimensions(label, ifc_w, ifc_h)

            fixtures_info.append({
                "label": label,
                "ftype": ftype,
                "w_mm": ifc_w,
                "h_mm": ifc_h,
                "frame_w": frame_w,
                "frame_h": frame_h,
                "dim_error": dim_error,
            })

            pieces = generate_frame_pieces(ftype, frame_w, frame_h, wall_type,
                                          ifc_w=ifc_w, ifc_h=ifc_h)

            # depth_dir = 壁面の法線方向（表面から外部に向かう方向）
            # 額縁の裏表間隔 = 120mm
            # origin=左下手前、depth_dir=手前→奥方向
            # FRAME_RULES: front(表)=外壁面、back(裏)=室内面
            front_offset = depth_dir * 0.12    # 表面（外壁側、120mm奥）
            back_offset = depth_dir * 0.0      # 裏面（室内側、原点側）

            # デバッグ用: 基点を球体で表示
            o = center  # origin = 左下手前
            all_frames.append({
                "kind": "基点",
                "points": [
                    [round(o[0], 3), round(o[2], 3), round(-o[1], 3)],
                    [round(o[0], 3), round(o[2], 3), round(-o[1], 3)]
                ],
                "length": 0,
                "fixture": label,
                "debug_origin": True,
                "dirs": {
                    "width": [round(width_dir[0], 3), round(width_dir[2], 3), round(-width_dir[1], 3)],
                    "height": [round(height_dir[0], 3), round(height_dir[2], 3), round(-height_dir[1], 3)],
                    "depth": [round(depth_dir[0], 3), round(depth_dir[2], 3), round(-depth_dir[1], 3)],
                },
            })

            for kind, piece_name, length_mm, side in pieces:
                # 額縁は裏表(front/back)で位置が分かれる
                # T-bar/木口/額縁受け/霧除けはKIND_DEPTHで絶対位置指定（side_offset=0）
                if kind == "額縁":
                    side_offset = front_offset if side == "front" else back_offset
                else:
                    side_offset = depth_dir * 0.0  # KIND_DEPTHが絶対位置を指定

                # center=左下手前, width_dir=幅方向, height_dir=高さ方向
                p1, p2 = compute_frame_line_3d(
                    kind, piece_name,
                    center, width_dir, height_dir, depth_dir,
                    w_m, h_m, length_mm, side_offset
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


def compute_frame_line_3d(kind, piece_name, origin, x_dir, y_dir, depth_dir,
                          w_m, h_m, length_mm, side_offset):
    """額縁部材の3Dライン座標を計算（実寸法 + 部材別オフセット）

    すべてIFC寸法ベースで統一。各部材を実際の長さで描画し、
    部材種別に応じて壁面内で位置をずらす。

    Args:
        kind: 部材種別（"額縁","霧除け","木口","T-bar","額縁受け"）
        piece_name: パート名（"上","下","左","右","横","縦1","縦2"）
        origin: 建具中心座標 (IFC座標系)
        x_dir: 幅方向ベクトル
        y_dir: 高さ方向ベクトル
        depth_dir: 壁面法線方向ベクトル（表面から外部に向かう方向）
        w_m: 開口幅 (m, IFC寸法)
        h_m: 開口高さ (m, IFC寸法)
        length_mm: 部材の実長さ (mm)
        side_offset: 表裏オフセットベクトル（front/backの位置ずれ）
    """
    # すべてIFC寸法ベースで統一
    # origin = 建具の左下手前 → (0,0) = 左下、+x_dir = 右、+y_dir = 上
    W = w_m     # IFC開口幅 (m)
    H = h_m     # IFC開口高さ (m)
    L = length_mm / 1000.0  # 部材実長 (m)

    # --- 部材種別ごとのdepth方向オフセット ---
    # 額縁: side_offsetで裏表が決まる(front=0.12外壁, back=0.0室内)
    # その他: side_offset=0、KIND_DEPTHが絶対位置
    KIND_DEPTH = {
        "額縁":     0.0,      # side_offsetで裏表が決まる
        "霧除け":   0.14,     # 外壁面(0.12)のさらに外側
        "木口":     0.04,     # 裏表の中心寄り（室内側）
        "T-bar":    0.06,     # 裏表の中心
        "額縁受け": 0.08,     # 裏表の中心寄り（外壁側）
    }
    KIND_Y_SHIFT = {
        "額縁受け": 0.03,   # 上部に30mm
    }

    d_off = KIND_DEPTH.get(kind, 0.0)
    y_shift = KIND_Y_SHIFT.get(kind, 0.0)
    base = origin + side_offset + depth_dir * d_off + y_dir * y_shift

    # --- 左下基点: origin=(左,下) → 右方向=+x_dir, 上方向=+y_dir ---
    cx = W / 2   # 開口幅の中心(左端からの距離)
    cy = H / 2   # 開口高さの中心(下端からの距離)

    if kind == "額縁" and piece_name == "上":
        p1 = base + (cx - L/2) * x_dir + H * y_dir
        p2 = base + (cx + L/2) * x_dir + H * y_dir
    elif kind == "額縁" and piece_name == "下":
        p1 = base + (cx - L/2) * x_dir
        p2 = base + (cx + L/2) * x_dir
    elif kind == "額縁" and piece_name == "左":
        # 上端=H固定、下にL分伸びる
        p1 = base + (H - L) * y_dir
        p2 = base + H * y_dir
    elif kind == "額縁" and piece_name == "右":
        p1 = base + W * x_dir + (H - L) * y_dir
        p2 = base + W * x_dir + H * y_dir

    elif kind == "霧除け":
        p1 = base + (cx - L/2) * x_dir + H * y_dir
        p2 = base + (cx + L/2) * x_dir + H * y_dir

    elif kind == "額縁受け":
        p1 = base + (cx - L/2) * x_dir + H * y_dir
        p2 = base + (cx + L/2) * x_dir + H * y_dir

    elif kind == "T-bar" and piece_name == "縦1":
        # T-bar左: 左端(x=0)位置
        p1 = base + (cy - L/2) * y_dir
        p2 = base + (cy + L/2) * y_dir
    elif kind == "T-bar" and piece_name == "縦2":
        # T-bar右: 右端(x=W)位置
        p1 = base + W * x_dir + (cy - L/2) * y_dir
        p2 = base + W * x_dir + (cy + L/2) * y_dir

    elif kind == "木口" and piece_name == "上":
        p1 = base + (cx - L/2) * x_dir + H * y_dir
        p2 = base + (cx + L/2) * x_dir + H * y_dir
    elif kind == "木口" and piece_name == "左":
        p1 = base + (cy - L/2) * y_dir
        p2 = base + (cy + L/2) * y_dir
    elif kind == "木口" and piece_name == "右":
        p1 = base + W * x_dir + (cy - L/2) * y_dir
        p2 = base + W * x_dir + (cy + L/2) * y_dir

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


def generate_html(model_name, meshes, frames, fixtures_info, type_totals, dim_errors=None):
    """Three.js 3Dビューア付きHTMLを生成（廻り縁・巾木ビューアと統一UI）"""

    fixture_count = len(fixtures_info)
    meshes_json = json.dumps(meshes, separators=(",", ":"))
    frames_json = json.dumps(frames, separators=(",", ":"))
    totals_json = json.dumps(type_totals, separators=(",", ":"), ensure_ascii=False)

    # 寸法エラー警告パネル
    dim_warn_html = ""
    if dim_errors:
        rows = ""
        for fi in dim_errors:
            e = fi["dim_error"]
            sug = ", ".join(e["suggestion"]) if e["suggestion"] else "該当なし"
            rows += (f'<tr><td style="font-weight:700;color:#ff4444">{e["marker"]}</td>'
                     f'<td>{e["spec_w"]}</td><td>{e["spec_h"]}</td>'
                     f'<td style="color:#ff4444;font-weight:700">{e["actual_w"]}</td>'
                     f'<td style="color:#ff4444;font-weight:700">{e["actual_h"]}</td>'
                     f'<td style="font-weight:700">{sug}</td></tr>')
        dim_warn_html = (
            f'<div id="dim-warn" style="position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);'
            f'z-index:300;background:rgba(30,0,0,0.95);border:2px solid #ff4444;border-radius:12px;'
            f'padding:20px 24px;color:#fff;font-size:13px;max-width:600px;backdrop-filter:blur(4px)">'
            f'<div style="font-size:16px;font-weight:700;color:#ff4444;margin-bottom:10px">'
            f'⚠ CAD寸法エラー検出 ({len(dim_errors)}件)</div>'
            f'<table style="width:100%;border-collapse:collapse;color:#eee">'
            f'<tr style="color:#ff8888;font-size:11px;border-bottom:1px solid #444">'
            f'<th style="text-align:left;padding:4px">品番</th>'
            f'<th>規格W</th><th>規格H</th><th>実W</th><th>実H</th><th>推定正解</th></tr>'
            f'{rows}</table>'
            f'<div style="font-size:11px;color:#ff8888;margin-top:8px">'
            f'マーカー名に対して実際の寸法が規格と一致しません</div>'
            f'<button onclick="document.getElementById(\'dim-warn\').style.display=\'none\'" '
            f'style="margin-top:10px;background:#ff4444;color:#fff;border:none;padding:6px 16px;'
            f'border-radius:4px;cursor:pointer;font-size:13px">閉じる</button></div>'
        )

    return f'''<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>額縁拾い 3Dビューア - {model_name}</title>
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
#frame-info {{ position:fixed; bottom:10px; left:10px; color:#fff; background:rgba(0,0,0,0.85);
  padding:14px 18px; border-radius:8px; font-size:13px; z-index:100; line-height:1.6; }}
#controls {{ position:fixed; bottom:10px; right:10px; z-index:100; }}
#controls button {{ background:rgba(255,255,255,0.15); color:#fff; border:1px solid rgba(255,255,255,0.3);
  padding:8px 14px; border-radius:6px; cursor:pointer; margin:2px; font-size:12px; }}
#controls button:hover {{ background:rgba(255,255,255,0.3); }}
#controls button.active {{ background:rgba(255,100,100,0.5); border-color:#ff6b6b; }}
#tooltip {{ position:fixed; display:none; background:rgba(0,0,0,0.85); color:#fff;
  padding:6px 10px; border-radius:4px; font-size:12px; pointer-events:none; z-index:200; }}
</style>
</head>
<body>
<div id="info">
  <h3>額縁拾い 3Dビューア - {model_name}</h3>
  <div>左ドラッグ: 回転 / 右ドラッグ: 移動 / ホイール: ズーム</div>
  <div style="margin-top:4px;color:#aaa;">対象建具: {fixture_count}箇所</div>
  <div id="sel-info" style="margin-top:6px;color:#aaa;">ホバーで部材情報表示</div>
</div>
<div id="frame-info"></div>
<div id="legend"></div>
<div id="controls"></div>
<div id="tooltip"></div>
{dim_warn_html}

<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/lines/LineSegmentsGeometry.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/lines/LineGeometry.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/lines/LineMaterial.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/lines/LineSegments2.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/lines/Line2.js"></script>
<script>
const MESHES={meshes_json};
const FRAMES={frames_json};
const TYPE_TOTALS={totals_json};
const FRAME_COLORS={{"額縁":0xff4444,"額縁受け":0x44ff44,"T-bar":0x4488ff,"霧除け":0xffaa00,"木口":0xff44ff}};
const FRAME_CSS={{"額縁":"#ff4444","額縁受け":"#44ff44","T-bar":"#4488ff","霧除け":"#ffaa00","木口":"#ff44ff"}};
const MESH_CSS={{
  "壁":"#cc7733","ドア":"#5588cc","窓":"#5588cc","梁":"#ddaa22",
  "柱":"#cc8844","床":"#ddbb77","手摺":"#888888","屋根":"#cc3333",
  "天窓":"#66aaff","部材":"#999999","階段":"#bbaa88"
}};
const HIDDEN_CATS={{"屋根":true}};

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

// カスタムOrbitクラス（タッチ対応 - 廻り縁・巾木ビューアと共通）
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

// 建物メッシュ
const buildingGroup=new THREE.Group();
const catGroups={{}};const bbox=new THREE.Box3();
MESHES.forEach(m=>{{
  const g=new THREE.BufferGeometry();
  g.setAttribute('position',new THREE.Float32BufferAttribute(m.verts,3));
  g.setIndex(m.faces);g.computeVertexNormals();
  const color=new THREE.Color(MESH_CSS[m.cat]||'#999');
  const mat=new THREE.MeshLambertMaterial({{color,transparent:true,opacity:0.2,side:THREE.DoubleSide,depthWrite:false}});
  const mesh=new THREE.Mesh(g,mat);
  if(!catGroups[m.cat]){{catGroups[m.cat]=new THREE.Group();}}
  catGroups[m.cat].add(mesh);
  bbox.expandByObject(mesh);
}});
// 屋根・天窓はbuildingGroupの外（独立制御）
Object.keys(catGroups).forEach(c=>{{
  if(c==='屋根'||c==='天窓'){{scene.add(catGroups[c]);}}
  else{{buildingGroup.add(catGroups[c]);}}
}});
scene.add(buildingGroup);
Object.keys(HIDDEN_CATS).forEach(c=>{{if(catGroups[c])catGroups[c].visible=false;}});

// 額縁ライン
const frameGroups={{}};
const kindOrder=["額縁","額縁受け","T-bar","霧除け","木口"];
kindOrder.forEach(k=>{{frameGroups[k]=new THREE.Group();scene.add(frameGroups[k]);}});
const debugGroup=new THREE.Group();scene.add(debugGroup);
FRAMES.forEach(f=>{{
  // デバッグ: 基点球体と方向矢印
  if(f.debug_origin){{
    const p=f.points[0];
    // 赤い球体（基点）
    const sg=new THREE.SphereGeometry(0.03,8,8);
    const sm=new THREE.MeshBasicMaterial({{color:0xff0000}});
    const sphere=new THREE.Mesh(sg,sm);
    sphere.position.set(p[0],p[1],p[2]);
    debugGroup.add(sphere);
    // 方向矢印: width=赤, height=緑, depth=青（各0.2m）
    if(f.dirs){{
      const len=0.2;
      const dirs=[['width',0xff0000],['height',0x00ff00],['depth',0x0000ff]];
      dirs.forEach(([k,c])=>{{
        const d=f.dirs[k];
        const ag=new THREE.BufferGeometry().setFromPoints([
          new THREE.Vector3(p[0],p[1],p[2]),
          new THREE.Vector3(p[0]+d[0]*len,p[1]+d[1]*len,p[2]+d[2]*len)
        ]);
        const am=new THREE.LineBasicMaterial({{color:c,linewidth:2}});
        debugGroup.add(new THREE.Line(ag,am));
      }});
    }}
    return;
  }}
  const color=FRAME_COLORS[f.kind]||0xffffff;
  const pts=f.points;
  const g=new THREE.LineGeometry();
  g.setPositions([pts[0][0],pts[0][1],pts[0][2],pts[1][0],pts[1][1],pts[1][2]]);
  const mat=new THREE.LineMaterial({{color,linewidth:4,resolution:new THREE.Vector2(innerWidth,innerHeight)}});
  const line=new THREE.Line2(g,mat);
  line.userData={{kind:f.kind,length:f.length,fixture:f.fixture}};
  if(frameGroups[f.kind])frameGroups[f.kind].add(line);
}});

// 集計パネル（左下）
let infoHtml='<b style="font-size:14px;">額縁集計（部材種別）</b><br><br>';
let grandTotal=0;
kindOrder.forEach(k=>{{
  const t=TYPE_TOTALS[k]||0;
  if(t===0) return;
  const cc=FRAME_CSS[k]||"#fff";
  grandTotal+=t;
  infoHtml+=`<span style="color:${{cc}}">━━ ${{k}}: ${{t.toFixed(2)}}m</span><br>`;
}});
infoHtml+=`<br><b style="font-size:15px;">合計: ${{grandTotal.toFixed(2)}}m</b>`;
document.getElementById('frame-info').innerHTML=infoHtml;

// カメラ設定
const center=new THREE.Vector3();bbox.getCenter(center);
const size=bbox.getSize(new THREE.Vector3());const maxDim=Math.max(size.x,size.y,size.z);
camera.position.set(center.x+maxDim*0.8,center.y+maxDim*0.6,center.z+maxDim*0.8);
const controls=new Orbit(camera,renderer.domElement);
controls.t.copy(center);camera.lookAt(center);

function resetCam(){{
  camera.position.set(center.x+maxDim*0.8,center.y+maxDim*0.6,center.z+maxDim*0.8);
  controls.t.copy(center);camera.lookAt(center);
}}

const grid=new THREE.GridHelper(20,40,0x444444,0x333333);grid.position.copy(center);grid.position.y=0;scene.add(grid);

// 凡例パネル（右上）: 額縁種別 + 建物カテゴリ
const leg=document.getElementById('legend');
kindOrder.forEach(k=>{{
  const t=TYPE_TOTALS[k]||0;
  if(t===0) return;
  const cc=FRAME_CSS[k]||"#fff";
  const d=document.createElement('div');
  d.innerHTML='<span class="cb" style="background:'+cc+'"></span>'+k+' ('+t.toFixed(1)+'m)';
  d.style.fontWeight='bold';
  d.onclick=()=>{{if(frameGroups[k]){{frameGroups[k].visible=!frameGroups[k].visible;d.style.opacity=frameGroups[k].visible?1:0.3;}}}};
  leg.appendChild(d);
}});
const sep=document.createElement('div');sep.style.borderTop='1px solid rgba(255,255,255,0.3)';sep.style.margin='6px 0';leg.appendChild(sep);
[...new Set(MESHES.map(m=>m.cat))].forEach(cat=>{{
  const d=document.createElement('div');
  d.innerHTML='<span class="cb" style="background:'+(MESH_CSS[cat]||'#999')+'"></span>'+cat;
  d.style.opacity=HIDDEN_CATS[cat]?0.3:1;
  d.onclick=()=>{{if(catGroups[cat]){{catGroups[cat].visible=!catGroups[cat].visible;d.style.opacity=catGroups[cat].visible?1:0.3;}}}};
  leg.appendChild(d);
}});

// コントロールボタン（右下）
const ctrlDiv=document.getElementById('controls');
kindOrder.forEach(k=>{{
  const t=TYPE_TOTALS[k]||0;
  if(t===0) return;
  const btn=document.createElement('button');
  btn.id='btn-'+k;btn.textContent=k;btn.className='active';
  btn.onclick=()=>{{if(frameGroups[k]){{frameGroups[k].visible=!frameGroups[k].visible;btn.classList.toggle('active',frameGroups[k].visible);
  const legItems=leg.querySelectorAll('div');legItems.forEach(d=>{{if(d.textContent.startsWith(k))d.style.opacity=frameGroups[k].visible?1:0.3;}});}}}};
  ctrlDiv.appendChild(btn);
}});
const btnB=document.createElement('button');btnB.id='btn-building';btnB.textContent='建物表示';btnB.className='active';
btnB.onclick=()=>{{let showB=buildingGroup.visible;showB=!showB;buildingGroup.visible=showB;btnB.classList.toggle('active',showB);}};
ctrlDiv.appendChild(btnB);
const btnR=document.createElement('button');btnR.textContent='リセット';
btnR.onclick=resetCam;ctrlDiv.appendChild(btnR);

// ツールチップ（ホバーで額縁情報表示）
const tooltip=document.getElementById('tooltip');
const raycaster=new THREE.Raycaster();raycaster.params.Line={{threshold:0.05}};
const mouse=new THREE.Vector2();
renderer.domElement.addEventListener('mousemove',e=>{{
  mouse.x=(e.clientX/innerWidth)*2-1;mouse.y=-(e.clientY/innerHeight)*2+1;
  raycaster.setFromCamera(mouse,camera);
  let all=[];kindOrder.forEach(k=>{{if(frameGroups[k])frameGroups[k].children.forEach(c=>all.push(c));}});
  const hits=raycaster.intersectObjects(all);
  if(hits.length>0){{
    const d=hits[0].object.userData;
    tooltip.style.display='block';
    tooltip.style.left=(e.clientX+12)+'px';tooltip.style.top=(e.clientY+12)+'px';
    tooltip.innerHTML=d.fixture+'<br>'+d.kind+': '+d.length+'mm';
    document.getElementById('sel-info').textContent=d.fixture+' / '+d.kind+': '+d.length+'mm';
  }}else{{tooltip.style.display='none';document.getElementById('sel-info').textContent='ホバーで部材情報表示';}}
}});

(function anim(){{requestAnimationFrame(anim);renderer.render(scene,camera);}})();
addEventListener('resize',()=>{{camera.aspect=innerWidth/innerHeight;camera.updateProjectionMatrix();renderer.setSize(innerWidth,innerHeight);scene.traverse(c=>{{if(c.material&&c.material.resolution)c.material.resolution.set(innerWidth,innerHeight);}});}});
</script>
</body>
</html>'''


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

    # 規格寸法チェック
    dim_errors = [fi for fi in fixtures_info if fi.get("dim_error")]
    if dim_errors:
        print(f"\n  ⚠ CAD寸法エラー検出: {len(dim_errors)}件")
        for fi in dim_errors:
            e = fi["dim_error"]
            sug = ", ".join(e["suggestion"]) if e["suggestion"] else "該当なし"
            print(f"    {e['marker']}: 規格={e['spec_w']}×{e['spec_h']} "
                  f"→ 実際={e['actual_w']}×{e['actual_h']} 推定正解={sug}")
        print()

    type_totals = defaultdict(float)
    for f in frames:
        if f.get("debug_origin"):
            continue
        type_totals[f["kind"]] += f["length"] / 1000.0
    type_totals = {k: round(v, 2) for k, v in type_totals.items()}

    print("Generating HTML viewer...")
    html = generate_html(model_name, meshes, frames, fixtures_info, type_totals, dim_errors)

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

    # === JSON集計出力（部材一覧Excel統合用） ===
    json_path = os.path.splitext(output_path)[0] + "_summary.json"
    summary = {
        "tool": "額縁拾い",
        "model": model_name,
        "type_totals": {k: round(v, 2) for k, v in type_totals.items()},
        "grand_total": round(total, 2),
        "dim_errors": [
            fi["dim_error"] for fi in fixtures_info if fi.get("dim_error")
        ],
        "fixtures": [
            {
                "label": fi["label"],
                "ftype": fi["ftype"],
                "w_mm": fi["w_mm"],
                "h_mm": fi["h_mm"],
                "frame_w": fi["frame_w"],
                "frame_h": fi["frame_h"],
            }
            for fi in fixtures_info
        ],
        "lines": [
            {
                "kind": f["kind"],
                "fixture": f["fixture"],
                "length_mm": f["length"],
            }
            for f in frames
            if not f.get("debug_origin")
        ],
    }
    with open(json_path, 'w', encoding='utf-8') as f_json:
        json.dump(summary, f_json, ensure_ascii=False, indent=2)
    print(f"JSON集計: {json_path}")
    print("\nSuccess!")


if __name__ == "__main__":
    main()
