#!/usr/bin/env python3
import ifcopenshell
import ifcopenshell.geom
import ifcopenshell.util.element
import json, re, os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import Counter, defaultdict

IFC_PATH = "/sessions/focused-exciting-ride/mnt/TALO IFC sekisan/ie4d1_2023.ifc"
OUTPUT_PATH = "/sessions/focused-exciting-ride/mnt/TALO IFC sekisan/部材一覧_ie4d1.xlsx"

# キット積算テンプレート（空ならOUTPUT_PATHと同じフォルダの「キット積算・価格計算.xlsx」を自動検索）
KIT_TEMPLATE = ""
KIT_OUTPUT   = ""   # 空なら「キット積算_<モデル名>.xlsx」を自動生成

TYPE_NAMES = {
    "IfcWall": "壁", "IfcWallStandardCase": "壁",
    "IfcSlab": "スラブ", "IfcColumn": "柱", "IfcBeam": "梁",
    "IfcDoor": "ドア", "IfcWindow": "窓", "IfcStair": "階段",
    "IfcRailing": "手摺", "IfcRoof": "屋根",
    "IfcMember": "部材",
    "IfcPlate": "板", "IfcCurtainWall": "カーテンウォール",
    "IfcFooting": "基礎",
}

# 除外するIFCタイプ（積算不要）
SKIP_TYPES = {"IfcBuildingElementProxy", "IfcCovering"}

# 壁のサブ分類（要素名ベース）
WALL_SUB = {
    "log": "ログ壁",
    "majikiri": "間仕切壁",
    "kiso": "基礎",
    "dodai": "土台",
}

# 除外する壁の要素名プレフィクス（断熱壁など積算不要）
SKIP_WALL_PREFIX = {"I"}

def classify_wall(elem_name):
    """壁の要素名からサブ分類を返す"""
    if not elem_name:
        return "壁（その他）"
    for key, label in WALL_SUB.items():
        if elem_name.startswith(key):
            return label
    return "壁（その他）"

# スラブのサブ分類（要素名 + レイヤーベース）
def classify_slab(elem_name, layer_name):
    """スラブの要素名・レイヤーからサブ分類を返す"""
    n = (elem_name or "").lower()
    l = (layer_name or "").lower()
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

def get_majikiri_panel_info(element):
    """間仕切壁のパネル枚数と芯材(Stud)情報を取得。
    材質レイヤーから: パネル(20mm)=面材, 胴縁(95mm)=Stud
    フォールバック: 壁の厚み(Width)で判定 → 135mm=両面(2枚), 115mm=片面(1枚)
    Returns: (panel_count, stud_thickness_mm)
    """
    panel_count = 0
    stud_thick = 0
    detected_by_material = False
    try:
        material = ifcopenshell.util.element.get_material(element)
        if material:
            layers = []
            if material.is_a("IfcMaterialLayerSetUsage"):
                layers = material.ForLayerSet.MaterialLayers
            elif material.is_a("IfcMaterialLayerSet"):
                layers = material.MaterialLayers
            if layers:
                for layer in layers:
                    if not layer.Material:
                        continue
                    name = layer.Material.Name or ""
                    thick = layer.LayerThickness or 0
                    if "パネル" in name or "panel" in name.lower():
                        panel_count += 1
                    elif "胴縁" in name or "stud" in name.lower():
                        stud_thick = thick
                detected_by_material = True
            elif material.is_a("IfcMaterialList"):
                # IfcMaterialListの場合は名前でカウント
                for m in material.Materials:
                    if m and m.Name:
                        if "パネル" in m.Name or "panel" in m.Name.lower():
                            panel_count += 1
                        elif "胴縁" in m.Name or "stud" in m.Name.lower():
                            stud_thick = 95  # default
                detected_by_material = True
    except:
        pass

    # 材質から判定できなかった場合 → 壁の厚み(Width)でフォールバック
    # 135mm = パネル(20)+胴縁(95)+パネル(20) → 両面(2枚)
    # 115mm = 胴縁(95)+パネル(20)            → 片面(1枚)
    if not detected_by_material or panel_count == 0:
        try:
            psets = ifcopenshell.util.element.get_psets(element)
            bq = psets.get('BaseQuantities', {})
            aq = psets.get('ArchiCADQuantities', {})
            width = bq.get('Width') or aq.get('幅') or aq.get('厚さ')
            if isinstance(width, (int, float)) and width > 0:
                w_mm = width if width > 10 else width * 1000  # m→mm変換
                if abs(w_mm - 135) < 5:
                    panel_count = 2
                    stud_thick = 95
                elif abs(w_mm - 115) < 5:
                    panel_count = 1
                    stud_thick = 95
        except:
            pass

    return panel_count, stud_thick

def get_material_info(element):
    try:
        material = ifcopenshell.util.element.get_material(element)
        if material:
            if material.is_a("IfcMaterial"):
                return material.Name or ""
            elif material.is_a("IfcMaterialLayerSetUsage"):
                return " / ".join(l.Material.Name for l in material.ForLayerSet.MaterialLayers if l.Material)
            elif material.is_a("IfcMaterialLayerSet"):
                return " / ".join(l.Material.Name for l in material.MaterialLayers if l.Material)
            elif material.is_a("IfcMaterialList"):
                return " / ".join(m.Name for m in material.Materials if m)
    except:
        pass
    return ""

def get_storey(element):
    try:
        container = ifcopenshell.util.element.get_container(element)
        if container and container.is_a("IfcBuildingStorey"):
            return container.Name or ""
    except:
        pass
    return ""

def extract_key_props(psets, ifc_type):
    """重要なプロパティを抽出"""
    arch = psets.get('ArchiCADProperties', {})
    bq = psets.get('BaseQuantities', {})
    aq = psets.get('ArchiCADQuantities', {})

    marker = arch.get('マーカーテキスト', '')
    nominal = arch.get('公称幅x高さ', '')
    nominal_thick = arch.get('公称幅x高さx厚さ', '')
    wall_struct = arch.get('壁構造', '')
    layer = arch.get('レイヤー', '')
    parent_id = arch.get('親ID', '')

    # AC_Pset_* から建具情報
    maker = ''
    model_no = ''
    kind = ''
    has_gakubuchi = ''
    has_kiryoke = ''
    gakubuchi_w = ''
    gakubuchi_t = ''
    for pn, pp in psets.items():
        if pn.startswith('AC_Pset_') and pn != 'AC_Pset_RenovationAndPhasing':
            maker = pp.get('メーカー', pp.get('Maker', maker))
            model_no = pp.get('型番', pp.get('Model', model_no))
            kind = pp.get('建具種類', kind)
            if '額縁の有無' in pp:
                has_gakubuchi = 'あり' if pp['額縁の有無'] else 'なし'
            if '霧除けの有無' in pp:
                has_kiryoke = 'あり' if pp['霧除けの有無'] else 'なし'
    for pn, pp in psets.items():
        if pn.startswith('AC_Equantity_'):
            gakubuchi_w = pp.get('額縁幅', '')
            gakubuchi_t = pp.get('額縁の厚み', '')

    # BaseQuantities → ArchiCADQuantities フォールバック
    # 壁: BQに GrossSideArea/NetSideArea あり、Area なし
    # 柱/Proxy: BQ自体なし、AQのみ
    # ドア/窓: BQに Area あり
    def pick(bq_keys, aq_keys):
        for k in bq_keys:
            v = bq.get(k)
            if v is not None and v != '':
                return v
        for k in aq_keys:
            v = aq.get(k)
            if v is not None and v != '':
                return v
        return ''

    bq_width = pick(['Width'], ['幅', '厚さ'])
    bq_height = pick(['Height'], ['高さ'])
    bq_depth = pick(['Depth', 'Thickness'], [])
    bq_length = pick(['Length', 'NetLength'], ['長さ(A)', '基準線長さ', '3D長さ'])
    bq_perimeter = pick(['Perimeter'], ['平面図外周'])

    # 面積: 壁はNetSideArea(開口除く壁面面積)、ドア/窓はArea、柱/ProxyはAQ表面積
    bq_area = pick(
        ['Area', 'GrossArea', 'NetArea', 'NetSideArea', 'GrossSideArea'],
        ['表面積'])
    bq_volume = pick(
        ['Volume', 'NetVolume', 'GrossVolume'],
        ['正味体積'])

    gl_height = aq.get('GLからの高度', '')
    # 壁の開口面積（窓+ドア）
    aq_window_area = aq.get('窓面積', '')
    aq_door_area = aq.get('ドア面積', '')
    opening_area = ''
    if isinstance(aq_window_area, (int, float)) and isinstance(aq_door_area, (int, float)):
        opening_area = aq_window_area + aq_door_area
    elif isinstance(aq_window_area, (int, float)):
        opening_area = aq_window_area
    elif isinstance(aq_door_area, (int, float)):
        opening_area = aq_door_area

    return {
        'マーカー': marker,
        'メーカー': maker,
        '型番': model_no,
        '建具種類': kind,
        'レイヤー': layer,
        '親ID': parent_id,
        '壁構造': wall_struct,
        '公称寸法': nominal_thick or nominal,
        '幅(mm)': bq_width,
        '高さ(mm)': bq_height,
        '厚さ/奥行(mm)': bq_depth,
        '長さ(mm)': bq_length,
        '面積(m²)': bq_area,
        '体積(m³)': bq_volume,
        '周長(mm)': bq_perimeter,
        'GL高度(mm)': gl_height,
        '開口面積(m²)': opening_area,
        '額縁': has_gakubuchi,
        '霧除け': has_kiryoke,
        '額縁幅(mm)': gakubuchi_w,
        '額縁厚(mm)': gakubuchi_t,
    }

def main():
    print("IFC読み込み中...")
    ifc_file = ifcopenshell.open(IFC_PATH)
    target_types = [t for t in TYPE_NAMES.keys() if t not in SKIP_TYPES]
    all_elements = []
    seen_ids = set()  # IfcWall/IfcWallStandardCase 重複除去用

    for ifc_type in target_types:
        elements = ifc_file.by_type(ifc_type)
        if not elements:
            continue
        added = 0
        for elem in elements:
            # IfcWallStandardCaseはIfcWallのサブタイプ → 重複除去
            gid = elem.GlobalId
            if gid in seen_ids:
                continue
            seen_ids.add(gid)
            added += 1

            type_name_jp = TYPE_NAMES.get(ifc_type, ifc_type)

            etype = ifcopenshell.util.element.get_type(elem)
            elem_type_name = etype.Name if etype else ""
            material = get_material_info(elem)
            storey = get_storey(elem)
            psets = ifcopenshell.util.element.get_psets(elem)
            kp = extract_key_props(psets, ifc_type)

            # 壁はサブ分類（ログ壁/間仕切壁/基礎/土台）
            if type_name_jp == "壁":
                # 除外対象の壁プレフィクスをスキップ
                ename = elem.Name or ""
                if any(ename.startswith(p) for p in SKIP_WALL_PREFIX):
                    continue
                type_name_jp = classify_wall(ename)

            # スラブはサブ分類（1F床/2F床/テラス/バルコニー/屋根）
            if type_name_jp == "スラブ":
                type_name_jp = classify_slab(elem.Name, kp['レイヤー'])

            # 間仕切壁のパネル情報
            panel_count = ''
            panel_area = ''
            stud_area = ''
            if type_name_jp == "間仕切壁":
                pc, st = get_majikiri_panel_info(elem)
                panel_count = pc if pc > 0 else 0
                # GrossSideArea = 片面の壁面面積
                area_val = kp['面積(m²)']
                if isinstance(area_val, (int, float)) and area_val > 0:
                    stud_area = area_val  # Stud面積 = 片面面積
                    panel_area = area_val * pc  # パネル面積 = 片面面積 × パネル枚数

            row = {
                "部材分類": type_name_jp,
                "要素名": elem.Name or "",
                "マーカー": kp['マーカー'],
                "型式名": elem_type_name,
                "階": storey,
                "材質": material,
                "メーカー": kp['メーカー'],
                "型番": kp['型番'],
                "建具種類": kp['建具種類'],
                "レイヤー": kp['レイヤー'],
                "壁構造": kp['壁構造'],
                "公称寸法": kp['公称寸法'],
                "幅(mm)": kp['幅(mm)'],
                "高さ(mm)": kp['高さ(mm)'],
                "厚さ/奥行(mm)": kp['厚さ/奥行(mm)'],
                "長さ(mm)": kp['長さ(mm)'],
                "面積(m²)": kp['面積(m²)'],
                "体積(m³)": kp['体積(m³)'],
                "パネル枚数": panel_count,
                "パネル面積(m²)": panel_area,
                "Stud面積(m²)": stud_area,
                "周長(mm)": kp['周長(mm)'],
                "GL高度(mm)": kp['GL高度(mm)'],
                "開口面積(m²)": kp['開口面積(m²)'],
                "額縁": kp['額縁'],
                "霧除け": kp['霧除け'],
                "額縁幅(mm)": kp['額縁幅(mm)'],
                "額縁厚(mm)": kp['額縁厚(mm)'],
                "親ID": kp['親ID'],
                "GlobalId": elem.GlobalId,
            }
            all_elements.append(row)
        if added:
            print(f"  {ifc_type}: {added}")

    print(f"\n合計 {len(all_elements)} 部材（重複除去済み）")

    # === Excel出力 ===
    wb = Workbook()
    hfont = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    hfill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    halign = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cfont = Font(name="Arial", size=10)
    border = Border(
        left=Side(style='thin', color='D0D0D0'), right=Side(style='thin', color='D0D0D0'),
        top=Side(style='thin', color='D0D0D0'), bottom=Side(style='thin', color='D0D0D0'))
    alt_fill = PatternFill(start_color="F2F7FC", end_color="F2F7FC", fill_type="solid")
    num_right = Alignment(horizontal="right")

    # ---------- Sheet 1: 部材一覧 ----------
    ws = wb.active
    ws.title = "部材一覧"
    headers = list(all_elements[0].keys())
    num_cols = {"幅(mm)", "高さ(mm)", "厚さ/奥行(mm)", "長さ(mm)", "面積(m²)", "体積(m³)",
                "パネル枚数", "パネル面積(m²)", "Stud面積(m²)",
                "周長(mm)", "GL高度(mm)", "開口面積(m²)", "額縁幅(mm)", "額縁厚(mm)"}

    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font = hfont; c.fill = hfill; c.alignment = halign; c.border = border

    for ri, elem in enumerate(all_elements, 2):
        for ci, key in enumerate(headers, 1):
            val = elem.get(key, "")
            if val == '' or val is None:
                val = None
            c = ws.cell(row=ri, column=ci, value=val)
            c.font = cfont; c.border = border
            if ri % 2 == 0:
                c.fill = alt_fill
            if key in num_cols and isinstance(val, (int, float)):
                c.alignment = num_right
                if "m²" in key:
                    c.number_format = '#,##0.0000'
                elif "m³" in key:
                    c.number_format = '#,##0.000000'
                else:
                    c.number_format = '#,##0.0'

    col_w = {"部材分類":12, "要素名":14, "マーカー":16, "型式名":22, "階":12, "材質":25,
             "メーカー":10, "型番":10, "建具種類":8, "レイヤー":14, "壁構造":20, "公称寸法":20,
             "幅(mm)":10, "高さ(mm)":10, "厚さ/奥行(mm)":12, "長さ(mm)":10,
             "面積(m²)":12, "体積(m³)":14,
             "パネル枚数":8, "パネル面積(m²)":14, "Stud面積(m²)":14,
             "周長(mm)":10, "GL高度(mm)":12,
             "開口面積(m²)":12, "額縁":6, "霧除け":6, "額縁幅(mm)":10, "額縁厚(mm)":10,
             "親ID":12, "GlobalId":36}
    for ci, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(ci)].width = col_w.get(h, 12)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(all_elements)+1}"
    ws.freeze_panes = "A2"

    # ---------- Sheet 2: 建具一覧 ----------
    ws2 = wb.create_sheet("建具一覧")
    fittings = [e for e in all_elements if e["部材分類"] in ("ドア", "窓")]
    f_headers = ["部材分類", "要素名", "マーカー", "建具種類", "メーカー", "型番", "型式名",
                 "幅(mm)", "高さ(mm)", "厚さ/奥行(mm)", "面積(m²)", "体積(m³)",
                 "GL高度(mm)", "開口面積(m²)", "額縁", "霧除け", "額縁幅(mm)", "額縁厚(mm)",
                 "壁構造", "親ID", "階"]

    for ci, h in enumerate(f_headers, 1):
        c = ws2.cell(row=1, column=ci, value=h)
        c.font = hfont; c.fill = hfill; c.alignment = halign; c.border = border

    # マーカーでソート
    fittings.sort(key=lambda x: (x["部材分類"], x.get("マーカー") or "zzz"))
    for ri, elem in enumerate(fittings, 2):
        for ci, key in enumerate(f_headers, 1):
            val = elem.get(key)
            if val == '' or val is None:
                val = None
            c = ws2.cell(row=ri, column=ci, value=val)
            c.font = cfont; c.border = border
            if ri % 2 == 0:
                c.fill = alt_fill
            if key in num_cols and isinstance(val, (int, float)):
                c.alignment = num_right
                if "m²" in key:
                    c.number_format = '#,##0.0000'
                elif "m³" in key:
                    c.number_format = '#,##0.000000'
                else:
                    c.number_format = '#,##0.0'

    f_col_w = {"部材分類":8, "要素名":8, "マーカー":16, "建具種類":8, "メーカー":10, "型番":10,
               "型式名":22, "幅(mm)":10, "高さ(mm)":10, "厚さ/奥行(mm)":12, "面積(m²)":12,
               "体積(m³)":14, "GL高度(mm)":12, "開口面積(m²)":12, "額縁":6, "霧除け":6,
               "額縁幅(mm)":10, "額縁厚(mm)":10, "壁構造":18, "親ID":12, "階":12}
    for ci, h in enumerate(f_headers, 1):
        ws2.column_dimensions[get_column_letter(ci)].width = f_col_w.get(h, 10)
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(f_headers))}{len(fittings)+1}"
    ws2.freeze_panes = "A2"

    # ---------- Sheet 3: 分類別集計 ----------
    # 集計ルール:
    #   ログ壁        → 面積(m²)
    #   間仕切壁      → パネル面積 / Stud面積
    #   基礎          → 長さ(m)
    #   土台          → 長さ(m)
    #   屋根(yane1)   → 型式名ごとに集計。破風・鼻隠しは長さ(m)、屋根本体は面積(m²)
    #   床            → 要素名・型式名ごとに面積(m²)
    #   柱(チャネル柱) → 木口カバー総長さ(m)
    #   梁            → 体積(m³)
    #   ドア・窓      → 集計不要
    #   手摺・階段    → 数量のみ
    ws3 = wb.create_sheet("分類別集計")

    s_headers = ["分類", "詳細", "数量", "集計値", "単位", "備考"]
    for ci, h in enumerate(s_headers, 1):
        c = ws3.cell(row=1, column=ci, value=h)
        c.font = hfont; c.fill = hfill; c.alignment = halign; c.border = border

    summary_rows = []  # (分類, 詳細, 数量, 値, 単位, 備考)

    # --- ログ壁 ---
    log_elems = [e for e in all_elements if e["部材分類"] == "ログ壁"]
    if log_elems:
        total_area = sum(e["面積(m²)"] for e in log_elems if isinstance(e.get("面積(m²)"), (int, float)))
        summary_rows.append(("ログ壁", "", len(log_elems), round(total_area, 2), "m²", ""))

    # --- 間仕切壁 ---
    maji_elems = [e for e in all_elements if e["部材分類"] == "間仕切壁"]
    if maji_elems:
        total_panel = sum(e["パネル面積(m²)"] for e in maji_elems if isinstance(e.get("パネル面積(m²)"), (int, float)))
        total_stud = sum(e["Stud面積(m²)"] for e in maji_elems if isinstance(e.get("Stud面積(m²)"), (int, float)))
        summary_rows.append(("間仕切壁", "パネル面積", len(maji_elems), round(total_panel, 2), "m²", ""))
        summary_rows.append(("", "Stud面積", "", round(total_stud, 2), "m²", ""))

    # --- 基礎 → 長さ ---
    kiso_elems = [e for e in all_elements if e["部材分類"] == "基礎"]
    if kiso_elems:
        total_len = sum(e["長さ(mm)"] for e in kiso_elems if isinstance(e.get("長さ(mm)"), (int, float)))
        summary_rows.append(("基礎", "", len(kiso_elems), round(total_len / 1000, 2), "m", ""))

    # --- 土台 → 長さ ---
    dodai_elems = [e for e in all_elements if e["部材分類"] == "土台"]
    if dodai_elems:
        total_len = sum(e["長さ(mm)"] for e in dodai_elems if isinstance(e.get("長さ(mm)"), (int, float)))
        summary_rows.append(("土台", "", len(dodai_elems), round(total_len / 1000, 2), "m", ""))

    # --- 屋根 → 型式名ごと。破風・鼻隠しは長さ、屋根本体は面積 ---
    yane_elems = [e for e in all_elements if e["部材分類"] == "屋根"]
    if yane_elems:
        # 型式名でグループ化
        yane_by_type = defaultdict(list)
        for e in yane_elems:
            tn = e.get("型式名") or "（不明）"
            yane_by_type[tn].append(e)

        for tn, elems in sorted(yane_by_type.items()):
            is_hafu = "破風" in tn or "鼻隠" in tn
            if is_hafu:
                # 長さ = GrossArea / (Width/1000)
                total_len = 0
                for e in elems:
                    area = e.get("面積(m²)")
                    width = e.get("幅(mm)")
                    if isinstance(area, (int, float)) and isinstance(width, (int, float)) and width > 0:
                        total_len += area / (width / 1000)
                summary_rows.append(("屋根", tn, len(elems), round(total_len, 2), "m", "長さ"))
            else:
                total_area = sum(e["面積(m²)"] for e in elems if isinstance(e.get("面積(m²)"), (int, float)))
                summary_rows.append(("屋根", tn, len(elems), round(total_area, 2), "m²", "面積"))

    # --- 床 → 要素名・型式名ごとに面積 ---
    floor_cats = {"1F床", "2F床", "テラス", "バルコニー"}
    floor_elems = [e for e in all_elements if e["部材分類"] in floor_cats]
    if floor_elems:
        # (要素名, 型式名) でグループ化
        floor_by_key = defaultdict(list)
        for e in floor_elems:
            key = (e.get("要素名") or "", e.get("型式名") or "")
            floor_by_key[key].append(e)

        for (ename, tname), elems in sorted(floor_by_key.items()):
            total_area = sum(e["面積(m²)"] for e in elems if isinstance(e.get("面積(m²)"), (int, float)))
            cat = elems[0]["部材分類"]
            summary_rows.append(("床", f"{ename} / {tname}", len(elems), round(total_area, 2), "m²", cat))

    # --- 柱(チャネル柱) → 木口カバー総長さ ---
    col_elems = [e for e in all_elements if e["部材分類"] == "柱"]
    if col_elems:
        total_h = sum(e["高さ(mm)"] for e in col_elems if isinstance(e.get("高さ(mm)"), (int, float)))
        summary_rows.append(("木口カバー", "チャネル柱", len(col_elems), round(total_h / 1000, 2), "m", "総長さ"))

    # --- 梁 → m³ ---
    beam_elems = [e for e in all_elements if e["部材分類"] == "梁"]
    if beam_elems:
        total_vol = sum(e["体積(m³)"] for e in beam_elems if isinstance(e.get("体積(m³)"), (int, float)))
        summary_rows.append(("梁", "", len(beam_elems), round(total_vol, 4), "m³", ""))

    # --- 手摺・階段 → 数量のみ ---
    for cat in ["手摺", "階段"]:
        cat_elems = [e for e in all_elements if e["部材分類"] == cat]
        if cat_elems:
            summary_rows.append((cat, "", len(cat_elems), "", "", ""))

    # --- その他（スラブその他等） ---
    shown_cats = {"ログ壁", "間仕切壁", "基礎", "土台", "屋根",
                  "1F床", "2F床", "テラス", "バルコニー",
                  "柱", "梁", "ドア", "窓", "手摺", "階段"}
    for e in all_elements:
        cat = e["部材分類"]
        if cat not in shown_cats:
            shown_cats.add(cat)
            cat_elems = [x for x in all_elements if x["部材分類"] == cat]
            total_area = sum(x["面積(m²)"] for x in cat_elems if isinstance(x.get("面積(m²)"), (int, float)))
            summary_rows.append((cat, "", len(cat_elems), round(total_area, 2) if total_area else "", "m²" if total_area else "", ""))

    # 書き出し
    bf = Font(name="Arial", bold=True, size=10)
    cat_fill = PatternFill(start_color="E8EEF7", end_color="E8EEF7", fill_type="solid")
    prev_cat = None
    for ri, (cat, detail, count, val, unit, note) in enumerate(summary_rows, 2):
        is_cat_row = (cat != "" and cat != prev_cat)
        ws3.cell(row=ri, column=1, value=cat).font = bf if is_cat_row else cfont
        ws3.cell(row=ri, column=2, value=detail).font = cfont
        ws3.cell(row=ri, column=3, value=count if count != "" else None).font = cfont
        c4 = ws3.cell(row=ri, column=4, value=val if val != "" else None)
        c4.font = cfont
        if isinstance(val, (int, float)):
            if unit == "m³":
                c4.number_format = '#,##0.0000'
            else:
                c4.number_format = '#,##0.00'
        ws3.cell(row=ri, column=5, value=unit).font = cfont
        ws3.cell(row=ri, column=6, value=note).font = cfont
        for ci in range(1, 7):
            ws3.cell(row=ri, column=ci).border = border
            if is_cat_row:
                ws3.cell(row=ri, column=ci).fill = cat_fill
        if cat:
            prev_cat = cat

    ws3.column_dimensions["A"].width = 14
    ws3.column_dimensions["B"].width = 36
    ws3.column_dimensions["C"].width = 8
    ws3.column_dimensions["D"].width = 14
    ws3.column_dimensions["E"].width = 6
    ws3.column_dimensions["F"].width = 10

    # ---------- Sheet 4: マーカー別集計 ----------
    ws4 = wb.create_sheet("マーカー別集計")
    marker_stats = defaultdict(lambda: {"count": 0, "category": "", "maker": "", "model": "",
                                         "sizes": set()})
    for e in all_elements:
        m = e.get("マーカー")
        if not m:
            continue
        marker_stats[m]["count"] += 1
        marker_stats[m]["category"] = e["部材分類"]
        if e.get("メーカー"):
            marker_stats[m]["maker"] = e["メーカー"]
        if e.get("型番"):
            marker_stats[m]["model"] = e["型番"]
        w = e.get("幅(mm)")
        h = e.get("高さ(mm)")
        if w and h:
            marker_stats[m]["sizes"].add(f"{w}×{h}")

    m_headers = ["マーカー", "部材分類", "数量", "メーカー", "型番", "サイズ(mm)"]
    for ci, h in enumerate(m_headers, 1):
        c = ws4.cell(row=1, column=ci, value=h)
        c.font = hfont; c.fill = hfill; c.alignment = halign; c.border = border

    sorted_markers = sorted(marker_stats.items(), key=lambda x: (x[1]["category"], x[0]))
    for ri, (mname, stats) in enumerate(sorted_markers, 2):
        ws4.cell(row=ri, column=1, value=mname).font = cfont
        ws4.cell(row=ri, column=2, value=stats["category"]).font = cfont
        ws4.cell(row=ri, column=3, value=stats["count"]).font = cfont
        ws4.cell(row=ri, column=4, value=stats["maker"]).font = cfont
        ws4.cell(row=ri, column=5, value=stats["model"]).font = cfont
        ws4.cell(row=ri, column=6, value=", ".join(sorted(stats["sizes"]))).font = cfont
        for ci in range(1, 7):
            ws4.cell(row=ri, column=ci).border = border
            if ri % 2 == 0:
                ws4.cell(row=ri, column=ci).fill = alt_fill

    ws4.column_dimensions["A"].width = 18
    ws4.column_dimensions["B"].width = 10
    ws4.column_dimensions["C"].width = 8
    ws4.column_dimensions["D"].width = 10
    ws4.column_dimensions["E"].width = 10
    ws4.column_dimensions["F"].width = 30
    ws4.auto_filter.ref = f"A1:F{len(sorted_markers)+1}"
    ws4.freeze_panes = "A2"

    wb.save(OUTPUT_PATH)
    print(f"Excel保存: {OUTPUT_PATH}")

    # サマリ表示
    print("\n=== 分類別集計 ===")
    for cat, detail, count, val, unit, note in summary_rows:
        count_s = f"{count:3d}個" if isinstance(count, int) else "   "
        val_s = f"{val:10.2f}" if isinstance(val, (int, float)) else "          "
        print(f"  {cat:12s} {detail:36s} {count_s}  {val_s} {unit:4s} {note}")

    print("\n=== 建具マーカー ===")
    for mname, stats in sorted_markers:
        sizes = ", ".join(sorted(stats["sizes"]))
        print(f"  {mname:16s} {stats['category']:4s} ×{stats['count']}  {stats['maker']:8s} {sizes}")

    # ---------- キット積算 Excel 自動生成 ----------
    kit_path = generate_kit_estimate(all_elements, marker_stats, summary_rows)
    if kit_path:
        print(f"\nキット積算Excel: {kit_path}")


# ============================================================
# キット積算・価格計算 Excel 自動入力
# ============================================================

# カスタムシートの行マッピング（F列に数量を書き込む行番号）
KIT_ROW_MAP = {
    'log_wall_m2':    10,   # ログ壁（LogWall）[m2]
    'floor1f_m2':     11,   # 床板1F（FloorBoard 1F）[m2]
    'floor2f_m2':     12,   # 床板2F（FloorBoard 2F）[m2]
    'ceil1f_m2':      13,   # 天井板1F（CeilingPanel 1F）[m2]
    'ceil2f_m2':      14,   # 天井板2F（CeilingPanel 2F）[m2]
    'beam_m3':        15,   # 集成梁（LaminatedBeam）[m3]
    'panel_m2':       16,   # パネル（LogPanel）[m2]
    'terrace_m2':     21,   # テラス（Terrace）[m2]
    'euro_door_cnt':  22,   # FINドア（EURO Door）[本]
}
KIT_VMW_START, KIT_VMW_END       = 44, 79
KIT_NAGAI_START, KIT_NAGAI_END   = 84, 102
KIT_VELUX_START, KIT_VELUX_END   = 107, 115
KIT_COL_F = 6  # F列

# VMSD-1 → 1810x2019 のような特殊マッピング
_VMSD_MAP = {'VMSD-1': '1810X2019', 'VMSD-2': '2419X2019'}


def _normalize_mk(text):
    """マーカー正規化"""
    return re.sub(r'\s+', '', str(text)).upper().strip() if text else ''


def _extract_base_code(label):
    """Excel品番ラベルから品番ベースコード抽出（サイズ情報除去）"""
    s = str(label).strip()
    s = re.sub(r'\s*\(\d+[xX×]\d+\)\s*', ' ', s)
    s = re.sub(r'\s+\d+幅$', '', s)
    return _normalize_mk(s)


def _build_item_map(ws, r_start, r_end):
    """品番→行番号マッピング: (exact_map, base_map)"""
    exact, base = {}, defaultdict(list)
    for r in range(r_start, r_end + 1):
        label = ws.cell(r, 2).value
        if label:
            ls = str(label).strip()
            exact[_normalize_mk(ls)] = r
            base[_extract_base_code(ls)].append(r)
    return exact, base


def _match_row(marker, exact, base):
    """マーカー→行番号マッチング"""
    mk = _normalize_mk(marker)
    if not mk:
        return None
    # 完全一致
    if mk in exact:
        return exact[mk]
    # VMSD特殊
    if mk in _VMSD_MAP:
        sz = _VMSD_MAP[mk]
        for k, r in exact.items():
            if 'VMSD' in k and sz in k:
                return r
    # ベースコード一致
    if mk in base:
        rows = base[mk]
        if len(rows) == 1:
            return rows[0]
        for r in rows:
            for k, v in exact.items():
                if v == r and '(LF-3)' not in k.upper() and '幅' not in k:
                    return r
        return rows[0]
    # 先頭一致
    for k, r in exact.items():
        if k.startswith(mk + '(') or k.startswith(mk + ' '):
            return r
    # 部分一致
    for k, r in exact.items():
        if mk and len(mk) >= 4 and mk == k[:len(mk)]:
            return r
    return None


def _collect_kit_quantities(all_elements, marker_stats):
    """all_elements からキット積算用の数量を収集"""
    q = {
        'log_wall_m2': 0.0, 'floor1f_m2': 0.0, 'floor2f_m2': 0.0,
        'ceil1f_m2': 0.0, 'ceil2f_m2': 0.0, 'beam_m3': 0.0,
        'panel_m2': 0.0, 'terrace_m2': 0.0, 'euro_door_cnt': 0,
    }

    # ログ壁面積
    for e in all_elements:
        if e["部材分類"] == "ログ壁":
            a = e.get("面積(m²)")
            if isinstance(a, (int, float)):
                q['log_wall_m2'] += a

    # 間仕切壁 パネル面積
    for e in all_elements:
        if e["部材分類"] == "間仕切壁":
            a = e.get("パネル面積(m²)")
            if isinstance(a, (int, float)):
                q['panel_m2'] += a

    # 床（1F/2F/テラス/バルコニー）
    has_2f = False
    for e in all_elements:
        cat = e["部材分類"]
        a = e.get("面積(m²)")
        if not isinstance(a, (int, float)):
            continue
        if cat == "1F床":
            q['floor1f_m2'] += a
        elif cat == "2F床":
            q['floor2f_m2'] += a
            has_2f = True
        elif cat == "テラス":
            q['terrace_m2'] += a

    # 梁体積
    for e in all_elements:
        if e["部材分類"] == "梁":
            v = e.get("体積(m³)")
            if isinstance(v, (int, float)):
                q['beam_m3'] += v

    # 屋根本体面積 → 天井板推定
    roof_body_area = 0.0
    for e in all_elements:
        if e["部材分類"] == "屋根":
            tn = e.get("型式名") or ""
            if "破風" not in tn and "鼻隠" not in tn:
                a = e.get("面積(m²)")
                if isinstance(a, (int, float)):
                    roof_body_area += a

    if has_2f:
        q['ceil1f_m2'] = q['floor2f_m2']
        q['ceil2f_m2'] = roof_body_area
    else:
        q['ceil1f_m2'] = roof_body_area
        q['ceil2f_m2'] = 0.0

    # ドア・窓マーカー別カウント
    door_window_counts = {}
    for mname, stats in marker_stats.items():
        cat = stats.get("category", "")
        if cat in ("ドア", "窓"):
            door_window_counts[re.sub(r'\s+', '', mname)] = stats["count"]
            if mname.upper().startswith("EURO"):
                q['euro_door_cnt'] += stats["count"]

    q['door_window_counts'] = door_window_counts

    # 丸め
    for k in ['log_wall_m2', 'floor1f_m2', 'floor2f_m2', 'ceil1f_m2', 'ceil2f_m2',
              'beam_m3', 'panel_m2', 'terrace_m2']:
        q[k] = round(q[k], 2)

    return q


def generate_kit_estimate(all_elements, marker_stats, summary_rows):
    """キット積算・価格計算 Excel を自動生成。成功時はパスを返す。"""
    out_dir = os.path.dirname(os.path.abspath(OUTPUT_PATH))
    model_name = os.path.splitext(os.path.basename(OUTPUT_PATH))[0].replace("部材一覧_", "")

    # テンプレート検索: KIT_TEMPLATE → OUTPUT_PATHフォルダ → IFC_PATHフォルダ
    template = KIT_TEMPLATE
    if not template:
        for d in [out_dir, os.path.dirname(os.path.abspath(IFC_PATH))]:
            cand = os.path.join(d, "キット積算・価格計算.xlsx")
            if os.path.exists(cand):
                template = cand
                break
    kit_out  = KIT_OUTPUT   or os.path.join(out_dir, f"キット積算_{model_name}.xlsx")

    if not template or not os.path.exists(template):
        print(f"\n⚠ キット積算テンプレートが見つかりません")
        print("  「キット積算・価格計算.xlsx」をIFCファイルと同じフォルダに置いてください")
        print("  キット積算Excelの自動生成をスキップします")
        return None

    print(f"\n=== キット積算 Excel 生成 ===")

    # 数量収集
    q = _collect_kit_quantities(all_elements, marker_stats)

    # テンプレート読み込み
    wb = load_workbook(template)
    ws_name = '積算入力（カスタム）'
    if ws_name not in wb.sheetnames:
        print(f"  ERROR: シート '{ws_name}' が見つかりません")
        return None
    ws = wb[ws_name]

    # スタイル
    inp_fill = PatternFill("solid", start_color="FFFDE7", end_color="FFFDE7")
    inp_font = Font(name='Arial', size=10, bold=True, color="1565C0")
    inp_align = Alignment(horizontal='right', vertical='center')

    def write_val(row, col, value):
        c = ws.cell(row=row, column=col)
        c.value = value
        c.fill = inp_fill
        c.font = inp_font
        c.alignment = inp_align

    # 材料数量書き込み
    for key, row in KIT_ROW_MAP.items():
        val = q.get(key, 0)
        if val and val > 0:
            write_val(row, KIT_COL_F, val)
            label = ws.cell(row, 2).value or key
            print(f"  {label[:30]:30s}: {val}")

    # VMW / Nagai / VELUX 書き込み
    dwc = q.get('door_window_counts', {})
    if dwc:
        vmw_ex, vmw_bs = _build_item_map(ws, KIT_VMW_START, KIT_VMW_END)
        nag_ex, nag_bs = _build_item_map(ws, KIT_NAGAI_START, KIT_NAGAI_END)
        vlx_ex, vlx_bs = _build_item_map(ws, KIT_VELUX_START, KIT_VELUX_END)

        unmatched = []
        for marker, count in sorted(dwc.items()):
            mu = marker.upper()
            if mu.startswith('EURO'):
                continue

            matched = False
            if any(k in mu for k in ['VMW', 'VMSD', 'MTF']):
                r = _match_row(marker, vmw_ex, vmw_bs)
                if r:
                    write_val(r, KIT_COL_F, count)
                    print(f"  VMW  {ws.cell(r,2).value or marker:30s}: {count}")
                    matched = True
            elif any(k in mu for k in ['NV', 'NVS', 'NW', 'NHW']):
                r = _match_row(marker, nag_ex, nag_bs)
                if r:
                    write_val(r, KIT_COL_F, count)
                    print(f"  Nagai {ws.cell(r,2).value or marker:30s}: {count}")
                    matched = True
            elif 'VELUX' in mu:
                r = _match_row(marker, vlx_ex, vlx_bs)
                if r:
                    write_val(r, KIT_COL_F, count)
                    print(f"  VELUX {ws.cell(r,2).value or marker:30s}: {count}")
                    matched = True

            if not matched:
                unmatched.append(f"{marker} x{count}")

        if unmatched:
            print(f"  ⚠ マッチなし: {', '.join(unmatched)}")

    # プロジェクト名
    ws.cell(5, 3).value = model_name

    wb.save(kit_out)
    print(f"  ✓ 保存: {kit_out}")
    return kit_out


if __name__ == "__main__":
    main()
