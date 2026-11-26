import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.colors import COLOR_INDEX
from copy import copy
from io import BytesIO
import xml.etree.ElementTree as ET
import xlrd
import re

st.set_page_config(page_title="BOM to Trial Template Filler", layout="wide")
st.title("BOM to Trial Template Auto-Filler")


def clean_header(text):
    return re.sub(r"\s+", " ", str(text)).strip().lower() if text is not None else ""


def rgb_to_hex(rgb):
    return "{:02X}{:02X}{:02X}".format(*rgb) if rgb and len(rgb) == 3 else None


def apply_tint(hex_color, tint):
    if not hex_color or len(hex_color) != 6 or not tint:
        return hex_color
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    if tint < 0:
        r, g, b = int(r * (1 + tint)), int(g * (1 + tint)), int(b * (1 + tint))
    else:
        r = int(r + (255 - r) * tint)
        g = int(g + (255 - g) * tint)
        b = int(b + (255 - b) * tint)
    return "{:02X}{:02X}{:02X}".format(max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b)))


def get_xlsx_theme_colors(wb):
    theme_colors = [
        "FFFFFF", "000000", "E7E6E6", "44546A", "4472C4",
        "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47",
    ]
    try:
        if wb.loaded_theme:
            root = ET.fromstring(wb.loaded_theme)
            ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
            clr_scheme = root.find(".//a:clrScheme", ns)
            if clr_scheme is not None:
                order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]
                extracted = []
                for tag in order:
                    node = clr_scheme.find(f"a:{tag}", ns)
                    if node is not None:
                        srgb = node.find("a:srgbClr", ns)
                        sysclr = node.find("a:sysClr", ns)
                        if srgb is not None and "val" in srgb.attrib:
                            extracted.append(srgb.attrib["val"].upper())
                        elif sysclr is not None and "lastClr" in sysclr.attrib:
                            extracted.append(sysclr.attrib["lastClr"].upper())
                if len(extracted) == 10:
                    theme_colors = extracted
    except Exception:
        pass
    return theme_colors


def resolve_openpyxl_color(color_obj, theme_colors, is_bg=True):
    if not color_obj:
        return None
    ctype = getattr(color_obj, "type", None)
    if ctype == "rgb":
        rgb = str(color_obj.rgb)
        if len(rgb) == 8:
            alpha, hex_val = rgb[:2], rgb[2:].upper()
            if alpha == "00" and hex_val == "000000":
                return None
            if is_bg and hex_val == "FFFFFF":
                return None
            if not is_bg and hex_val == "000000":
                return None
            return hex_val
    elif ctype == "theme":
        idx = color_obj.theme
        if 0 <= idx < len(theme_colors):
            final_hex = apply_tint(theme_colors[idx], getattr(color_obj, "tint", 0.0))
            if is_bg and final_hex == "FFFFFF":
                return None
            if not is_bg and final_hex == "000000":
                return None
            return final_hex
    elif ctype == "indexed":
        idx = color_obj.indexed
        if 0 <= idx < len(COLOR_INDEX):
            rgb = COLOR_INDEX[idx]
            hex_val = rgb[2:].upper() if len(rgb) == 8 else rgb.upper()
            if is_bg and hex_val in ("FFFFFF", "000000"):
                return None
            if not is_bg and hex_val == "000000":
                return None
            return hex_val
    return None


def get_cell_style(wb, sh, r, c, is_xls, theme_colors=None):
    bg_hex, font_hex = None, None
    try:
        if is_xls:
            xf = wb.xf_list[sh.cell_xf_index(r, c)]
            if xf.background.fill_pattern != 0:
                bg_idx = xf.background.pattern_colour_index
                b_col = rgb_to_hex(wb.colour_map.get(bg_idx))
                if b_col and b_col != "FFFFFF" and bg_idx not in (64, 65):
                    bg_hex = b_col
            font = wb.font_list[xf.font_index]
            f_col = rgb_to_hex(wb.colour_map.get(font.colour_index))
            if f_col and f_col != "000000" and font.colour_index not in (32767, 8, 64):
                font_hex = f_col
        else:
            cell = sh.cell(row=r + 1, column=c + 1)
            if cell.fill and cell.fill.fill_type not in (None, "none"):
                bg_hex = resolve_openpyxl_color(cell.fill.start_color, theme_colors, is_bg=True)
            if cell.font and cell.font.color:
                font_hex = resolve_openpyxl_color(cell.font.color, theme_colors, is_bg=False)
    except Exception:
        pass
    return bg_hex, font_hex


def get_row_styles_map(wb, sh, r, cmap, is_xls, theme_colors=None):
    cell_styles = {}
    for key, col_idx in cmap.items():
        bg, fg = get_cell_style(wb, sh, r, col_idx, is_xls, theme_colors)
        cell_styles[key] = (bg, fg)
    cell_styles["_row_default"] = (None, None)
    return cell_styles


def extract_bom_data(bom_file):
    fb = bom_file.getvalue()
    is_xls = bom_file.name.lower().endswith(".xls")
    xls_pd = pd.ExcelFile(BytesIO(fb))
    sname = next((s for s in xls_pd.sheet_names if "vse_bom" in s.lower() or "bom" in s.lower()), xls_pd.sheet_names[0])
    df = pd.read_excel(xls_pd, sheet_name=sname, header=None)

    theme_colors = None
    if is_xls:
        wb_fmt = xlrd.open_workbook(file_contents=fb, formatting_info=True)
        sh_fmt = wb_fmt.sheet_by_name(sname)
    else:
        wb_fmt = openpyxl.load_workbook(BytesIO(fb), data_only=True)
        sh_fmt = wb_fmt[sname]
        theme_colors = get_xlsx_theme_colors(wb_fmt)

    meta = {"customer_id": "", "ship_site": ""}
    hdr_idx = None
    for r in range(min(20, len(df))):
        rvals = [clean_header(v) for v in df.iloc[r]]
        for c, val in enumerate(rvals):
            for key, label in [("customer_id", "customer id"), ("ship_site", "ship site")]:
                if label in val:
                    for off in range(1, 4):
                        if c + off < len(df.columns) and pd.notna(df.iloc[r, c + off]):
                            meta[key] = str(df.iloc[r, c + off]).strip()
                            break
        if any("vse item" in v for v in rvals) and any("manufacturer" in v for v in rvals):
            hdr_idx = r
            break

    if hdr_idx is None:
        raise ValueError("Could not find 'VSE Item' and 'Manufacturer' headers in the BOM file.")

    bh = [clean_header(v) for v in df.iloc[hdr_idx]]
    cmap = {}
    rules = [
        ("lvl", lambda h: h == "lvl"),
        ("vse_item", lambda h: h == "vse item"),
        ("cust_item", lambda h: h == "cust item"),
        ("description", lambda h: h == "description"),
        ("manufacturer", lambda h: h == "manufacturer"),
        ("mpn", lambda h: "manufacturer p/n" in h or "mpn" in h),
        ("mfg_id", lambda h: "manufacturer id" in h),
        ("planner", lambda h: "planner" in h),
        ("vendor_id", lambda h: "vendor id" in h),
        ("unit_cost", lambda h: "unit cost" in h),
    ]
    for idx, h in enumerate(bh):
        for k, fn in rules:
            if k not in cmap and fn(h):
                cmap[k] = idx

    last_row = hdr_idx
    for r in range(hdr_idx + 1, len(df)):
        if any(pd.notna(df.iloc[r, cmap[k]]) and str(df.iloc[r, cmap[k]]).strip() != "" for k in ["lvl", "vse_item", "manufacturer", "mpn"] if k in cmap):
            last_row = r

    records, parent = [], {}
    for r in range(hdr_idx + 1, last_row + 1):
        row = df.iloc[r]

        def val(k):
            if k in cmap and pd.notna(row[cmap[k]]):
                s = str(row[cmap[k]]).strip()
                return "" if (s == "0" and k == "vendor_id") else s
            return ""

        lvl, vse, mfg, mpn = val("lvl"), val("vse_item"), val("manufacturer"), val("mpn")
        styles_map = get_row_styles_map(wb_fmt, sh_fmt, r, cmap, is_xls, theme_colors)

        if not lvl and not vse and not mfg and not mpn:
            parent = {}
            records.append({"_styles": styles_map})
            continue

        if lvl or vse:
            rc = row[cmap["unit_cost"]] if "unit_cost" in cmap and pd.notna(row[cmap["unit_cost"]]) else ""
            parent = {
                "vse item": vse,
                "cust item": val("cust_item"),
                "description": val("description"),
                "planner codes": val("planner"),
                "vendor id": val("vendor_id"),
                "unit cost": rc if rc != 0 else "",
            }

        has_part = bool(vse or mpn or mfg)
        records.append({
            "manufacturer id": val("mfg_id"),
            "manufacturer p/n": mpn,
            "manufatures": mfg,
            "manufacturer": mfg,
            "vse item": parent.get("vse item", ""),
            "description": parent.get("description", ""),
            "unit cost": parent.get("unit cost", ""),
            "planner codes": parent.get("planner codes", ""),
            "customer id": meta["customer_id"] if has_part else "",
            "cust item": parent.get("cust item", ""),
            "vendor id pricing from": parent.get("vendor id", ""),
            "vendor id": parent.get("vendor id", ""),
            "stock location": meta["ship_site"] if has_part else "",
            "_styles": styles_map,
        })
    return records


def map_header_to_key(hc):
    if "manufacturer id" in hc:
        return "mfg_id"
    if "manufacturer p/n" in hc or "mpn" in hc:
        return "mpn"
    if "manufature" in hc or "manufacturer" in hc:
        return "manufacturer"
    if "vse item" in hc:
        return "vse_item"
    if "description" in hc:
        return "description"
    if "unit cost" in hc:
        return "unit_cost"
    if "planner" in hc:
        return "planner"
    if "cust item" in hc:
        return "cust_item"
    if "vendor id" in hc:
        return "vendor_id"
    return "_row_default"


def fill_trial_template(trial_file, records):
    wb = openpyxl.load_workbook(trial_file, data_only=False)
    ws = wb.active
    th = [clean_header(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]
    sec = {"main", "planning", "controls", "vse", "item coset", "item cost", "customer contract", "item mfgr"}
    bf = copy(ws.cell(2, 1).font) if ws.max_row >= 2 else Font(name="Calibri", size=11)

    base_formulas = {}
    if ws.max_row >= 2:
        for c_idx in range(1, ws.max_column + 1):
            val = ws.cell(2, c_idx).value
            if isinstance(val, str) and val.startswith("="):
                base_formulas[c_idx] = val

    for r_idx, rec in enumerate(records, start=2):
        styles_map = rec.get("_styles", {})

        for c_idx, hc in enumerate(th, start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            existing_val = cell.value

            if isinstance(existing_val, str) and existing_val.startswith("="):
                pass
            elif c_idx in base_formulas:
                cell.value = re.sub(r"(?<=[A-Z])2\b", str(r_idx), base_formulas[c_idx])
            else:
                if hc in sec or not hc:
                    v = ""
                elif "vendor id" in hc:
                    v = rec.get("vendor id", "")
                else:
                    v = rec.get(hc, next((val for k, val in rec.items() if k != "_styles" and k and k in hc), ""))
                cell.value = v

            skey = map_header_to_key(hc)
            bg, fg = styles_map.get(skey, (None, None))

            cf = copy(bf)
            cf.color = fg if fg else "000000"
            cell.font = cf

            if bg and bg.upper() not in ("FFFFFF", "00000000"):
                cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
            else:
                cell.fill = PatternFill(fill_type=None)

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


c1, c2 = st.columns(2)
with c1:
    t_up = st.file_uploader("1. Upload Trial Template File (.xlsx)", type=["xlsx"])
with c2:
    b_up = st.file_uploader("2. Upload BOM File (.xls / .xlsx)", type=["xls", "xlsx"])

if t_up and b_up and st.button("Fill Template & Download", type="primary"):
    try:
        with st.spinner("Extracting BOM items and preserving exact colors & formulas..."):
            recs = extract_bom_data(b_up)
            f_xl = fill_trial_template(t_up, recs)
            st.success(f"Successfully populated {len(recs)} rows with exact source colors & formulas!")
            st.dataframe(pd.DataFrame(recs).drop(columns=["_styles"], errors="ignore").head(20), use_container_width=True)
            st.download_button(
                "Download Filled Excel",
                f_xl,
                b_up.name.rsplit(".", 1)[0] + "_Trial_Filled.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error(f"Error: {e}")
