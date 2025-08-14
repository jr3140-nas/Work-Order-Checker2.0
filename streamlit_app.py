# streamlit_app.py — v10f (no preview, robust PDF pies)
import io, re, math
from xml.sax.saxutils import escape as xml_escape
import pandas as pd
import streamlit as st
from datetime import datetime, date
from typing import Dict, Any, List, Tuple, Optional

PDF_ENGINE = "reportlab"
try:
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.lib.utils import ImageReader
except Exception:
    PDF_ENGINE = "fpdf"
    from fpdf import FPDF, HTMLMixin  # type: ignore

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

st.set_page_config(page_title="Craft-Based Daily Report", layout="wide")

EXPECTED_TIME_COLS = [
    "AddressBookNumber","Name","Production Date","OrderNumber","Sum of Hours.","Hours Estimated",
    "Status","Type","PMFrequency","Description","Problem","Lead Area","Craft","CostCenter","UnitNumber","StructureTag"
]

TYPE_MAP = {
    "0": "Break In","1": "Maintenance Order","2": "Material Repair TMJ Order","3": "Capital Project",
    "4": "Urgent Corrective","5": "Emergency Order","6": "PM Restore/Replace","7": "Inspection Maintenance Order",
    "8": "Follow Up Maintenance Order","9": "Standing W.O. - Do not Delete","B": "Marketing","C": "Cost Improvement",
    "D": "Design Work - ETO","E": "Plant Work - ETO","G": "Governmental/Regulatory","M": "Model W.O. - Eq Mgmt",
    "N": "Template W.O. - CBM Alerts","P": "Project","R": "Rework Order","S": "Shop Order","T": "Tool Order",
    "W": "Case","X": "General Work Request","Y": "Follow Up Work Request","Z": "System Work Request",
}

SPECIAL_COLORS = {
    "inspection maintenance order": "#2ca02c",
    "pm restore/replace": "#228b22",
    "emergency order": "#d62728",
    "break in": "#b22222",
}
OTHER_COLOR = "#b0b0b0"
PALETTE = list(plt.get_cmap("tab20").colors)

def type_to_desc(v: Any) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)): return ""
    s = str(v).strip()
    if s == "": return ""
    try:
        if isinstance(v, (int, float)) or s.replace('.', '', 1).isdigit():
            s_num = str(int(float(s)))
            return TYPE_MAP.get(s_num, s_num)
    except Exception:
        pass
    return TYPE_MAP.get(s.upper(), s)

def normalize_excel_date(v) -> Optional[str]:
    if v is None or (isinstance(v, float) and pd.isna(v)) or v == "": return None
    if isinstance(v, (datetime, date)):
        return datetime(v.year, v.month, v.day).strftime("%m/%d/%Y")
    if isinstance(v, (int, float)):
        for unit, origin in [("D","1899-12-30"), ("ms","unix")]:
            try:
                d = pd.to_datetime(v, unit=unit, origin=origin)
                return d.strftime("%m/%d/%Y")
            except Exception:
                pass
    try:
        d = pd.to_datetime(str(v), errors="coerce")
        if pd.notnull(d): return d.strftime("%m/%d/%Y")
    except Exception:
        pass
    return None

def numberish(v) -> float:
    if isinstance(v, (int, float)): return float(v)
    if isinstance(v, str):
        try: return float("".join(ch for ch in v if (ch.isdigit() or ch in ".-")))
        except Exception: return 0.0
    return 0.0

def name_key(s: str) -> str:
    return " ".join(str(s).strip().upper().split())

def build_name_to_craft(addr_df: pd.DataFrame) -> Tuple[Dict[str, str], List[str]]:
    addr_df = addr_df.rename(columns=lambda c: str(c).strip())
    col_map = {str(c).strip().lower(): str(c).strip() for c in addr_df.columns}
    name_col = col_map.get("name")
    craft_desc_col = col_map.get("craft description") or col_map.get("craft_description")
    if not name_col or not craft_desc_col:
        missing = []
        if not name_col: missing.append("Name")
        if not craft_desc_col: missing.append("Craft Description")
        raise ValueError(f"Address Book missing required columns: {missing}. Found columns: {list(addr_df.columns)}")
    ab = addr_df[[name_col, craft_desc_col]].dropna(how="any").copy()
    ab[name_col] = ab[name_col].astype(str).map(name_key)
    ab[craft_desc_col] = ab[craft_desc_col].astype(str).str.strip()
    conflicts = []
    mapping: Dict[str, str] = {}
    for _, row in ab.iterrows():
        nm = row[name_col]; cd = row[craft_desc_col]
        if nm in mapping and mapping[nm] != cd: conflicts.append(f"{nm}: '{mapping[nm]}' vs '{cd}'")
        else: mapping[nm] = cd
    return mapping, conflicts

def build_report(df: pd.DataFrame, selected_date: str, name_to_craft: Dict[str, str]):
    df = df.copy(); df["__ProdDate"] = df["Production Date"].apply(normalize_excel_date)
    df = df[df["__ProdDate"] == selected_date]
    df["__NameKey"] = df["Name"].astype(str).map(name_key)
    df["__CraftDesc"] = df["__NameKey"].map(name_to_craft).fillna("(Unmapped Name)")
    groups = {}
    for _, r in df.iterrows():
        k = (r["__CraftDesc"], r["__NameKey"], r["OrderNumber"])
        if k not in groups:
            groups[k] = {"Craft": r["__CraftDesc"], "Name": r["Name"], "OrderNumber": r["OrderNumber"],
                        "SumOfHours": 0.0, "Type": set(), "Description": set(), "Problem": set()}
        g = groups[k]
        g["SumOfHours"] += numberish(r.get("Sum of Hours.", 0))
        t_desc = type_to_desc(r.get("Type", ""))
        if t_desc: g["Type"].add(t_desc)
        d = str(r.get("Description", "")).strip()
        p = str(r.get("Problem", "")).strip()
        if d: g["Description"].add(d)
        if p: g["Problem"].add(p)
    crafts: Dict[str, List[Dict[str, Any]]] = {}
    for v in groups.values():
        crafts.setdefault(v["Craft"], []).append({
            "Name": v["Name"],
            "Work Order #": v["OrderNumber"],
            "Sum of Hours": round(v["SumOfHours"], 2),
            "Type": "; ".join(sorted(v["Type"])),
            "Description": "; ".join(sorted(v["Description"])),
            "Problem": "; ".join(sorted(v["Problem"])),
        })
    def wo_val(s) -> float:
        s = str(s); return float(s) if s.isdigit() else float("inf")
    for c in list(crafts.keys()): crafts[c].sort(key=lambda r: wo_val(r.get("Work Order #", "")))
    day_names = set(df["__NameKey"].tolist()); mapped_names = set(name_to_craft.keys())
    unmapped = sorted(n for n in day_names if n not in mapped_names)
    return crafts, unmapped, df

def summarize_hours_by_type_per_area(df: pd.DataFrame) -> Dict[str, Dict[str, float]]:
    tmp = df.copy()
    tmp["__TypeDesc"] = tmp["Type"].map(lambda x: type_to_desc(x) or "Unknown Type")
    tmp["__H"] = tmp["Sum of Hours."].map(numberish)
    pivot = tmp.groupby(["__CraftDesc", "__TypeDesc"])["__H"].sum().reset_index()
    result: Dict[str, Dict[str, float]] = {}
    for _, r in pivot.iterrows():
        area = r["__CraftDesc"]; t = r["__TypeDesc"]; h = float(r["__H"] or 0.0)
        result.setdefault(area, {}); result[area][t] = result[area].get(t, 0.0) + h
    return result

def build_global_color_map(summary: Dict[str, Dict[str, float]]) -> Dict[str, str]:
    labels = set(); [labels.update(typemap.keys()) for typemap in summary.values()]
    labels = {(l or '').strip() for l in labels}
    cmap: Dict[str, str] = {}
    for lab in list(labels):
        if lab.lower() in SPECIAL_COLORS: cmap[lab] = SPECIAL_COLORS[lab.lower()]
    labels = {l for l in labels if l not in cmap}
    it = iter(PALETTE)
    for lab in sorted(labels, key=str.lower):
        try: cmap[lab] = next(it)
        except StopIteration:
            it = iter(PALETTE); cmap[lab] = next(it)
    cmap["Other"] = OTHER_COLOR; cmap["other"] = OTHER_COLOR; return cmap

def _place_outside_labels(ax, wedges, labels: List[str], values: List[float], total: float, label_mode: str, color_map: Dict[str,str]):
    entries = []
    for w, lab, val in zip(wedges, labels, values):
        theta = 0.5*(w.theta1+w.theta2); ang = math.radians(theta)
        x,y = math.cos(ang), math.sin(ang); side = "right" if x>=0 else "left"
        pct = (val/total*100.0) if total>0 else 0.0
        text = f"{lab} — {pct:.0f}%" if label_mode=="percent" else f"{lab} — {val:.1f}h"
        entries.append({"x":x,"y":y,"side":side,"text":text,"label":lab})
    for side in ("right","left"):
        se = [e for e in entries if e["side"]==side]
        if not se: continue
        se.sort(key=lambda e: -e["y"]); min_dy = 0.12; last_y = 1.2
        for e in se:
            y_target = e["y"]
            if last_y < 1.0: y_target = min(y_target, last_y - min_dy)
            last_y = y_target; x_text = 1.2 if side=="right" else -1.2
            ax.annotate("", xy=(e["x"],e["y"]), xytext=(x_text,y_target),
                        arrowprops=dict(arrowstyle="-", connectionstyle="arc3", lw=0.6, color="#444"),
                        annotation_clip=False)
            ax.text(x_text,y_target,e["text"], ha="left" if side=="right" else "right", va="center",
                    fontsize=7.5, bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=0.75), zorder=3)

def render_pie_pages(summary: Dict[str, Dict[str, float]], selected_date: str, label_mode: str) -> List[bytes]:
    if not summary:
        fig, ax = plt.subplots(figsize=(11, 8.5)); ax.axis("off")
        ax.text(0.5, 0.5, f"No hours found for {selected_date}", ha="center", va="center", fontsize=16, fontweight="bold")
        buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=150); plt.close(fig); return [buf.getvalue()]
    color_map = build_global_color_map(summary)
    items = sorted(summary.items(), key=lambda kv: sum(kv[1].values()), reverse=True)
    pies_per_page = 6; pages: List[bytes] = []
    for i in range(0, len(items), pies_per_page):
        chunk = items[i:i+pies_per_page]
        fig, axs = plt.subplots(2,3,figsize=(11,8.5), constrained_layout=True)
        title_unit = "%" if label_mode == "percent" else "hours"
        fig.suptitle(f"Hours by Work Order Type per Area — {selected_date} ({title_unit})", fontsize=14)
        axs = axs.flatten(); [ax.axis("off") for ax in axs]
        for ax, (area, typemap) in zip(axs, chunk):
            labels = list(typemap.keys()); sizes = [max(0.0, float(typemap[k])) for k in labels]
            if sum(sizes) <= 0:
                ax.text(0.5, 0.5, f"{area}\n(no hours)", ha="center", va="center", fontsize=10, fontweight="bold"); continue
            total0 = sum(sizes); lbl2, sz2, other = [], [], 0.0
            for l, s in zip(labels, sizes):
                if total0 > 0 and (s/total0) < 0.02: other += s
                else: lbl2.append(l); sz2.append(s)
            if other > 0: lbl2.append("Other"); sz2.append(other)
            total = sum(sz2); cols = [color_map.get(l, OTHER_COLOR) for l in lbl2]
            wedges, _ = ax.pie(sz2, startangle=90, counterclock=False, colors=cols, radius=1.0, labels=None)
            ax.set_title(f"{area} — Total: {total:.2f} h", fontsize=11, fontweight="bold", pad=10)
            _place_outside_labels(ax, wedges, lbl2, sz2, total, label_mode, color_map)
        buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=150); plt.close(fig); pages.append(buf.getvalue())
    return pages

def _soft_wrap_for_pdf(s: str, max_chars: int = 600, unbroken_chunk: int = 40) -> str:
    if s is None: return ""
    s = str(s)
    if not s: return ""
    def _break_token(m):
        token = m.group(0); return " ".join(token[i:i+unbroken_chunk] for i in range(0, len(token), unbroken_chunk))
    s = re.sub(r"\S{" + str(unbroken_chunk) + r",}", _break_token, s)
    if len(s) > max_chars: s = s[:max_chars - 1] + "…"
    return s

def _compute_rl_col_widths(rows: List[List[str]], page_inner_width: float) -> List[float]:
    minw = [120, 90, 90, 140, 240, 240]; pad = 14; naturals = []
    for col_idx in range(len(rows[0])):
        max_w = 0.0
        for r in rows:
            txt = str(r[col_idx]) if r[col_idx] is not None else ""
            max_w = max(max_w, stringWidth(txt, "Helvetica", 8))
        naturals.append(max(max_w + pad, minw[col_idx]))
    total = sum(naturals)
    if total <= page_inner_width: return naturals
    over = total - page_inner_width
    shrinkable = [max(0.0, naturals[i] - minw[i]) for i in range(len(naturals))]
    total_shrinkable = sum(shrinkable)
    if total_shrinkable <= 0:
        scale = page_inner_width / total if total > 0 else 1.0
        return [w * scale for w in naturals]
    widths = []
    for i, w in enumerate(naturals):
        reduce = over * (shrinkable[i] / total_shrinkable) if total_shrinkable > 0 else 0.0
        widths.append(max(minw[i], w - reduce))
    if sum(widths) > page_inner_width:
        scale = (page_inner_width - 0.01 * page_inner_width) / sum(widths)
        widths = [w * scale for w in widths]
    return widths

def make_pdf(selected_date: str, crafts: Dict[str, List[Dict[str, Any]]], cover_summary: Dict[str, Dict[str, float]], label_mode: str, cap_desc=600, cap_prob=600) -> bytes:
    if PDF_ENGINE == "reportlab":
        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=landscape(letter), leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24)
        styles = getSampleStyleSheet(); title_style = styles["Title"]; header_style = styles["Heading2"]
        body8 = ParagraphStyle("Body8", parent=styles["BodyText"], fontName="Helvetica", fontSize=8, leading=10)
        table_style = TableStyle([("GRID", (0,0), (-1,-1), 0.25, colors.grey),("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
                                  ("ALIGN", (0,0), (-1,0), "LEFT"),("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                                  ("FONTSIZE", (0,0), (-1,0), 9),("FONTSIZE", (0,1), (-1,-1), 8),("VALIGN", (0,0), (-1,-1), "TOP")])
        story: List = []
        # Cover pages: use BytesIO + explicit size (avoid filename checks)
        pie_pages = render_pie_pages(cover_summary, selected_date, label_mode)
        for idx, png in enumerate(pie_pages):
            data = bytes(png)
            # Two streams: one to query size, one to feed RLImage
            bio_size = io.BytesIO(data); bio_img = io.BytesIO(data)
            reader = ImageReader(bio_size)
            iw, ih = reader.getSize()
            max_w, max_h = doc.width, doc.height
            scale = min(max_w/iw, max_h/ih) * 0.95
            img = RLImage(bio_img, width=iw*scale, height=ih*scale)
            story.append(img)
            if idx < len(pie_pages) - 1: story.append(PageBreak())
        # Details
        story += [Paragraph(f"Daily Report — {selected_date}", title_style), Spacer(1, 6),
                  Paragraph("Sorted by Work Order # within each craft", styles["Normal"]), Spacer(1, 12)]
        page_inner_width = doc.width
        for craft, rows in crafts.items():
            story.append(Paragraph(str(craft), header_style))
            matrix = [["Name","Work Order #","Sum of Hours","Type","Description","Problem"]]
            for r in rows:
                matrix.append([str(r.get("Name","")), str(r.get("Work Order #","")), f'{float(r.get("Sum of Hours",0)):.2f}',
                               str(r.get("Type","")), str(r.get("Description","")), str(r.get("Problem",""))])
            col_widths = _compute_rl_col_widths(matrix, page_inner_width)
            data_rows = [matrix[0]]
            for raw in matrix[1:]:
                name = xml_escape(_soft_wrap_for_pdf(raw[0],180)); wo = xml_escape(_soft_wrap_for_pdf(raw[1],60)); hrs = xml_escape(_soft_wrap_for_pdf(raw[2],32))
                typ = xml_escape(_soft_wrap_for_pdf(raw[3],220)); desc = xml_escape(_soft_wrap_for_pdf(raw[4],cap_desc)); prob = xml_escape(_soft_wrap_for_pdf(raw[5],cap_prob))
                data_rows.append([Paragraph(name, body8), Paragraph(wo, body8), Paragraph(hrs, body8), Paragraph(typ, body8), Paragraph(desc, body8), Paragraph(prob, body8)])
            tbl = Table(data_rows, repeatRows=1, colWidths=col_widths); tbl.setStyle(table_style); story.append(tbl); story.append(Spacer(1,10))
        doc.build(story); pdf = buf.getvalue(); buf.close(); return pdf

    # FPDF fallback
    class PDF(FPDF, HTMLMixin): pass
    margin = 24; pdf = PDF(orientation="L", unit="pt", format="Letter")
    pie_pages = render_pie_pages(cover_summary, selected_date, label_mode)
    for png in pie_pages:
        pdf.add_page()
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp.write(png); tmp.flush(); img_path = tmp.name
        max_w = pdf.w - 2*margin; max_h = pdf.h - 2*margin
        from PIL import Image as PILImage
        im = PILImage.open(img_path); iw, ih = im.size; scale = min(max_w/iw, max_h/ih) * 0.95; w = iw * scale; h = ih * scale
        pdf.image(img_path, x=(pdf.w-w)/2, y=(pdf.h-h)/2, w=w, h=h)
    pdf.add_page(); pdf.set_left_margin(margin); pdf.set_right_margin(margin)
    pdf.set_auto_page_break(auto=True, margin=margin)
    pdf.set_font("Helvetica", "B", 16); pdf.cell(0, 18, f"Daily Report — {selected_date}", ln=1)
    pdf.set_font("Helvetica", "", 10); pdf.cell(0, 14, "Sorted by Work Order # within each craft", ln=1)
    page_inner_width = pdf.w - pdf.l_margin - pdf.r_margin
    def compute_fpdf_widths(rows: List[List[str]]) -> List[float]:
        minw = [120, 90, 90, 140, 240, 240]; pad = 12; naturals = []; pdf.set_font("Helvetica", "", 8)
        for col_idx in range(len(rows[0])):
            mx = 0.0
            for r in rows:
                txt = str(r[col_idx]) if r[col_idx] is not None else ""
                mx = max(mx, pdf.get_string_width(txt))
            naturals.append(max(mx + pad, minw[col_idx]))
        total = sum(naturals)
        if total <= page_inner_width: return naturals
        over = total - page_inner_width
        shrinkable = [max(0.0, naturals[i] - minw[i]) for i in range(len(naturals))]
        total_shrink = sum(shrinkable)
        if total_shrink <= 0:
            scale = page_inner_width / total if total > 0 else 1.0
            return [w * scale for w in naturals]
        widths = []
        for i, w in enumerate(naturals):
            reduce = over * (shrinkable[i] / total_shrink) if total_shrink > 0 else 0.0
            widths.append(max(minw[i], w - reduce))
        if sum(widths) > page_inner_width:
            scale = (page_inner_width - 0.01 * page_inner_width) / sum(widths)
            widths = [w * scale for w in widths]
        return widths
    th = 14; pdf.set_font("Helvetica", "", 8)
    for craft, rows in crafts.items():
        pdf.ln(6); pdf.set_font("Helvetica", "B", 13); pdf.cell(0, 16, str(craft), ln=1); pdf.set_font("Helvetica", "", 8)
        matrix = [["Name","Work Order #","Sum of Hours","Type","Description","Problem"]]
        for r in rows:
            matrix.append([str(r.get("Name","")), str(r.get("Work Order #","")), f'{float(r.get("Sum of Hours",0)):.2f}',
                           str(r.get("Type","")), str(r.get("Description","")), str(r.get("Problem",""))])
        col_widths = compute_fpdf_widths(matrix)
        pdf.set_font("Helvetica", "B", 9)
        for w, txt in zip(col_widths, matrix[0]): pdf.cell(w, th, txt, border=1)
        pdf.ln(th)
        pdf.set_font("Helvetica", "", 8)
        for raw in matrix[1:]:
            fields = [(raw[0],180),(raw[1],60),(raw[2],32),(raw[3],220),(raw[4],600),(raw[5],600)]
            clipped = []
            for s, cap in fields:
                s = "" if s is None else str(s)
                if len(s) > cap: s = s[:cap-1] + "…"
                clipped.append(s)
            for w, txt in zip(col_widths, clipped): pdf.cell(w, th, txt, border=1)
            pdf.ln(th)
    return bytes(pdf.output(dest="S").encode("latin1"))

# ------------------------ UI ------------------------
st.title("Craft-Based Daily Report (Excel → PDF) — v10f")

with st.sidebar:
    st.markdown("**Instructions**")
    st.markdown("1) Upload the **Address Book** (.xlsx) to build the Name → Craft Description mapping.")
    st.markdown("2) Upload the **Time on Work Order** (.xlsx).")
    st.markdown("3) Pick a **Production Date** (MM/DD/YYYY).")
    st.markdown("4) Choose **Cover Labels**: Percent or Hours.")
    st.markdown("5) Download PDF. The first page(s) show pies by area; details follow.")

col1, col2, col3 = st.columns([1,1,1])
with col1:
    addr_file = st.file_uploader("Upload Address Book (.xlsx)", type=["xlsx"], key="addr")
with col2:
    time_file = st.file_uploader("Upload Time on Work Order (.xlsx)", type=["xlsx"], key="time")
with col3:
    label_mode = st.radio("Cover Labels", options=["percent", "hours"], index=0, horizontal=False)

cap_choice = st.selectbox("PDF text length (Description/Problem)", ["Compact (450)", "Standard (600)", "Verbose (800)"], index=1)
cap_map = {"Compact (450)": 450, "Standard (600)": 600, "Verbose (800)": 800}
cap_val = cap_map[cap_choice]

addr_map: Dict[str, str] | None = None
addr_conflicts: List[str] = []

if addr_file is not None:
    try:
        addr_df = pd.read_excel(addr_file)
        addr_map, addr_conflicts = build_name_to_craft(addr_df)
        st.success(f"Address Book loaded: {len(addr_map)} names mapped.")
        if addr_conflicts:
            st.warning("Conflicting craft descriptions for the same name:\n- " + "\n- ".join(addr_conflicts))
    except Exception as e:
        st.error(f"Failed to read Address Book: {e}")

df = None; dates: List[str] = []
if time_file is not None:
    try:
        df = pd.read_excel(time_file, header=2)
        df.columns = [str(c).strip() for c in df.columns]
        missing = [c for c in EXPECTED_TIME_COLS if c not in df.columns]
        if missing: st.error(f"Time sheet missing expected columns: {missing}")
        else:
            dates = sorted({d for d in (df["Production Date"].apply(normalize_excel_date).dropna().tolist())})
            st.caption(f"Detected dates: {(dates[0] if dates else '—')} → {(dates[-1] if dates else '—')} • Unique dates: {len(dates)}")
    except Exception as e:
        st.exception(e)

selected_date = st.selectbox("Production Date", options=(dates if dates else [""]), index=(len(dates)-1 if dates else 0))

if df is not None and addr_map is not None and selected_date:
    crafts, unmapped_names, df_filtered = build_report(df, selected_date, addr_map)
    if unmapped_names: st.error("Unmapped Names (from selected date):\n- " + "\n- ".join(unmapped_names))
    cover_summary = summarize_hours_by_type_per_area(df_filtered)

    for craft, rows in crafts.items():
        st.subheader(craft)
        st.dataframe(pd.DataFrame(rows, columns=["Name","Work Order #","Sum of Hours","Type","Description","Problem"]))

    pdf_bytes = make_pdf(selected_date, crafts, cover_summary, label_mode, cap_desc=cap_val, cap_prob=cap_val)
    st.download_button("Download PDF", data=pdf_bytes, file_name=f"nas_report_{selected_date.replace('/', '-')}.pdf", mime="application/pdf")
elif (df is not None) and (addr_map is None):
    st.info("Upload the **Address Book** to generate the Name → Craft Description mapping.")
elif (addr_map is not None) and (df is None):
    st.info("Upload the **Time on Work Order** sheet to continue.")
