
import io
import re
import itertools
from xml.sax.saxutils import escape as xml_escape
import pandas as pd
import streamlit as st
from datetime import datetime, date
from typing import Dict, Any, List, Tuple
GROUP_FILTER_OPTIONS = ["Electrical", "Mechanical", "Turns", "Segment Shop", "Utilities"]
CRAFT_ORDER = ["Turns", "EAF Mech Days", "EAF Elec Days", "AOD Mech Days", "AOD Elec Days", "Alloy Mech Days", "Caster Mech Days", "Caster Elec Days", "WTP Mech Days", "Baghouse Mech Days", "Preheater Elec Days", "Segment Shop", "Utilities Mech Days", "HVAC Elec Days"]
GROUP_ORDER = ["Turns","Electrical","Segment Shop","Utilities","Mechanical"]

def _load_group_order(file) -> List[str]:
    """
    Accepts an uploaded Excel file and returns a canonical group order list.
    First non-empty column is used. Duplicates removed preserving order.
    """
    try:
        df = pd.read_excel(file)
        col = df.columns[0]
        vals = [str(x).strip() for x in df[col].tolist() if str(x).strip()]
        seen = set()
        out = []
        for v in vals:
            key = v.lower()
            if key not in seen:
                seen.add(key)
                # Canonicalize to our known labels where possible
                canon = v.strip()
                low = canon.lower()
                if "elec" in low: canon = "Electrical"
                elif "mech" in low: canon = "Mechanical"
                elif "turn" in low: canon = "Turns"
                elif "segment" in low or "seg shop" in low: canon = "Segment Shop"
                elif "util" in low: canon = "Utilities"
                out.append(canon)
        return out
    except Exception:
        return []


# Prefer ReportLab; fall back to fpdf2 if ReportLab isn't available.
PDF_ENGINE = "reportlab"
try:
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.pdfbase.pdfmetrics import stringWidth
except Exception:
    PDF_ENGINE = "fpdf"
    from fpdf import FPDF, HTMLMixin  # type: ignore

# Charts
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

st.set_page_config(page_title="Craft-Based Daily Report", layout="wide")

EXPECTED_TIME_COLS = [
    "AddressBookNumber","Name","Production Date","OrderNumber","Sum of Hours.","Hours Estimated",
    "Status","Type","PMFrequency","Description","Problem","Lead Area","Craft","CostCenter","UnitNumber","StructureTag"
]

# --- Work Order Type Mapping ---
TYPE_MAP = {
    "0": "Break In","1": "Maintenance Order","2": "Material Repair TMJ Order","3": "Capital Project",
    "4": "Urgent Corrective","5": "Emergency Order","6": "PM Restore/Replace","7": "Inspection Maintenance Order",
    "8": "Follow Up Maintenance Order","9": "Standing W.O. - Do not Delete","B": "Marketing","C": "Cost Improvement",
    "D": "Design Work - ETO","E": "Plant Work - ETO","G": "Governmental/Regulatory","M": "Model W.O. - Eq Mgmt",
    "N": "Template W.O. - CBM Alerts","P": "Project","R": "Rework Order","S": "Shop Order","T": "Tool Order",
    "W": "Case","X": "General Work Request","Y": "Follow Up Work Request","Z": "System Work Request",
}

GREEN1 = "#2ca02c"   # inspection
GREEN2 = "#228b22"   # restore/replace
RED1   = "#d62728"   # emergency
RED2   = "#b22222"   # break in
# fallback palette (cycled for all other types)
FALLBACK = list(plt.get_cmap("tab20").colors)

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

# ------------------------ Utilities ------------------------
def normalize_excel_date(v) -> str | None:
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

def build_name_to_craft_from_work(df: pd.DataFrame) -> Dict[str, str]:
    cols = {c.lower(): c for c in df.columns}
    name_col = cols.get("name")
    craft_col = cols.get("craft")
    lead_col = cols.get("lead area") or cols.get("lead_area")
    if not name_col:
        return {}
    source = craft_col or lead_col
    if not source:
        return {}
    tmp = df[[name_col, source]].copy()
    tmp[name_col] = tmp[name_col].astype(str).str.strip()
    tmp[source] = tmp[source].astype(str).str.strip()
    mapping = {}
    for nm, g in tmp.groupby(name_col):
        vs = [v for v in g[source].tolist() if str(v).strip()]
        if not vs:
            continue
        most = max(set(vs), key=vs.count)
        mapping[name_key(nm)] = str(most)
    return mapping
    return " ".join(str(s).strip().upper().split())

# ------------------------ Mapping from Address Book ------------------------
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
        nm = row[name_col]
        cd = row[craft_desc_col]
        if nm in mapping and mapping[nm] != cd:
            conflicts.append(f"{nm}: '{mapping[nm]}' vs '{cd}'")
        else:
            mapping[nm] = cd

    return mapping, conflicts

# ------------------------ Area categorization & filter ------------------------
def craft_category(desc: str) -> str:
    s = (desc or "").lower()
    if any(k in s for k in ["hvac", "utility", "utilities", "wtp", "water treatment"]):
        return "HVAC & Utilities"
    if "elec" in s or "elect" in s:
        return "Electrical"
    if "mech" in s:
        return "Mechanical"
    return "Other"

def area_passes(desc: str, mode: str) -> bool:
    if mode == "All": return True
    return craft_category(desc) == mode

# ------------------------ Report Logic ------------------------
def build_report(df: pd.DataFrame, selected_date: str, name_to_craft: Dict[str, str], selected_groups: list[str] | list):
    df = df.copy()
    df["__ProdDate"] = df["Production Date"].apply(normalize_excel_date)
    df = df[df["__ProdDate"] == selected_date]

    df["__NameKey"] = df["Name"].astype(str).map(name_key)
    df["__CraftDesc"] = df["__NameKey"].map(name_to_craft).fillna("(Unmapped Name)")

    # Filter by area selection
    df = df[df["__CraftDesc"].map(lambda x: area_passes(x, selected_groups))]

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
    for (_, _, _), v in groups.items():
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
    for c in list(crafts.keys()):
        crafts[c].sort(key=lambda r: wo_val(r.get("Work Order #", "")))

    day_names = set(df["__NameKey"].tolist())
    mapped_names = set(name_to_craft.keys())
    unmapped = sorted(n for n in day_names if n not in mapped_names)

    return crafts, unmapped, df

# ------------------------ Summary for Cover Page ------------------------
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

# ------------------------ Coloring logic ------------------------
def label_color(label: str) -> str | None:
    l = (label or "").strip().lower()
    if l == "inspection maintenance order":
        return "#2ca02c"  # green
    if l == "pm restore/replace":
        return "#228b22"  # green (darker)
    if l == "emergency order":
        return "#d62728"  # red
    if l == "break in":
        return "#b22222"  # dark red
    return None

def colors_for_labels(labels: List[str]) -> List[str]:
    out = []
    cyc = itertools.cycle(plt.get_cmap("tab20").colors)
    for lab in labels:
        c = label_color(lab)
        if c is None:
            c = next(cyc)
        out.append(c)
    return out

def render_pie_pages(summary: Dict[str, Dict[str, float]], selected_date: str, label_mode: str, craft_order: List[str]) -> List[bytes]:
    order_map = {g.lower(): i for i, g in enumerate(craft_order)}
    items = sorted(summary.items(), key=lambda kv: (order_map.get(craft_category(kv[0]).lower(), 999), kv[0]))
    pies_per_page = 6
    pages: List[bytes] = []
    for i in range(0, len(items), pies_per_page):
        chunk = items[i:i+pies_per_page]
        fig, axs = plt.subplots(2, 3, figsize=(11, 8.5), constrained_layout=True)
        title_unit = "%" if label_mode == "percent" else "hours"
        fig.suptitle(f"Hours by Work Order Type per Area — {selected_date} ({title_unit})", fontsize=14)
        axs = axs.flatten()
        for ax in axs: ax.axis("off")
        for ax, (area, typemap) in zip(axs, chunk):
            labels = list(typemap.keys())
            sizes = [max(0.0, float(typemap[k])) for k in labels]
            if sum(sizes) <= 0:
                ax.text(0.5, 0.5, f"{area}\n(no hours)", ha="center", va="center", fontsize=10)
                continue
            total0 = sum(sizes)
            lbl2, sz2, other = [], [], 0.0
            for l, s in zip(labels, sizes):
                if total0 > 0 and (s/total0) < 0.02: other += s
                else: lbl2.append(l); sz2.append(s)
            if other > 0: lbl2.append("Other"); sz2.append(other)
            total = sum(sz2)

            # Colors with special rules
            cols = colors_for_labels(lbl2)

            if label_mode == "percent": autopct = "%1.0f%%"
            else:
                def _autopct(pct, _total=total): return f"{_total * pct / 100.0:.1f}h"
                autopct = _autopct
            wedges, texts, autotexts = ax.pie(sz2, autopct=autopct, startangle=90, counterclock=False, colors=cols)
            ax.set_title(f"{area} — {total:.2f} h", fontsize=10)
            if len(lbl2) <= 6: ax.legend(wedges, lbl2, fontsize=7, loc="lower center", bbox_to_anchor=(0.5, -0.15), ncol=2)
        buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=150); plt.close(fig); pages.append(buf.getvalue())
    return pages

# ------------------------ PDF Helpers ------------------------
def _soft_wrap_for_pdf(s: str, max_chars: int = 600, unbroken_chunk: int = 40) -> str:
    if s is None: return ""
    s = str(s)
    if not s: return ""
    def _break_token(m):
        token = m.group(0)
        return " ".join(token[i:i+unbroken_chunk] for i in range(0, len(token), unbroken_chunk))
    s = re.sub(r"\S{" + str(unbroken_chunk) + r",}", _break_token, s)
    if len(s) > max_chars:
        s = s[:max_chars - 1] + "…"
    return s

# ------------------------ Column Auto-sizing ------------------------
def _compute_rl_col_widths(rows: List[List[str]], page_inner_width: float) -> List[float]:
    minw = [120, 90, 90, 140, 240, 240]; pad = 14
    naturals = []
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

# ------------------------ PDF Output ------------------------
def make_pdf(selected_date: str, crafts: Dict[str, List[Dict[str, Any]]], cover_summary: Dict[str, Dict[str, float]], label_mode: str, craft_order: List[str], cap_desc=600, cap_prob=600) -> bytes:
    if PDF_ENGINE == "reportlab":
        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=landscape(letter), leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24)
        styles = getSampleStyleSheet()
        title_style = styles["Title"]; header_style = styles["Heading2"]
        body8 = ParagraphStyle("Body8", parent=styles["BodyText"], fontName="Helvetica", fontSize=8, leading=10)
        table_style = TableStyle([
            ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
            ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
            ("ALIGN", (0,0), (-1,0), "LEFT"),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,0), 9),
            ("FONTSIZE", (0,1), (-1,-1), 8),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
        ])

        story: List = []

        # --- COVER PAGES ---
        pie_pages = render_pie_pages(cover_summary, selected_date, label_mode, craft_order)
        for idx, png in enumerate(pie_pages):
            img = RLImage(io.BytesIO(png))
            iw, ih = getattr(img, 'imageWidth', None), getattr(img, 'imageHeight', None)
            if iw is None or ih is None:
                iw, ih = img.wrap(0, 0)
            max_w, max_h = doc.width, doc.height
            scale = min(max_w / iw, max_h / ih) * 0.95  # headroom
            img.drawWidth = iw * scale; img.drawHeight = ih * scale
            story.append(img)
            if idx < len(pie_pages) - 1:
                story.append(PageBreak())

        # --- DETAILED TABLES ---
        story += [Paragraph(f"Daily Report — {selected_date}", title_style), Spacer(1, 6),
                  Paragraph("Sorted by Work Order # within each craft", styles["Normal"]), Spacer(1, 12)]

        page_inner_width = doc.width
        for craft, rows in crafts.items():
            story.append(Paragraph(str(craft), header_style))
            matrix = [["Name", "Work Order #", "Sum of Hours", "Type", "Description", "Problem"]]
            for r in rows:
                matrix.append([
                    str(r.get("Name","")),
                    str(r.get("Work Order #","")),
                    f'{float(r.get("Sum of Hours",0)):.2f}',
                    str(r.get("Type","")),
                    str(r.get("Description","")),
                    str(r.get("Problem","")),
                ])
            col_widths = _compute_rl_col_widths(matrix, page_inner_width)
            data = [matrix[0]]
            for raw in matrix[1:]:
                name = xml_escape(_soft_wrap_for_pdf(raw[0], max_chars=180))
                wo   = xml_escape(_soft_wrap_for_pdf(raw[1], max_chars=60))
                hrs  = xml_escape(_soft_wrap_for_pdf(raw[2], max_chars=32))
                typ  = xml_escape(_soft_wrap_for_pdf(raw[3], max_chars=220))
                desc = xml_escape(_soft_wrap_for_pdf(raw[4], max_chars=cap_desc))
                prob = xml_escape(_soft_wrap_for_pdf(raw[5], max_chars=cap_prob))
                data.append([
                    Paragraph(name, body8),
                    Paragraph(wo,   body8),
                    Paragraph(hrs,  body8),
                    Paragraph(typ,  body8),
                    Paragraph(desc, body8),
                    Paragraph(prob, body8),
                ])
            tbl = Table(data, repeatRows=1, colWidths=col_widths)
            tbl.setStyle(table_style)
            story.append(tbl)
            story.append(Spacer(1, 10))

        doc.build(story)
        pdf = buf.getvalue(); buf.close(); return pdf

    # --- FPDF fallback ---
    class PDF(FPDF, HTMLMixin): pass

    margin = 24
    pdf = PDF(orientation="L", unit="pt", format="Letter")

    # COVER PAGES
    pie_pages = render_pie_pages(cover_summary, selected_date, label_mode, craft_order)
    for png in pie_pages:
        pdf.add_page()
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp.write(png); tmp.flush(); img_path = tmp.name
        # Fit both width and height with margin
        max_w = pdf.w - 2*margin; max_h = pdf.h - 2*margin
        from PIL import Image as PILImage
        im = PILImage.open(img_path)
        iw, ih = im.size
        scale = min(max_w/iw, max_h/ih) * 0.95
        w = iw * scale; h = ih * scale
        pdf.image(img_path, x=(pdf.w-w)/2, y=(pdf.h-h)/2, w=w, h=h)

    # DETAIL PAGES
    pdf.add_page(); pdf.set_left_margin(margin); pdf.set_right_margin(margin)
    pdf.set_auto_page_break(auto=True, margin=margin)
    pdf.set_font("Helvetica", "B", 16); pdf.cell(0, 18, f"Daily Report — {selected_date}", ln=1)
    pdf.set_font("Helvetica", "", 10); pdf.cell(0, 14, "Sorted by Work Order # within each craft", ln=1)

    page_inner_width = pdf.w - pdf.l_margin - pdf.r_margin

    def compute_fpdf_widths(rows: List[List[str]]) -> List[float]:
        minw = [120, 90, 90, 140, 240, 240]; pad = 12
        naturals = []; pdf.set_font("Helvetica", "", 8)
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
        matrix = [["Name", "Work Order #", "Sum of Hours", "Type", "Description", "Problem"]]
        for r in rows:
            matrix.append([
                str(r.get("Name","")), str(r.get("Work Order #","")),
                f'{float(r.get("Sum of Hours",0)):.2f}', str(r.get("Type","")),
                str(r.get("Description","")), str(r.get("Problem","")),
            ])
        col_widths = compute_fpdf_widths(matrix)

        # header
        pdf.set_font("Helvetica", "B", 9)
        for w, txt in zip(col_widths, matrix[0]): pdf.cell(w, th, txt, border=1)
        pdf.ln(th)

        # rows (truncate to fit)
        pdf.set_font("Helvetica", "", 8)
        for raw in matrix[1:]:
            fields = [
                (raw[0], 180), (raw[1], 60), (raw[2], 32),
                (raw[3], 220), (raw[4], 600), (raw[5], 600)
            ]
            clipped = []
            for s, cap in fields:
                s = "" if s is None else str(s)
                if len(s) > cap: s = s[:cap-1] + "…"
                clipped.append(s)
            for w, txt in zip(col_widths, clipped): pdf.cell(w, th, txt, border=1)
            pdf.ln(th)

    return bytes(pdf.output(dest="S").encode("latin1"))

# ------------------------ UI ------------------------
st.title("Craft-Based Daily Report (Excel → PDF) — Cover Page & Area Filter")

with st.sidebar:
    st.markdown("**Instructions**")
    st.markdown("1) Upload the **Address Book** (.xlsx) to build the Name → Craft Description mapping.")
    st.markdown("1) Upload the **Time on Work Order** (.xlsx).")
    st.markdown("2) Pick a **Production Date** (MM/DD/YYYY).")
    st.markdown("3) Choose **Cover Labels**: Percent or Hours.")
    st.markdown("4) Use **Group Filter** to focus the display.")
    st.markdown("6) Download PDF. The first page(s) show pies by area; details follow.")

col1, col2, col3 = st.columns([1,1,1])
with col1:
    pass
with col2:
time_file = st.file_uploader("Upload Time on Work Order (.xlsx)", type=["xlsx"], key="time")
with col3:
    label_mode = st.radio("Cover Labels", options=["percent", "hours"], index=0, horizontal=False)

selected_groups = st.multiselect("Group Filter", options=GROUP_FILTER_OPTIONS, default=GROUP_FILTER_OPTIONS)

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
    crafts, unmapped_names, df_filtered = build_report(df, selected_date, addr_map, selected_groups)
    if unmapped_names: st.error("Unmapped Names (from selected date & filter):\n- " + "\n- ".join(unmapped_names))

    cover_summary = summarize_hours_by_type_per_area(df_filtered)

    with st.expander("Show tabular summary"):
        rows = []
        for area, typemap in sorted(cover_summary.items(), key=lambda kv: sum(kv[1].values()), reverse=True):
            total = sum(typemap.values())
            for t, h in sorted(typemap.items(), key=lambda kv: kv[1], reverse=True):
                rows.append({"Area": area, "Type": t, "Hours": round(h, 2), "Area Total": round(total, 2)})
        st.table(pd.DataFrame(rows))

    for craft, rows in crafts.items():
        st.subheader(f"{craft}")
        st.table(pd.DataFrame(rows, columns=["Name","Work Order #","Sum of Hours","Type","Description","Problem"]))

    pdf_bytes = make_pdf(selected_date, crafts, cover_summary, label_mode, CRAFT_ORDER, cap_desc=cap_val, cap_prob=cap_val)
    st.download_button("Download PDF", data=pdf_bytes, file_name=f"nas_report_{selected_date.replace('/', '-')}.pdf", mime="application/pdf")
elif False:
    pass
elif (addr_map is not None) and (df is None):
    st.info("Upload the **Time on Work Order** sheet to continue.")
