
import io
import math
from xml.sax.saxutils import escape as xml_escape
import pandas as pd
import streamlit as st
from datetime import datetime, date
from typing import Dict, Any, List, Tuple

# Prefer ReportLab; fall back to fpdf2 if ReportLab isn't available.
PDF_ENGINE = "reportlab"
try:
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.pagesizes import letter, portrait
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.pdfbase.pdfmetrics import stringWidth
except Exception:
    PDF_ENGINE = "fpdf"
    from fpdf import FPDF, HTMLMixin  # type: ignore

st.set_page_config(page_title="Craft-Based Daily Report", layout="wide")

EXPECTED_TIME_COLS = [
    "AddressBookNumber","Name","Production Date","OrderNumber","Sum of Hours.","Hours Estimated",
    "Status","Type","PMFrequency","Description","Problem","Lead Area","Craft","CostCenter","UnitNumber","StructureTag"
]
EXPECTED_ADDR_COLS = ["Name", "Craft Description"]

# ------------------------ Utilities ------------------------
def normalize_excel_date(v) -> str | None:
    if v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
        return None
    if isinstance(v, (datetime, date)):
        return datetime(v.year, v.month, v.day).strftime("%m/%d/%Y")
    if isinstance(v, (int, float)):
        # Excel serial first
        try:
            d = pd.to_datetime(v, unit="D", origin="1899-12-30")
            return d.strftime("%m/%d/%Y")
        except Exception:
            pass
        # epoch ms fallback
        try:
            d = pd.to_datetime(v, unit="ms", origin="unix")
            return d.strftime("%m/%d/%Y")
        except Exception:
            pass
    # strings
    try:
        d = pd.to_datetime(str(v), errors="coerce")
        if pd.notnull(d):
            return d.strftime("%m/%d/%Y")
    except Exception:
        pass
    return None

def numberish(v) -> float:
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float("".join(ch for ch in v if (ch.isdigit() or ch in ".-")))
        except Exception:
            return 0.0
    return 0.0

def name_key(s: str) -> str:
    return " ".join(str(s).strip().upper().split())

# ------------------------ Mapping from Address Book ------------------------
def build_name_to_craft(addr_df: pd.DataFrame) -> Tuple[Dict[str, str], List[str]]:
    """Return NAME->Craft Description (uppercased names). Also return duplicate/conflict messages."""
    # Normalize columns (strip whitespace); handles non-string column labels safely.
    addr_df = addr_df.rename(columns=lambda c: str(c).strip())

    # Build a lookup for case-insensitive column matching.
    col_map = {str(c).strip().lower(): str(c).strip() for c in addr_df.columns}
    name_col = col_map.get("name")
    craft_desc_col = col_map.get("craft description") or col_map.get("craft_description")

    if not name_col or not craft_desc_col:
        missing = []
        if not name_col:
            missing.append("Name")
        if not craft_desc_col:
            missing.append("Craft Description")
        raise ValueError(
            f"Address Book missing required columns: {missing}. "
            f"Found columns: {list(addr_df.columns)}"
        )

    # Drop rows with missing name or craft description
    ab = addr_df[[name_col, craft_desc_col]].dropna(how="any").copy()

    # Normalize names to a stable match key (UPPER + single-spaced)
    def name_key(s: str) -> str:
        return " ".join(str(s).strip().upper().split())

    ab[name_col] = ab[name_col].astype(str).map(name_key)
    ab[craft_desc_col] = ab[craft_desc_col].astype(str).str.strip()

    # Build mapping and detect conflicts
    mapping: Dict[str, str] = {}
    conflicts: List[str] = []
    for _, row in ab.iterrows():
        nm = row[name_col]
        cd = row[craft_desc_col]
        if nm in mapping and mapping[nm] != cd:
            conflicts.append(f"{nm}: '{mapping[nm]}' vs '{cd}'")
        else:
            mapping[nm] = cd

    return mapping, conflicts

# ------------------------ Report Logic ------------------------
def build_report(df: pd.DataFrame, selected_date: str, name_to_craft: Dict[str, str]):
    df = df.copy()
    df["__ProdDate"] = df["Production Date"].apply(normalize_excel_date)
    df = df[df["__ProdDate"] == selected_date]

    # craft by name (ignore craft number)
    df["__NameKey"] = df["Name"].astype(str).map(name_key)
    df["__CraftDesc"] = df["__NameKey"].map(name_to_craft).fillna("(Unmapped Name)")

    groups = {}
    for _, r in df.iterrows():
        k = (r["__CraftDesc"], r["__NameKey"], r["OrderNumber"])
        if k not in groups:
            groups[k] = {
                "Craft": r["__CraftDesc"],
                "Name": r["Name"],
                "OrderNumber": r["OrderNumber"],
                "SumOfHours": 0.0,
                "Type": set(),
                "Description": set(),
                "Problem": set(),
            }
        g = groups[k]
        g["SumOfHours"] += numberish(r.get("Sum of Hours.", 0))
        t = str(r.get("Type", "")).strip()
        d = str(r.get("Description", "")).strip()
        p = str(r.get("Problem", "")).strip()
        if t: g["Type"].add(t)
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

    # sort within craft by numeric WO
    def wo_val(s) -> float:
        s = str(s)
        return float(s) if s.isdigit() else float("inf")
    for c in list(crafts.keys()):
        crafts[c].sort(key=lambda r: wo_val(r.get("Work Order #", "")))

    # unmapped names for the selected date
    day_names = set(df["__NameKey"].tolist())
    mapped_names = set(name_to_craft.keys())
    unmapped = sorted(n for n in day_names if n not in mapped_names)

    return crafts, unmapped

# ------------------------ Column Auto-sizing ------------------------
def _compute_rl_col_widths(rows: List[List[str]], page_inner_width: float) -> List[float]:
    """Compute column widths for ReportLab based on content metrics, scaled to fit page width."""
    # minimal widths (pt) to keep columns readable
    minw = [90, 80, 80, 90, 150, 150]  # Name, WO, Sum, Type, Desc, Problem
    pad = 12  # per-cell padding estimate

    # measure natural widths
    naturals = []
    for col_idx in range(len(rows[0])):
        max_w = 0.0
        for r in rows:
            txt = str(r[col_idx]) if r[col_idx] is not None else ""
            # assume body size 8pt for measurement
            max_w = max(max_w, stringWidth(txt, "Helvetica", 8))
        naturals.append(max(max_w + pad, minw[col_idx]))

    total = sum(naturals)
    if total <= page_inner_width:
        return naturals

    # shrink above minimums proportionally
    over = total - page_inner_width
    shrinkable = [max(0.0, naturals[i] - minw[i]) for i in range(len(naturals))]
    total_shrinkable = sum(shrinkable)
    if total_shrinkable <= 0:
        # nothing to shrink; force equal scaling
        scale = page_inner_width / total if total > 0 else 1.0
        return [w * scale for w in naturals]

    widths = []
    for i, w in enumerate(naturals):
        reduce = over * (shrinkable[i] / total_shrinkable) if total_shrinkable > 0 else 0.0
        widths.append(max(minw[i], w - reduce))
    return widths

# ------------------------ PDF Output ------------------------
def make_pdf(selected_date: str, crafts: Dict[str, List[Dict[str, Any]]]) -> bytes:
    if PDF_ENGINE == "reportlab":
        buf = io.BytesIO()
        # Force portrait orientation explicitly
        doc = SimpleDocTemplate(buf, pagesize=portrait(letter), leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
        styles = getSampleStyleSheet()
        title_style = styles["Title"]
        header_style = styles["Heading2"]
        table_style = TableStyle([
            ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
            ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
            ("ALIGN", (0,0), (-1,0), "LEFT"),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,0), 9),
            ("FONTSIZE", (0,1), (-1,-1), 8),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
        ])

        # smaller body paragraph style for wrapping
        body8 = ParagraphStyle(
            "Body8",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=8,
            leading=10,
        )

        story: List = []
        story += [Paragraph(f"Daily Report — {selected_date}", title_style), Spacer(1, 6),
                  Paragraph("Sorted by Work Order # within each craft", styles["Normal"]), Spacer(1, 12)]

        page_inner_width = doc.width  # available width after margins

        for craft, rows in crafts.items():
            story.append(Paragraph(str(craft), header_style))

            # Build raw rows (strings) for width measurement
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

            # Convert data cells to Paragraph for wrapping
            data = [matrix[0]]  # header row as plain text
            for raw in matrix[1:]:
                data.append([
                    Paragraph(xml_escape(raw[0]), body8),
                    Paragraph(xml_escape(raw[1]), body8),
                    Paragraph(xml_escape(raw[2]), body8),
                    Paragraph(xml_escape(raw[3]), body8),
                    Paragraph(xml_escape(raw[4]), body8),
                    Paragraph(xml_escape(raw[5]), body8),
                ])

            tbl = Table(data, repeatRows=1, colWidths=col_widths)
            tbl.setStyle(table_style)
            story.append(tbl)
            story.append(Spacer(1, 10))

        doc.build(story)
        pdf = buf.getvalue()
        buf.close()
        return pdf

    # fpdf2 fallback — force portrait and attempt width hints
    class PDF(FPDF, HTMLMixin):
        pass

    pdf = PDF(orientation="P", unit="pt", format="Letter")
    pdf.set_auto_page_break(auto=True, margin=36)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 18, f"Daily Report — {selected_date}", ln=1)
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 14, "Sorted by Work Order # within each craft", ln=1)

    page_inner_width = pdf.w - 72  # 36pt margins left+right

    def compute_fpdf_widths(rows: List[List[str]]) -> List[float]:
        # minimal widths to avoid unreadable columns
        minw = [90, 80, 80, 90, 150, 150]
        pad = 12
        naturals = []
        pdf.set_font("Helvetica", "", 8)
        for col_idx in range(len(rows[0])):
            max_w = 0.0
            for r in rows:
                txt = str(r[col_idx]) if r[col_idx] is not None else ""
                max_w = max(max_w, pdf.get_string_width(txt))
            naturals.append(max(max_w + pad, minw[col_idx]))
        total = sum(naturals)
        if total <= page_inner_width:
            return naturals
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
        return widths

    for craft, rows in crafts.items():
        pdf.ln(6)
        pdf.set_font("Helvetica", "B", 13)
        pdf.cell(0, 16, str(craft), ln=1)

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
        col_widths = compute_fpdf_widths(matrix)

        # Render table header
        pdf.set_font("Helvetica", "B", 9)
        th = 14
        x0 = pdf.get_x()
        y0 = pdf.get_y()
        headers = matrix[0]
        for w, txt in zip(col_widths, headers):
            pdf.cell(w, th, txt, border=1)
        pdf.ln(th)

        # Render rows (simple, no wrapping to keep code compact)
        pdf.set_font("Helvetica", "", 8)
        for raw in matrix[1:]:
            for w, txt in zip(col_widths, raw):
                # truncate long text for fallback to keep layout; primary engine (ReportLab) handles wrapping
                s = str(txt)
                max_chars = int(max(5, w / 4.5))  # rough fit heuristic
                if len(s) > max_chars:
                    s = s[:max_chars-1] + "…"
                pdf.cell(w, th, s, border=1)
            pdf.ln(th)

    return bytes(pdf.output(dest="S").encode("latin1"))

# ------------------------ UI ------------------------
st.title("Craft-Based Daily Report (Excel → PDF) — Portrait & Auto-sized Columns")

with st.sidebar:
    st.markdown("**Instructions**")
    st.markdown("1) Upload the **Address Book** (.xlsx) to build the Name → Craft Description mapping.")
    st.markdown("2) Upload the **Time on Work Order** (.xlsx).")
    st.markdown("3) Pick a **Production Date** (MM/DD/YYYY).")
    st.markdown("4) Review and **Download PDF**.")
    st.markdown(f"PDF engine in use: **{PDF_ENGINE}** (portrait forced; columns auto-sized)")


col1, col2 = st.columns(2)
with col1:
    addr_file = st.file_uploader("Upload Address Book (.xlsx)", type=["xlsx"], key="addr")
with col2:
    time_file = st.file_uploader("Upload Time on Work Order (.xlsx)", type=["xlsx"], key="time")

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

df = None
dates: List[str] = []
if time_file is not None:
    try:
        df = pd.read_excel(time_file, header=2)  # 3rd row as header
        df.columns = [str(c).strip() for c in df.columns]
        missing = [c for c in EXPECTED_TIME_COLS if c not in df.columns]
        if missing:
            st.error(f"Time sheet missing expected columns: {missing}")
        else:
            dates = sorted({d for d in (df["Production Date"].apply(normalize_excel_date).dropna().tolist())})
            st.caption(f"Detected dates: {(dates[0] if dates else '—')} → {(dates[-1] if dates else '—')} • Unique dates: {len(dates)}")
    except Exception as e:
        st.exception(e)

selected_date = st.selectbox("Production Date", options=(dates if dates else [""]), index=(len(dates)-1 if dates else 0))

if df is not None and addr_map is not None and selected_date:
    crafts, unmapped_names = build_report(df, selected_date, addr_map)

    if unmapped_names:
        st.error("Unmapped Names (from selected date):\n- " + "\n- ".join(unmapped_names))

    for craft, rows in crafts.items():
        st.subheader(craft)
        st.dataframe(pd.DataFrame(rows, columns=["Name","Work Order #","Sum of Hours","Type","Description","Problem"]))

    pdf_bytes = make_pdf(selected_date, crafts)
    st.download_button("Download PDF", data=pdf_bytes, file_name=f"nas_report_{selected_date.replace('/', '-')}.pdf", mime="application/pdf")
elif (df is not None) and (addr_map is None):
    st.info("Upload the **Address Book** to generate the Name → Craft Description mapping.")
elif (addr_map is not None) and (df is None):
    st.info("Upload the **Time on Work Order** sheet to continue.")
