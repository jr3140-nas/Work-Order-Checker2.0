# Craft-Based Daily Report (Excel → PDF) — Portrait & Auto-sized Columns

This Streamlit app builds the **Name → Craft Description** mapping at runtime from your Address Book (.xlsx).  
**PDF export is forced to portrait**, and **columns auto-size** to fit the page.

## Workflow
1. Upload **Address Book** (`Name`, `Craft Description`).
2. Upload **Time on Work Order** (third row is headers).
3. Pick **Production Date** (MM/DD/YYYY).
4. Review and **Download PDF**.

### Notes
- ReportLab engine: exact auto-sizing using text metrics and wrapping (`Paragraph`), scaled to fit the printable area.
- FPDF fallback: portrait enforced; widths estimated and applied (wrap/truncation simplified).

## Run locally
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

