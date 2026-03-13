import io
import re
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="ASN PDF Carton Calculator", layout="wide")

ASN_RE = re.compile(r"ASN\s*No\s*:\s*((?:CH|CR)\d+)", re.IGNORECASE)


def clean(v):
    return "" if v is None else str(v).replace("\r", "").strip()


def safe_int(v):
    s = clean(v).replace(",", "")
    if not s:
        return None
    try:
        return int(float(s))
    except Exception:
        return None


def norm_rev(v):
    s = clean(v)
    return s.zfill(2) if s else ""


@st.cache_data(show_spinner=False)
def load_packing_db_from_bytes(xlsx_bytes: bytes):
    wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)

    if "LOOKUP_TABLE" in wb.sheetnames:
        ws = wb["LOOKUP_TABLE"]
    elif "MASTER_DB" in wb.sheetnames:
        ws = wb["MASTER_DB"]
    else:
        ws = wb[wb.sheetnames[0]]

    lookup = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        item = clean(row[0] if len(row) > 0 else "")
        rev = norm_rev(row[1] if len(row) > 1 else "")
        pcs = safe_int(row[2] if len(row) > 2 else "")

        if item and pcs:
            lookup[(item, rev)] = pcs

    return lookup


def detect_asn(text, fallback_name):
    m = ASN_RE.search(text or "")
    return m.group(1) if m else fallback_name


def table_has_required_headers(header_row):
    hdr = [clean(c).lower().replace("\n", " ") for c in header_row]
    return ("item no." in hdr) and ("quantity" in hdr)


def extract_rows_from_pdf(uploaded_file, packing_lookup):
    pdf_bytes = uploaded_file.read()
    pdf_name = uploaded_file.name
    all_rows = []
    asn = ""
    total_qty = None

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            if not asn:
                asn = detect_asn(text, Path(pdf_name).stem)

            for table in page.extract_tables() or []:
                if not table or not table[0]:
                    continue

                if not table_has_required_headers(table[0]):
                    continue

                for raw in table[1:]:
                    if not raw:
                        continue

                    row = [clean(c) for c in raw]
                    first = clean(row[0]).lower()

                    if first.startswith("total quantity"):
                        for cell in row[1:]:
                            q = safe_int(cell)
                            if q is not None:
                                total_qty = q
                                break
                        continue

                    if not clean(row[0]).isdigit():
                        continue

                    item = clean(row[2] if len(row) > 2 else "")
                    rev = norm_rev(row[3] if len(row) > 3 else "")
                    qty = safe_int(row[4] if len(row) > 4 else "") or 0
                    uom = clean(row[5] if len(row) > 5 else "")
                    net_weight = clean(row[6] if len(row) > 6 else "")
                    lot_invoice = clean(row[9] if len(row) > 9 else "").replace("\n", " | ")
                    line_no = clean(row[10] if len(row) > 10 else "")

                    pcs_ctn = packing_lookup.get((item, rev))
                    cartons = qty // pcs_ctn if pcs_ctn else None
                    loose = qty % pcs_ctn if pcs_ctn else None
                    status = "OK" if pcs_ctn else "NO PACKING DB"

                    all_rows.append({
                        "ASN": asn,
                        "Seq": safe_int(row[0]),
                        "PO No": clean(row[1] if len(row) > 1 else ""),
                        "Item No": item,
                        "Rev": rev,
                        "Quantity": qty,
                        "Uom": uom,
                        "Net Weight (KG)": net_weight,
                        "PCS/CTN": pcs_ctn if pcs_ctn else "",
                        "Cartons": cartons if pcs_ctn else "",
                        "Loose PCS": loose if pcs_ctn else "",
                        "Status": status,
                        "Lot/Invoice": lot_invoice,
                        "Line No": line_no,
                        "Source PDF": pdf_name,
                        "Page": page_num,
                    })

    if not asn:
        asn = Path(pdf_name).stem

    return asn, all_rows, total_qty


def build_excel(line_df, summary_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        line_df.to_excel(writer, index=False, sheet_name="ASN_Line_Data")
        summary_df.to_excel(writer, index=False, sheet_name="ASN_Summary")
    output.seek(0)
    return output


st.title("ASN PDF Carton Calculator")
st.caption("Upload Delivery Note PDF + Packing DB -> tu tinh so thung chan, PCS le, Line No.")

col1, col2 = st.columns(2)

with col1:
    pdf_files = st.file_uploader(
        "Upload Delivery Note PDF",
        type=["pdf"],
        accept_multiple_files=True
    )

with col2:
    packing_db = st.file_uploader(
        "Upload Packing DB (Auto_Carton_Calculator.xlsx)",
        type=["xlsx"]
    )

run = st.button("Run", use_container_width=True)

if run:
    if not pdf_files:
        st.error("Ban chua upload PDF.")
        st.stop()

    if not packing_db:
        st.error("Ban chua upload file Packing DB.")
        st.stop()

    with st.spinner("Dang doc Packing DB..."):
        packing_lookup = load_packing_db_from_bytes(packing_db.read())

    if not packing_lookup:
        st.error("Khong doc duoc Packing DB.")
        st.stop()

    all_rows = []
    summary_rows = []

    with st.spinner("Dang doc PDF va tinh cartons..."):
        for pdf in pdf_files:
            pdf.seek(0)
            asn, rows, total_qty = extract_rows_from_pdf(pdf, packing_lookup)

            all_rows.extend(rows)

            total_cartons = sum(int(r["Cartons"] or 0) for r in rows)
            total_loose = sum(int(r["Loose PCS"] or 0) for r in rows)
            missing_lines = sum(1 for r in rows if r["Status"] != "OK")
            total_lines = len(rows)

            summary_rows.append({
                "ASN": asn,
                "Source PDF": pdf.name,
                "Lines": total_lines,
                "Total Quantity": total_qty if total_qty is not None else "",
                "Total Cartons": total_cartons,
                "Total Loose PCS": total_loose,
                "Missing Packing Lines": missing_lines,
            })

    line_df = pd.DataFrame(all_rows)
    summary_df = pd.DataFrame(summary_rows)

    st.success("Da xu ly xong.")

    st.subheader("ASN Summary")
    st.dataframe(summary_df, use_container_width=True)

    st.subheader("ASN Line Data")
    st.dataframe(line_df, use_container_width=True)

    excel_file = build_excel(line_df, summary_df)

    st.download_button(
        "Download ASN_Result.xlsx",
        data=excel_file,
        file_name="ASN_Result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with st.expander("Yeu cau Packing DB"):
    st.write("Lookup theo: Item No + Rev")
    st.write("Can co 3 cot chinh:")
    st.write("- Item No")
    st.write("- Rev")
    st.write("- Packing / PCS per carton")
