import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.title("ASN PDF → Excel")

uploaded_files = st.file_uploader(
    "Upload Delivery Note PDF",
    type="pdf",
    accept_multiple_files=True
)

data = []

if uploaded_files:
    for file in uploaded_files:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                lines = text.split("\n")

                for line in lines:
                    if "PC" in line:
                        parts = line.split()
                        if len(parts) > 5:
                            item = parts[2]
                            rev = parts[3]
                            qty = parts[4]

                            data.append({
                                "Item": item,
                                "Rev": rev,
                                "Qty": qty
                            })

df = pd.DataFrame(data)

if not df.empty:
    st.dataframe(df)

    output = BytesIO()
    df.to_excel(output, index=False)

    st.download_button(
        "Download Excel",
        output.getvalue(),
        "ASN_Result.xlsx"
    )
