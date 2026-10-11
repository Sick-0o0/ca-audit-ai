import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="CA Audit AI",
    page_icon="📊",
    layout="wide"
)

st.title("CA Audit AI")
st.caption("Accounting & Audit Analytics")

st.subheader("Duplicate Invoice Detection")
st.write(
    "Upload a purchase register in Excel format "
    "to identify potentially duplicated invoices."
)

uploaded_file = st.file_uploader(
    "Upload Purchase Register",
    type=["xlsx"]
)

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    required_columns = [
        "Supplier", "Invoice_Number", "Invoice_Date", "Amount"
    ]

    if not all(col in df.columns for col in required_columns):
        st.error(
            "Required columns: Supplier, Invoice_Number, "
            "Invoice_Date, Amount"
        )
    else:
        if st.button("Analyse Invoices"):
            df["Potential_Duplicate"] = df.duplicated(
                subset=["Supplier", "Invoice_Number"],
                keep=False
            )

            exceptions = df[df["Potential_Duplicate"]].copy()

            st.metric("Total Invoices", len(df))
            st.metric("Potential Duplicate Entries", len(exceptions))

            if exceptions.empty:
                st.success("No potential duplicates detected.")
            else:
                st.warning(
                    "Potential duplicates found. "
                    "Verify supporting documents before concluding."
                )
                st.dataframe(exceptions, use_container_width=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                exceptions.to_excel(
                    writer, sheet_name="Exceptions", index=False
                )

            st.download_button(
                "Download Audit Report",
                data=output.getvalue(),
                file_name="Duplicate_Invoice_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
