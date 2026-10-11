
import streamlit as st
import pandas as pd
from io import BytesIO
from google import genai

st.set_page_config(
    page_title="CA Audit AI",
    page_icon="📊",
    layout="wide"
)

st.title("CA Audit AI")
st.caption("Accounting & Audit Analytics")

st.header("1. Duplicate Invoice Detection")

uploaded_file = st.file_uploader(
    "Upload Purchase Register",
    type=["xlsx"]
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)

        required = [
            "Supplier",
            "Invoice_Number",
            "Invoice_Date",
            "Amount"
        ]

        if not all(col in df.columns for col in required):
            st.error(
                "Excel must contain these columns: "
                "Supplier, Invoice_Number, Invoice_Date, Amount"
            )

        else:
            df["Potential_Duplicate"] = df.duplicated(
                subset=["Supplier", "Invoice_Number"],
                keep=False
            )

            exceptions = df[
                df["Potential_Duplicate"]
            ].copy()

            st.metric("Total Invoices", len(df))
            st.metric(
                "Potential Duplicate Entries",
                len(exceptions)
            )

            if exceptions.empty:
                st.success(
                    "No potential duplicate invoices detected."
                )
            else:
                st.warning(
                    "Potential duplicates found. "
                    "Verify supporting documents before concluding."
                )
                st.dataframe(
                    exceptions,
                    use_container_width=True
                )

            output = BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl"
            ) as writer:
                df.to_excel(
                    writer,
                    sheet_name="All Invoices",
                    index=False
                )
                exceptions.to_excel(
                    writer,
                    sheet_name="Exceptions",
                    index=False
                )

            st.download_button(
                label="Download Audit Report",
                data=output.getvalue(),
                file_name="Duplicate_Invoice_Report.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                )
            )

            st.divider()
            st.header("2. Ask Your AI Audit Assistant")

            question = st.text_area(
                "Ask a question about the audit findings",
                placeholder=(
                    "Explain the risks and suggest audit procedures "
                    "for these duplicate invoices."
                )
            )

            if st.button("Ask AI"):
                if not question.strip():
                    st.warning("Please enter a question.")

                else:
                    try:
                        api_key = st.secrets["GEMINI_API_KEY"]

                        client = genai.Client(
                            api_key=api_key
                        )

                        findings = exceptions[
                            required
                        ].to_dict(orient="records")

                        prompt = f"""
You are an audit assistant supporting a Chartered Accountant.

Analyse these potential duplicate invoice findings:
{findings}

User question:
{question}

Provide:
1. Explanation of the findings.
2. Potential audit risks.
3. Relevant audit assertions.
4. Practical follow-up audit procedures.

Do not assume fraud or conclude an error without evidence.
Recommend checking original invoices, purchase orders,
goods receipt notes, ledger postings and payment records
where relevant.

Do not invent facts or legal references.
Clearly distinguish potential exceptions from confirmed errors.
"""

                        with st.spinner(
                            "AI is analysing the findings..."
                        ):
                            response = client.models.generate_content(
                                model="gemini-3.8-flash",
                                contents=prompt
                            )

                        if response.text:
                            st.subheader("AI Analysis")
                            st.markdown(response.text)
                        else:
                            st.warning(
                                "The AI returned an empty response."
                            )

                    except Exception as e:
                        st.error("AI request failed.")
                        st.code(
                            f"{type(e).__name__}: {e}"
                        )

    except Exception as e:
        st.error(
            "Could not read the Excel file. "
            "Please check that it is a valid .xlsx purchase register."
        )
        st.code(f"{type(e).__name__}: {e}")
