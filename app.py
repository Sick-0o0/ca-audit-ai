
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
            "Supplier", "Invoice_Number",
            "Invoice_Date", "Amount"
        ]

        if not all(col in df.columns for col in required):
            st.error(
                "Excel must contain: Supplier, Invoice_Number, "
                "Invoice_Date and Amount."
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
                st.success("No potential duplicates detected.")
            else:
                st.warning(
                    "Potential duplicates require verification."
                )
                st.dataframe(
                    exceptions,
                    use_container_width=True
                )

            output = BytesIO()
            with pd.ExcelWriter(
                output, engine="openpyxl"
            ) as writer:
                exceptions.to_excel(
                    writer,
                    sheet_name="Exceptions",
                    index=False
                )

            st.download_button(
                "Download Audit Report",
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
                    "Example: Explain the risks and suggest "
                    "audit procedures for these invoices."
                )
            )

            if st.button("Ask AI"):
                if not question.strip():
                    st.warning("Please enter a question.")
                else:
                    try:
                        client = genai.Client(
                            api_key=st.secrets["GEMINI_API_KEY"]
                        )

                        # Send only the flagged rows and
                        # the user's question to the model.
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
1. What the findings indicate.
2. Potential audit risks.
3. Relevant audit assertions.
4. Practical follow-up audit procedures.

Do not assume fraud or conclude an error without evidence.
Explain that supporting documents and payment records
must be checked. Do not invent facts or legal references.
"""

                        with st.spinner("AI is analysing..."):
                            response = client.models.generate_content(
                                model="gemini-2.5-flash",
                                contents=prompt
                            )

                        st.subheader("AI Analysis")
                        st.write(response.text)

                    except Exception:
                        st.error(
                            "AI request failed. Check your API "
                            "key, model availability, API quota "
                            "and Streamlit app logs."
                        )

    except Exception:
        st.error(
            "Could not read the Excel file. Please check "
            "that it is a valid .xlsx purchase register."
        )
