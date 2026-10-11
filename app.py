
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

# ---------- COMMON AI FUNCTION ----------

def ask_ai(question, findings, module_name):
    try:
        client = genai.Client(
            api_key=st.secrets["GEMINI_API_KEY"]
        )

        prompt = f"""
You are an audit assistant supporting a Chartered Accountant.

Audit module: {module_name}
Potential exceptions:
{findings}

User question:
{question}

Provide:
1. Explanation of findings.
2. Potential audit risks.
3. Relevant audit assertions.
4. Practical further audit procedures.

Do not assume fraud or conclude an error without evidence.
Distinguish potential exceptions from confirmed errors.
Do not invent facts or legal references.
"""

        response = client.models.generate_content(
            model="gemini-3.8-flash",
            contents=prompt
        )

        if response.text:
            st.subheader("AI Analysis")
            st.markdown(response.text)
        else:
            st.warning("AI returned an empty response.")

    except Exception as e:
        st.error("AI request failed.")
        st.code(f"{type(e).__name__}: {e}")


# ---------- MODULE 1: DUPLICATE INVOICES ----------

st.header("1. Duplicate Invoice Detection")

purchase_file = st.file_uploader(
    "Upload Purchase Register",
    type=["xlsx"],
    key="purchase_file"
)

if purchase_file is not None:
    try:
        purchase = pd.read_excel(purchase_file)

        purchase_columns = [
            "Supplier",
            "Invoice_Number",
            "Invoice_Date",
            "Amount"
        ]

        if not all(c in purchase.columns for c in purchase_columns):
            st.error(
                "Required columns: Supplier, Invoice_Number, "
                "Invoice_Date, Amount"
            )
        else:
            purchase["Potential_Duplicate"] = purchase.duplicated(
                subset=["Supplier", "Invoice_Number"],
                keep=False
            )

            duplicates = purchase[
                purchase["Potential_Duplicate"]
            ].copy()

            st.metric("Total Invoices", len(purchase))
            st.metric(
                "Potential Duplicate Entries",
                len(duplicates)
            )

            if duplicates.empty:
                st.success("No potential duplicates detected.")
            else:
                st.warning(
                    "Potential duplicates require verification."
                )
                st.dataframe(
                    duplicates,
                    use_container_width=True
                )

            report = BytesIO()

            with pd.ExcelWriter(report, engine="openpyxl") as writer:
                purchase.to_excel(
                    writer, sheet_name="All Invoices", index=False
                )
                duplicates.to_excel(
                    writer, sheet_name="Exceptions", index=False
                )

            st.download_button(
                "Download Invoice Audit Report",
                data=report.getvalue(),
                file_name="Invoice_Audit_Report.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                key="download_invoice_report"
            )

            st.subheader("Ask AI about Invoice Findings")

            invoice_question = st.text_area(
                "Your question",
                key="invoice_question",
                placeholder="Explain the risks and audit procedures."
            )

            if st.button("Analyse Invoice Findings with AI"):
                if not invoice_question.strip():
                    st.warning("Enter a question first.")
                elif duplicates.empty:
                    st.info("There are no flagged invoice entries to analyse.")
                else:
                    ask_ai(
                        invoice_question,
                        duplicates[purchase_columns].to_dict(
                            orient="records"
                        ),
                        "Duplicate Invoice Testing"
                    )

    except Exception as e:
        st.error("Could not process the purchase register.")
        st.code(f"{type(e).__name__}: {e}")


st.divider()

# ---------- MODULE 2: JOURNAL ENTRY TESTING ----------

st.header("2. Journal Entry Testing")

journal_file = st.file_uploader(
    "Upload Journal Entries Excel File",
    type=["xlsx"],
    key="journal_file"
)

if journal_file is not None:
    try:
        journals = pd.read_excel(journal_file)

        journal_columns = [
            "Entry_ID",
            "Posting_Date",
            "Debit_Account",
            "Credit_Account",
            "Amount",
            "Description",
            "Prepared_By",
            "Approved_By"
        ]

        if not all(c in journals.columns for c in journal_columns):
            st.error(
                "Required columns: Entry_ID, Posting_Date, "
                "Debit_Account, Credit_Account, Amount, "
                "Description, Prepared_By, Approved_By"
            )
        else:
            journals["Posting_Date"] = pd.to_datetime(
                journals["Posting_Date"], errors="coerce"
            )

            journals["Amount"] = pd.to_numeric(
                journals["Amount"], errors="coerce"
            )

            debit_account = (
                journals["Debit_Account"].fillna("").astype(str).str.lower()
            )
            credit_account = (
                journals["Credit_Account"].fillna("").astype(str).str.lower()
            )

            prepared = (
                journals["Prepared_By"].fillna("").astype(str).str.strip()
            )
            approved = (
                journals["Approved_By"].fillna("").astype(str).str.strip()
            )

            journals["Round_Figure"] = (
                journals["Amount"].gt(0)
                & journals["Amount"].mod(10000).eq(0)
            )

            journals["Period_End"] = (
                journals["Posting_Date"].dt.day.ge(28)
            )

            journals["Weekend_Posting"] = (
                journals["Posting_Date"].dt.dayofweek.ge(5)
            )

            journals["Missing_Approval"] = approved.eq("")

            journals["Self_Approval"] = (
                prepared.ne("")
                & approved.ne("")
                & prepared.str.casefold().eq(approved.str.casefold())
            )

            journals["Suspense_Account"] = (
                debit_account.str.contains("suspense", na=False)
                | credit_account.str.contains("suspense", na=False)
            )

            rule_columns = [
                "Round_Figure",
                "Period_End",
                "Weekend_Posting",
                "Missing_Approval",
                "Self_Approval",
                "Suspense_Account"
            ]

            journals["Flag_Count"] = journals[rule_columns].sum(axis=1)

            flagged = journals[
                journals["Flag_Count"] > 0
            ].copy()

            st.metric("Total Journal Entries", len(journals))
            st.metric("Potential Exceptions", len(flagged))

            st.caption(
                "Flags are screening indicators, not proof of "
                "misstatement or fraud."
            )

            if flagged.empty:
                st.success("No entries were flagged by these rules.")
            else:
                st.dataframe(
                    flagged,
                    use_container_width=True
                )

            st.subheader("Exception Summary")

            for rule in rule_columns:
                st.write(
                    f"{rule.replace('_', ' ')}: "
                    f"{int(journals[rule].sum())} entries"
                )

            journal_report = BytesIO()

            with pd.ExcelWriter(
                journal_report, engine="openpyxl"
            ) as writer:
                journals.to_excel(
                    writer, sheet_name="All Entries", index=False
                )
                flagged.to_excel(
                    writer, sheet_name="Exceptions", index=False
                )

            st.download_button(
                "Download Journal Audit Report",
                data=journal_report.getvalue(),
                file_name="Journal_Entry_Audit_Report.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                key="download_journal_report"
            )

            st.subheader("Ask AI about Journal Findings")

            journal_question = st.text_area(
                "Your question",
                key="journal_question",
                placeholder=(
                    "Explain the flagged entries and suggest "
                    "further audit procedures."
                )
            )

            if st.button("Analyse Journal Findings with AI"):
                if not journal_question.strip():
                    st.warning("Enter a question first.")
                elif flagged.empty:
                    st.info("No flagged entries to analyse.")
                else:
                    safe_findings = flagged[
                        journal_columns + rule_columns + ["Flag_Count"]
                    ].to_dict(orient="records")

                    ask_ai(
                        journal_question,
                        safe_findings,
                        "Journal Entry Testing"
                    )

    except Exception as e:
        st.error("Could not process the journal entries file.")
        st.code(f"{type(e).__name__}: {e}")
