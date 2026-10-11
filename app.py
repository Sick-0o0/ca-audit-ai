
import time
from io import BytesIO

import pandas as pd
import streamlit as st
from google import genai


# --------------------------------------------------
# PAGE CONFIGURATION
# --------------------------------------------------

st.set_page_config(
    page_title="CA Audit AI",
    page_icon="📊",
    layout="wide"
)

st.title("CA Audit AI")
st.caption("Excel-based audit checks and AI-assisted analysis")


# --------------------------------------------------
# GEMINI AI WITH AUTOMATIC RETRIES
# --------------------------------------------------

MODEL_NAME = "gemini-3.8-flash"


def generate_ai_response(client, prompt):
    """
    Retry temporary Gemini errors up to 4 times.
    Waits 2, 4 and 8 seconds between attempts.
    """

    last_error = None

    for attempt in range(4):
        try:
            response = client.models.generate_content(
                model=MODEL_NAME,
                contents=prompt
            )
            return response

        except Exception as error:
            last_error = error
            error_text = str(error).upper()

            temporary_error = any(
                code in error_text
                for code in [
                    "503",
                    "UNAVAILABLE",
                    "429",
                    "RESOURCE_EXHAUSTED",
                    "500",
                    "INTERNAL",
                    "502",
                    "504",
                    "TIMEOUT"
                ]
            )

            if not temporary_error:
                raise

            if attempt < 3:
                wait_seconds = 2 ** (attempt + 1)

                st.warning(
                    f"AI server is busy. "
                    f"Retrying in {wait_seconds} seconds "
                    f"(attempt {attempt + 2} of 4)..."
                )

                time.sleep(wait_seconds)

    raise RuntimeError(
        "Gemini remained unavailable after 4 attempts. "
        "Please try again later. Details: "
        + str(last_error)
    )


def ask_ai(question, findings, module_name):
    """Send audit findings to Gemini and display its response."""

    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        client = genai.Client(api_key=api_key)

        findings_text = findings.to_string(index=False)

        prompt = f"""
You are an audit assistant supporting a Chartered Accountant.

Module: {module_name}

User's question:
{question}

Audit findings:
{findings_text}

Provide a concise, structured response with:
1. Explanation of the findings.
2. Potential audit risks.
3. Relevant financial statement assertions, where applicable.
4. Suggested audit procedures and supporting documents to verify.
5. Practical next steps.

Important:
- A flagged transaction is not automatically fraud.
- Do not claim that fraud has occurred without evidence.
- Do not invent accounting standards, legal provisions or facts.
- Clearly distinguish observations from conclusions.
- Treat these findings as preliminary and requiring verification.
"""

        with st.spinner("Generating AI analysis..."):
            response = generate_ai_response(client, prompt)

        if response.text:
            st.subheader("AI Analysis")
            st.markdown(response.text)
        else:
            st.warning(
                "The AI returned an empty response. "
                "Please try again."
            )

    except Exception as error:
        st.error(
            "AI analysis could not be completed. "
            "Your Excel-based checks can still be reviewed."
        )
        st.code(f"{type(error).__name__}: {error}")


# --------------------------------------------------
# EXCEL DOWNLOAD HELPER
# --------------------------------------------------

def convert_df_to_excel(dataframe, sheet_name):
    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:
        dataframe.to_excel(
            writer,
            index=False,
            sheet_name=sheet_name
        )

    return output.getvalue()


# --------------------------------------------------
# SIDEBAR MODULE SELECTION
# --------------------------------------------------

module = st.sidebar.selectbox(
    "Select Audit Module",
    [
        "Duplicate Invoice Detection",
        "Journal Entry Testing"
    ]
)


# ==================================================
# MODULE 1: DUPLICATE INVOICE DETECTION
# ==================================================

if module == "Duplicate Invoice Detection":

    st.header("Duplicate Invoice Detection")

    st.write(
        "Upload a purchase register to identify repeated "
        "supplier and invoice number combinations."
    )

    uploaded_file = st.file_uploader(
        "Upload Purchase Register Excel",
        type=["xlsx", "xls"],
        key="purchase_register"
    )

    if uploaded_file is not None:

        try:
            df = pd.read_excel(uploaded_file)

            st.subheader("Uploaded Data")
            st.dataframe(df, use_container_width=True)

            required_columns = [
                "Supplier",
                "Invoice_Number",
                "Invoice_Date",
                "Amount"
            ]

            missing_columns = [
                column
                for column in required_columns
                if column not in df.columns
            ]

            if missing_columns:
                st.error(
                    "Missing required columns: "
                    + ", ".join(missing_columns)
                )

                st.info(
                    "Required columns: Supplier, Invoice_Number, "
                    "Invoice_Date, Amount. Rename your Excel "
                    "headers to match these names."
                )

            else:
                df["Supplier"] = (
                    df["Supplier"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                df["Invoice_Number"] = (
                    df["Invoice_Number"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                df["Amount"] = pd.to_numeric(
                    df["Amount"],
                    errors="coerce"
                )

                valid_invoice = (
                    df["Supplier"].ne("")
                    & df["Invoice_Number"].ne("")
                )

                duplicate_mask = (
                    valid_invoice
                    & df.duplicated(
                        subset=[
                            "Supplier",
                            "Invoice_Number"
                        ],
                        keep=False
                    )
                )

                duplicates = df.loc[
                    duplicate_mask
                ].copy()

                total_records = len(df)
                duplicate_rows = len(duplicates)
                duplicate_groups = (
                    duplicates.groupby(
                        ["Supplier", "Invoice_Number"],
                        dropna=False
                    ).ngroups
                    if not duplicates.empty
                    else 0
                )

                col1, col2, col3 = st.columns(3)

                col1.metric(
                    "Total Records",
                    total_records
                )

                col2.metric(
                    "Flagged Rows",
                    duplicate_rows
                )

                col3.metric(
                    "Duplicate Groups",
                    duplicate_groups
                )

                st.subheader("Duplicate Invoice Results")

                if duplicates.empty:
                    st.success(
                        "No duplicate supplier and invoice "
                        "number combinations were detected."
                    )

                else:
                    st.warning(
                        "Potential duplicate invoices detected. "
                        "Verify the supporting documents before "
                        "drawing any conclusion."
                    )

                    st.dataframe(
                        duplicates,
                        use_container_width=True
                    )

                report = df.copy()
                report["Duplicate_Flag"] = duplicate_mask

                st.download_button(
                    label="Download Duplicate Invoice Report",
                    data=convert_df_to_excel(
                        report,
                        "Duplicate_Report"
                    ),
                    file_name="Duplicate_Invoice_Report.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    )
                )

                st.divider()
                st.subheader("Ask AI About These Findings")

                question = st.text_area(
                    "Enter your audit question",
                    value=(
                        "Explain the potential risks in these "
                        "duplicate invoice findings and suggest "
                        "audit procedures."
                    ),
                    key="duplicate_question"
                )

                if st.button(
                    "Analyse Duplicate Invoices with AI",
                    key="duplicate_ai"
                ):

                    if duplicates.empty:
                        st.info(
                            "No duplicate rows were detected. "
                            "You can still review the data manually."
                        )

                    elif not question.strip():
                        st.warning(
                            "Please enter a question first."
                        )

                    else:
                        ask_ai(
                            question,
                            duplicates,
                            "Duplicate Invoice Detection"
                        )

        except Exception as error:
            st.error("Unable to read or process the Excel file.")
            st.code(f"{type(error).__name__}: {error}")


# ==================================================
# MODULE 2: JOURNAL ENTRY TESTING
# ==================================================

elif module == "Journal Entry Testing":

    st.header("Journal Entry Testing")

    st.write(
        "Upload a journal entry listing to identify entries "
        "that may require further audit investigation."
    )

    uploaded_file = st.file_uploader(
        "Upload Journal Entries Excel",
        type=["xlsx", "xls"],
        key="journal_entries"
    )

    if uploaded_file is not None:

        try:
            df = pd.read_excel(uploaded_file)

            st.subheader("Uploaded Data")
            st.dataframe(df, use_container_width=True)

            required_columns = [
                "Entry_ID",
                "Posting_Date",
                "Debit_Account",
                "Credit_Account",
                "Amount",
                "Description",
                "Prepared_By",
                "Approved_By"
            ]

            missing_columns = [
                column
                for column in required_columns
                if column not in df.columns
            ]

            if missing_columns:
                st.error(
                    "Missing required columns: "
                    + ", ".join(missing_columns)
                )

                st.info(
                    "Rename your Excel headers to match the "
                    "required column names shown above."
                )

            else:
                data = df.copy()

                data["Posting_Date"] = pd.to_datetime(
                    data["Posting_Date"],
                    errors="coerce"
                )

                data["Amount"] = pd.to_numeric(
                    data["Amount"],
                    errors="coerce"
                )

                prepared = (
                    data["Prepared_By"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .str.casefold()
                )

                approved = (
                    data["Approved_By"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .str.casefold()
                )

                debit_account = (
                    data["Debit_Account"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .str.casefold()
                )

                credit_account = (
                    data["Credit_Account"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .str.casefold()
                )

                # Rule 1: Round figure amounts
                data["Round_Figure"] = (
                    data["Amount"].notna()
                    & data["Amount"].gt(0)
                    & data["Amount"].mod(10000).eq(0)
                )

                # Rule 2: Posting date falls on day 28, 29, 30 or 31
                data["Period_End"] = (
                    data["Posting_Date"].notna()
                    & data["Posting_Date"].dt.day.ge(28)
                )

                # Rule 3: Weekend postings
                data["Weekend_Posting"] = (
                    data["Posting_Date"].notna()
                    & data["Posting_Date"].dt.dayofweek.ge(5)
                )

                # Rule 4: Missing approval
                data["Missing_Approval"] = approved.eq("")

                # Rule 5: Same person prepared and approved
                data["Self_Approval"] = (
                    prepared.ne("")
                    & approved.ne("")
                    & prepared.eq(approved)
                )

                # Rule 6: Suspense account used
                data["Suspense_Account"] = (
                    debit_account.str.contains(
                        "suspense",
                        regex=False
                    )
                    | credit_account.str.contains(
                        "suspense",
                        regex=False
                    )
                )

                flag_columns = [
                    "Round_Figure",
                    "Period_End",
                    "Weekend_Posting",
                    "Missing_Approval",
                    "Self_Approval",
                    "Suspense_Account"
                ]

                data["Flag_Count"] = (
                    data[flag_columns]
                    .astype(int)
                    .sum(axis=1)
                )

                flagged = data.loc[
                    data["Flag_Count"] > 0
                ].copy()

                col1, col2, col3 = st.columns(3)

                col1.metric(
                    "Total Journal Entries",
                    len(data)
                )

                col2.metric(
                    "Flagged Entries",
                    len(flagged)
                )

                col3.metric(
                    "Flagged Rule Occurrences",
                    int(data[flag_columns].astype(int).sum().sum())
                )

                st.subheader("Rule-wise Results")

                rule_counts = pd.DataFrame({
                    "Audit Rule": flag_columns,
                    "Flagged Entries": [
                        int(data[column].sum())
                        for column in flag_columns
                    ]
                })

                st.dataframe(
                    rule_counts,
                    use_container_width=True,
                    hide_index=True
                )

                st.subheader("Flagged Journal Entries")

                if flagged.empty:
                    st.success(
                        "No entries were flagged by the selected "
                        "rules. This does not guarantee that all "
                        "journal entries are free from error."
                    )

                else:
                    st.warning(
                        "Review the flagged entries and verify "
                        "supporting documents, business purpose "
                        "and approval requirements."
                    )

                    st.dataframe(
                        flagged,
                        use_container_width=True
                    )

                st.download_button(
                    label="Download Journal Entry Audit Report",
                    data=convert_df_to_excel(
                        data,
                        "Journal_Audit_Report"
                    ),
                    file_name="Journal_Entry_Audit_Report.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    )
                )

                st.divider()
                st.subheader("Ask AI About Journal Entry Findings")

                question = st.text_area(
                    "Enter your audit question",
                    value=(
                        "Explain the flagged journal entries, "
                        "their potential audit risks and the "
                        "procedures I should perform."
                    ),
                    key="journal_question"
                )

                if st.button(
                    "Analyse Journal Entries with AI",
                    key="journal_ai"
                ):

                    if flagged.empty:
                        st.info(
                            "No entries were flagged. You can "
                            "still ask a general audit question "
                            "about the uploaded journal listing."
                        )

                        findings_for_ai = data.head(100)

                    else:
                        findings_for_ai = flagged.head(100)

                    if not question.strip():
                        st.warning(
                            "Please enter a question first."
                        )

                    else:
                        ask_ai(
                            question,
                            findings_for_ai,
                            "Journal Entry Testing"
                        )

                st.caption(
                    "Rule limitation: Posting dates on day 28 or "
                    "later are only preliminary indicators, not "
                    "confirmed month-end entries. The round-figure "
                    "threshold is ₹10,000. Adjust these rules to "
                    "your audit methodology and client data."
                )

        except Exception as error:
            st.error("Unable to read or process the Excel file.")
            st.code(f"{type(error).__name__}: {error}")


# --------------------------------------------------
# FOOTER
# --------------------------------------------------

st.divider()

st.caption(
    "CA Audit AI | Automated flags are preliminary indicators "
    "and must be validated using appropriate audit evidence."
)
