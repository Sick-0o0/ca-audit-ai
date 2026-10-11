
import time
from io import BytesIO

import pandas as pd
import streamlit as st
from google import genai


# ==================================================
# PAGE CONFIGURATION
# ==================================================

st.set_page_config(
    page_title="CA Audit AI",
    page_icon="📊",
    layout="wide"
)

st.title("CA Audit AI")
st.caption("Excel-based audit checks and AI-assisted analysis")


# ==================================================
# GEMINI AI WITH RETRIES
# ==================================================

MODEL_NAME = "gemini-3.8-flash"


def generate_ai_response(client, prompt):
    last_error = None

    for attempt in range(4):
        try:
            return client.models.generate_content(
                model=MODEL_NAME,
                contents=prompt
            )

        except Exception as error:
            last_error = error
            error_text = str(error).upper()

            temporary_error = any(
                code in error_text
                for code in [
                    "503", "UNAVAILABLE",
                    "429", "RESOURCE_EXHAUSTED",
                    "500", "INTERNAL",
                    "502", "504", "TIMEOUT"
                ]
            )

            if not temporary_error:
                raise

            if attempt < 3:
                wait_seconds = 2 ** (attempt + 1)
                st.warning(
                    f"AI server busy. Retrying in "
                    f"{wait_seconds} seconds..."
                )
                time.sleep(wait_seconds)

    raise RuntimeError(
        "Gemini is still unavailable after 4 attempts. "
        + str(last_error)
    )


def ask_ai(question, findings, module_name):
    try:
        client = genai.Client(
            api_key=st.secrets["GEMINI_API_KEY"]
        )

        findings_text = findings.to_string(index=False)

        # Avoid sending an excessively large prompt
        findings_text = findings_text[:20000]

        prompt = f"""
You are an audit assistant supporting a Chartered Accountant.

Audit module: {module_name}

User question:
{question}

Data and findings:
{findings_text}

Explain:
1. What the findings mean.
2. Potential audit risks.
3. Relevant financial statement assertions.
4. Further audit procedures and documents to inspect.
5. Practical next steps.

Rules:
- A flagged item is not automatically an error or fraud.
- Do not invent facts, laws, standards or accounting references.
- Clearly distinguish indicators from conclusions.
- Explain limitations where data is insufficient.
- Recommendations must be verified using audit evidence.
"""

        with st.spinner("Generating AI analysis..."):
            response = generate_ai_response(client, prompt)

        if response.text:
            st.subheader("AI Analysis")
            st.markdown(response.text)
        else:
            st.warning("AI returned an empty response.")

    except Exception as error:
        st.error("AI analysis could not be completed.")
        st.code(f"{type(error).__name__}: {error}")


# ==================================================
# EXCEL HELPER
# ==================================================

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


def download_report(dataframe, label, filename, sheet_name):
    st.download_button(
        label=label,
        data=convert_df_to_excel(dataframe, sheet_name),
        file_name=filename,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


def read_excel_file(uploaded_file):
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file)


# ==================================================
# SIDEBAR
# ==================================================

module = st.sidebar.selectbox(
    "Select Audit Module",
    [
        "Duplicate Invoice Detection",
        "Journal Entry Testing",
        "Trial Balance Anomaly Detection"
    ]
)


# ==================================================
# MODULE 1: DUPLICATE INVOICE DETECTION
# ==================================================

if module == "Duplicate Invoice Detection":

    st.header("Duplicate Invoice Detection")

    st.write(
        "Identify repeated supplier and invoice number "
        "combinations in a purchase register."
    )

    uploaded_file = st.file_uploader(
        "Upload Purchase Register Excel",
        type=["xlsx", "xls"],
        key="purchase_register"
    )

    if uploaded_file is not None:
        try:
            df = read_excel_file(uploaded_file)

            st.subheader("Uploaded Data")
            st.dataframe(df, use_container_width=True)

            required = [
                "Supplier",
                "Invoice_Number",
                "Invoice_Date",
                "Amount"
            ]

            missing = [c for c in required if c not in df.columns]

            if missing:
                st.error("Missing columns: " + ", ".join(missing))
            else:
                df["Supplier"] = (
                    df["Supplier"].fillna("").astype(str)
                    .str.strip().str.upper()
                )
                df["Invoice_Number"] = (
                    df["Invoice_Number"].fillna("").astype(str)
                    .str.strip().str.upper()
                )
                df["Amount"] = pd.to_numeric(
                    df["Amount"], errors="coerce"
                )

                valid = (
                    df["Supplier"].ne("")
                    & df["Invoice_Number"].ne("")
                )

                duplicate_mask = (
                    valid
                    & df.duplicated(
                        ["Supplier", "Invoice_Number"],
                        keep=False
                    )
                )

                duplicates = df.loc[duplicate_mask].copy()

                c1, c2, c3 = st.columns(3)
                c1.metric("Total Records", len(df))
                c2.metric("Flagged Rows", len(duplicates))
                c3.metric(
                    "Duplicate Groups",
                    duplicates.groupby(
                        ["Supplier", "Invoice_Number"]
                    ).ngroups if not duplicates.empty else 0
                )

                st.subheader("Duplicate Invoice Results")

                if duplicates.empty:
                    st.success("No duplicate combinations detected.")
                else:
                    st.warning(
                        "Potential duplicates found. Verify invoice "
                        "copies, purchase orders, GRNs and payments."
                    )
                    st.dataframe(
                        duplicates, use_container_width=True
                    )

                report = df.copy()
                report["Duplicate_Flag"] = duplicate_mask

                download_report(
                    report,
                    "Download Duplicate Invoice Report",
                    "Duplicate_Invoice_Report.xlsx",
                    "Duplicate Report"
                )

                st.divider()
                question = st.text_area(
                    "Ask AI about these findings",
                    value=(
                        "Explain the potential risks and suggest "
                        "audit procedures for these duplicate invoices."
                    ),
                    key="duplicate_question"
                )

                if st.button(
                    "Analyse Duplicate Invoices with AI",
                    key="duplicate_ai"
                ):
                    if not question.strip():
                        st.warning("Enter a question first.")
                    elif duplicates.empty:
                        st.info("No duplicate rows to analyse.")
                    else:
                        ask_ai(
                            question,
                            duplicates,
                            "Duplicate Invoice Detection"
                        )

        except Exception as error:
            st.error("Unable to process the Excel file.")
            st.code(f"{type(error).__name__}: {error}")


# ==================================================
# MODULE 2: JOURNAL ENTRY TESTING
# ==================================================

elif module == "Journal Entry Testing":

    st.header("Journal Entry Testing")

    st.write(
        "Identify journal entries requiring further audit review."
    )

    uploaded_file = st.file_uploader(
        "Upload Journal Entries Excel",
        type=["xlsx", "xls"],
        key="journal_entries"
    )

    if uploaded_file is not None:
        try:
            df = read_excel_file(uploaded_file)

            st.subheader("Uploaded Data")
            st.dataframe(df, use_container_width=True)

            required = [
                "Entry_ID",
                "Posting_Date",
                "Debit_Account",
                "Credit_Account",
                "Amount",
                "Description",
                "Prepared_By",
                "Approved_By"
            ]

            missing = [c for c in required if c not in df.columns]

            if missing:
                st.error("Missing columns: " + ", ".join(missing))
            else:
                data = df.copy()

                data["Posting_Date"] = pd.to_datetime(
                    data["Posting_Date"], errors="coerce"
                )
                data["Amount"] = pd.to_numeric(
                    data["Amount"], errors="coerce"
                )

                prepared = (
                    data["Prepared_By"].fillna("").astype(str)
                    .str.strip().str.casefold()
                )
                approved = (
                    data["Approved_By"].fillna("").astype(str)
                    .str.strip().str.casefold()
                )
                debit_account = (
                    data["Debit_Account"].fillna("").astype(str)
                    .str.strip().str.casefold()
                )
                credit_account = (
                    data["Credit_Account"].fillna("").astype(str)
                    .str.strip().str.casefold()
                )

                data["Round_Figure"] = (
                    data["Amount"].notna()
                    & data["Amount"].gt(0)
                    & data["Amount"].mod(10000).eq(0)
                )

                data["Period_End"] = (
                    data["Posting_Date"].notna()
                    & data["Posting_Date"].dt.day.ge(28)
                )

                data["Weekend_Posting"] = (
                    data["Posting_Date"].notna()
                    & data["Posting_Date"].dt.dayofweek.ge(5)
                )

                data["Missing_Approval"] = approved.eq("")

                data["Self_Approval"] = (
                    prepared.ne("")
                    & approved.ne("")
                    & prepared.eq(approved)
                )

                data["Suspense_Account"] = (
                    debit_account.str.contains("suspense", regex=False)
                    | credit_account.str.contains("suspense", regex=False)
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
                    data[flag_columns].astype(int).sum(axis=1)
                )

                flagged = data[data["Flag_Count"] > 0].copy()

                c1, c2, c3 = st.columns(3)
                c1.metric("Total Entries", len(data))
                c2.metric("Flagged Entries", len(flagged))
                c3.metric(
                    "Rule Flag Occurrences",
                    int(data[flag_columns].astype(int).sum().sum())
                )

                st.subheader("Rule-wise Results")

                rule_counts = pd.DataFrame({
                    "Audit Rule": flag_columns,
                    "Flagged Entries": [
                        int(data[c].sum()) for c in flag_columns
                    ]
                })

                st.dataframe(
                    rule_counts,
                    use_container_width=True,
                    hide_index=True
                )

                st.subheader("Flagged Journal Entries")

                if flagged.empty:
                    st.success("No entries flagged by these rules.")
                else:
                    st.warning(
                        "Review supporting documents, business purpose "
                        "and the applicable approval policy."
                    )
                    st.dataframe(
                        flagged, use_container_width=True
                    )

                download_report(
                    data,
                    "Download Journal Entry Audit Report",
                    "Journal_Entry_Audit_Report.xlsx",
                    "Journal Audit Report"
                )

                st.caption(
                    "Period_End is a simple day-28-or-later heuristic, "
                    "not a confirmed month-end test. Round-figure "
                    "threshold is 10,000."
                )

                st.divider()

                question = st.text_area(
                    "Ask AI about journal entry findings",
                    value=(
                        "Explain the flagged entries, potential audit "
                        "risks and further audit procedures."
                    ),
                    key="journal_question"
                )

                if st.button(
                    "Analyse Journal Entries with AI",
                    key="journal_ai"
                ):
                    if not question.strip():
                        st.warning("Enter a question first.")
                    else:
                        findings = (
                            flagged.head(100)
                            if not flagged.empty
                            else data.head(100)
                        )
                        ask_ai(
                            question, findings, "Journal Entry Testing"
                        )

        except Exception as error:
            st.error("Unable to process the Excel file.")
            st.code(f"{type(error).__name__}: {error}")


# ==================================================
# MODULE 3: TRIAL BALANCE ANOMALY DETECTION
# ==================================================

elif module == "Trial Balance Anomaly Detection":

    st.header("Trial Balance Anomaly Detection")

    st.write(
        "Review debit-credit movements, closing balances, "
        "suspense accounts and unusually large balances."
    )

    st.info(
        "Required columns: Account_Code, Account_Name, "
        "Opening_Balance, Debit, Credit. Opening_Balance must "
        "use a signed convention: debit positive, credit negative."
    )

    uploaded_file = st.file_uploader(
        "Upload Trial Balance Excel",
        type=["xlsx", "xls"],
        key="trial_balance"
    )

    if uploaded_file is not None:
        try:
            df = read_excel_file(uploaded_file)

            st.subheader("Uploaded Trial Balance")
            st.dataframe(df, use_container_width=True)

            required = [
                "Account_Code",
                "Account_Name",
                "Opening_Balance",
                "Debit",
                "Credit"
            ]

            missing = [c for c in required if c not in df.columns]

            if missing:
                st.error("Missing columns: " + ", ".join(missing))
                st.info(
                    "Rename the Excel headers to match the required "
                    "column names. You may keep additional columns."
                )

            else:
                data = df.copy()

                for column in [
                    "Opening_Balance", "Debit", "Credit"
                ]:
                    data[column] = pd.to_numeric(
                        data[column], errors="coerce"
                    )

                numeric_columns = [
                    "Opening_Balance", "Debit", "Credit"
                ]

                invalid_rows = data[numeric_columns].isna().any(axis=1)

                if invalid_rows.any():
                    st.warning(
                        f"{int(invalid_rows.sum())} row(s) have blank "
                        "or non-numeric balance values. These need "
                        "correction before the report can be relied on."
                    )

                data[numeric_columns] = data[numeric_columns].fillna(0)

                # Signed convention:
                # Debit balances are positive; credit balances negative.
                data["Closing_Balance"] = (
                    data["Opening_Balance"]
                    + data["Debit"]
                    - data["Credit"]
                )

                data["Closing_Debit_Balance"] = (
                    data["Closing_Balance"].clip(lower=0)
                )

                data["Closing_Credit_Balance"] = (
                    -data["Closing_Balance"].clip(upper=0)
                )

                data["Suspense_Account"] = (
                    data["Account_Name"]
                    .fillna("")
                    .astype(str)
                    .str.contains("suspense", case=False, regex=False)
                )

                data["Zero_Closing_Balance"] = (
                    data["Closing_Balance"].abs() < 0.005
                )

                threshold = st.number_input(
                    "Large balance threshold (₹)",
                    min_value=1,
                    value=500000,
                    step=50000,
                    key="tb_threshold"
                )

                data["Large_Closing_Balance"] = (
                    data["Closing_Balance"].abs() >= threshold
                )

                data["Missing_Account_Name"] = (
                    data["Account_Name"].isna()
                    | data["Account_Name"].astype(str).str.strip().eq("")
                )

                data["Duplicate_Account_Code"] = (
                    data["Account_Code"].notna()
                    & data["Account_Code"].astype(str).duplicated(
                        keep=False
                    )
                )

                data["TB_Flag_Count"] = data[
                    [
                        "Suspense_Account",
                        "Large_Closing_Balance",
                        "Missing_Account_Name",
                        "Duplicate_Account_Code"
                    ]
                ].astype(int).sum(axis=1)

                debit_total = data["Debit"].sum()
                credit_total = data["Credit"].sum()
                movement_difference = debit_total - credit_total

                opening_net = data["Opening_Balance"].sum()
                closing_net = data["Closing_Balance"].sum()

                col1, col2, col3 = st.columns(3)

                col1.metric(
                    "Total Debit Movements",
                    f"₹{debit_total:,.2f}"
                )

                col2.metric(
                    "Total Credit Movements",
                    f"₹{credit_total:,.2f}"
                )

                col3.metric(
                    "Debit-Credit Difference",
                    f"₹{movement_difference:,.2f}"
                )

                col4, col5, col6 = st.columns(3)

                col4.metric(
                    "Opening Net Balance",
                    f"₹{opening_net:,.2f}"
                )

                col5.metric(
                    "Closing Net Balance",
                    f"₹{closing_net:,.2f}"
                )

                col6.metric(
                    "Accounts Flagged",
                    int((data["TB_Flag_Count"] > 0).sum())
                )

                st.subheader("Trial Balance Checks")

                if abs(movement_difference) < 0.005:
                    st.success(
                        "Debit and credit movements agree within "
                        "the rounding tolerance."
                    )
                else:
                    st.warning(
                        "Debit and credit movements do not agree. "
                        "Check the source report, missing accounts, "
                        "sign convention and extraction completeness."
                    )

                if abs(closing_net - (
                    opening_net + debit_total - credit_total
                )) < 0.005:
                    st.caption(
                        "Closing net balance reconciles mathematically "
                        "to opening net balance plus debits minus credits."
                    )

                st.subheader("Potential Anomalies")

                flagged = data[data["TB_Flag_Count"] > 0].copy()

                if flagged.empty:
                    st.success(
                        "No accounts met the selected anomaly rules."
                    )
                else:
                    st.warning(
                        "These accounts need review. A large or "
                        "suspense balance is not automatically an error."
                    )
                    st.dataframe(
                        flagged, use_container_width=True
                    )

                st.subheader("Rule-wise Summary")

                tb_rules = [
                    "Suspense_Account",
                    "Large_Closing_Balance",
                    "Missing_Account_Name",
                    "Duplicate_Account_Code",
                    "Zero_Closing_Balance"
                ]

                rule_summary = pd.DataFrame({
                    "Audit Rule": tb_rules,
                    "Accounts Flagged": [
                        int(data[c].sum()) for c in tb_rules
                    ]
                })

                st.dataframe(
                    rule_summary,
                    use_container_width=True,
                    hide_index=True
                )

                download_report(
                    data,
                    "Download Trial Balance Audit Report",
                    "Trial_Balance_Audit_Report.xlsx",
                    "TB Audit Report"
                )

                st.divider()
                st.subheader("Ask AI About Trial Balance")

                question = st.text_area(
                    "Enter your audit question",
                    value=(
                        "Analyse these Trial Balance checks. Explain "
                        "the potential risks and suggest further "
                        "audit procedures. Also explain any debit-credit "
                        "difference without assuming fraud."
                    ),
                    key="tb_question"
                )

                if st.button(
                    "Analyse Trial Balance with AI",
                    key="tb_ai"
                ):
                    if not question.strip():
                        st.warning("Enter a question first.")
                    else:
                        ai_findings = pd.concat(
                            [
                                data.head(20),
                                flagged.head(80)
                            ]
                        ).drop_duplicates()

                        ask_ai(
                            question,
                            ai_findings,
                            "Trial Balance Anomaly Detection"
                        )

                st.caption(
                    "This module uses signed opening balances: debit "
                    "positive and credit negative. It calculates "
                    "closing balances as Opening + Debit - Credit. "
                    "The movement difference is Debit total - Credit "
                    "total. Confirm that the uploaded source follows "
                    "this convention before relying on results."
                )

        except Exception as error:
            st.error("Unable to process the Trial Balance Excel file.")
            st.code(f"{type(error).__name__}: {error}")


# ==================================================
# FOOTER
# ==================================================

st.divider()

st.caption(
    "CA Audit AI | Automated flags are preliminary indicators. "
    "Verify all findings using appropriate audit evidence."
)
