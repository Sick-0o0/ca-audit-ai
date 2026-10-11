
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
        "Gemini unavailable after 4 attempts. "
        + str(last_error)
    )


def ask_ai(question, findings, module_name):
    try:
        client = genai.Client(
            api_key=st.secrets["GEMINI_API_KEY"]
        )

        prompt = f"""
You are an audit assistant supporting a Chartered Accountant.

Module: {module_name}
Question: {question}

Data and findings:
{findings.to_string(index=False)[:20000]}

Explain:
1. What the findings mean.
2. Potential audit risks.
3. Relevant financial statement assertions.
4. Further audit procedures and supporting documents.
5. Practical next steps.

A flagged item is not automatically an error or fraud.
Do not invent facts, laws, standards or accounting references.
Clearly distinguish observations from conclusions.
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
# EXCEL HELPERS
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

        ws = writer.sheets[sheet_name]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for column_cells in ws.columns:
            max_length = max(
                len(str(cell.value or ""))
                for cell in column_cells
            )

            ws.column_dimensions[
                column_cells[0].column_letter
            ].width = min(max_length + 3, 35)

    return output.getvalue()


def download_report(data, label, filename, sheet_name):
    st.download_button(
        label=label,
        data=convert_df_to_excel(data, sheet_name),
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
                        "Potential duplicates detected. Verify invoices, "
                        "purchase orders, GRNs and payment records."
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

    uploaded_file = st.file_uploader(
        "Upload Journal Entries Excel",
        type=["xlsx", "xls"],
        key="journal_entries"
    )

    if uploaded_file is not None:
        try:
            df = read_excel_file(uploaded_file)

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

                st.dataframe(
                    pd.DataFrame({
                        "Audit Rule": flag_columns,
                        "Flagged Entries": [
                            int(data[c].sum()) for c in flag_columns
                        ]
                    }),
                    use_container_width=True,
                    hide_index=True
                )

                st.subheader("Flagged Journal Entries")
                st.dataframe(flagged, use_container_width=True)

                download_report(
                    data,
                    "Download Journal Entry Audit Report",
                    "Journal_Entry_Audit_Report.xlsx",
                    "Journal Audit Report"
                )

                st.caption(
                    "Period_End is a day-28-or-later heuristic. "
                    "Round-figure threshold is 10,000."
                )

                question = st.text_area(
                    "Ask AI about journal entry findings",
                    value=(
                        "Explain the flagged entries, potential risks "
                        "and further audit procedures."
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
                            question,
                            findings,
                            "Journal Entry Testing"
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
        "account-type mismatches, suspense accounts and large balances."
    )

    st.info(
        "Required columns: Account_Code, Account_Name, Account_Type, "
        "Opening_Balance, Debit, Credit. Opening balance convention: "
        "debit positive and credit negative."
    )

    uploaded_file = st.file_uploader(
        "Upload Trial Balance Excel",
        type=["xlsx", "xls"],
        key="trial_balance"
    )

    if uploaded_file is not None:
        try:
            df = read_excel_file(uploaded_file)

            required = [
                "Account_Code",
                "Account_Name",
                "Account_Type",
                "Opening_Balance",
                "Debit",
                "Credit"
            ]

            missing = [c for c in required if c not in df.columns]

            if missing:
                st.error("Missing columns: " + ", ".join(missing))
            else:
                data = df.copy()

                data["Account_Type"] = (
                    data["Account_Type"].fillna("").astype(str)
                    .str.strip().str.title()
                )

                numeric_columns = [
                    "Opening_Balance", "Debit", "Credit"
                ]

                for column in numeric_columns:
                    data[column] = pd.to_numeric(
                        data[column], errors="coerce"
                    )

                invalid_rows = data[numeric_columns].isna().any(axis=1)

                if invalid_rows.any():
                    st.warning(
                        f"{int(invalid_rows.sum())} row(s) contain "
                        "blank or non-numeric amounts. Review them."
                    )

                data[numeric_columns] = data[numeric_columns].fillna(0)

                # Signed convention:
                # Debit balance positive; credit balance negative.
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

                threshold = st.number_input(
                    "Large balance threshold (₹)",
                    min_value=1,
                    value=500000,
                    step=50000,
                    key="tb_threshold"
                )

                data["Suspense_Account"] = (
                    data["Account_Type"].eq("Suspense")
                    | data["Account_Name"].fillna("").astype(str)
                    .str.contains("suspense", case=False, regex=False)
                )

                data["Large_Closing_Balance"] = (
                    data["Closing_Balance"].abs() >= threshold
                )

                # Expected normal balance direction by account type.
                # These are review indicators, not proof of error.
                normal_debit_types = ["Asset", "Expense"]
                normal_credit_types = [
                    "Liability", "Equity", "Revenue"
                ]

                data["Unexpected_Balance_Direction"] = (
                    (
                        data["Account_Type"].isin(normal_debit_types)
                        & data["Closing_Balance"].lt(-0.005)
                    )
                    |
                    (
                        data["Account_Type"].isin(normal_credit_types)
                        & data["Closing_Balance"].gt(0.005)
                    )
                )

                data["Unknown_Account_Type"] = ~data[
                    "Account_Type"
                ].isin(
                    normal_debit_types
                    + normal_credit_types
                    + ["Suspense"]
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

                # Large balances are shown separately; they do not
                # automatically make an account an anomaly.
                anomaly_columns = [
                    "Suspense_Account",
                    "Unexpected_Balance_Direction",
                    "Unknown_Account_Type",
                    "Missing_Account_Name",
                    "Duplicate_Account_Code"
                ]

                data["TB_Flag_Count"] = (
                    data[anomaly_columns].astype(int).sum(axis=1)
                )

                flagged = data[
                    data["TB_Flag_Count"] > 0
                ].copy()

                debit_total = data["Debit"].sum()
                credit_total = data["Credit"].sum()
                movement_difference = debit_total - credit_total

                opening_net = data["Opening_Balance"].sum()
                closing_net = data["Closing_Balance"].sum()

                c1, c2, c3 = st.columns(3)
                c1.metric(
                    "Total Debit Movements",
                    f"₹{debit_total:,.2f}"
                )
                c2.metric(
                    "Total Credit Movements",
                    f"₹{credit_total:,.2f}"
                )
                c3.metric(
                    "Debit-Credit Difference",
                    f"₹{movement_difference:,.2f}"
                )

                c4, c5, c6 = st.columns(3)
                c4.metric(
                    "Opening Net Balance",
                    f"₹{opening_net:,.2f}"
                )
                c5.metric(
                    "Closing Net Balance",
                    f"₹{closing_net:,.2f}"
                )
                c6.metric(
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
                        "Verify source completeness, period, account "
                        "mapping and the report's sign convention."
                    )

                st.subheader("Potential Anomalies")

                if flagged.empty:
                    st.success(
                        "No accounts met the selected anomaly rules."
                    )
                else:
                    st.warning(
                        "Review these accounts. Unexpected direction "
                        "and suspense indicators require investigation; "
                        "they are not proof of error or fraud."
                    )
                    st.dataframe(
                        flagged,
                        use_container_width=True
                    )

                st.subheader("Large Balances for Review")

                large_balances = data[
                    data["Large_Closing_Balance"]
                ].copy()

                if large_balances.empty:
                    st.info(
                        "No closing balances meet the selected threshold."
                    )
                else:
                    st.dataframe(
                        large_balances,
                        use_container_width=True
                    )

                st.subheader("Rule-wise Summary")

                summary_rules = (
                    anomaly_columns
                    + ["Large_Closing_Balance"]
                )

                st.dataframe(
                    pd.DataFrame({
                        "Audit Rule": summary_rules,
                        "Accounts Flagged": [
                            int(data[c].sum())
                            for c in summary_rules
                        ]
                    }),
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

                question = st.text_area(
                    "Ask AI about Trial Balance findings",
                    value=(
                        "Explain the Trial Balance findings, unexpected "
                        "balance directions, suspense balances and "
                        "suggest further audit procedures."
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
                                flagged.head(80),
                                large_balances.head(20)
                            ]
                        ).drop_duplicates()

                        if ai_findings.empty:
                            ai_findings = data.head(100)

                        ask_ai(
                            question,
                            ai_findings,
                            "Trial Balance Anomaly Detection"
                        )

                st.caption(
                    "Assumption: Opening_Balance uses debit-positive and "
                    "credit-negative signs. Closing = Opening + Debit - "
                    "Credit. Normal balance direction is a review rule "
                    "only; contra accounts and legitimate exceptions "
                    "may require different treatment."
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
    "Validate findings using appropriate audit evidence."
)
