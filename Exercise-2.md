# Exercise 2 — AI-Powered Financing Dues, Collections & Ageing Analysis

**Tools:** Claude AI and Microsoft Excel  
**Application Area:** Financing, Receivables Management and Financial Operations  
**Difficulty Level:** Beginner to Intermediate

## 1. Business Scenario

A finance company maintains records of scheduled instalments and payments received from its borrowers. Some instalments are fully paid, some are partially paid, some are overdue, and others are payable in the future.

The management requires an accurate and regularly updated view of outstanding dues, overdue amounts, collection performance and borrower-wise payment status.

In this exercise, you will use Claude AI and Microsoft Excel to automate receivables analysis, overdue ageing, dashboard generation and management reporting.

## 2. Learning Objectives

- Analyse financing and repayment records.
- Calculate outstanding and overdue amounts.
- Perform borrower-wise and ageing-wise analysis.
- Identify collection priorities and data discrepancies.
- Develop a dynamic, formula-driven Excel workbook.
- Generate a financial dashboard and management report.
- Validate AI-generated financial calculations.

## 3. Input Dataset

Create an Excel workbook named **Financing_Dues_Input.xlsx** containing a worksheet named `Dues`.

| Instalment ID | Borrower | Due Date | Scheduled Amount (₹) | Receipts Allocated (₹) |
|---|---|---|---:|---:|
| I001 | Borrower A | 26-Aug-2026 | 10000 | 5000 |
| I002 | Borrower B | 27-Aug-2026 | 12000 | 2000 |
| I003 | Borrower C | 25-Sep-2026 | 15000 | 0 |
| I004 | Borrower D | 26-Sep-2026 | 20000 | 5000 |
| I005 | Borrower D | 25-Oct-2026 | 8000 | 3000 |
| I006 | Borrower E | 26-Oct-2026 | 18000 | 10000 |
| I007 | Borrower F | 24-Nov-2026 | 11000 | 1000 |
| I008 | Borrower F | 25-Nov-2026 | 16000 | 6000 |
| I009 | Borrower A | 30-Nov-2026 | 14000 | 0 |
| I010 | Borrower B | 15-Nov-2026 | 9000 | 9000 |

### Additional Information

| Parameter | Value |
|---|---|
| Analysis date | 25-Nov-2026 |
| Currency | INR |
| Ageing buckets | 1–30, 31–60, 61–90 and 91+ days |
| Record granularity | One row per scheduled instalment |

Receipts have already been allocated to their respective instalments as of the analysis date. Outstanding amounts represent unpaid scheduled dues, not necessarily the total outstanding loan principal.

All data is fictional and provided for training purposes.

## 4. Task 1 — Data Processing and Financial Analysis

Open Claude AI and attach the prepared Excel workbook. Alternatively, use the Claude add-in within Microsoft Excel, if available.

**Prompt 1 — Analyse financing and collection records**

Act as a financial operations analyst.

Analyse the attached financing instalment dataset using 25 November 2026 as the analysis date.

Perform the following activities:

1. Examine the dataset for missing values, duplicate records, invalid dates, overpayments and other inconsistencies.
2. Calculate the outstanding amount and current days past due for each instalment.
3. Classify instalments as Settled, Overdue, Due Today or Future Due.
4. Classify overdue amounts into the following ageing buckets:
   - 1–30 days
   - 31–60 days
   - 61–90 days
   - 91+ days
5. Calculate total scheduled amounts, receipts, outstanding dues and overdue amounts.
6. Calculate borrower-wise collection percentages, outstanding balances and overdue balances.
7. Prepare borrower-wise and ageing-wise summary tables.
8. Identify significant collection exposures and records requiring review.

Classify fully paid instalments as settled regardless of their due dates. Flag invalid records rather than silently correcting them.

Preserve the original data and explain the calculation methodology. Do not invent interest, penalties, repayment commitments or borrower risk ratings.

## 5. Task 2 — Generate a Dynamic Excel Workbook

**Prompt 2 — Create a formula-driven financial analysis workbook**

Using the original dataset and your previous analysis, generate a professionally formatted Excel workbook named:

**Financing_Dues_Analysis.xlsx**

Organise the workbook into the following worksheets:

| Worksheet | Purpose |
|---|---|
| Dues | Original financing and repayment records |
| Analysis | Instalment-wise calculations and payment status |
| Borrower_Summary | Borrower-wise financial and collection analysis |
| Ageing_Summary | Overdue amounts classified by ageing buckets |
| Dashboard | Visual summary of financing dues and collections |

Apply the following requirements:

- Preserve the original source records.
- Provide an editable analysis-date cell instead of using TODAY().
- Use Excel formulas for all calculated columns and summaries.
- Use structured tables that accommodate additional records.
- Apply Indian currency formatting and appropriate conditional formatting.
- Highlight overdue instalments and data-quality exceptions.
- Ensure that calculations update when valid inputs or the analysis date change.
- Clearly flag any manual refresh requirements.
- Verify all formulas, totals and financial control checks.

Return the completed workbook as a downloadable .xlsx file.

## 6. Task 3 — Generate a Financial Dashboard

**Prompt 3 — Create a financing and collections dashboard**

Create a professional, single-page financial dashboard within the generated Excel workbook.

Include the following key performance indicators:

- Total scheduled amount.
- Total receipts collected.
- Total outstanding amount.
- Total overdue amount.
- Amount due today.
- Future dues.
- Overall collection percentage.
- Overdue percentage of outstanding dues.
- Number of borrowers with overdue instalments.

Include suitable visualisations for:

- Scheduled amounts versus collections.
- Borrower-wise outstanding and overdue amounts.
- Ageing-wise distribution of overdue amounts.
- Instalment payment-status distribution.
- Borrower-wise collection percentages.

Include a table of overdue instalments, sorted by days past due.

Use a clean corporate colour scheme, readable labels, formula-driven indicators and appropriate conditional formatting.

Clearly define the denominator used for every percentage. Ensure that the dashboard remains consistent with the underlying calculations.

Return the updated downloadable Excel workbook.

## 7. Task 4 — Generate a Management Report

**Prompt 4 — Prepare a financial management report**

Act as a financial operations analyst.

Using the completed financing dues analysis, prepare a concise management report covering:

1. Executive summary.
2. Overall financing and collection position.
3. Outstanding and overdue dues.
4. Borrower-wise collection analysis.
5. Overdue ageing analysis.
6. Significant collection exposures.
7. Data discrepancies and exceptions.
8. Suggested follow-up priorities and process improvements.
9. Conclusion.

Include relevant financial figures, percentages, tables and charts wherever appropriate.

Base all observations on the analysed dataset. Distinguish factual findings from recommendations. Do not infer creditworthiness or reasons for non-payment without supporting evidence.

Generate the report as a professionally formatted, downloadable Microsoft Word document named **Financing_Collections_Report.docx**.

## 8. Task 5 — Validate the Generated Results

Independently verify:

- Scheduled amounts, receipts and outstanding balances.
- Days past due and payment-status classifications.
- Ageing bucket boundaries.
- Borrower-wise calculations and collection percentages.
- Excel formulas and dashboard indicators.
- Consistency between the Excel workbook and management report.

Ensure that the following financial control checks are satisfied for valid records:

**Total Scheduled Amount − Total Receipts = Total Outstanding Amount**

**Overdue Amount + Amount Due Today + Future Dues = Total Outstanding Amount**

**Sum of All Overdue Ageing Buckets = Total Overdue Amount**

Investigate and resolve any discrepancies.

## 9. Independent Practical Assignment

Create a separate copy of the completed workbook and perform the following modifications:

1. Increase the receipts allocated against instalment I007 from ₹1,000 to ₹5,000.
2. Record an additional receipt of ₹5,000 against instalment I001.
3. Add the following new instalments:

| Instalment ID | Borrower | Due Date | Scheduled Amount (₹) | Receipts Allocated (₹) |
|---|---|---|---:|---:|
| I011 | Borrower G | 10-Nov-2026 | 25000 | 8000 |
| I012 | Borrower H | 15-Dec-2026 | 30000 | 0 |
| I013 | Borrower C | 20-Nov-2026 | 12000 | 12000 |

4. Change the analysis date to 30-Nov-2026. Assume all recorded receipts have been received by this date.
5. Refresh the analysis and dashboard.
6. Compare the revised financial position with the original results.
7. Explain the changes in outstanding dues, overdue amounts, ageing distribution and collection percentages.
8. Generate an updated management report.

## 10. Expected Deliverables

| Sl. No. | Deliverable | File Format |
|---:|---|---|
| 1 | Original financing dataset | .xlsx |
| 2 | Consolidated financing analysis workbook with summaries and dashboard | .xlsx |
| 3 | Financial management report | .docx |
| 4 | Independent assignment workbook | .xlsx |
| 5 | Updated management report | .docx |

## 11. Key Takeaway

Artificial Intelligence can transform routine financing records into meaningful financial insights, dynamic dashboards and actionable management reports.

**AI-assisted financial analysis should be accurate, transparent, verifiable and repeatable.**
