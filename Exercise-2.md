# ✅ **Exercise 2 — AI-Powered Financing Dues, Collections & Ageing Analysis**

**Tools:** Claude AI and Microsoft Excel\
**Application Area:** Financing, Receivables Management and Financial Operations\
**Difficulty Level:** Beginner to Intermediate

## ☑️ **1. Business Scenario**

A finance company maintains records of scheduled instalments and payments received from its borrowers. Some instalments are fully paid, some are partially paid, some are overdue, and others are payable in the future.

The management requires an accurate and regularly updated view of outstanding dues, overdue amounts, collection performance and borrower-wise payment status.

In this exercise, you will use Claude AI and Microsoft Excel to automate receivables analysis, overdue ageing, dashboard generation and management reporting.

## ☑️ **2. Learning Objectives**

- Analyse financing and repayment records.
- Calculate outstanding and overdue amounts.
- Perform borrower-wise and ageing-wise analysis.
- Identify collection priorities and data discrepancies.
- Develop a dynamic, formula-driven Excel workbook.
- Generate a financial dashboard and management report.
- Validate AI-generated financial calculations.

## ☑️ **3. Input Dataset**

Create a Microsoft Excel workbook named:

`Financing_Dues_Input.xlsx`

The workbook should contain a worksheet named `Dues`.

### 🟤 **Worksheet: Dues**

| Instalment ID | Borrower   | Due Date    | Scheduled Amount (₹) | Receipts Allocated (₹) |
| ------------- | ---------- | ----------- | -------------------- | ---------------------- |
| I001          | Borrower A | 26-Aug-2026 | 10000                | 5000                   |
| I002          | Borrower B | 27-Aug-2026 | 12000                | 2000                   |
| I003          | Borrower C | 25-Sep-2026 | 15000                | 0                      |
| I004          | Borrower D | 26-Sep-2026 | 20000                | 5000                   |
| I005          | Borrower D | 25-Oct-2026 | 8000                 | 3000                   |
| I006          | Borrower E | 26-Oct-2026 | 18000                | 10000                  |
| I007          | Borrower F | 24-Nov-2026 | 11000                | 1000                   |
| I008          | Borrower F | 25-Nov-2026 | 16000                | 6000                   |
| I009          | Borrower A | 30-Nov-2026 | 14000                | 0                      |
| I010          | Borrower B | 15-Nov-2026 | 9000                 | 9000                   |

### 🟤 **Additional Information**

| Parameter          | Value                            |
| ------------------ | -------------------------------- |
| Analysis date      | 25-Nov-2026                      |
| Currency           | INR                              |
| Ageing buckets     | 1–30, 31–60, 61–90 and 91+ days  |
| Record granularity | One row per scheduled instalment |

Receipts have already been allocated to their respective instalments as of the analysis date. Outstanding amounts represent unpaid scheduled dues, not necessarily the total outstanding loan principal.

All data is fictional and provided for training purposes.

## ☑️ **4. Task 1 — Data Processing and Financial Analysis**

Open Claude AI and attach the prepared Excel workbook.

Alternatively, if Claude for Excel is available, open the workbook in Microsoft Excel and use the Claude add-in.

Copy and execute the following prompt.

### 🔴 **📋 Prompt 1 — Analyse Financing and Collection Records**

```
Act as a financial operations analyst.

Analyse the attached financing instalment dataset using
25 November 2026 as the analysis date.

Perform the following activities:

1. Examine the dataset for missing values, duplicate
   records, invalid dates, overpayments and other
   inconsistencies.

2. Calculate the outstanding amount and current days
   past due for each instalment.

3. Classify instalments as:
   - Settled
   - Overdue
   - Due Today
   - Future Due

4. Classify overdue amounts into ageing buckets:
   - 1–30 days
   - 31–60 days
   - 61–90 days
   - 91+ days

5. Calculate total scheduled amounts, receipts,
   outstanding dues and overdue amounts.

6. Calculate borrower-wise collection percentages,
   outstanding balances and overdue balances.

7. Prepare borrower-wise and ageing-wise summary tables.

8. Identify significant collection exposures and
   records requiring review.

Classify fully paid instalments as settled regardless
of their due dates.

Flag invalid records rather than silently correcting them.

Preserve the original data and explain the calculation
methodology.

Do not invent interest, penalties, repayment commitments
or borrower risk ratings.

Present the results using clearly formatted tables and
provide a concise summary of your findings.
```

**Expected output:** Instalment-wise financial analysis, borrower-wise summary, overdue ageing analysis and identification of data discrepancies.

## ☑️ **5. Task 2 — Generate a Dynamic Excel Workbook**

Using the results generated in Task 1, create a reusable, formula-driven financial analysis workbook.

Execute the following prompt in the same Claude conversation.

### 🔴 **📋 Prompt 2 — Create a Formula-Driven Financial Analysis Workbook**

```
Using the original dataset and your previous analysis,
generate a professionally formatted Microsoft Excel
workbook named:

Financing_Dues_Analysis.xlsx

Organise the workbook into the following worksheets:

1. Dues
   Original financing and repayment records.

2. Analysis
   Instalment-wise calculations and payment status.

3. Borrower_Summary
   Borrower-wise financial and collection analysis.

4. Ageing_Summary
   Overdue amounts classified by ageing buckets.

5. Dashboard
   Visual summary of financing dues and collections.

Apply the following requirements:

- Preserve the original source records.
- Provide an editable analysis-date cell instead of
  using TODAY().
- Use Excel formulas for all calculated columns
  and summaries.
- Use structured tables that accommodate additional
  records.
- Apply Indian currency formatting and appropriate
  conditional formatting.
- Highlight overdue instalments and data-quality
  exceptions.
- Ensure that calculations update when valid inputs
  or the analysis date change.
- Clearly flag any manual refresh requirements.
- Verify all formulas, totals and financial
  control checks.

Return the completed workbook as a downloadable
.xlsx file.
```

**Expected output:** A professionally formatted, dynamic Excel workbook named `Financing_Dues_Analysis.xlsx`.

## ☑️ **6. Task 3 — Generate a Financial Dashboard**

Develop a professional, single-page dashboard presenting the overall financing and collections position.

### 🔴 **📋 Prompt 3 — Create a Financing and Collections Dashboard**

```
Create a professional, single-page financial dashboard
within the generated Excel workbook.

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

Include a table of overdue instalments, sorted by
days past due in descending order.

Use a clean corporate colour scheme, readable labels,
formula-driven indicators and appropriate conditional
formatting.

Clearly define the denominator used for every percentage.

If a collection percentage includes future scheduled
instalments in its denominator, label it accordingly.
Do not confuse it with the collection rate for amounts
already due.

Ensure that the dashboard remains consistent with
the underlying calculations.

Return the updated downloadable Excel workbook.
```

**Expected output:** A single-page financing and collections dashboard integrated into the Excel workbook.

## ☑️ **7. Task 4 — Generate a Management Report**

Prepare a concise management report highlighting the important findings from the financing and collections analysis.

### 🔴 **📋 Prompt 4 — Prepare a Financial Management Report**

```
Act as a financial operations analyst.

Using the completed financing dues analysis, prepare
a professional management report covering:

1. Executive Summary
2. Overall Financing and Collection Position
3. Outstanding and Overdue Dues
4. Borrower-wise Collection Analysis
5. Overdue Ageing Analysis
6. Significant Collection Exposures
7. Data Discrepancies and Exceptions
8. Suggested Follow-up Priorities and Process Improvements
9. Conclusion

Include relevant financial figures, percentages, tables
and charts wherever appropriate.

Use clear and concise business language.

Base all observations on the analysed dataset.

Distinguish factual findings from recommendations.

Do not infer creditworthiness or reasons for non-payment
without supporting evidence.

Ensure that all figures are consistent with the Excel
workbook and dashboard.

Generate the report as a professionally formatted,
downloadable Microsoft Word document named:

Financing_Collections_Report.docx
```

**Expected output:** A professional financial management report in Microsoft Word format.

## ☑️ **8. Task 5 — Validate the Generated Results**

Independently verify the accuracy and consistency of the generated analysis, workbook, dashboard and management report.

**Financial Validation Checklist**

0 of 9 verified

Scheduled amounts, receipts and outstanding balances

Days past due calculations

Instalment payment-status classifications

Overdue ageing bucket boundaries

Borrower-wise financial calculations

Collection and overdue percentages

Accuracy of Excel formulas

Dashboard indicators and visualisations

Consistency between the Excel workbook and management report

Reset checklist

### 🟤 **Financial Control Checks**

Ensure that the following equations are satisfied for all valid records.

Control Check 1 — Outstanding Amount

Total Scheduled Amount − Total Receipts

\=

## ☑️ **Total Outstanding Amount**

Control Check 2 — Outstanding Classification

Overdue Amount + Amount Due Today + Future Dues

\=

## ☑️ **Total Outstanding Amount**

Control Check 3 — Ageing Classification

Sum of All Overdue Ageing Buckets

\=

## ☑️ **Total Overdue Amount**

Investigate and resolve any discrepancies before accepting the generated results.

## ☑️ **9. Independent Practical Assignment**

Create a separate copy of the completed Excel workbook and perform the following modifications.

### 🟤 **Modification 1 — Update Existing Receipts**

| Instalment ID | Modification                                           |
| ------------- | ------------------------------------------------------ |
| I007          | Increase the receipts allocated from ₹1,000 to ₹5,000. |
| I001          | Record an additional receipt of ₹5,000.                |

### 🟤 **Modification 2 — Add New Instalments**

Add the following records to the `Dues` worksheet.

| Instalment ID | Borrower   | Due Date    | Scheduled Amount (₹) | Receipts Allocated (₹) |
| ------------- | ---------- | ----------- | -------------------- | ---------------------- |
| I011          | Borrower G | 10-Nov-2026 | 25000                | 8000                   |
| I012          | Borrower H | 15-Dec-2026 | 30000                | 0                      |
| I013          | Borrower C | 20-Nov-2026 | 12000                | 12000                  |

### 🟤 **Modification 3 — Update the Analysis Date**

Change the analysis date to **30-Nov-2026**.

Assume all recorded receipts have been received by this date.

### 🟤 **Activities to Perform**

1. Refresh the financial analysis and dashboard.
2. Verify whether the formulas automatically accommodate the additional records.
3. Compare the revised financial position with the original results.
4. Explain the changes in outstanding dues, overdue amounts, ageing distribution and collection percentages.
5. Repeat the financial control checks.
6. Generate an updated Excel workbook and management report.

## ☑️ **10. Expected Deliverables**

| Sl. No. | Deliverable                              | File Format     |
| ------- | ---------------------------------------- | --------------- |
| 1       | Original financing dataset               | .xlsx           |
| 2       | Instalment-wise financial analysis       | Excel worksheet |
| 3       | Borrower-wise collection summary         | Excel worksheet |
| 4       | Overdue ageing summary                   | Excel worksheet |
| 5       | Financing and collections dashboard      | Excel worksheet |
| 6       | Consolidated financing analysis workbook | .xlsx           |
| 7       | Financial management report              | .docx           |
| 8       | Independent assignment workbook          | .xlsx           |
| 9       | Updated management report                | .docx           |

## ☑️ **11. Key Takeaway**

Artificial Intelligence can transform routine financing records into meaningful financial insights, dynamic dashboards and actionable management reports.

By combining AI capabilities with structured financial data and Excel formulas, organisations can improve the efficiency of receivables management, overdue monitoring and management reporting.
