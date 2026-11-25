# **Exercise 1 — AI-Powered Bank Reconciliation**

**Tools:** Claude AI and Microsoft Excel\
**Application Area:** Finance, Bookkeeping and Backend Operations\
**Difficulty Level:** Beginner to Intermediate

## **1. Business Scenario**

An organisation maintains its financial transactions in an internal bank book. The bank independently maintains a statement of transactions processed through the corresponding bank account.

At the end of a reporting period, the closing balance appearing in the company's bank book may differ from the closing balance appearing in the bank statement.

Bank reconciliation is the process of comparing these two sets of financial records, identifying discrepancies, investigating their possible causes and determining the adjustments necessary to reconcile the balances.

In this exercise, you will use Claude AI and Microsoft Excel to automate the analysis and preparation of a bank reconciliation statement.

## **2. Learning Objectives**

Upon completing this exercise, you should be able to:

- Compare two financial datasets using AI.
- Automatically match corresponding financial transactions.
- Detect duplicate, missing and unmatched transactions.
- Investigate discrepancies between financial records.
- Generate a bank reconciliation statement.
- Create an exception report for management review.
- Develop a formula-driven Excel workbook.
- Generate a financial dashboard and analytical report.
- Validate AI-generated financial calculations and conclusions.

## **3. Input Datasets**

Create a Microsoft Excel workbook named:

`Bank_Reconciliation_Input.xlsx`

The workbook should contain the following two worksheets.

### **Worksheet 1: Bank_Statement**

| Row ID | Date        | Reference | Description         | Amount (₹) |
| ------ | ----------- | --------- | ------------------- | ---------- |
| BK001  | 20-Nov-2026 | RCP101    | Customer receipt    | 10000      |
| BK002  | 21-Nov-2026 | PAY201    | Supplier A payment  | -3000      |
| BK003  | 22-Nov-2026 | PAY202    | Supplier B payment  | -2000      |
| BK004  | 25-Nov-2026 | CHG2511   | Bank service charge | -500       |

### **Worksheet 2: Company_Bank_Book**

| Row ID | Date        | Reference | Description        | Amount (₹) |
| ------ | ----------- | --------- | ------------------ | ---------- |
| LG001  | 20-Nov-2026 | RCP101    | Customer receipt   | 10000      |
| LG002  | 21-Nov-2026 | PAY201    | Supplier A payment | -3000      |
| LG003  | 21-Nov-2026 | PAY201    | Supplier A payment | -3000      |
| LG004  | 22-Nov-2026 | PAY202    | Supplier B payment | -2000      |
| LG005  | 24-Nov-2026 | CHQ301    | Issued cheque      | -5000      |

### **Additional Information**

| Parameter                           | Value       |
| ----------------------------------- | ----------- |
| Opening balance — Bank Statement    | ₹1,00,000   |
| Opening balance — Company Bank Book | ₹1,00,000   |
| Reconciliation date                 | 25-Nov-2026 |
| Currency                            | INR         |
| Positive amount                     | Receipt     |
| Negative amount                     | Payment     |

Both datasets relate to the same fictional bank account. All data provided in this exercise is simulated for training purposes.

## **4. Task 1 — Data Analysis and Discrepancy Identification**

Open Claude AI and attach the prepared Excel workbook.

Alternatively, if Claude for Excel is available, open the workbook in Microsoft Excel and use the Claude add-in.

Copy and execute the following prompt.

### **📋 Prompt 1 — Analyse and Compare Financial Transactions**

```
Analyse the attached Excel workbook containing two worksheets
named Bank_Statement and Company_Bank_Book.

Both worksheets represent transactions associated with the
same bank account.

The opening balance in both records is ₹1,00,000, and the
reconciliation date is 25 November 2026. Positive amounts
represent receipts and negative amounts represent payments.

Perform the following activities:

1. Examine the structure, completeness and quality of both
   datasets.
2. Calculate the total receipts, total payments and closing
   balance for each dataset.
3. Automatically match corresponding transactions using
   appropriate combinations of transaction reference, amount,
   date and description.
4. Apply one-to-one transaction matching wherever appropriate.
5. Identify duplicate or potentially duplicate transactions.
6. Identify transactions appearing in one dataset but not
   in the other.
7. Identify inconsistent transaction details, if any.
8. Calculate the difference between the closing balances.
9. Investigate and explain the possible reasons for each
   discrepancy.
10. Classify each identified discrepancy into an appropriate
    category.
11. Identify any additional information or supporting
    documents required to confirm your findings.

Present the results using clearly formatted tables.

Preserve the original data and include the relevant source
row IDs for every finding.

Distinguish confirmed observations from probable explanations
and assumptions. Do not fabricate information or make
unsupported conclusions.

Provide a concise summary of your overall findings.
```

**Expected output:** Transaction matching analysis, discrepancy identification and a summary of findings.

## **5. Task 2 — Automated Bank Reconciliation**

Using the results generated in Task 1, prepare a comprehensive bank reconciliation statement.

Execute the following prompt in the same Claude conversation.

### **📋 Prompt 2 — Generate a Bank Reconciliation Statement**

```
Based on your previous analysis of the Bank_Statement and
Company_Bank_Book worksheets, independently determine the
adjustments that may be required to reconcile the two
closing balances.

Perform the following activities:

1. Identify the most probable explanation for each
   discrepancy.
2. Determine whether each discrepancy represents a possible
   accounting error, an unrecorded transaction, a timing
   difference or another reconciliation issue.
3. Determine the appropriate adjustment amount and direction
   for each discrepancy.
4. Distinguish adjustments that may require corrections in
   the company bank book from adjustments made only for
   reconciliation purposes.
5. Prepare a detailed provisional bank reconciliation
   statement.
6. Calculate the adjusted closing balance for both records.
7. Calculate any remaining unreconciled difference.
8. Explain the financial reasoning behind every proposed
   adjustment.
9. Identify assumptions and findings requiring supporting
   evidence or human verification.
10. Provide appropriate recommendations for resolving the
    identified discrepancies.

Reference the original transaction IDs wherever applicable.

Do not introduce artificial balancing entries or make
unsupported adjustments simply to force the balances to agree.

If the available information is insufficient to establish
a definitive reconciliation, clearly identify the unresolved
items and present the reconciliation as provisional.
```

**Expected output:** Provisional bank reconciliation statement with proposed adjustments, supporting explanations and items requiring verification.

## **6. Task 3 — Generate a Formula-Driven Excel Workbook**

Generate a new Excel workbook containing the complete analysis and reconciliation results.

### **📋 Prompt 3 — Create the Final Excel Workbook**

```
Using the original datasets and your completed reconciliation
analysis, create a professionally formatted and downloadable
Microsoft Excel workbook named:

Bank_Reconciliation_Analysis.xlsx

Organise the workbook into the following worksheets:

1. Bank_Statement
   Original bank transaction records.

2. Company_Bank_Book
   Original company transaction records.

3. Transaction_Matching
   Matched, unmatched and ambiguous transactions.

4. Exception_Report
   Identified discrepancies, probable causes and review
   requirements.

5. Reconciliation
   Detailed bank reconciliation statement.

6. Dashboard
   Visual summary of the reconciliation results.

Apply the following requirements:

- Preserve the original source datasets without modifying
  or deleting any records.
- Use Excel formulas for financial calculations wherever
  applicable.
- Use structured tables and appropriate column headings.
- Reference source transaction IDs in the analysis.
- Apply Indian currency formatting.
- Use appropriate conditional formatting to highlight
  discrepancies.
- Clearly distinguish verified findings, proposed
  adjustments and unresolved items.
- Provide editable controls for opening balances,
  reconciliation date and proposed adjustments.
- Ensure that calculated totals and reconciliation
  balances update when the relevant inputs change.
- Clearly indicate whether transaction matching must be
  rerun when source records change.
- Validate all calculations and check for formula errors.

Provide the completed Excel workbook as a downloadable
.xlsx file.
```

**Expected output:** A professionally formatted, formula-driven Excel workbook named `Bank_Reconciliation_Analysis.xlsx`.

## **7. Task 4 — Generate a Bank Reconciliation Dashboard**

Develop a professional, single-page dashboard that presents the important findings from the reconciliation process.

### **📋 Prompt 4 — Create a Financial Dashboard**

```
Using the completed bank reconciliation analysis, create
a professional single-page dashboard within the Excel
workbook.

Include the following key performance indicators:

- Opening balance.
- Bank statement closing balance.
- Company bank-book closing balance.
- Total matched transactions.
- Total unmatched or ambiguous transactions.
- Number of identified exceptions.
- Total proposed adjustments.
- Adjusted closing balances.
- Remaining unreconciled difference.
- Overall reconciliation and review status.

Include suitable visualisations for:

- Comparison of original closing balances.
- Comparison of adjusted closing balances.
- Distribution of identified discrepancies by category.
- Financial impact of proposed adjustments.
- Transaction matching and exception status.

Use a clean corporate colour scheme, readable labels
and appropriate conditional formatting.

Use formula-driven dashboard indicators wherever possible.

Ensure that a zero remaining difference does not
automatically imply that all proposed adjustments have
been verified or approved.

Return the updated downloadable Excel workbook.
```

**Expected output:** A single-page financial dashboard integrated into the reconciliation workbook.

## **8. Task 5 — Generate a Management Report**

Prepare a concise management report based on the completed reconciliation analysis.

### **📋 Prompt 5 — Generate a Management Summary**

```
Act as a financial operations analyst.

Using the completed bank reconciliation analysis, prepare
a professional management report containing the following
sections:

1. Executive Summary
2. Scope of Reconciliation
3. Comparison of Financial Records
4. Key Discrepancies Identified
5. Probable Causes of Discrepancies
6. Proposed Reconciliation Adjustments
7. Final Reconciliation Status
8. Items Requiring Verification
9. Recommendations for Improving Financial Record Management
10. Conclusion

Use clear and concise business language.

Include relevant financial figures and transaction
references.

Distinguish factual observations from assumptions and
proposed corrective actions.

Do not invent explanations, supporting evidence or
management approvals.

Generate the report as a professionally formatted,
downloadable Microsoft Word document named:

Bank_Reconciliation_Report.docx
```

**Expected output:** A professional management report in Microsoft Word format.

## **9. Task 6 — Validate the Generated Results**

Independently verify the following:

**Financial Validation Checklist**

0 of 10 verified

Opening and closing balances

Total receipts and payments

Transaction matching accuracy

Identified discrepancies

Proposed adjustment amounts

Adjusted closing balances

Remaining unreconciled difference

Accuracy of Excel formulas

Consistency between the dashboard and the underlying data

Consistency between the management report and the Excel workbook

Reset checklist

Ensure that all proposed accounting corrections are supported by appropriate reasoning and that any unverified assumptions are clearly documented.

## **10. Independent Practical Assignment**

Create a separate copy of the original input workbook and perform the following modifications:

| Sl. No. | Modification                                                                                            |
| ------- | ------------------------------------------------------------------------------------------------------- |
| 1       | Change the bank service charge from ₹500 to ₹750.                                                       |
| 2       | Add a new customer receipt of ₹12,000 to the bank statement only, using a unique transaction reference. |
| 3       | Add a supplier payment of ₹4,500 to both datasets, using the same transaction reference.                |
| 4       | Add another payment entry of ₹2,000 to the company bank book only, using a new transaction reference.   |

Repeat the complete reconciliation process using Claude AI.

Compare the revised results with the original analysis and explain how the additional transactions affect the reconciliation.

Generate an updated Excel workbook, dashboard and management report.

## **11. Expected Deliverables**

| Sl. No. | Deliverable                          | File Format     |
| ------- | ------------------------------------ | --------------- |
| 1       | Original input workbook              | .xlsx           |
| 2       | Transaction matching analysis        | Excel worksheet |
| 3       | Exception report                     | Excel worksheet |
| 4       | Bank reconciliation statement        | Excel worksheet |
| 5       | Financial reconciliation dashboard   | Excel worksheet |
| 6       | Consolidated reconciliation workbook | .xlsx           |
| 7       | Management report                    | .docx           |
| 8       | Independent assignment output        | .xlsx and .docx |

## **12. Key Takeaway**

Artificial Intelligence can significantly accelerate financial transaction matching, discrepancy identification, bank reconciliation, dashboard preparation and management reporting.

However, financial accuracy depends on reliable source data, appropriate business rules, correct calculations and human verification.
