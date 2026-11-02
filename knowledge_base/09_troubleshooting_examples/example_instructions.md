# Excel Examples and AI Agent Instructions

## 1. Overview

This document provides reusable examples and task instructions for an Excel AI Agent.

It covers common Excel operations such as:

- Data cleaning
- Missing-value handling
- Duplicate removal
- Sorting and filtering
- Data transformation
- Formula creation
- Lookup operations
- PivotTable creation
- Data analysis
- Chart generation
- Dashboard creation
- KPI calculations
- Report generation
- Troubleshooting
- Workbook and worksheet operations

The AI Agent should use these examples as guidance when interpreting user requests and planning Excel operations.

**Important distinction:**

- Knowledge Base retrieval provides instructions, examples, and explanations.
- Excel tools perform actual workbook operations.
- Retrieving an instruction does not mean the workbook has been modified.

---

## 2. Standard AI Agent Workflow

For every Excel task, follow this workflow.

### Step 1: Understand the User Request

Identify:

- What the user wants.
- Which workbook or worksheet is involved.
- Which columns or data ranges are relevant.
- Whether the task requires reading, analysis, or modification.
- What output the user expects.

If important information is missing, ask a clarification question.

### Step 2: Inspect the Workbook

Use the available Excel tools to inspect:

- Workbook sheets.
- Worksheet names.
- Column headers.
- Data ranges.
- Sample records.
- Data types.
- Existing formulas.
- Existing formatting, when relevant.

Never assume a worksheet or column exists without checking.

### Step 3: Plan the Operation

Break complex requests into smaller steps.

Example:

User request:

"Clean the sales data, remove duplicates, calculate total revenue, and create a monthly sales chart."

Possible execution plan:

1. Inspect the sales worksheet.
2. Identify missing and invalid values.
3. Clean the relevant columns.
4. Identify and remove confirmed duplicates.
5. Calculate revenue.
6. Group revenue by month.
7. Create a chart.
8. Verify the results.

### Step 4: Execute Using Excel Tools

Select the appropriate tool for each operation.

Examples:

- Read workbook data using reader tools.
- Clean values using cleaning tools.
- Write values using writer tools.
- Apply formulas using formula tools.
- Create charts using chart tools.
- Generate dashboards using dashboard tools.

Use only tools that are available in the application.

### Step 5: Validate the Result

After execution, verify:

- The requested operation completed.
- The correct worksheet and columns were used.
- Row counts and calculations are reasonable.
- Formulas return expected results.
- Charts use the correct data.
- Original data was preserved when required.

### Step 6: Report the Result

Provide a concise summary:

- What operation was performed.
- Which worksheet was affected.
- What changed.
- Any warnings or unresolved issues.

Never claim an operation succeeded unless the tools confirm it.

---

## 3. Data Cleaning Examples

### Example 1: Clean a Dataset

**User request:**

"Clean my sales data."

**Expected workflow:**

1. Inspect the worksheet.
2. Identify missing values.
3. Identify duplicate records.
4. Check inconsistent text.
5. Check incorrect data types.
6. Check date and number formats.
7. Apply appropriate cleaning operations.
8. Validate the cleaned dataset.

**Expected output:**

A cleaned dataset with a summary of the corrections.

**Important:** Do not delete records automatically when the correct treatment is unclear.

### Example 2: Remove Duplicate Records

**User request:**

"Remove duplicate customer records."

**Expected workflow:**

1. Identify the customer identifier.
2. Determine which columns define a duplicate.
3. Count duplicate records.
4. Preserve the first occurrence unless another rule is specified.
5. Remove confirmed duplicates.
6. Verify the remaining record count.

**Clarification:** Two customers with the same name are not necessarily duplicates. Prefer a unique customer ID when available.

### Example 3: Handle Missing Values

**User request:**

"Fill missing values in the sales column."

**Expected workflow:**

1. Identify missing cells.
2. Check the column data type.
3. Understand what missing values represent.
4. Ask whether to use zero, mean, median, or another method when the correct choice is unclear.
5. Apply the selected method.
6. Verify the result.

Do not automatically replace every missing value with zero.

---

## 4. Sorting and Filtering Examples

### Example 4: Sort Sales

**User request:**

"Sort sales from highest to lowest."

**Expected workflow:**

1. Identify the sales column.
2. Select the complete dataset.
3. Sort by sales in descending order.
4. Verify that entire rows remain together.

**Expected result:** Highest sales appear first.

### Example 5: Filter High-Value Orders

**User request:**

"Show orders above 10,000."

**Expected workflow:**

1. Identify the order amount column.
2. Apply the condition `Order Amount > 10000`.
3. Display matching records.
4. Report the number of matching orders.

Do not delete non-matching records unless explicitly requested.

---

## 5. Data Transformation Examples

### Example 6: Create a Full Name Column

**User request:**

"Combine first name and last name."

Example data:

| First Name | Last Name |
|---|---|
| Arun | Kumar |
| Priya | Devi |

Formula:

```excel
=A2&" "&B2
```

Expected output:

| Full Name |
|---|
| Arun Kumar |
| Priya Devi |

### Example 7: Extract Year from a Date

**User request:**

"Create a year column from the order date."

Formula:

```excel
=YEAR(A2)
```

Verify that the source column contains valid Excel dates.

### Example 8: Calculate Revenue

**User request:**

"Calculate revenue for every product."

Formula:

```excel
=B2*C2
```

Assumption:

- Column B = Quantity
- Column C = Unit Price

Expected output: A new Revenue column.

---

## 6. Formula and Lookup Examples

### Example 9: Calculate Total Sales

**User request:**

"Calculate total sales."

Formula:

```excel
=SUM(B2:B100)
```

The actual range must be determined from the workbook.

### Example 10: Calculate Average Sales

Formula:

```excel
=AVERAGE(B2:B100)
```

### Example 11: Find the Highest Sale

Formula:

```excel
=MAX(B2:B100)
```

### Example 12: Find Customer Details

**User request:**

"Find the customer name for customer ID C102."

Example formula:

```excel
=XLOOKUP(E2,A2:A100,B2:B100,"Not Found")
```

Assumption:

- Column A contains Customer ID.
- Column B contains Customer Name.
- E2 contains the requested customer ID.

If XLOOKUP is unavailable, consider INDEX and MATCH or another supported lookup method.

### Example 13: Categorize Sales

**User request:**

"Classify sales as High or Low."

Formula:

```excel
=IF(B2>=10000,"High","Low")
```

The threshold should follow the user's requirement.

---

## 7. PivotTable Examples

### Example 14: Sales by Region

**User request:**

"Show total sales for each region."

**Expected PivotTable configuration:**

- Rows: Region
- Values: Sum of Sales

**Expected output:**

| Region | Total Sales |
|---|---:|
| North | 50000 |
| South | 42000 |
| East | 38000 |

Values above are illustrative.

### Example 15: Monthly Sales Summary

**User request:**

"Show monthly sales."

**Expected configuration:**

- Rows: Order Date, grouped by Month and Year.
- Values: Sum of Sales.

If dates span multiple years, retain the year in the grouping to avoid combining different years incorrectly.

### Example 16: Product Performance

**User request:**

"Find the top-performing products."

**Expected configuration:**

- Rows: Product.
- Values: Sum of Revenue.
- Sort: Revenue descending.

The agent should explain the ranking metric used.

---

## 8. Data Analysis Examples

### Example 17: Sales Summary

**User request:**

"Analyze my sales data."

**Expected analysis:**

- Total revenue.
- Total number of orders.
- Average order value.
- Highest-value order.
- Lowest-value order.
- Sales by region.
- Sales by product.
- Monthly sales trends.

Use only metrics supported by the available columns.

### Example 18: Compare Two Regions

**User request:**

"Compare North and South region sales."

**Expected analysis:**

1. Calculate total sales for both regions.
2. Calculate the absolute difference.
3. Calculate percentage difference.
4. Identify the higher-performing region.

Percentage difference formula:

```excel
=(North_Sales-South_Sales)/South_Sales*100
```

Handle a zero comparison value to avoid division errors.

### Example 19: Identify Trends

**User request:**

"Find the monthly sales trend."

**Expected analysis:**

1. Group sales by month and year.
2. Calculate monthly totals.
3. Compare consecutive periods.
4. Identify increases and decreases.
5. Highlight unusual changes when supported by the data.

Do not claim that a trend will continue without sufficient evidence.

---

## 9. Charts and Visualization Examples

### Example 20: Create a Sales Chart

**User request:**

"Create a chart showing sales by region."

**Recommended chart:** Column or bar chart.

**Workflow:**

1. Prepare regional sales totals.
2. Select the correct data range.
3. Create the chart.
4. Add a meaningful title.
5. Verify category labels and values.

### Example 21: Monthly Sales Trend Chart

**User request:**

"Create a monthly sales trend chart."

**Recommended chart:** Line chart.

Use chronological month ordering.

### Example 22: Product Contribution Chart

**User request:**

"Show each product's contribution to total sales."

**Recommended chart:** Bar chart or pie chart, depending on the number of categories.

For many products, prefer a sorted bar chart.

---

## 10. Dashboard Examples

### Example 23: Create a Sales Dashboard

**User request:**

"Create a dashboard for my sales workbook."

**Possible dashboard components:**

**KPI cards:**
- Total Revenue.
- Total Orders.
- Average Order Value.
- Total Customers.

**Charts:**
- Monthly Revenue Trend.
- Sales by Region.
- Top 5 Products.
- Sales by Category.

**Filters:**
- Date.
- Region.
- Product Category.

**Workflow:**

1. Inspect the workbook.
2. Identify relevant fields.
3. Calculate KPIs.
4. Prepare summary tables.
5. Create charts.
6. Arrange dashboard components.
7. Validate calculations and filters.

Do not create KPIs when the required source data is unavailable.

---

## 11. KPI and Reporting Examples

### Example 24: Calculate Average Order Value

Formula:

```excel
=Total_Revenue/Total_Orders
```

Handle zero orders appropriately.

### Example 25: Calculate Profit Margin

Formula:

```excel
=IFERROR((Revenue-Cost)/Revenue,0)
```

Confirm that revenue and cost use compatible units and definitions.

### Example 26: Generate a Monthly Report

**User request:**

"Generate a monthly sales report."

**Expected report:**

1. Reporting period.
2. Total revenue.
3. Total orders.
4. Average order value.
5. Regional performance.
6. Product performance.
7. Monthly comparison.
8. Important observations.

Clearly identify the reporting period and source worksheet.

---

## 12. Workbook Operation Examples

### Example 27: List Worksheets

**User request:**

"Show all sheets in my workbook."

Expected operation:

1. Inspect workbook metadata.
2. Retrieve worksheet names.
3. Display the names.

### Example 28: Read a Worksheet

**User request:**

"Show the first 10 rows of the Sales sheet."

Expected operation:

1. Confirm that the Sales sheet exists.
2. Read the first 10 data rows.
3. Display the results.

### Example 29: Create a New Worksheet

**User request:**

"Create a sheet called Monthly Summary."

Expected operation:

1. Check whether the sheet already exists.
2. Create it only if necessary.
3. Confirm the result.

### Example 30: Save a Workbook

**User request:**

"Save the updated workbook."

Expected operation:

1. Identify the active workbook.
2. Save to the intended location.
3. Verify that the save operation succeeded.

Never report a successful save without confirmation from the tool.

---

## 13. Multi-Step Task Examples

### Example 31: Complete Sales Analysis

**User request:**

"Clean my sales data, calculate revenue, analyze monthly performance, and create a chart."

**Execution plan:**

1. Inspect workbook and worksheet.
2. Identify required columns.
3. Check missing values and duplicates.
4. Clean the data according to appropriate rules.
5. Calculate revenue.
6. Aggregate revenue by month and year.
7. Analyze monthly performance.
8. Create a line chart.
9. Validate calculations and chart data.
10. Report completed operations.

### Example 32: Customer Analysis

**User request:**

"Analyze customer data and identify the top customers."

**Execution plan:**

1. Inspect customer and transaction data.
2. Identify customer identifiers.
3. Calculate total purchase value per customer.
4. Sort by purchase value.
5. Identify the top customers.
6. Present a summary table.
7. Create a chart if requested.

### Example 33: Inventory Analysis

**User request:**

"Find products with low stock."

**Execution plan:**

1. Inspect the inventory worksheet.
2. Identify product and stock columns.
3. Ask for the low-stock threshold if it is not defined.
4. Filter products below the threshold.
5. Display the matching products.
6. Report the number of products requiring attention.

---

## 14. Follow-Up Question Examples

The agent should maintain task context during follow-up requests.

Example conversation:

**User:** Analyze my sales data.

**Agent:** I found the Sales worksheet. It contains Order Date, Region, Product, Quantity, and Revenue columns. Would you like a general summary or a regional comparison?

**User:** Regional comparison.

**Agent:** I will compare total revenue across regions.

**User:** Create a chart too.

**Agent:** I will use the regional summary to create a chart.

**User:** Show only the top three regions.

**Agent:** I will filter the regional results to the top three based on total revenue.

The agent should preserve the existing analysis context instead of starting an unrelated task.

---

## 15. AI Agent Instruction Guidelines

Use these principles when processing Excel requests.

### 15.1 Understand Intent

Classify the request as one or more of:

- Information retrieval.
- Data reading.
- Data analysis.
- Data cleaning.
- Data transformation.
- Formula creation.
- Workbook modification.
- Chart creation.
- Dashboard creation.
- Troubleshooting.

### 15.2 Inspect Before Acting

Before modifying a workbook:

- Inspect its structure.
- Confirm worksheet names.
- Confirm column names.
- Determine the relevant data range.
- Check whether the requested operation is possible.

### 15.3 Preserve Original Data

- Avoid overwriting original data unnecessarily.
- Prefer creating a new worksheet or output column when appropriate.
- Create backups before high-impact operations when supported.
- Confirm destructive operations such as deleting sheets or removing records.
- Do not silently discard information.

### 15.4 Select the Correct Tool

Choose tools based on the operation.

| Task | Tool category |
|---|---|
| Read data | Reader |
| Write data | Writer |
| Apply formulas | Formula |
| Clean data | Cleaning |
| Sort and filter | Reader or manipulation tools |
| Create charts | Charts |
| Create dashboards | Dashboard |
| Retrieve guidance | Knowledge Base / RAG |

Use the actual tools implemented in the application. Do not assume a tool exists just because it appears in this table.

### 15.5 Validate Every Modification

After writing or modifying data:

- Re-read the affected range when possible.
- Check row and column counts.
- Verify formulas.
- Confirm the output location.
- Check for errors.

### 15.6 Communicate Clearly

A useful completion message should state:

- Operation performed.
- Worksheet affected.
- Important results.
- Output location, if applicable.
- Any remaining issues.

Do not claim that an operation was completed when it was only planned.

---

## 16. Read-Only and Write Operations

The agent must distinguish between operations that only inspect data and operations that modify it.

**Read-only operations:**
- Display worksheet names.
- Read rows and columns.
- Calculate statistics without writing to the workbook.
- Analyze trends.
- Answer questions about existing data.

**Write operations:**
- Add or delete rows.
- Modify cell values.
- Insert formulas.
- Change formatting.
- Create worksheets.
- Create charts.
- Build dashboards.
- Remove duplicates.

For write operations, inspect the target and validate the result. Ask for confirmation when the operation is destructive, ambiguous, or difficult to reverse.

---

## 17. Error Recovery Examples

### Example 34: Worksheet Not Found

**Problem:** Requested worksheet does not exist.

**Agent response:**

"I couldn't find a worksheet with that name. These are the available worksheets: [worksheet names]. Which one should I use?"

### Example 35: Column Not Found

**Problem:** Requested column does not exist.

**Agent response:**

"I couldn't find the requested column. The available columns are [column names]. Please confirm which column contains the required information."

### Example 36: Formula Error

**Problem:** Formula returns an error.

**Recovery:**

1. Inspect the formula.
2. Check cell references.
3. Check data types.
4. Check division by zero.
5. Correct the formula if the intended calculation is clear.
6. Recalculate and validate.

### Example 37: Invalid Data

**Problem:** Numeric calculations encounter text values.

**Recovery:**

1. Identify invalid cells.
2. Inspect representative values.
3. Determine whether conversion is safe.
4. Convert only when the meaning is clear.
5. Report values that cannot be interpreted reliably.

---

## 18. Knowledge Base Usage Instructions

The Knowledge Base is a reference source for Excel concepts, procedures, examples, and troubleshooting.

When a user asks an Excel-related question:

1. Identify the relevant topic.
2. Retrieve relevant Knowledge Base content.
3. Use the retrieved content to understand the task.
4. Inspect the workbook when the task depends on actual data.
5. Select the required Excel tools.
6. Execute the operation when requested.
7. Validate the result.
8. Explain the outcome.

### Important Rules

- Do not answer using unrelated Knowledge Base content.
- Do not invent columns, worksheets, formulas, or data.
- Do not assume an Excel operation has been executed merely because instructions were retrieved.
- Do not use RAG as a replacement for Excel tools.
- Do not modify data when the user only requests an explanation.
- Ask clarification questions when important details are missing.

**Final principle:**

The Knowledge Base explains how to perform Excel tasks. The AI Agent understands the user's request and coordinates the workflow. Excel tools perform the actual workbook operations.