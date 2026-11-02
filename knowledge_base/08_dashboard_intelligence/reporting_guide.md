# Excel Reporting Guide

## 1. Overview

Excel reporting is the process of organizing, analyzing, and presenting data in a structured format to communicate useful information.

Reports help businesses:
- Monitor performance.
- Track important metrics.
- Identify trends and problems.
- Compare results across periods.
- Communicate findings.
- Support business decisions.

A good report should be accurate, readable, structured, and relevant to its intended audience.

## 2. Types of Excel Reports

### 2.1 Sales Report

Used to analyze sales performance.

Common information:
- Total revenue.
- Number of orders.
- Product-wise sales.
- Region-wise sales.
- Monthly revenue.
- Salesperson performance.
- Target achievement.

### 2.2 Financial Report

Used to monitor financial performance.

Common information:
- Revenue.
- Expenses.
- Gross profit.
- Net profit.
- Budget versus actual.
- Cash flow.

### 2.3 HR Report

Used to analyze workforce information.

Common information:
- Employee count.
- Department-wise headcount.
- Attendance.
- Employee turnover.
- New hires.
- Leave information.

### 2.4 Inventory Report

Used to monitor stock levels.

Common information:
- Product name.
- Opening stock.
- Purchased quantity.
- Sold quantity.
- Closing stock.
- Reorder level.
- Stock status.

### 2.5 Operational Report

Used to monitor business operations.

Common information:
- Production output.
- Defect count.
- Delivery performance.
- Processing time.
- Productivity.
- Pending work.

### 2.6 Management Report

Provides a high-level summary for decision-makers.

Common information:
- Business KPIs.
- Performance against targets.
- Important trends.
- Risks and issues.
- Recommended actions.

## 3. Report Structure

A professional report generally contains the following sections.

### 3.1 Report Header

Include:
- Report title.
- Organization or department, when relevant.
- Reporting period.
- Report preparation date.
- Data source, when useful.

Example:

**Monthly Sales Performance Report**

Reporting Period: October 2026

### 3.2 Executive Summary

Provides a short explanation of the most important findings.

Example:

- Total revenue reached ₹8,50,000.
- Revenue increased by 12% compared with the previous month.
- The South region generated the highest revenue.
- Product A contributed the largest share of sales.

Only include findings supported by the actual data.

### 3.3 KPI Summary

| KPI | Actual | Target | Status |
|---|---:|---:|---|
| Revenue | ₹8,50,000 | ₹10,00,000 | Below Target |
| Orders | 1,250 | 1,200 | Above Target |
| Customer Satisfaction | 92% | 95% | Below Target |

The values are illustrative.

### 3.4 Detailed Analysis

Include relevant breakdowns:
- Category-wise performance.
- Region-wise performance.
- Monthly trends.
- Top and bottom performers.
- Comparisons with previous periods.

### 3.5 Findings and Recommendations

Explain:
- What happened?
- Where did it happen?
- How significant was it?
- What may have contributed to it?
- What action should be considered?

Do not present correlation as proven causation.

## 4. Preparing Data for Reporting

Reliable reports require reliable source data.

### 4.1 Inspect the Dataset

Check:
- Column names.
- Data types.
- Missing values.
- Duplicate records.
- Invalid dates.
- Inconsistent categories.
- Incorrect numerical values.

### 4.2 Convert Data into an Excel Table

Steps:

1. Select the dataset.
2. Press `Ctrl + T` on Windows or use Insert → Table.
3. Confirm that the table has headers.
4. Assign a meaningful table name.

Example:

`SalesData`

Excel Tables provide structured references and automatically expand when new rows are added.

### 4.3 Clean the Data

Common cleaning operations:
- Remove unwanted spaces.
- Standardize category names.
- Correct date formats.
- Handle missing values.
- Remove duplicates when appropriate.
- Validate numerical fields.

Do not remove records automatically without checking whether they are genuinely duplicates or invalid.

## 5. Excel Formulas for Reporting

### 5.1 Total Revenue

```excel
=SUM(SalesData[Revenue])
```

### 5.2 Total Orders

```excel
=COUNTA(SalesData[Order_ID])
```

This assumes each valid order ID appears once per order.

### 5.3 Region-wise Revenue

```excel
=SUMIFS(SalesData[Revenue],SalesData[Region],"South")
```

### 5.4 Product-wise Sales

```excel
=SUMIFS(SalesData[Revenue],SalesData[Product],"Laptop")
```

### 5.5 Monthly Revenue

```excel
=SUMIFS(
    SalesData[Revenue],
    SalesData[Order_Date],">="&DATE(2026,1,1),
    SalesData[Order_Date],"<"&DATE(2026,2,1)
)
```

This calculates January 2026 revenue when the date column contains valid Excel dates.

### 5.6 Average Revenue

```excel
=AVERAGE(SalesData[Revenue])
```

### 5.7 Highest Revenue

```excel
=MAX(SalesData[Revenue])
```

### 5.8 Lowest Revenue

```excel
=MIN(SalesData[Revenue])
```

### 5.9 Count Records by Status

```excel
=COUNTIF(SalesData[Status],"Completed")
```

## 6. PivotTables for Reporting

PivotTables are useful for summarizing large datasets without writing many individual formulas.

### Example: Region-wise Sales Report

Steps:

1. Select the source data.
2. Choose Insert → PivotTable.
3. Select the source range or Table.
4. Place Region in Rows.
5. Place Revenue in Values.
6. Confirm that the calculation is Sum.
7. Apply currency formatting.

Example result:

| Region | Total Revenue |
|---|---:|
| North | ₹2,50,000 |
| South | ₹3,50,000 |
| East | ₹1,50,000 |
| West | ₹1,00,000 |

### Example: Monthly Sales Report

1. Place Order Date in Rows.
2. Place Revenue in Values.
3. Group dates by Months and Years when appropriate.
4. Apply number formatting.
5. Add a chart if required.

### PivotTable Reporting Best Practices

- Use meaningful field names.
- Verify aggregation methods.
- Refresh after source data changes.
- Check date grouping.
- Avoid counting blank records unintentionally.
- Validate totals against the source data.

## 7. Report Visualization

Charts help readers understand information quickly.

| Reporting requirement | Suitable chart |
|---|---|
| Monthly performance | Line chart |
| Category comparison | Bar chart |
| Regional revenue | Column chart |
| Actual vs budget | Clustered bar or column chart |
| Market contribution | Bar or doughnut chart |
| Distribution | Histogram |
| KPI performance | KPI cards |

### Chart Reporting Guidelines

- Add descriptive titles.
- Include units where necessary.
- Use readable labels.
- Avoid excessive colors.
- Keep chart scales appropriate.
- Do not use 3D charts unnecessarily.
- Highlight important findings without distorting the data.

## 8. Periodic Reporting

Businesses often prepare reports on a regular schedule.

### 8.1 Daily Reports

Examples:
- Daily sales.
- Daily production.
- Daily attendance.
- Daily transactions.

### 8.2 Weekly Reports

Examples:
- Weekly sales performance.
- Weekly operational status.
- Weekly customer support activity.

### 8.3 Monthly Reports

Examples:
- Monthly revenue.
- Monthly expenses.
- Monthly employee turnover.
- Monthly KPI performance.

### 8.4 Quarterly Reports

Examples:
- Quarterly revenue growth.
- Quarterly budget performance.
- Regional comparisons.
- Business performance review.

### 8.5 Yearly Reports

Examples:
- Annual revenue.
- Annual profit.
- Year-over-year growth.
- Annual workforce analysis.

Always confirm the reporting period and date boundaries.

## 9. Report Comparison and Variance Analysis

Reports become more valuable when current results are compared with meaningful reference values.

### 9.1 Actual vs Target

```excel
=B2-C2
```

Where:
- B2 = Actual.
- C2 = Target.

### 9.2 Variance Percentage

```excel
=IFERROR((B2-C2)/C2,0)
```

### 9.3 Month-over-Month Growth

```excel
=IFERROR((B2-C2)/C2,0)
```

Where:
- B2 = Current month.
- C2 = Previous month.

### 9.4 Year-over-Year Growth

```excel
=IFERROR((B2-C2)/C2,0)
```

Where:
- B2 = Current year value.
- C2 = Previous year value for the comparable period.

Clearly label the comparison period. Do not compare incomplete periods with complete periods without explaining the limitation.

## 10. Report Formatting

A report should be easy to read and interpret.

### 10.1 Header Formatting

- Use a clear report title.
- Show the reporting period.
- Use consistent font styles.
- Keep headings visually distinct.

### 10.2 Table Formatting

- Use meaningful column names.
- Apply consistent number formats.
- Align numbers consistently.
- Use filters when useful.
- Highlight totals and subtotals.
- Avoid unnecessary merged cells.

### 10.3 Number Formatting

Examples:

| Data | Recommended display |
|---|---|
| Revenue | ₹8,50,000 |
| Percentage | 92.5% |
| Quantity | 1,250 units |
| Date | 01-Oct-2026 |
| Average time | 4.5 hours |

### 10.4 Conditional Formatting

Use conditional formatting to highlight:
- Negative growth.
- Values below target.
- High expenses.
- Delayed deliveries.
- Unusually high defect rates.

Use rules that match the business meaning of each metric.

## 11. Automating Excel Reports with Python

Python can help read, summarize, and export reporting data.

### 11.1 Read an Excel File

```python
import pandas as pd

df = pd.read_excel("sales.xlsx")

print(df.head())
```

### 11.2 Generate a Summary

```python
total_revenue = df["Revenue"].sum()
total_orders = df["Order_ID"].nunique()

summary = pd.DataFrame({
    "Metric": ["Total Revenue", "Total Orders"],
    "Value": [total_revenue, total_orders]
})

print(summary)
```

### 11.3 Generate Region-wise Report

```python
region_report = (
    df.groupby("Region")["Revenue"]
    .sum()
    .reset_index()
    .sort_values("Revenue", ascending=False)
)

print(region_report)
```

### 11.4 Generate Monthly Report

```python
df["Order_Date"] = pd.to_datetime(
    df["Order_Date"],
    errors="coerce"
)

df = df.dropna(subset=["Order_Date"])

df["Month"] = df["Order_Date"].dt.to_period("M")

monthly_report = (
    df.groupby("Month")["Revenue"]
    .sum()
    .reset_index()
)

print(monthly_report)
```

### 11.5 Export Reports to Excel

```python
with pd.ExcelWriter("sales_report.xlsx") as writer:
    summary.to_excel(
        writer,
        sheet_name="KPI Summary",
        index=False
    )

    region_report.to_excel(
        writer,
        sheet_name="Region Report",
        index=False
    )

    monthly_report.to_excel(
        writer,
        sheet_name="Monthly Report",
        index=False
    )
```

This creates an Excel workbook containing three report sheets. It does not automatically create a fully formatted interactive Excel dashboard.

## 12. Real-World Reporting Examples

### 12.1 Sales Performance Report

Source columns:
- Order ID.
- Order Date.
- Product.
- Region.
- Salesperson.
- Quantity.
- Revenue.

Possible report outputs:
- Total revenue.
- Total orders.
- Monthly sales.
- Region-wise revenue.
- Product performance.
- Salesperson ranking.

### 12.2 Employee Attendance Report

Source columns:
- Employee ID.
- Employee Name.
- Department.
- Attendance Date.
- Attendance Status.

Possible report outputs:
- Total employees.
- Present count.
- Absent count.
- Attendance percentage.
- Department-wise attendance.

### 12.3 Inventory Report

Source columns:
- Product ID.
- Product Name.
- Opening Stock.
- Purchased Quantity.
- Sold Quantity.
- Reorder Level.

Closing stock:

```excel
=B2+C2-D2
```

Where:
- B2 = Opening stock.
- C2 = Purchased quantity.
- D2 = Sold quantity.

Possible report outputs:
- Available stock.
- Low-stock products.
- Reorder requirements.
- Product-wise stock value.

### 12.4 Financial Report

Source columns:
- Transaction Date.
- Category.
- Revenue.
- Expenses.
- Department.

Possible report outputs:
- Total revenue.
- Total expenses.
- Net profit.
- Monthly profit.
- Department-wise expenses.
- Budget variance.

## 13. Report Validation

Before sharing a report, verify:

1. The source data is correct.
2. The reporting period is correct.
3. The formulas use the correct ranges.
4. The aggregation methods are appropriate.
5. The totals match the source data.
6. Duplicate records are handled correctly.
7. Missing values are understood.
8. Filters do not exclude required records.
9. Charts match the summarized data.
10. The report is readable and complete.

For financial or management reports, use additional review and approval where required.

## 14. Common Reporting Problems

| Problem | Possible solution |
|---|---|
| Incorrect totals | Verify source ranges and formulas |
| Missing records | Check filters and source data |
| Wrong monthly results | Verify date values and boundaries |
| PivotTable not updated | Refresh the PivotTable |
| Duplicate transactions | Validate unique identifiers |
| Incorrect percentages | Check numerator and denominator |
| Unreadable report | Improve layout and formatting |
| Inconsistent categories | Standardize source values |
| Exported report is cut off | Adjust print area and page setup |
| Formula errors | Inspect references and data types |

## 15. Best Practices

1. Define the report's purpose before creating it.
2. Identify the intended audience.
3. Use clean and structured source data.
4. Maintain consistent report layouts.
5. Use clear column headings.
6. Apply correct formulas and aggregations.
7. Include the reporting period.
8. Use meaningful charts.
9. Validate totals and calculations.
10. Separate source data from report outputs.
11. Document important calculation definitions.
12. Avoid unsupported conclusions.
13. Preserve the original workbook.
14. Review the report before sharing.

## 16. Excel AI Agent Instructions

When a user requests an Excel report, the AI Agent should:

1. Identify the report type and business objective.
2. Inspect workbook sheets and available data.
3. Identify required columns and reporting periods.
4. Determine which calculations and summaries are needed.
5. Ask for clarification when report requirements are ambiguous.
6. Clean or prepare data only when required and authorized.
7. Use appropriate Excel formulas, PivotTables, or Python data analysis.
8. Create the requested report structure.
9. Apply readable formatting.
10. Validate calculations and report totals.
11. Preserve original source data.
12. Explain the report sheets and important findings.
13. Never invent missing records, targets, or results.

### Example User Requests

- "Create a monthly sales report."
- "Generate a region-wise revenue report."
- "Prepare an employee attendance report."
- "Create a financial summary."
- "Compare this month's sales with last month."
- "Generate an inventory report."
- "Create a management summary."
- "Export the analysis into a new Excel workbook."
- "Find the top five products by revenue."

### Important Agent Rule

Reporting knowledge helps the AI Agent determine what information belongs in a report and how it should be presented. Actual report creation must be performed using the appropriate workbook tools. The agent must verify the generated results before confirming completion.