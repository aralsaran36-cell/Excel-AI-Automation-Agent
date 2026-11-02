# Comparative Analysis in Excel

## 1. Overview and Fundamentals

Comparative analysis is the process of comparing two or more sets of data to identify differences, similarities, performance gaps, and opportunities for improvement.

Businesses use comparative analysis to understand:

- Which product generates more revenue?
- Which region performs better?
- How much did sales increase compared with last month?
- Are actual results meeting the target?
- Which department has the highest expenses?
- Did performance improve after a business change?

### 1.1 Comparative Analysis vs Trend Analysis

| Comparative Analysis | Trend Analysis |
|---|---|
| Compares two or more groups or values | Studies movement over time |
| Focuses on differences and performance gaps | Focuses on direction and patterns |
| Example: Compare Chennai and Coimbatore sales | Example: Study monthly sales growth |
| Uses variance, ratios, rankings and percentage differences | Uses time-based patterns and growth rates |

Both techniques can be used together.

### 1.2 Types of Comparative Analysis

**1. Period-to-Period Comparison**

Compare results across two periods.

Example:
- January sales: ₹80,000
- February sales: ₹95,000

**2. Month-over-Month (MoM) Comparison**

Compare a month with the previous month.

Example: Compare March sales with February sales.

**3. Year-over-Year (YoY) Comparison**

Compare a period with the same period in the previous year.

Example: Compare January 2026 sales with January 2025 sales.

**4. Actual vs Target Comparison**

Compare actual performance with a planned target.

Example:
- Target: ₹100,000
- Actual: ₹90,000

**5. Group or Segment Comparison**

Compare categories such as:
- Sales by region
- Revenue by product
- Expenses by department
- Employee performance by team

**6. Before-and-After Comparison**

Compare results before and after a business change.

Example: Compare monthly sales before and after a marketing campaign.

---

## 2. Core Concepts, Methods and Techniques

### 2.1 Absolute Difference

Absolute difference measures how much one value differs from another.

Formula:

`Difference = Current Value - Previous Value`

Example:

| Month | Sales |
|---|---:|
| January | 80,000 |
| February | 95,000 |

Difference = 95,000 - 80,000 = ₹15,000.

A positive result means an increase. A negative result means a decrease.

### 2.2 Percentage Change

Percentage change measures the relative increase or decrease compared with the original value.

Formula:

`Percentage Change = ((New Value - Old Value) / Old Value) * 100`

Example:

- Old sales = ₹80,000
- New sales = ₹95,000
- Difference = ₹15,000

Percentage Change = (15,000 / 80,000) × 100 = 18.75%.

The sales increased by 18.75%.

**Important:** If the original value is zero, percentage change is undefined. Do not divide by zero or automatically report the result as 0%.

### 2.3 Variance Analysis

Variance measures the difference between actual performance and expected or planned performance.

Formula:

`Variance = Actual Value - Target Value`

Example:

| Metric | Value |
|---|---:|
| Target | 100,000 |
| Actual | 90,000 |
| Variance | -10,000 |

The negative variance indicates that actual sales are ₹10,000 below the target.

Variance can be:
- Positive: Actual is greater than target.
- Negative: Actual is below target.
- Zero: Actual equals target.

Whether a positive variance is good or bad depends on the metric. Higher revenue is usually positive, while higher expenses may be negative.

### 2.4 Percentage Variance

Percentage variance expresses the variance relative to the target.

Formula:

`Percentage Variance = ((Actual - Target) / Target) * 100`

Example:

- Target = ₹100,000
- Actual = ₹90,000
- Variance = -₹10,000
- Percentage variance = -10%.

### 2.5 Target Achievement Percentage

Measures how much of the target has been achieved.

Formula:

`Achievement % = (Actual / Target) * 100`

Example:

Actual = ₹90,000 and Target = ₹100,000.

Achievement = 90%.

### 2.6 Percentage Change vs Percentage-Point Difference

These are not the same.

**Percentage change** measures relative growth or decline.

**Percentage-point difference** measures the direct subtraction of two percentages.

Example:

- Previous profit margin = 20%
- Current profit margin = 25%

Percentage-point difference = 25% - 20% = 5 percentage points.

Percentage change in the margin = ((25% - 20%) / 20%) × 100 = 25%.

Therefore:
- Increase = 5 percentage points.
- Relative increase = 25%.

### 2.7 Ratio Comparison

Ratios compare two related quantities.

Example:

- Product A sales = ₹200,000
- Product B sales = ₹100,000

Sales ratio = 2:1.

Product A generated twice the sales of Product B.

### 2.8 Index and Base Comparison

An index expresses values relative to a selected base value.

Formula:

`Index = (Current Value / Base Value) * 100`

Example:

- Base-year sales = ₹100,000
- Current-year sales = ₹125,000

Index = 125.

An index of 125 means the current value is 25% higher than the base value.

### 2.9 Ranking

Ranking arranges categories according to their performance.

Examples:
- Highest-selling products.
- Best-performing regions.
- Departments with the highest expenses.

Excel functions:
- `RANK.EQ`
- `SORT`
- `SORTBY`

---

## 3. Step-by-Step Implementation in Excel

### 3.1 Compare Monthly Sales

Create the following worksheet:

| Month | Previous Sales | Current Sales |
|---|---:|---:|
| January | 80000 | 95000 |
| February | 90000 | 85000 |
| March | 75000 | 100000 |
| April | 110000 | 120000 |

Add three columns:
- Difference
- Percentage Change
- Performance

Assume:
- Column B = Previous Sales
- Column C = Current Sales
- Column D = Difference
- Column E = Percentage Change
- Column F = Performance

**Step 1: Calculate Difference**

In D2:

`=C2-B2`

**Step 2: Calculate Percentage Change**

In E2:

`=IF(B2=0,"N/A",(C2-B2)/B2)`

Format column E as Percentage.

**Step 3: Identify Performance**

In F2:

`=IF(C2>B2,"Increased",IF(C2<B2,"Decreased","No Change"))`

**Step 4: Apply formulas to remaining rows**

Drag the formulas down.

### 3.2 Compare Actual Sales with Targets

Create the following table:

| Region | Target | Actual |
|---|---:|---:|
| Chennai | 100000 | 120000 |
| Coimbatore | 90000 | 85000 |
| Salem | 80000 | 80000 |
| Madurai | 75000 | 60000 |

Add these columns:

**Variance**

`=C2-B2`

**Variance Percentage**

`=IF(B2=0,"N/A",(C2-B2)/B2)`

**Achievement Percentage**

`=IF(B2=0,"N/A",C2/B2)`

**Performance Status**

`=IF(C2>B2,"Above Target",IF(C2<B2,"Below Target","Target Achieved"))`

Format percentage columns as percentages.

### 3.3 Compare Product Performance

Example:

| Product | Sales |
|---|---:|
| Laptop | 250000 |
| Mobile | 180000 |
| Tablet | 90000 |
| Monitor | 120000 |

To rank products, enter this formula in C2:

`=RANK.EQ(B2,$B$2:$B$5,0)`

The highest sales value receives rank 1.

### 3.4 Conditional Formatting

Conditional formatting makes differences easier to understand.

Steps:

1. Select the Variance column.
2. Open Home.
3. Select Conditional Formatting.
4. Choose Highlight Cells Rules.
5. Apply formatting for positive and negative values.

Useful formatting:
- Green: Positive variance.
- Red: Negative variance.
- Yellow: Values requiring attention.

For percentage achievement:
- Green: At or above 100%.
- Red: Below 100%.

### 3.5 Create Comparison Charts

Choose charts based on the comparison:

| Requirement | Recommended Chart |
|---|---|
| Compare sales by region | Clustered Column |
| Compare actual vs target | Clustered Bar or Column |
| Compare product performance | Bar Chart |
| Compare percentage achievement | Bar Chart |
| Compare multiple categories | Clustered Column |
| Show positive and negative variance | Bar Chart |

Steps:

1. Select the comparison data.
2. Open Insert.
3. Choose the appropriate chart.
4. Add a meaningful chart title.
5. Add axis labels when useful.
6. Check that categories and values are displayed correctly.

### 3.6 Use PivotTables for Group Comparisons

When a dataset contains many records, use a PivotTable.

Example dataset:

| Date | Region | Product | Sales |
|---|---|---|---:|
| 01-01-2026 | Chennai | Laptop | 50000 |
| 02-01-2026 | Salem | Mobile | 30000 |
| 03-01-2026 | Chennai | Mobile | 40000 |
| 04-01-2026 | Salem | Laptop | 60000 |

Steps:

1. Select the dataset.
2. Open Insert → PivotTable.
3. Drag Region into Rows.
4. Drag Product into Columns.
5. Drag Sales into Values.
6. Compare sales across regions and products.

This creates a summarized comparison without manually calculating every category.

---

## 4. Python Integration and Real-World Business Use Cases

Python Pandas can automate comparative analysis across large Excel datasets.

### 4.1 Load Excel Data

```python
import pandas as pd

df = pd.read_excel("sales.xlsx")

print(df.head())
```

### 4.2 Compare Sales by Region

```python
region_sales = df.groupby("Region")["Sales"].sum()

print(region_sales)
```

This calculates total sales for each region.

### 4.3 Compare Sales by Product

```python
product_sales = df.groupby("Product")["Sales"].sum()

print(product_sales.sort_values(ascending=False))
```

This identifies products with the highest and lowest sales.

### 4.4 Create a Region and Product Comparison

```python
comparison = pd.pivot_table(
    df,
    index="Region",
    columns="Product",
    values="Sales",
    aggfunc="sum",
    fill_value=0
)

print(comparison)
```

This produces a comparison table with regions as rows and products as columns.

### 4.5 Compare Two Periods

Assume the dataset contains `Date`, `Region`, and `Sales`.

```python
df["Date"] = pd.to_datetime(df["Date"])

df["Month"] = df["Date"].dt.to_period("M")

monthly_sales = df.groupby("Month")["Sales"].sum()

comparison = pd.DataFrame({
    "Sales": monthly_sales
})

comparison["Previous Month"] = comparison["Sales"].shift(1)

comparison["Difference"] = (
    comparison["Sales"] - comparison["Previous Month"]
)

comparison["Percentage Change"] = (
    comparison["Difference"] /
    comparison["Previous Month"].replace(0, float("nan"))
) * 100

print(comparison)
```

This calculates monthly sales differences and percentage changes.

### 4.6 Compare Actual vs Target

Assume a DataFrame contains `Region`, `Target`, and `Actual`.

```python
df["Variance"] = df["Actual"] - df["Target"]

df["Achievement"] = (
    df["Actual"] / df["Target"].replace(0, float("nan"))
) * 100

df["Status"] = "Below Target"

df.loc[
    df["Actual"] > df["Target"],
    "Status"
] = "Above Target"

df.loc[
    df["Actual"] == df["Target"],
    "Status"
] = "Target Achieved"

print(df)
```

This calculates target variance and achievement percentage.

### 4.7 Compare Two DataFrames

When comparing two datasets, use `merge()` to align records by a common key.

```python
old_data = pd.DataFrame({
    "Product": ["Laptop", "Mobile", "Tablet"],
    "Sales": [80000, 50000, 30000]
})

new_data = pd.DataFrame({
    "Product": ["Laptop", "Mobile", "Tablet"],
    "Sales": [95000, 45000, 40000]
})

comparison = old_data.merge(
    new_data,
    on="Product",
    suffixes=("_Old", "_New")
)

comparison["Difference"] = (
    comparison["Sales_New"] - comparison["Sales_Old"]
)

print(comparison)
```

This compares the same products across two datasets.

Important:
- Use a reliable matching key.
- Check for duplicate keys before merging.
- Check for missing records.
- Avoid treating unmatched records as zero unless that is the intended business rule.

### 4.8 Real-World Business Use Cases

**Sales Analysis**
- Compare monthly sales.
- Identify high-performing regions.
- Compare product revenue.
- Measure target achievement.

**Finance**
- Compare budget and actual expenses.
- Identify cost overruns.
- Compare quarterly revenue.
- Measure profit differences.

**Human Resources**
- Compare employee performance.
- Compare department headcount.
- Analyze attendance differences.
- Compare recruitment costs.

**Inventory**
- Compare stock across warehouses.
- Identify differences between expected and actual inventory.
- Compare product demand.
- Identify slow-moving products.

**Marketing**
- Compare campaign performance.
- Compare customer acquisition costs.
- Measure conversion-rate differences.
- Compare revenue generated by campaigns.

**Operations**
- Compare production output.
- Compare delivery performance.
- Measure defect-rate differences.
- Compare operational costs.

---

## 5. Best Practices, Common Mistakes and Troubleshooting

### 5.1 Best Practices

1. Compare values that use the same units and definitions.
2. Use equivalent periods when comparing time-based data.
3. Select an appropriate baseline.
4. Calculate both absolute and percentage differences when useful.
5. Clearly distinguish percentage changes from percentage-point differences.
6. Check missing values before calculating comparisons.
7. Handle zero denominators safely.
8. Use consistent number formats.
9. Use PivotTables for large category-based comparisons.
10. Explain whether a positive variance is favorable for the metric being analyzed.

### 5.2 Common Mistakes

**Mistake 1: Using the Wrong Baseline**

Comparing January 2026 with December 2025 may not be appropriate when seasonal effects are important.

Solution: Use a suitable comparison period, such as January 2025 for a YoY comparison.

**Mistake 2: Confusing Percentage Change and Percentage Points**

A change from 20% to 25% is a 5-percentage-point increase, not a 5% increase.

Solution: State both measurements correctly when necessary.

**Mistake 3: Dividing by Zero**

Percentage change cannot be calculated using a zero baseline.

Solution: Return N/A or another clearly defined result.

**Mistake 4: Comparing Different Units**

Comparing revenue in rupees with revenue in dollars without conversion produces misleading results.

Solution: Standardize units before comparison.

**Mistake 5: Ignoring Missing Categories**

A product missing from one dataset may indicate a new or discontinued product rather than zero sales.

Solution: Identify unmatched records and apply an explicit business rule.

**Mistake 6: Misinterpreting Positive Variance**

Higher expenses are not necessarily good.

Solution: Interpret variance according to the metric.

### 5.3 Troubleshooting

| Problem | Possible Cause | Solution |
|---|---|---|
| Percentage change shows an error | Original value is zero | Use IF or handle zero denominators |
| Incorrect comparison results | Wrong matching key | Check the key and data alignment |
| Duplicate comparison records | Duplicate keys in source data | Check duplicates before merging |
| Unexpected negative variance | Actual is below target | Verify target and actual columns |
| PivotTable shows incorrect totals | Wrong aggregation method | Check Sum, Count or Average |
| Percentage values look incorrect | Cell formatting issue | Apply Percentage format |
| Some categories are missing | Source datasets contain different categories | Identify unmatched records |

---

## 6. Excel AI Agent: Comparative Analysis Instructions

The Excel AI Agent should identify comparison requests, select the correct calculation, and use the appropriate Excel tools.

### 6.1 Natural-Language Requests

Examples:

- "Compare this month's sales with last month's sales."
- "Show the percentage difference between Chennai and Salem."
- "Compare actual sales with the target."
- "Find the top five products by revenue."
- "Compare department expenses."
- "Show the difference between 2025 and 2026 sales."
- "Identify products whose sales decreased."
- "Create an actual vs target comparison chart."

### 6.2 Agent Workflow

1. Understand what the user wants to compare.
2. Identify the relevant worksheet, columns, categories and periods.
3. Read the Excel data.
4. Validate the data and identify missing or zero values.
5. Select the appropriate comparison method.
6. Calculate differences, percentage changes, variances or rankings.
7. Write the results into the workbook when requested.
8. Create a comparison chart when requested.
9. Summarize the important findings in simple language.

### 6.3 Choosing the Correct Method

| User Request | Required Method |
|---|---|
| Compare two months | Period comparison |
| Compare this year with last year | YoY comparison |
| Compare actual and target | Variance analysis |
| Compare two regions | Group comparison |
| Find the highest-selling product | Ranking |
| Calculate sales growth | Percentage change |
| Compare profit margins | Percentage-point difference and relative change |
| Compare two Excel files | Dataset merge and comparison |

### 6.4 Important Agent Rules

- Do not assume which columns represent actual values and targets.
- Ask for clarification when the comparison period or metric is unclear.
- Do not divide by zero.
- Do not confuse percentage change with percentage-point difference.
- Preserve the original workbook data unless the user requests changes.
- Explain the calculation and result clearly.
- Distinguish between explaining comparative analysis and actually modifying an Excel workbook.

**Final Objective:** The Excel AI Agent should compare data accurately, identify performance gaps, calculate meaningful differences, and present clear business insights.