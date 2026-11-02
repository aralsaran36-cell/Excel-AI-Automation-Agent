# Excel Chart Selection Guide

## 1. Overview

Choosing the correct chart is an important part of data analysis. Different charts communicate different types of information.

A chart should be selected based on:
- The question being analyzed.
- The type of data available.
- The number of categories.
- The relationship between variables.
- The message the user wants to communicate.

The objective is not simply to create a chart. It is to choose a visualization that makes the data easy to understand.

---

## 2. Understand the Type of Analysis

Before selecting a chart, identify the analytical requirement.

### 2.1 Comparison

Comparison answers questions such as:
- Which region has the highest sales?
- Which product generates the most revenue?
- Which department has the highest expenses?

Recommended charts:
- Column Chart
- Bar Chart
- Clustered Bar Chart

Example:

| Region | Sales |
|---|---:|
| Chennai | 90000 |
| Coimbatore | 120000 |
| Salem | 75000 |

**Best choice:** Bar or Column Chart.

### 2.2 Trend Analysis

Trend analysis identifies how values change over time.

Questions:
- Are monthly sales increasing?
- How has revenue changed over the year?
- Is customer growth improving?

Recommended charts:
- Line Chart
- Area Chart
- Column Chart for discrete period comparisons

Example:

| Month | Revenue |
|---|---:|
| Jan | 50000 |
| Feb | 65000 |
| Mar | 70000 |
| Apr | 85000 |

**Best choice:** Line Chart.

### 2.3 Composition

Composition explains how individual categories contribute to a total.

Questions:
- What percentage of expenses belongs to each department?
- Which product contributes the most revenue?
- How is the budget distributed?

Recommended charts:
- Pie Chart
- Doughnut Chart
- Stacked Column Chart
- Treemap

**Best choice:** Pie or Doughnut Chart for a small number of categories representing one whole.

### 2.4 Distribution

Distribution explains how numerical values are spread.

Questions:
- What is the salary distribution?
- How are customer ages distributed?
- Are delivery times concentrated around a particular value?

Recommended charts:
- Histogram
- Box and Whisker Chart

**Best choice:** Histogram for frequency distribution.

### 2.5 Relationship

Relationship analysis examines how two or more numerical variables are associated.

Questions:
- Does advertising expenditure relate to sales?
- Is employee experience associated with salary?
- Does production volume relate to operating cost?

Recommended charts:
- Scatter Chart
- Bubble Chart

**Best choice:** Scatter Chart for two numerical variables.

### 2.6 Ranking

Ranking identifies the highest- and lowest-performing categories.

Questions:
- Who are the top five sales employees?
- Which products generate the least revenue?
- Which branches perform best?

Recommended charts:
- Horizontal Bar Chart
- Column Chart

**Best choice:** Sorted Bar Chart.

### 2.7 Financial Change Analysis

Financial change analysis explains how positive and negative values contribute to a final result.

Questions:
- How does revenue become net profit?
- What caused the budget increase?
- How did cash flow change?

Recommended chart:
- Waterfall Chart.

### 2.8 Hierarchical Analysis

Hierarchical analysis displays data organized into levels.

Example:

```text
Company
 ├── Electronics
 │    ├── Laptop
 │    └── Mobile
 └── Furniture
      ├── Chair
      └── Table
```

Recommended charts:
- Treemap
- Sunburst Chart.

---

## 3. Chart Selection Reference Table

| Business Requirement | Recommended Chart |
|---|---|
| Compare products | Bar or Column |
| Compare regions | Bar or Column |
| Monthly sales trends | Line |
| Yearly revenue trends | Line |
| Expense contribution | Pie or Doughnut |
| Department budget share | Doughnut |
| Salary distribution | Histogram |
| Employee salary spread | Box and Whisker |
| Advertising vs Sales | Scatter |
| Sales, Profit and Market Size | Bubble |
| Actual vs Target | Combo |
| Revenue to Net Profit | Waterfall |
| Category and Subcategory contribution | Treemap |
| Regional geographical performance | Map |
| Cumulative performance | Area |
| Top 10 products | Sorted Bar |
| Multiple categories over time | Line or Stacked Area, depending on the question |

---

## 4. Understand Data Structure Before Selecting a Chart

### 4.1 Categorical Data

Categorical data represents groups or labels.

Examples:
- Region.
- Product.
- Department.
- Gender.
- Customer segment.

Useful charts:
- Bar.
- Column.
- Pie, when showing a small number of parts of a whole.

### 4.2 Numerical Data

Numerical data represents measurable quantities.

Examples:
- Revenue.
- Salary.
- Quantity.
- Profit.
- Age.

Useful charts:
- Histogram.
- Scatter.
- Box and Whisker.
- Column or Bar for aggregated categories.

### 4.3 Time-Based Data

Time-based data represents observations across dates or periods.

Examples:
- Daily sales.
- Monthly revenue.
- Quarterly profit.
- Yearly customer growth.

Useful charts:
- Line.
- Area.
- Column.

### 4.4 Multiple Numerical Variables

When a dataset contains multiple numerical measurements, choose a chart based on the relationship being studied.

Example:

| Advertising Cost | Sales | Profit |
|---:|---:|---:|
| 10000 | 50000 | 12000 |
| 15000 | 65000 | 18000 |
| 20000 | 85000 | 25000 |

Possible visualizations:
- Scatter Chart: Advertising Cost versus Sales.
- Bubble Chart: Advertising Cost, Sales, and Profit.

---

## 5. Step-by-Step Chart Selection Process

Use this process before creating any chart.

### Step 1: Understand the User's Question

Identify what the user wants to know.

Example:
"Show which region has the highest sales."

Analytical requirement: Comparison and ranking.

### Step 2: Identify the Required Columns

For regional sales:
- Region.
- Sales.

### Step 3: Inspect the Data

Check:
- Data types.
- Missing values.
- Duplicate records.
- Incorrect values.
- Number of categories.

### Step 4: Determine Whether Aggregation Is Required

If multiple rows exist for each region, calculate the appropriate summary.

Example:

| Region | Sales |
|---|---:|
| Chennai | 30000 |
| Chennai | 40000 |
| Salem | 25000 |
| Salem | 35000 |

Summarized result:

| Region | Total Sales |
|---|---:|
| Chennai | 70000 |
| Salem | 60000 |

### Step 5: Select the Chart

For this example, use a Bar Chart or Column Chart.

### Step 6: Configure the Chart

- Categories: Region.
- Values: Total Sales.
- Title: Total Sales by Region.
- Axis: Sales amount.

### Step 7: Verify the Visualization

Check that:
- The correct values are displayed.
- Categories are readable.
- The chart communicates the requested comparison.
- No important data is missing.

---

## 6. Common Chart Selection Mistakes

### Mistake 1: Using Pie Charts for Too Many Categories

**Problem:** A pie chart with many slices becomes difficult to understand.

**Better approach:** Use a sorted Bar Chart.

### Mistake 2: Using Column Charts for Long Category Names

**Problem:** Long labels overlap or become difficult to read.

**Better approach:** Use a horizontal Bar Chart.

### Mistake 3: Using Pie Charts for Trends

**Problem:** Pie charts cannot communicate changes over time effectively.

**Better approach:** Use a Line Chart.

### Mistake 4: Using Line Charts for Unordered Categories

**Problem:** Connecting unrelated categories can imply a false sequence.

**Better approach:** Use a Bar or Column Chart.

### Mistake 5: Using a Histogram for Ordinary Categories

**Problem:** Histograms represent numerical distributions, not arbitrary category counts.

**Better approach:** Use a Bar Chart.

### Mistake 6: Using a Secondary Axis Without Explanation

**Problem:** Different scales can make comparisons misleading.

**Better approach:** Use a clearly labeled secondary axis only when necessary.

### Mistake 7: Using 3D Charts Without a Need

**Problem:** 3D effects can distort visual comparisons.

**Better approach:** Prefer simple 2D charts.

### Mistake 8: Showing Too Many Series

**Problem:** Too many lines, colors, or bars create visual clutter.

**Better approach:** Filter the data or create separate charts.

---

## 7. Advanced Chart Selection

### 7.1 Clustered vs Stacked Column Chart

**Clustered Column Chart**

Compares multiple series side by side.

Example: Compare Laptop and Mobile sales across regions.

**Stacked Column Chart**

Shows how multiple series contribute to a total.

Example: Show total regional sales split into Laptop and Mobile sales.

Choose based on whether the priority is comparing individual series or understanding their combined total.

### 7.2 Line vs Area Chart

**Line Chart**
- Emphasizes trends.
- Works well for comparing multiple time series.
- Keeps individual series visually distinct.

**Area Chart**
- Emphasizes magnitude and cumulative contribution.
- Can become difficult to read when many series overlap.

### 7.3 Bar vs Column Chart

**Bar Chart**
- Suitable for many categories.
- Works well with long labels.
- Useful for rankings.

**Column Chart**
- Suitable for fewer categories.
- Useful for comparisons across discrete periods.

### 7.4 Scatter vs Line Chart

**Scatter Chart**
- Examines relationships between numerical variables.
- Uses numerical X and Y values.

**Line Chart**
- Shows ordered progression, commonly over time.
- Connects observations in sequence.

### 7.5 Combo Chart

Use a Combo Chart when two related measurements need different visual representations.

Example:

| Month | Revenue | Profit Margin |
|---|---:|---:|
| Jan | 100000 | 15 |
| Feb | 120000 | 18 |
| Mar | 150000 | 20 |

Possible configuration:
- Revenue: Column.
- Profit Margin: Line.
- Profit Margin: Secondary axis, if needed.

Clearly label the secondary axis and its unit.

---

## 8. Real-World Business Examples

### Sales Performance

Requirement: Identify the top-performing products.

- Aggregate sales by product.
- Sort products by total sales.
- Use a horizontal Bar Chart.
- Display the top products.

### Financial Reporting

Requirement: Understand the movement from revenue to net profit.

- Identify revenue and expense components.
- Calculate the financial changes.
- Use a Waterfall Chart.

### HR Analytics

Requirement: Understand employee salary distribution.

- Extract salary values.
- Validate numerical data.
- Create a Histogram to show salary frequency.
- Use a Box and Whisker Chart to examine spread and potential outliers.

### Marketing Analytics

Requirement: Examine whether advertising expenditure is associated with sales.

- Use advertising cost and sales as numerical variables.
- Create a Scatter Chart.
- Add a trendline if useful.
- Avoid claiming causation from correlation alone.

### Inventory Analytics

Requirement: Compare stock distribution across product categories.

- Aggregate stock quantities by category.
- Use a Bar or Column Chart.
- Use a Treemap when hierarchical category contribution is important.

---

## 9. Excel AI Agent Instructions

When a user requests a chart, the AI Agent must identify the analytical purpose before choosing a chart type.

### Chart Selection Decision Rules

1. Comparison → Bar or Column.
2. Trend over time → Line.
3. Composition → Pie, Doughnut, or Stacked Chart.
4. Numerical distribution → Histogram or Box and Whisker.
5. Numerical relationship → Scatter.
6. Three-variable relationship → Bubble.
7. Ranking → Sorted Bar.
8. Financial changes → Waterfall.
9. Two different measurements → Combo.
10. Hierarchical contribution → Treemap or Sunburst.
11. Geographical analysis → Map.
12. Cumulative trends → Area.

### Agent Workflow

1. Understand the user's request.
2. Identify the analytical purpose.
3. Locate the required source columns.
4. Inspect the data types and quality.
5. Aggregate data if necessary.
6. Select the most appropriate chart.
7. Create the chart.
8. Add a meaningful title and labels.
9. Format the chart.
10. Verify that the visualization answers the original question.

### Agent Restrictions

- Do not select a chart based only on the user's mention of a chart name if it would misrepresent the data. Explain the issue and suggest a suitable alternative.
- Do not aggregate data without understanding the required calculation.
- Do not assume that all numerical columns should be summed.
- Do not use a histogram for ordinary categorical comparisons.
- Do not use a pie chart for a large number of categories.
- Do not confuse correlation with causation.
- Do not invent missing data.
- Preserve the original dataset.
- Do not claim a chart was created until the workbook operation succeeds.

### Example User Requests

- "Which chart is best for comparing monthly sales?"
- "Show the top 10 products by revenue."
- "I want to visualize salary distribution."
- "Compare advertising cost and sales."
- "Show the contribution of each department to total expenses."
- "Create a chart comparing revenue and profit margin."

### Final Verification

Before completing the task, confirm:
- The chart matches the analytical purpose.
- The correct source data is used.
- Required calculations are correct.
- The chart is readable.
- Labels and units are accurate.
- The workbook is saved successfully.