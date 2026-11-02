# Excel Data Formatting

## 1. Overview and Fundamentals

### 1.1 What Is Data Formatting?

Data Formatting is the process of changing how data appears in Excel cells without necessarily changing the underlying values.

It helps make spreadsheets readable, consistent, professional and easier to analyze.

Examples:
- Displaying dates as DD-MM-YYYY.
- Displaying numbers with two decimal places.
- Showing currency values with a currency symbol.
- Highlighting important values.
- Applying consistent fonts, colors and borders.
- Adjusting column widths and row heights.

### 1.2 Why Is Data Formatting Important?

Data formatting helps to:

- Improve spreadsheet readability.
- Present information professionally.
- Make important values easier to identify.
- Maintain consistency across worksheets.
- Improve business reporting.
- Highlight trends and exceptions.
- Make dashboards easier to understand.

### 1.3 Data Formatting vs. Data Cleaning

| Data Formatting | Data Cleaning |
|---|---|
| Changes how data is displayed | Corrects data quality problems |
| Focuses on appearance and presentation | Focuses on accuracy and consistency |
| Uses number formats, fonts and styles | Uses duplicate removal, missing-value handling and corrections |
| May not change underlying values | Often changes underlying data |

Example:

Changing `45000` to display as `₹45,000.00` is formatting. Correcting an incorrect value of `450000` to `45000` is data cleaning.

### 1.4 Types of Excel Formatting

1. Number formatting
2. Date and time formatting
3. Currency and accounting formatting
4. Text formatting
5. Cell formatting
6. Conditional formatting
7. Table formatting
8. Row and column formatting
9. Alignment and text wrapping
10. Custom number formatting

---

## 2. Core Concepts, Methods and Techniques

### 2.1 Number Formatting

Number formatting controls how numeric values appear.

Common number formats:

| Format | Example Value | Display |
|---|---:|---:|
| General | 1250.5 | 1250.5 |
| Number | 1250.5 | 1,250.50 |
| Percentage | 0.25 | 25% |
| Currency | 1250 | ₹1,250.00 |
| Accounting | 1250 | Currency-aligned display |
| Scientific | 1250000 | 1.25E+06 |

Steps:

1. Select the required cells.
2. Open the Home tab.
3. Find the Number group.
4. Choose the required format.
5. Adjust decimal places if necessary.

### 2.2 Date Formatting

Excel supports different date display formats.

Examples:

| Format | Display |
|---|---|
| DD-MM-YYYY | 31-10-2026 |
| MM/DD/YYYY | 10/31/2026 |
| DD-MMM-YYYY | 31-Oct-2026 |
| MMMM DD, YYYY | October 31, 2026 |
| YYYY-MM-DD | 2026-10-31 |

Steps:

1. Select date cells.
2. Press Ctrl + 1.
3. Select Date or Custom.
4. Choose the required format.
5. Click OK.

Important: Formatting does not automatically convert text that looks like a date into a genuine Excel date value.

### 2.3 Currency and Accounting Formatting

Currency formatting displays financial values with currency symbols.

Example:

```text
12500 → ₹12,500.00
```

Accounting formatting aligns currency symbols and decimal values vertically, making financial statements easier to read.

Use:
- Currency for ordinary financial values.
- Accounting for financial reports and statements.

### 2.4 Percentage Formatting

Percentage formatting displays decimal values as percentages.

Examples:

| Stored Value | Display |
|---:|---:|
| 0.25 | 25% |
| 0.75 | 75% |
| 1 | 100% |
| 0.125 | 12.5% |

Important: Applying percentage formatting to 25 displays 2500%, because Excel treats 25 as 25.0 rather than 0.25.

### 2.5 Text Formatting

Text formatting controls how text appears.

Common options:

- Font type.
- Font size.
- Bold.
- Italic.
- Underline.
- Font color.
- Cell fill color.
- Text alignment.

Example:

Use bold text and a background color for column headings.

### 2.6 Alignment and Text Wrapping

Excel provides horizontal and vertical alignment options.

Common options:

- Left alignment.
- Center alignment.
- Right alignment.
- Top alignment.
- Middle alignment.
- Bottom alignment.
- Wrap Text.
- Merge and Center.

Wrap Text displays long content on multiple lines within a cell.

Use Merge and Center carefully because merged cells can make sorting, filtering and automation more difficult.

### 2.7 Cell Borders and Fill Colors

Borders help separate rows, columns and sections.

Fill colors help distinguish:
- Headers.
- Important totals.
- Input cells.
- Calculated fields.
- Warning values.

Avoid using too many colors because they reduce readability.

### 2.8 Conditional Formatting

Conditional Formatting automatically changes cell appearance based on specified conditions.

Common options:

- Highlight Cells Rules.
- Top/Bottom Rules.
- Data Bars.
- Color Scales.
- Icon Sets.
- Formula-based rules.

Example: Highlight sales values greater than 100000.

Steps:

1. Select the sales range.
2. Open Home.
3. Select Conditional Formatting.
4. Choose Highlight Cells Rules.
5. Select Greater Than.
6. Enter 100000.
7. Choose a formatting style.

### 2.9 Custom Number Formatting

Custom number formats allow users to control how numbers, dates and text are displayed.

Examples:

**Display numbers with two decimal places:**

```text
0.00
```

**Display numbers with thousand separators:**

```text
#,##0
```

**Display positive and negative values differently:**

```text
#,##0;[Red]-#,##0
```

**Display zero as a dash:**

```text
#,##0;[Red]-#,##0;-
```

**Display values with a percentage symbol:**

```text
0.00%
```

Custom number formats change display, not the underlying numeric value.

### 2.10 Format Painter

Format Painter copies formatting from one cell or range to another.

Steps:

1. Select the cell containing the desired formatting.
2. Click Format Painter.
3. Select the destination cells.

It copies formatting, not the underlying data.

### 2.11 Format as Table

Excel tables provide built-in formatting and structured data management.

Benefits:

- Consistent row formatting.
- Automatic filter buttons.
- Structured references.
- Automatic expansion of table ranges.
- Easier formula management.

Steps:

1. Select the dataset.
2. Press Ctrl + T.
3. Confirm the table range.
4. Select My table has headers if applicable.
5. Choose a table style.

---

## 3. Step-by-Step Implementation with Excel Examples

### 3.1 Example: Format an Employee Dataset

Dataset:

| Employee ID | Name | Salary | Joining Date |
|---|---|---:|---|
| E101 | Arun | 25000 | 2025-01-15 |
| E102 | Priya | 32000 | 2025-03-10 |
| E103 | Kumar | 28000 | 2025-05-20 |

Requirements:

- Make headings bold.
- Display salary as currency.
- Display dates in DD-MM-YYYY format.
- Adjust column widths.
- Apply borders.

Implementation:

1. Select the heading row and apply Bold.
2. Select the Salary column and choose Currency.
3. Select the Joining Date column and choose a date format.
4. Adjust column widths using AutoFit.
5. Apply borders to the dataset.

### 3.2 Example: Highlight High Sales

Dataset:

| Product | Sales |
|---|---:|
| Laptop | 85000 |
| Mobile | 45000 |
| Monitor | 120000 |
| Keyboard | 15000 |

Requirement: Highlight sales greater than 50000.

Steps:

1. Select the Sales column.
2. Open Conditional Formatting.
3. Select Highlight Cells Rules.
4. Choose Greater Than.
5. Enter 50000.
6. Select a formatting style.

Excel highlights Laptop and Monitor.

### 3.3 Example: Display Negative Values in Red

Select the numeric range and open Format Cells.

Choose Custom and use:

```text
#,##0;[Red]-#,##0
```

Positive values appear normally, while negative values appear in red.

### 3.4 Example: Use Data Bars

Requirement: Compare employee sales performance visually.

Steps:

1. Select the sales values.
2. Open Conditional Formatting.
3. Select Data Bars.
4. Choose a style.

Longer bars represent higher values.

### 3.5 Example: Apply Consistent Table Formatting

1. Select the complete dataset.
2. Press Ctrl + T.
3. Confirm the data range.
4. Enable headers if present.
5. Choose a suitable table style.
6. Rename the table using Table Design.

Example table name:

```text
EmployeeTable
```

### 3.6 Example: Improve a Dashboard's Formatting

For a business dashboard:

- Use a consistent font.
- Use clear section headings.
- Apply number formats to KPIs.
- Use limited colors.
- Align charts and KPI cards.
- Highlight important metrics.
- Avoid unnecessary borders and decoration.

The objective is to make important business information easy to understand.

---

## 4. Python Integration and Real-World Business Use Cases

Python can apply formatting to Excel workbooks programmatically.

### 4.1 Apply Formatting Using OpenPyXL

OpenPyXL supports fonts, fills, borders, alignment and number formats.

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

wb = Workbook()
ws = wb.active

ws.title = "Sales Report"

ws.append(["Product", "Sales"])
ws.append(["Laptop", 85000])
ws.append(["Mobile", 45000])
ws.append(["Monitor", 120000])

# Format header
for cell in ws[1]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )
    cell.alignment = Alignment(horizontal="center")

# Apply number formatting
for cell in ws["B"][1:]:
    cell.number_format = '#,##0.00'

# Set column widths
ws.column_dimensions["A"].width = 20
ws.column_dimensions["B"].width = 15

wb.save("formatted_sales.xlsx")
```

### 4.2 Apply Currency Formatting

```python
from openpyxl import load_workbook

wb = load_workbook("sales.xlsx")
ws = wb.active

for cell in ws["B"][1:]:
    cell.number_format = '₹#,##0.00'

wb.save("currency_formatted_sales.xlsx")
```

### 4.3 Apply Date Formatting

```python
from openpyxl import load_workbook

wb = load_workbook("employees.xlsx")
ws = wb.active

for cell in ws["D"][1:]:
    cell.number_format = "DD-MM-YYYY"

wb.save("formatted_employees.xlsx")
```

This changes the display format of actual date values. Text strings that resemble dates may need conversion first.

### 4.4 Apply Conditional Formatting

OpenPyXL can create conditional formatting rules.

```python
from openpyxl import load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

wb = load_workbook("sales.xlsx")
ws = wb.active

red_fill = PatternFill(
    start_color="FFC7CE",
    end_color="FFC7CE",
    fill_type="solid"
)

rule = CellIsRule(
    operator="greaterThan",
    formula=["50000"],
    fill=red_fill
)

ws.conditional_formatting.add("B2:B100", rule)

wb.save("conditional_sales.xlsx")
```

### 4.5 Apply Formatting Using Pandas

Pandas supports Excel output styling through its Styler interface.

```python
import pandas as pd

df = pd.read_excel("sales.xlsx")

with pd.ExcelWriter(
    "styled_sales.xlsx",
    engine="openpyxl"
) as writer:

    df.style.format({
        "Sales": "₹{:,.2f}"
    }).to_excel(
        writer,
        sheet_name="Sales",
        index=False
    )
```

Pandas is useful for generating formatted reports from DataFrames. OpenPyXL is more suitable when you need to modify existing workbook formatting directly.

### 4.6 Real-World Business Use Cases

**HR Reports**
- Format employee joining dates.
- Display salary as currency.
- Highlight missing information.
- Format employee tables.

**Sales Reports**
- Display revenue with currency symbols.
- Highlight high-performing products.
- Apply data bars to compare sales.
- Format monthly sales reports.

**Finance Reports**
- Apply accounting formats.
- Highlight negative balances.
- Display percentages.
- Format financial summaries.

**Inventory Reports**
- Highlight low stock.
- Format product quantities.
- Apply consistent table styles.
- Display product prices.

**Business Dashboards**
- Format KPI values.
- Highlight important metrics.
- Apply consistent chart colors.
- Improve report readability.

### 4.7 Data Formatting in an Excel AI Agent

An Excel AI Agent can understand formatting instructions and execute them using Python tools.

Example user instructions:

- "Format salary as Indian currency."
- "Make the header row bold."
- "Highlight sales above 100000."
- "Display dates in DD-MM-YYYY format."
- "Apply professional formatting to this report."
- "Adjust all column widths automatically."

Possible workflow:

1. Understand the user's instruction.
2. Identify the target worksheet and range.
3. Read the workbook structure.
4. Select the appropriate formatting operation.
5. Execute the operation using OpenPyXL.
6. Save the workbook.
7. Report the formatting changes.

The agent should distinguish between changing cell appearance and changing actual cell values.

---

## 5. Best Practices, Common Mistakes and Troubleshooting

### 5.1 Best Practices

- Maintain consistent formatting across worksheets.
- Use readable fonts and appropriate font sizes.
- Apply number formats based on data type.
- Use currency formatting for financial values.
- Use percentage formatting for ratios and rates.
- Use conditional formatting for important business rules.
- Avoid excessive colors and decorative styles.
- Keep column headings clear.
- Use tables for structured datasets.
- Preserve the original workbook before making large formatting changes.

### 5.2 Common Mistakes

**Mistake 1: Formatting Numbers as Text**

Converting numeric values into text can affect calculations and sorting.

Use numeric cells with appropriate number formats instead.

**Mistake 2: Confusing Display Values with Actual Values**

A cell displaying 25% may contain the underlying value 0.25.

Formatting changes how a value appears, not necessarily what it stores.

**Mistake 3: Using Too Many Colors**

Excessive colors make spreadsheets difficult to read.

Use a limited and consistent color scheme.

**Mistake 4: Using Merge Cells in Data Tables**

Merged cells can interfere with sorting, filtering and programmatic processing.

Avoid merging cells inside structured datasets.

**Mistake 5: Incorrect Date Formatting**

Changing a cell's number format does not convert an invalid text date into a real Excel date.

Convert text dates before applying date formatting.

**Mistake 6: Overwriting Existing Formatting**

Applying a new style across an entire worksheet can remove useful formatting.

Target only the required cells and preserve existing styles where possible.

### 5.3 Troubleshooting

| Problem | Possible Cause | Solution |
|---|---|---|
| Dates display as numbers | General number format | Apply a date format |
| Currency symbol is missing | Incorrect number format | Apply the required currency format |
| Percentage is too large | Stored value is not a decimal fraction | Correct the underlying value |
| Column content is hidden | Column width is insufficient | Use AutoFit |
| Text is cut off | Wrap Text is disabled or row height is insufficient | Enable Wrap Text and adjust row height |
| Conditional formatting does not work | Incorrect rule or range | Check the condition and target range |
| Numbers are not calculated | Values may be stored as text | Convert them into numeric values |
| Python formatting does not appear | Incorrect cell range or save operation | Verify the target cells and saved file |

### 5.4 Important Limitations

- Formatting does not automatically correct inaccurate data.
- Number formats affect display, not the underlying value.
- Conditional formatting depends on its rules and target ranges.
- OpenPyXL does not calculate Excel formulas.
- Some advanced Excel features may not be fully preserved when modifying complex workbooks programmatically.
- Always test the output workbook before distributing it.

### 5.5 Final Summary

Excel Data Formatting improves the readability, consistency and presentation of spreadsheets.

Key concepts:
- Number, date and currency formatting.
- Percentage and custom number formats.
- Fonts, fills, borders and alignment.
- Conditional formatting.
- Format Painter and Excel tables.
- Programmatic formatting using OpenPyXL.
- Report styling using Pandas.
- Formatting automation through an Excel AI Agent.

Good formatting makes business information easier to read and understand, while data cleaning and validation ensure that the underlying information is reliable.