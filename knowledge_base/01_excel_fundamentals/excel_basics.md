# Excel Basics – Complete Guide

## 1. Overview and Fundamentals

Microsoft Excel is a spreadsheet application used to store, organize, calculate, analyze, and visualize structured data.

Excel organizes information using:
- **Workbook:** An Excel file (`.xlsx`).
- **Worksheet:** An individual sheet inside a workbook.
- **Row:** A horizontal collection of cells, identified by numbers.
- **Column:** A vertical collection of cells, identified by letters.
- **Cell:** The intersection of a row and column, such as `A1`.
- **Range:** A collection of cells, such as `A1:C10`.

### Common Excel File Formats

| Format | Purpose |
|---|---|
| `.xlsx` | Standard Excel workbook |
| `.xls` | Legacy Excel format |
| `.csv` | Comma-separated values |
| `.xlsm` | Excel workbook with macros |
| `.xlsb` | Binary Excel workbook |

## 2. Core Concepts, Methods and Techniques

### Cell References

Excel supports three major reference types.

**Relative Reference**
```excel
=A1+B1
```
When copied to another row, the references change automatically.

**Absolute Reference**
```excel
=A1*$B$1
```
The `$` symbol locks the row and column reference.

**Mixed Reference**
```excel
=$A1
=A$1
```
- `$A1`: Locks the column.
- `A$1`: Locks the row.

### Basic Excel Operations

| Operation | Formula |
|---|---|
| Addition | `=A1+B1` |
| Subtraction | `=A1-B1` |
| Multiplication | `=A1*B1` |
| Division | `=A1/B1` |
| Sum | `=SUM(A1:A10)` |
| Average | `=AVERAGE(A1:A10)` |
| Minimum | `=MIN(A1:A10)` |
| Maximum | `=MAX(A1:A10)` |
| Count numbers | `=COUNT(A1:A10)` |
| Count non-empty cells | `=COUNTA(A1:A10)` |

### Important Excel Features

- AutoFill: Extends formulas and data patterns.
- Flash Fill: Detects and applies data patterns.
- Freeze Panes: Keeps selected rows or columns visible.
- Find and Replace: Searches and replaces cell values.
- Conditional Formatting: Highlights cells based on conditions.
- Data Validation: Restricts the values entered into cells.
- Format as Table: Converts a range into a structured Excel table.

## 3. Step-by-Step Implementation with Excel Examples

### Example: Employee Salary Analysis

Create the following dataset:

| Employee | Department | Salary |
|---|---|---:|
| Arun | IT | 25000 |
| Kumar | HR | 30000 |
| Priya | IT | 35000 |
| Divya | Sales | 28000 |
| Ravi | HR | 32000 |

### Step 1: Calculate Total Salary

```excel
=SUM(C2:C6)
```

### Step 2: Calculate Average Salary

```excel
=AVERAGE(C2:C6)
```

### Step 3: Find Highest Salary

```excel
=MAX(C2:C6)
```

### Step 4: Find Lowest Salary

```excel
=MIN(C2:C6)
```

### Step 5: Count Employees

```excel
=COUNTA(A2:A6)
```

### Step 6: Calculate Department-Wise Salary

```excel
=SUMIF(B2:B6,"IT",C2:C6)
```

This calculates the total salary of employees belonging to the IT department.

### Step 7: Create an Excel Table

1. Select the dataset.
2. Press `Ctrl + T`.
3. Enable My table has headers.
4. Click OK.
5. Use table filters to analyze employee records.

## 4. Python Integration and Real-World Business Use Cases

Python can automate Excel operations using libraries such as `openpyxl` and `pandas`.

### Read an Excel File Using Pandas

```python
import pandas as pd

df = pd.read_excel("employees.xlsx")

print(df.head())
print(df.info())
print(df.describe())
```

### Calculate Total Salary

```python
total_salary = df["Salary"].sum()

print("Total Salary:", total_salary)
```

### Calculate Department-Wise Salary

```python
department_salary = (
    df.groupby("Department")["Salary"]
    .sum()
)

print(department_salary)
```

### Export Results to Excel

```python
department_salary.to_excel(
    "department_salary.xlsx"
)
```

### Read Excel Using OpenPyXL

```python
from openpyxl import load_workbook

workbook = load_workbook("employees.xlsx")

sheet = workbook.active

for row in sheet.iter_rows(
    min_row=1,
    values_only=True
):
    print(row)
```

### Real-World Business Applications

1. Employee salary management.
2. Sales tracking and reporting.
3. Inventory management.
4. Financial calculations.
5. Customer information management.
6. Monthly business performance analysis.
7. Automated Excel report generation.

## 5. Best Practices, Common Mistakes and Troubleshooting

### Best Practices

- Use meaningful column names.
- Keep one type of information in each column.
- Avoid unnecessary blank rows and columns.
- Convert structured datasets into Excel Tables.
- Use appropriate cell formats.
- Maintain consistent date and number formats.
- Save important workbooks regularly.
- Keep raw data separate from analysis results.

### Common Excel Errors

| Error | Meaning | Possible Solution |
|---|---|---|
| `#DIV/0!` | Division by zero | Check the denominator |
| `#VALUE!` | Invalid data type | Check formula inputs |
| `#REF!` | Invalid cell reference | Correct the reference |
| `#NAME?` | Unrecognized function or name | Check spelling |
| `#N/A` | Value not available | Check lookup values |
| `#NUM!` | Invalid numeric calculation | Check numeric inputs |

### Troubleshooting Examples

**Problem 1: Numbers are stored as text**

Solution:
- Select the affected cells.
- Use Convert to Number if available.
- Alternatively use `=VALUE(A1)`.

**Problem 2: Formula is not calculating**

Solution:
- Check whether the cell is formatted as Text.
- Change the format to General.
- Re-enter the formula.
- Check whether calculation mode is Automatic.

**Problem 3: Excel displays `#####`**

Solution:
- Increase the column width.
- Check whether the cell contains a date or negative date value.

**Problem 4: Formula references change unexpectedly**

Solution:
- Use absolute references such as `$A$1` when a reference must remain fixed.

### Final Learning Objectives

After studying this document, the learner should understand:

- Excel workbook and worksheet structure.
- Rows, columns, cells and ranges.
- Cell references.
- Basic mathematical functions.
- Excel tables and formatting.
- Reading and writing Excel files using Python.
- Common Excel errors and troubleshooting techniques.