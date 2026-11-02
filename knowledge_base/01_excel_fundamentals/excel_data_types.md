# Excel Data Types – Complete Guide

## 1. Overview and Fundamentals

Excel stores different types of information in cells. Understanding data types is important for calculations, data analysis, filtering, sorting, and automation.

The main Excel data types are:

1. Text
2. Numbers
3. Dates and Times
4. Logical Values
5. Errors
6. Blank Values

### 1.1 Text Data

Text contains characters, words, names, codes, and descriptions.

Examples:

| Employee ID | Employee Name | Department |
|---|---|---|
| EMP001 | Arun | IT |
| EMP002 | Priya | HR |
| EMP003 | Kumar | Sales |

Text values are generally used for identification and descriptions.

### 1.2 Numeric Data

Numbers represent quantities that can be used in mathematical calculations.

Examples:

- Salary: 25000
- Quantity: 150
- Price: 499.50
- Percentage: 85%

### 1.3 Date and Time Data

Excel stores dates and times as numeric serial values.

Examples:

- `31-10-2026`
- `10:30 AM`
- `31-10-2026 10:30 AM`

Different display formats can represent the same underlying date value.

### 1.4 Logical Data

Logical values represent TRUE or FALSE conditions.

Example:

```excel
=A1>100
```

If A1 contains 150, the result is TRUE.

### 1.5 Error Values

Excel displays error values when a formula cannot be evaluated correctly.

Common errors:

- `#DIV/0!`
- `#VALUE!`
- `#REF!`
- `#NAME?`
- `#N/A`
- `#NUM!`
- `#NULL!`

## 2. Core Concepts, Methods and Techniques

### 2.1 Text vs Numbers

A value such as `100` may be stored as a number or as text.

| Value | Data Type | Behaviour |
|---|---|---|
| 100 | Number | Supports arithmetic |
| "100" | Text | May not behave as a number |
| EMP100 | Text | Used as an identifier |
| 100.50 | Number | Supports decimal calculations |

A common problem occurs when numeric values are imported as text.

### 2.2 Identifying Data Types

Use the following Excel functions:

**ISNUMBER**

```excel
=ISNUMBER(A2)
```

Returns TRUE if A2 contains a numeric value.

**ISTEXT**

```excel
=ISTEXT(A2)
```

Returns TRUE if A2 contains text.

**ISBLANK**

```excel
=ISBLANK(A2)
```

Checks whether a cell is empty.

**ISLOGICAL**

```excel
=ISLOGICAL(A2)
```

Checks whether a cell contains TRUE or FALSE.

**ISERROR**

```excel
=ISERROR(A2)
```

Checks whether a formula result is an error.

### 2.3 Converting Data Types

**Convert Text to Number**

```excel
=VALUE(A2)
```

**Convert Number to Text**

```excel
=TEXT(A2,"0")
```

**Convert Date to Display Text**

```excel
=TEXT(A2,"dd-mm-yyyy")
```

**Remove Extra Spaces from Text**

```excel
=TRIM(A2)
```

### 2.4 Date and Time Calculations

Excel stores dates as serial numbers and times as fractions of a day.

Examples:

```excel
=TODAY()
```

Returns the current date.

```excel
=NOW()
```

Returns the current date and time.

```excel
=A2-B2
```

Calculates the difference between two dates when A2 and B2 contain valid Excel dates.

## 3. Step-by-Step Implementation with Excel Examples

### Example Dataset

| Employee | Age | Salary | Joining Date | Active |
|---|---:|---:|---|---|
| Arun | 25 | 25000 | 01-01-2024 | TRUE |
| Priya | 28 | 35000 | 15-03-2023 | TRUE |
| Kumar | 30 | 30000 | 10-06-2022 | FALSE |
| Divya | 26 | 28000 | 20-08-2024 | TRUE |

Assume the columns are A to E and the data begins in row 2.

### Step 1: Check Whether Salary Is Numeric

```excel
=ISNUMBER(C2)
```

Expected result: TRUE for a numeric salary.

### Step 2: Check Employee Name Data Type

```excel
=ISTEXT(A2)
```

Expected result: TRUE.

### Step 3: Check Whether a Cell Is Empty

```excel
=ISBLANK(F2)
```

Expected result: TRUE if F2 is empty.

### Step 4: Convert Text Salary into Number

If C2 contains the text value `"25000"`:

```excel
=VALUE(C2)
```

### Step 5: Calculate Employee Service Period in Days

```excel
=TODAY()-D2
```

This calculates the number of days since the employee's joining date.

### Step 6: Extract Date Components

Year:

```excel
=YEAR(D2)
```

Month:

```excel
=MONTH(D2)
```

Day:

```excel
=DAY(D2)
```

### Step 7: Check Multiple Data Types

```excel
=IF(ISNUMBER(C2),"Valid Salary","Check Salary")
```

This identifies whether the salary is stored as a number.

## 4. Python Integration and Real-World Business Use Cases

Python can identify and convert Excel data types during data processing.

### 4.1 Inspect Data Types Using Pandas

```python
import pandas as pd

df = pd.read_excel("employees.xlsx")

print(df.dtypes)
```

This displays the detected data type of each column.

### 4.2 Identify Numeric Columns

```python
numeric_columns = df.select_dtypes(
    include="number"
)

print(numeric_columns.head())
```

### 4.3 Identify Text Columns

```python
text_columns = df.select_dtypes(
    include="object"
)

print(text_columns.head())
```

### 4.4 Convert a Column to Numeric

```python
df["Salary"] = pd.to_numeric(
    df["Salary"],
    errors="coerce"
)
```

`errors="coerce"` converts invalid numeric values into missing values (`NaN`).

### 4.5 Convert a Column to Date

```python
df["Joining Date"] = pd.to_datetime(
    df["Joining Date"],
    errors="coerce"
)
```

### 4.6 Identify Missing Values

```python
missing_values = df.isnull().sum()

print(missing_values)
```

### 4.7 Inspect Data Types with OpenPyXL

```python
from openpyxl import load_workbook

workbook = load_workbook("employees.xlsx")

sheet = workbook.active

for row in sheet.iter_rows(min_row=2):
    for cell in row:
        print(
            cell.coordinate,
            cell.value,
            cell.data_type
        )
```

OpenPyXL exposes cell values and their stored data types.

### Real-World Business Use Cases

1. **Sales Management:** Ensure sales amounts are numeric before calculating revenue.
2. **HR Analytics:** Convert joining dates into proper date values.
3. **Financial Analysis:** Prevent calculations on numbers stored as text.
4. **Inventory Management:** Validate product codes and quantities.
5. **Customer Analytics:** Identify invalid phone numbers and customer identifiers.
6. **Automated Reporting:** Detect unexpected data types before generating reports.

## 5. Best Practices, Common Mistakes and Troubleshooting

### Best Practices

- Maintain consistent data types within each column.
- Store dates as actual Excel dates instead of inconsistent text.
- Use numeric formats for quantities and financial values.
- Keep identifiers such as employee IDs as text when leading zeros matter.
- Validate imported data before calculations.
- Use `ISNUMBER`, `ISTEXT`, and `ISBLANK` for Excel validation.
- Inspect Pandas data types before performing analysis.

### Common Mistakes

**Mistake 1: Numbers Stored as Text**

Problem:

```excel
=SUM(A1:A5)
```

Some text-formatted numeric values may be ignored by SUM.

Solution:

```excel
=VALUE(A1)
```

For an entire column in Python:

```python
df["Salary"] = pd.to_numeric(
    df["Salary"],
    errors="coerce"
)
```

**Mistake 2: Incorrect Date Format**

Problem: A date imported as text cannot be used reliably in date calculations.

Solution:

```excel
=DATEVALUE(A2)
```

This works when Excel recognizes the text date format.

**Mistake 3: Leading Zeros Disappear**

Problem: Product code `00125` becomes `125` when treated as a number.

Solution: Store the code as Text or use a suitable custom number format when appropriate.

**Mistake 4: Blank Cells and Zero Are Confused**

A blank cell and a cell containing `0` are not the same.

Use:

```excel
=ISBLANK(A2)
```

to check for a truly empty cell.

**Mistake 5: Invalid Values During Python Conversion**

Problem: A column contains values such as `"N/A"` or `"unknown"`.

Solution:

```python
df["Salary"] = pd.to_numeric(
    df["Salary"],
    errors="coerce"
)
```

Review the resulting missing values before replacing or removing them.

### Final Learning Objectives

After studying this document, the learner should understand:

- Excel text, numeric, date, logical, blank, and error values.
- Differences between numbers and numbers stored as text.
- Excel data-type checking functions.
- Date and time handling.
- Data-type conversion using Excel formulas.
- Data-type inspection and conversion using Pandas and OpenPyXL.
- Common data-type problems in business datasets.