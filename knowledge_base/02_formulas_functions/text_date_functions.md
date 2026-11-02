

# Text and Date Functions in Excel

## 1. Overview and Fundamentals

Text and Date functions in Excel are used to manipulate text values, extract specific information, format strings, calculate dates, and perform time-based operations.

These functions are essential for cleaning datasets, preparing reports, managing employee records, processing customer information, and analyzing business timelines.

### Text Functions

Text functions help users:
- Extract characters from strings.
- Combine multiple text values.
- Remove unnecessary spaces.
- Change uppercase and lowercase.
- Replace specific characters.
- Find text within a string.
- Convert numbers into formatted text.

### Date and Time Functions

Date and time functions help users:
- Calculate the difference between two dates.
- Identify current dates and times.
- Extract day, month, and year.
- Calculate employee experience.
- Find deadlines and overdue periods.
- Calculate working days.
- Generate monthly and yearly reports.

### Common Text Functions

| Function | Purpose |
|---|---|
| LEFT | Extracts characters from the beginning |
| RIGHT | Extracts characters from the end |
| MID | Extracts characters from a specified position |
| LEN | Counts characters |
| TRIM | Removes unnecessary spaces |
| UPPER | Converts text to uppercase |
| LOWER | Converts text to lowercase |
| PROPER | Capitalizes the first letter of each word |
| CONCAT | Combines text values |
| TEXTJOIN | Combines text with a delimiter |
| FIND | Finds text and is case-sensitive |
| SEARCH | Finds text and is not case-sensitive |
| SUBSTITUTE | Replaces matching text |
| REPLACE | Replaces characters by position |
| TEXT | Formats a number or date as text |
| VALUE | Converts numeric text into a number |

### Common Date Functions

| Function | Purpose |
|---|---|
| TODAY | Returns the current date |
| NOW | Returns the current date and time |
| DAY | Extracts the day |
| MONTH | Extracts the month |
| YEAR | Extracts the year |
| DATE | Creates a date |
| DATEDIF | Calculates the difference between dates |
| EDATE | Adds or subtracts months |
| EOMONTH | Returns the last day of a month |
| WEEKDAY | Returns the weekday number |
| WORKDAY | Calculates a working date |
| NETWORKDAYS | Counts working days |

## 2. Core Concepts, Methods and Techniques

### 2.1 LEFT Function

LEFT extracts a specified number of characters from the beginning of a text string.

Syntax:

`=LEFT(text,[num_chars])`

Example:

`=LEFT("EXCEL AGENT",5)`

Result: `EXCEL`

Use cases:
- Extract country codes.
- Retrieve product prefixes.
- Extract employee ID prefixes.

### 2.2 RIGHT Function

RIGHT extracts characters from the end of a text string.

Syntax:

`=RIGHT(text,[num_chars])`

Example:

`=RIGHT("EMPLOYEE101",3)`

Result: `101`

Use cases:
- Extract the last digits of an ID.
- Retrieve file extensions.
- Extract product suffixes.

### 2.3 MID Function

MID extracts characters from a specified position.

Syntax:

`=MID(text,start_num,num_chars)`

Example:

`=MID("EXCEL AGENT",7,5)`

Result: `AGENT`

Use cases:
- Extract information from structured IDs.
- Retrieve codes from product names.
- Extract specific sections from reference numbers.

### 2.4 LEN Function

LEN counts the number of characters in a text string, including spaces.

Syntax:

`=LEN(text)`

Example:

`=LEN("Excel Agent")`

Result: `11`

Use cases:
- Validate phone number lengths.
- Check ID formats.
- Identify unexpected text lengths.

### 2.5 TRIM Function

TRIM removes leading and trailing spaces and reduces repeated internal spaces to single spaces.

Syntax:

`=TRIM(text)`

Example:

`=TRIM("  Arun   Kumar  ")`

Result: `Arun Kumar`

Use cases:
- Clean imported datasets.
- Normalize customer names.
- Remove unwanted spaces before lookup operations.

Note: TRIM does not remove every type of invisible whitespace, such as some non-breaking spaces.

### 2.6 UPPER, LOWER and PROPER

These functions change the capitalization of text.

Examples:

`=UPPER("excel agent")`

Result: `EXCEL AGENT`

`=LOWER("EXCEL AGENT")`

Result: `excel agent`

`=PROPER("saravanan kumar")`

Result: `Saravanan Kumar`

Use cases:
- Standardize customer names.
- Format employee records.
- Normalize text before comparison.

### 2.7 CONCAT Function

CONCAT combines text from multiple cells or strings.

Syntax:

`=CONCAT(text1,text2,...)`

Example:

`=CONCAT(A2," ",B2)`

If A2 contains Saravanan and B2 contains Kumar, the result is:

`Saravanan Kumar`

### 2.8 TEXTJOIN Function

TEXTJOIN combines multiple text values using a specified delimiter.

Syntax:

`=TEXTJOIN(delimiter,ignore_empty,text1,...)`

Example:

`=TEXTJOIN(", ",TRUE,A2:A5)`

This combines the values in A2:A5 with commas while ignoring empty cells.

Use cases:
- Combine customer addresses.
- Generate comma-separated product lists.
- Combine employee names.

### 2.9 FIND and SEARCH Functions

Both functions locate text inside another text string.

FIND is case-sensitive.

`=FIND("A","Excel Agent")`

SEARCH is not case-sensitive.

`=SEARCH("agent","Excel Agent")`

Both return the starting position of the matching text.

Use cases:
- Find keywords in descriptions.
- Identify specific codes.
- Check whether a cell contains a particular word.

### 2.10 SUBSTITUTE Function

SUBSTITUTE replaces matching text with another value.

Syntax:

`=SUBSTITUTE(text,old_text,new_text,[instance_num])`

Example:

`=SUBSTITUTE("Product-A-2026","-"," ")`

Result: `Product A 2026`

To replace only the second occurrence:

`=SUBSTITUTE("A-B-C","-","/",2)`

Result: `A-B/C`

### 2.11 REPLACE Function

REPLACE replaces characters based on their position.

Syntax:

`=REPLACE(old_text,start_num,num_chars,new_text)`

Example:

`=REPLACE("EMP12345",4,5,"99999")`

Result: `EMP99999`

Difference:

- SUBSTITUTE replaces matching text.
- REPLACE replaces text based on character position.

### 2.12 TEXT Function

TEXT converts numbers or dates into formatted text.

Syntax:

`=TEXT(value,format_text)`

Examples:

`=TEXT(A2,"dd-mm-yyyy")`

`=TEXT(B2,"₹#,##0.00")`

`=TEXT(C2,"0.00%")`

Use cases:
- Format dates for reports.
- Display currency values.
- Generate readable report labels.

Important: TEXT returns text, not a numeric value.

### 2.13 VALUE Function

VALUE converts a number stored as text into a numeric value.

Syntax:

`=VALUE(text)`

Example:

`=VALUE("1500")`

Result: `1500`

Use cases:
- Convert imported numeric text.
- Prepare data for calculations.
- Fix number-format inconsistencies.

### 2.14 TODAY and NOW

TODAY returns the current date.

`=TODAY()`

NOW returns the current date and time.

`=NOW()`

These functions update when Excel recalculates the workbook.

### 2.15 DAY, MONTH and YEAR

These functions extract date components.

Examples:

`=DAY(A2)`

`=MONTH(A2)`

`=YEAR(A2)`

If A2 contains 15-08-2026:

- DAY returns 15.
- MONTH returns 8.
- YEAR returns 2026.

### 2.16 DATE Function

DATE creates a valid Excel date using year, month, and day values.

Syntax:

`=DATE(year,month,day)`

Example:

`=DATE(2026,8,15)`

Result: 15-Aug-2026 when displayed with that date format.

### 2.17 DATEDIF Function

DATEDIF calculates the difference between two dates.

Syntax:

`=DATEDIF(start_date,end_date,unit)`

Common units:

- `"Y"` – Complete years.
- `"M"` – Complete months.
- `"D"` – Total days.
- `"YM"` – Remaining months after complete years.
- `"MD"` – Remaining days after complete months.
- `"YD"` – Remaining days after complete years.

Example:

`=DATEDIF(A2,B2,"Y")`

This calculates the completed years between two dates.

Note: DATEDIF is a legacy-compatible function and may not appear in Excel's formula autocomplete.

### 2.18 EDATE and EOMONTH

EDATE adds or subtracts a specified number of months.

`=EDATE(A2,3)`

This returns the date three months after A2.

EOMONTH returns the last day of a specified month offset.

`=EOMONTH(A2,0)`

This returns the last day of the same month as A2.

`=EOMONTH(A2,1)`

This returns the last day of the following month.

### 2.19 WEEKDAY Function

WEEKDAY returns a number representing the day of the week.

Syntax:

`=WEEKDAY(serial_number,[return_type])`

Example:

`=WEEKDAY(A2,2)`

With return type 2:
- Monday = 1
- Tuesday = 2
- ...
- Sunday = 7

### 2.20 WORKDAY and NETWORKDAYS

WORKDAY calculates a future or previous working date, excluding weekends and optionally specified holidays.

Syntax:

`=WORKDAY(start_date,days,[holidays])`

Example:

`=WORKDAY(A2,10,H2:H10)`

This calculates the date 10 working days after A2, excluding weekends and holidays listed in H2:H10.

NETWORKDAYS counts working days between two dates.

`=NETWORKDAYS(A2,B2,H2:H10)`

Use cases:
- Calculate employee working days.
- Track project timelines.
- Calculate business delivery deadlines.

## 3. Step-by-Step Implementation with Excel Examples

### Example 1: Clean Customer Names

Dataset:

| Customer Name |
|---|
| `  arun kumar  ` |
| `PRIYA SHARMA` |
| `kumar   raj` |

Goal: Standardize names.

Steps:

1. Insert a new column called Clean Name.
2. Enter the following formula in B2:

`=PROPER(TRIM(A2))`

3. Drag the formula down.

Results:

| Original Name | Clean Name |
|---|---|
| `  arun kumar  ` | Arun Kumar |
| `PRIYA SHARMA` | Priya Sharma |
| `kumar   raj` | Kumar Raj |

### Example 2: Extract Product Codes

Dataset:

| Product ID |
|---|
| LAP-2026-101 |
| MOB-2026-102 |
| TAB-2026-103 |

Goal: Extract product category and final ID.

Category:

`=LEFT(A2,3)`

Final ID:

`=RIGHT(A2,3)`

Extract the year:

`=MID(A2,5,4)`

### Example 3: Combine Employee First and Last Names

Dataset:

| First Name | Last Name |
|---|---|
| Arun | Kumar |
| Priya | Sharma |
| Kumar | Raj |

Formula:

`=TEXTJOIN(" ",TRUE,A2,B2)`

This generates the full employee name.

### Example 4: Calculate Employee Experience

Dataset:

| Employee | Joining Date |
|---|---|
| Arun | 15-06-2020 |
| Priya | 10-01-2022 |
| Kumar | 20-08-2023 |

Goal: Calculate completed years of experience as of today.

Formula:

`=DATEDIF(B2,TODAY(),"Y")`

To display years and remaining months:

`=DATEDIF(B2,TODAY(),"Y")&" Years "&DATEDIF(B2,TODAY(),"YM")&" Months"`

### Example 5: Calculate Project Deadlines

Suppose a project starts on 1 October 2026 and requires 15 working days.

Formula:

`=WORKDAY(A2,15)`

To exclude company holidays stored in D2:D5:

`=WORKDAY(A2,15,D2:D5)`

### Example 6: Generate Monthly Reports

Suppose a transaction date is stored in A2.

Extract the month:

`=TEXT(A2,"mmmm")`

Extract the year:

`=YEAR(A2)`

Generate a month-year label:

`=TEXT(A2,"mmm-yyyy")`

This helps group transactions by month and year.

## 4. Python Integration and Real-World Business Use Cases

Python Pandas provides text and date manipulation methods for automating Excel data processing.

### 4.1 Text Cleaning with Pandas

```python
import pandas as pd

df = pd.DataFrame({
    "Name": [
        "  arun kumar  ",
        "PRIYA SHARMA",
        "kumar   raj"
    ]
})

df["Clean_Name"] = (
    df["Name"]
    .str.strip()
    .str.replace(r"\s+", " ", regex=True)
    .str.title()
)

print(df)
```

Equivalent Excel operations:
- TRIM
- PROPER

### 4.2 Extract Text Using Python

```python
import pandas as pd

df = pd.DataFrame({
    "Product_ID": [
        "LAP-2026-101",
        "MOB-2026-102",
        "TAB-2026-103"
    ]
})

df["Category"] = df["Product_ID"].str[:3]
df["Year"] = df["Product_ID"].str[4:8]
df["ID"] = df["Product_ID"].str[-3:]

print(df)
```

Equivalent Excel operations:
- LEFT
- MID
- RIGHT

### 4.3 Combine Text Columns

```python
import pandas as pd

df = pd.DataFrame({
    "First_Name": ["Arun", "Priya"],
    "Last_Name": ["Kumar", "Sharma"]
})

df["Full_Name"] = (
    df["First_Name"] + " " + df["Last_Name"]
)

print(df)
```

### 4.4 Date Conversion and Extraction

```python
import pandas as pd

df = pd.DataFrame({
    "Joining_Date": [
        "2020-06-15",
        "2022-01-10",
        "2023-08-20"
    ]
})

df["Joining_Date"] = pd.to_datetime(
    df["Joining_Date"]
)

df["Year"] = df["Joining_Date"].dt.year
df["Month"] = df["Joining_Date"].dt.month
df["Day"] = df["Joining_Date"].dt.day

print(df)
```

### 4.5 Calculate Employee Experience

```python
import pandas as pd

df = pd.DataFrame({
    "Employee": ["Arun", "Priya"],
    "Joining_Date": ["2020-06-15", "2022-01-10"]
})

df["Joining_Date"] = pd.to_datetime(
    df["Joining_Date"]
)

today = pd.Timestamp.today().normalize()

df["Experience_Days"] = (
    today - df["Joining_Date"]
).dt.days

df["Experience_Years"] = (
    df["Experience_Days"] / 365.25
).round(2)

print(df)
```

This calculates an approximate experience duration in years. For exact completed years and months, use calendar-based date calculations.

### 4.6 Calculate Working Days

```python
import pandas as pd
import numpy as np

start_date = "2026-10-01"
end_date = "2026-10-31"

holidays = [
    "2026-10-02"
]

working_days = np.busday_count(
    start_date,
    end_date,
    holidays=holidays
)

print("Working Days:", working_days)
```

`np.busday_count` counts business days from the start date up to, but not including, the end date. Its default weekend is Saturday and Sunday.

### 4.7 Export Cleaned Data to Excel

```python
import pandas as pd

df = pd.read_excel("employees.xlsx")

df["Name"] = (
    df["Name"]
    .str.strip()
    .str.replace(r"\s+", " ", regex=True)
    .str.title()
)

df["Joining_Date"] = pd.to_datetime(
    df["Joining_Date"],
    errors="coerce"
)

df["Joining_Year"] = (
    df["Joining_Date"].dt.year
)

df.to_excel(
    "cleaned_employees.xlsx",
    index=False
)
```

### 4.8 Real-World Business Use Cases

**HR Management**
- Standardize employee names.
- Extract joining month and year.
- Calculate employee experience.
- Calculate working days and leave periods.

**Sales and Customer Management**
- Clean customer names and addresses.
- Extract product codes.
- Generate customer reference IDs.
- Group sales by month and year.

**Finance**
- Format transaction dates.
- Calculate invoice due dates.
- Identify overdue payments.
- Generate monthly financial reports.

**Project Management**
- Calculate working-day deadlines.
- Identify project completion dates.
- Calculate project duration.
- Track delayed tasks.

## 5. Best Practices, Common Mistakes and Troubleshooting

### Best Practices

1. Clean text before performing lookup operations.
2. Use TRIM to remove ordinary extra spaces.
3. Use consistent capitalization for names and categories.
4. Convert date values into valid Excel dates before calculations.
5. Use TODAY for dynamic current-date calculations.
6. Use WORKDAY and NETWORKDAYS for business-day calculations.
7. Use TEXT only when a formatted text representation is required.
8. Preserve original columns when cleaning important business data.
9. Validate date formats when importing data from external sources.

### Common Mistakes

**Problem 1: #VALUE! in MID or REPLACE**

Possible reason: Invalid start position or number of characters.

Solution: Check the start position and length arguments.

**Problem 2: TRIM Does Not Remove All Spaces**

Possible reason: The data contains non-breaking spaces.

Solution:

`=TRIM(SUBSTITUTE(A2,CHAR(160)," "))`

**Problem 3: Date Values Are Not Recognized**

Possible reason: Dates are imported as text or use an unexpected regional format.

Solution:
- Check the original date format.
- Convert text into valid dates.
- Use DATE when the year, month, and day are available separately.

**Problem 4: DATEDIF Returns an Error**

Possible reason: The start date is later than the end date.

Solution: Verify both dates and ensure the end date is not earlier than the start date.

**Problem 5: TEXT Output Cannot Be Used for Arithmetic**

Possible reason: TEXT converts numbers and dates into text.

Solution: Keep the original numeric or date value in a separate cell for calculations.

**Problem 6: Incorrect Working-Day Calculation**

Possible reason:
- Holidays were not included.
- Weekend configuration differs from the business calendar.
- Start and end date assumptions are incorrect.

Solution: Verify the holiday range and use WORKDAY.INTL or NETWORKDAYS.INTL when the weekend pattern differs from Saturday and Sunday.

### Troubleshooting Checklist

- Check for leading, trailing, and non-breaking spaces.
- Verify text extraction positions.
- Confirm whether a formula returns text or a number.
- Check whether date cells contain valid Excel date values.
- Verify date order and regional date formats.
- Confirm the holiday list.
- Check the weekend configuration.
- Test formulas using blank cells and invalid input values.

### Final Summary

Text and Date functions are essential for cleaning, transforming, formatting, and analyzing Excel datasets.

**Text functions:**
- LEFT, RIGHT, MID: Extract text.
- LEN: Count characters.
- TRIM: Remove unnecessary spaces.
- UPPER, LOWER, PROPER: Change capitalization.
- CONCAT, TEXTJOIN: Combine text.
- FIND, SEARCH: Locate text.
- SUBSTITUTE, REPLACE: Replace text.
- TEXT, VALUE: Convert and format values.

**Date functions:**
- TODAY, NOW: Current date and time.
- DAY, MONTH, YEAR: Extract date components.
- DATE: Create dates.
- DATEDIF: Calculate date differences.
- EDATE, EOMONTH: Perform month-based calculations.
- WEEKDAY: Identify weekdays.
- WORKDAY, NETWORKDAYS: Calculate working dates and days.

In Python, Pandas provides powerful string and datetime operations for automating Excel data cleaning and reporting.

