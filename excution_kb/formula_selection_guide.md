# Formula Selection Guide

## 1. Purpose

This document helps the Excel AI Agent identify the appropriate Excel formula based on the user's natural-language request.

The agent must understand the requested operation and select a suitable formula based on the calculation objective.

## 2. Basic Calculation Functions

| User Intent | Formula |
|---|---|
| Calculate total | SUM |
| Find average | AVERAGE |
| Find highest value | MAX |
| Find lowest value | MIN |
| Count numerical values | COUNT |
| Count non-empty cells | COUNTA |
| Find middle value | MEDIAN |
| Find most frequent value | MODE.SNGL |
| Calculate standard deviation | STDEV.S |
| Calculate variance | VAR.S |

## 3. Conditional Calculation Functions

| User Intent | Formula |
|---|---|
| Count records matching one condition | COUNTIF |
| Count records matching multiple conditions | COUNTIFS |
| Sum values matching one condition | SUMIF |
| Sum values matching multiple conditions | SUMIFS |
| Average values matching one condition | AVERAGEIF |
| Average values matching multiple conditions | AVERAGEIFS |

Examples:

- Count employees in Sales department → COUNTIF.
- Count employees in Sales department with salary above 50000 → COUNTIFS.
- Calculate North region sales → SUMIF.
- Calculate North region Laptop sales → SUMIFS.

## 4. Logical Functions

| User Intent | Formula |
|---|---|
| Check one condition | IF |
| Check multiple conditions using AND | AND |
| Check multiple conditions using OR | OR |
| Handle formula errors | IFERROR |
| Check whether a cell is blank | ISBLANK |
| Check whether a value is numeric | ISNUMBER |

Examples:

- Mark sales above 50000 as High → IF.
- Check whether two conditions are true → AND.
- Check whether at least one condition is true → OR.
- Replace formula errors with a message → IFERROR.

## 5. Lookup Functions

| User Intent | Formula |
|---|---|
| Find a matching value | XLOOKUP |
| Search vertically in a table | VLOOKUP |
| Search horizontally in a table | HLOOKUP |
| Find a matching position | MATCH |
| Return a value by row and column position | INDEX |
| Search using INDEX and MATCH | INDEX + MATCH |

Examples:

- Find employee salary using employee ID → XLOOKUP.
- Find product price using product code → XLOOKUP.
- Find a matching row position → MATCH.

Choose the function according to the workbook layout and supported Excel version.

## 6. Text Functions

| User Intent | Formula |
|---|---|
| Extract characters from the beginning | LEFT |
| Extract characters from the end | RIGHT |
| Extract characters from the middle | MID |
| Remove extra spaces | TRIM |
| Convert text to uppercase | UPPER |
| Convert text to lowercase | LOWER |
| Capitalize words | PROPER |
| Combine text | CONCAT |
| Find text position | FIND |
| Replace text | SUBSTITUTE |
| Count characters | LEN |

## 7. Date and Time Functions

| User Intent | Formula |
|---|---|
| Get current date | TODAY |
| Get current date and time | NOW |
| Extract year | YEAR |
| Extract month | MONTH |
| Extract day | DAY |
| Calculate days between dates | DAYS |
| Calculate working days | NETWORKDAYS |
| Add months to a date | EDATE |
| Find month-end date | EOMONTH |

## 8. Ranking and Sorting Functions

| User Intent | Formula |
|---|---|
| Find rank of a value | RANK.EQ |
| Return largest value at a position | LARGE |
| Return smallest value at a position | SMALL |
| Sort values dynamically | SORT |
| Sort using another column | SORTBY |
| Return unique values | UNIQUE |

These functions may require newer Excel versions.

## 9. Percentage and Growth Calculations

| User Intent | Formula |
|---|---|
| Calculate percentage | Part / Total |
| Calculate percentage change | (Current - Previous) / Previous |
| Calculate percentage contribution | Category Total / Overall Total |
| Calculate profit | Revenue - Cost |
| Calculate profit margin | Profit / Revenue |
| Calculate discount | Original Price × Discount Rate |

Example:

User request:

"Calculate monthly sales growth."

Formula pattern:

```excel
=(Current_Month-Previous_Month)/Previous_Month
```

Handle cases where the previous period is zero.

## 10. Formula Selection Decision Rules

Follow these steps:

1. Identify the calculation objective.
2. Determine whether the request requires aggregation, conditional calculation, logical evaluation, lookup, text processing, date processing, ranking, or percentage calculation.
3. Select the corresponding formula.
4. Identify all required arguments.
5. Inspect the actual workbook structure.
6. Generate the formula using actual cell references.
7. Validate the formula.
8. Execute through an appropriate available tool.

## 11. Multiple Formula Requirements

Some requests require multiple calculations.

Example:

"Calculate total sales, average sales and highest sales."

Required formulas:

- SUM
- AVERAGE
- MAX

The agent should identify all requested calculations and execute the necessary operations.

Avoid calculating unrelated metrics.

## 12. Formula Compatibility

Before using a formula, consider:

- Excel version.
- Supported functions.
- Workbook format.
- Data type compatibility.
- Formula argument requirements.

If a function is unsupported, select a compatible alternative only when it produces the intended result.

## 13. Important Rules

- Choose formulas based on user intent.
- Do not select formulas using keyword matching alone.
- Understand the complete request.
- Identify all conditions.
- Verify the required data columns.
- Consider formula compatibility.
- Avoid unnecessary calculations.
- Validate the generated formula before execution.