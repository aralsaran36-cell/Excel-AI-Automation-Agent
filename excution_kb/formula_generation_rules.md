# Formula Generation Rules

## 1. Purpose

This document defines how the Excel AI Agent should generate valid Excel formulas based on natural-language user requests and actual workbook information.

The agent must understand the requested calculation, identify the required data, construct the formula, validate its references, and prepare it for execution through an available Excel tool.

## 2. Formula Generation Principles

Follow these rules:

1. Understand the calculation requested by the user.
2. Identify the appropriate Excel function.
3. Inspect the actual workbook structure.
4. Identify the required columns and data ranges.
5. Determine the output location.
6. Generate the formula using actual cell references.
7. Validate the formula syntax and references.
8. Pass the formula to the appropriate available tool.
9. Verify the execution result.

Never generate formulas using assumed column letters, row numbers, or worksheet names.

## 3. Formula Construction

A formula generally contains:

- Excel function name.
- Required arguments.
- Cell references or ranges.
- Conditions when applicable.
- Operators when required.

Example:

```excel
=SUM(B2:B100)
```

This example illustrates a SUM formula. The actual range must be identified from the workbook.

## 4. Basic Formula Generation

### 4.1 SUM

User requests:

- Calculate total sales.
- Find total units.
- Add all expenses.

Formula pattern:

```excel
=SUM(range)
```

Generation requirements:

- Identify the numerical column.
- Identify the actual data range.
- Generate the SUM formula.

### 4.2 AVERAGE

User requests:

- Calculate average salary.
- Find average sales.
- Calculate average quantity.

Formula pattern:

```excel
=AVERAGE(range)
```

### 4.3 MAX

User requests:

- Find highest sales.
- Identify maximum salary.
- Find the largest quantity.

Formula pattern:

```excel
=MAX(range)
```

### 4.4 MIN

User requests:

- Find the lowest sales.
- Identify minimum salary.
- Find the smallest quantity.

Formula pattern:

```excel
=MIN(range)
```

### 4.5 COUNT

User requests:

- Count numerical entries.
- Count cells containing numbers.

Formula pattern:

```excel
=COUNT(range)
```

### 4.6 COUNTA

User requests:

- Count non-empty cells.
- Count entries in a column.

Formula pattern:

```excel
=COUNTA(range)
```

## 5. Conditional Formula Generation

### 5.1 COUNTIF

User request:

"Count employees in the HR department."

Formula pattern:

```excel
=COUNTIF(criteria_range,criteria)
```

### 5.2 SUMIF

User request:

"Calculate total sales for the North region."

Formula pattern:

```excel
=SUMIF(criteria_range,criteria,sum_range)
```

### 5.3 COUNTIFS

User request:

"Count Laptop orders in the North region."

Formula pattern:

```excel
=COUNTIFS(criteria_range1,criteria1,criteria_range2,criteria2)
```

### 5.4 SUMIFS

User request:

"Calculate Laptop sales in the North region."

Formula pattern:

```excel
=SUMIFS(sum_range,criteria_range1,criteria1,criteria_range2,criteria2)
```

The agent must identify each condition and map it to the correct workbook column.

## 6. Logical Formula Generation

### 6.1 IF

User request:

"Mark sales above 50000 as High and others as Low."

Formula pattern:

```excel
=IF(condition,value_if_true,value_if_false)
```

### 6.2 AND

User request:

"Check whether sales are above 50000 and quantity is above 10."

Formula pattern:

```excel
=AND(condition1,condition2)
```

### 6.3 OR

User request:

"Identify records where sales are above 50000 or quantity is above 10."

Formula pattern:

```excel
=OR(condition1,condition2)
```

## 7. Lookup Formula Generation

### 7.1 XLOOKUP

User request:

"Find the price for product ID 101."

Formula pattern:

```excel
=XLOOKUP(lookup_value,lookup_array,return_array)
```

### 7.2 VLOOKUP

User request:

"Find employee details using employee ID."

Formula pattern:

```excel
=VLOOKUP(lookup_value,table_array,col_index_num,FALSE)
```

Select the lookup function based on the workbook structure and Excel compatibility.

## 8. Text Formula Generation

Examples:

```excel
=LEFT(A2,3)
```

```excel
=RIGHT(A2,4)
```

```excel
=MID(A2,2,5)
```

```excel
=TRIM(A2)
```

```excel
=CONCAT(A2,B2)
```

Select the function according to the requested text operation.

## 9. Date Formula Generation

Examples:

```excel
=TODAY()
```

```excel
=YEAR(A2)
```

```excel
=MONTH(A2)
```

```excel
=DAY(A2)
```

```excel
=DATEDIF(A2,B2,"Y")
```

Verify that the referenced cells contain valid date values.

## 10. Formula Validation Rules

Before execution, verify:

- Function name is valid.
- Number of arguments is correct.
- Argument order is correct.
- Cell references exist.
- Referenced ranges are valid.
- Criteria are correctly represented.
- Text values are enclosed in quotation marks where required.
- Formula starts with `=`.
- Division by zero is handled when relevant.
- Formula compatibility is considered.

## 11. Dynamic Workbook References

The agent must never assume:

- The required data is always in column A or B.
- The dataset always starts in row 2.
- The workbook always contains a fixed number of records.
- The active worksheet is always the correct worksheet.

Inspect the actual workbook before generating cell references.

## 12. Formula Execution Rules

1. Generate the formula.
2. Validate its syntax.
3. Identify the appropriate available formula execution tool.
4. Inspect its actual MCP schema.
5. Map the formula and required inputs to the actual parameters.
6. Execute the tool.
7. Check the returned result.
8. Verify the result against the requested calculation.

Do not invent tool parameters or assume successful execution.

## 13. Error Handling

If formula generation or execution fails:

- Check the formula syntax.
- Check the referenced columns.
- Check the data types.
- Check the cell ranges.
- Check worksheet references.
- Check the selected tool's actual parameter schema.
- Correct the identified problem before retrying.

## 14. Important Rules

- Use actual workbook information.
- Never assume cell references.
- Never generate unsupported formulas.
- Select formulas according to the user's intent.
- Validate formulas before execution.
- Use available MCP tools for actual workbook modifications.
- Verify the final result.