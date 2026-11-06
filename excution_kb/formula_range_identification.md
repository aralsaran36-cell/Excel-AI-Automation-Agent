# Formula Range Identification

## 1. Purpose

This document defines how the Excel AI Agent identifies the correct worksheet, columns, cell references, data ranges, and output locations before generating or executing formulas.

The agent must use actual workbook information rather than assuming fixed cell locations.

## 2. Workbook Inspection

Before identifying a formula range, inspect:

- Workbook information.
- Worksheet names.
- Active worksheet.
- Column headers.
- Data types.
- Used range.
- First data row.
- Last data row.
- Empty rows and columns.
- Relevant data locations.

Use the available workbook inspection tools based on their actual MCP descriptions and schemas.

## 3. Column Identification

The agent must identify the correct column using the user's requested field.

Examples:

| User Request | Required Column |
|---|---|
| Total sales | Sales |
| Average salary | Salary |
| Total units | Units |
| Count employees | Employee ID or another appropriate field |
| Highest price | Price |
| Average quantity | Quantity |

Column identification should consider:

- Exact header names.
- Similar header names.
- Common business terminology.
- Column data types.
- Actual values in the column.

If multiple columns could match the request, ask the user for clarification.

## 4. Cell Range Identification

A cell range identifies the cells used in a calculation.

Examples:

```excel
A2:A100
```

```excel
B2:B500
```

```excel
A1:F100
```

These are illustrative examples only.

The agent must determine the actual first and last data rows from the workbook.

## 5. Header and Data Row Detection

When identifying a data range:

1. Locate the header row.
2. Identify the first data row.
3. Identify the last relevant data row.
4. Exclude headers from numerical calculations.
5. Check for blank rows within the dataset.
6. Determine whether the data contains multiple separate sections.
7. Verify that the range includes all relevant records.

Do not automatically assume that the header is in row 1 or data begins in row 2.

## 6. Range Identification for Basic Formulas

### SUM

User request:

"Calculate total sales."

Required information:

- Sales column.
- First data row.
- Last data row.

Formula pattern:

```excel
=SUM(actual_sales_range)
```

### AVERAGE

User request:

"Calculate average salary."

Required information:

- Salary column.
- Actual data range.

Formula pattern:

```excel
=AVERAGE(actual_salary_range)
```

### MAX and MIN

Identify the relevant numerical column and its actual data range.

Formula patterns:

```excel
=MAX(actual_range)
```

```excel
=MIN(actual_range)
```

## 7. Range Identification for Conditional Formulas

### SUMIF

User request:

"Calculate North region sales."

Identify:

- Region criteria range.
- Sales sum range.
- Criteria value: North.

Formula pattern:

```excel
=SUMIF(region_range,"North",sales_range)
```

### COUNTIF

User request:

"Count employees in HR."

Identify:

- Department criteria range.
- Criteria value: HR.

Formula pattern:

```excel
=COUNTIF(department_range,"HR")
```

### SUMIFS

User request:

"Calculate Laptop sales in North region."

Identify:

- Sales sum range.
- Region criteria range.
- Product criteria range.
- Required criteria values.

Formula pattern:

```excel
=SUMIFS(sales_range,region_range,"North",product_range,"Laptop")
```

Ensure that the relevant criteria ranges and sum range cover compatible records.

## 8. Multi-Column Range Identification

Some formulas require multiple columns.

Example:

User request:

"Calculate total sales for each product in the North region."

Required columns:

- Product.
- Region.
- Sales.

The agent must map each requested condition to the correct column and use the appropriate ranges.

Never use the same range for different fields unless the workbook structure confirms that they are the same.

## 9. Output Cell Identification

Before writing a formula, determine where the result should be placed.

Possible output locations:

- User-specified cell.
- Empty cell in the current worksheet.
- A new summary section.
- A newly created summary worksheet, if requested and supported.

Rules:

1. Respect the user's specified output location.
2. Check whether the target cell already contains data.
3. Avoid overwriting existing information without authorization.
4. If no output location is specified, follow the application's established output-location policy.
5. Ask the user when the location is important and cannot be determined safely.

## 10. Absolute and Relative References

The agent should understand:

Relative reference:

```excel
A2
```

Absolute reference:

```excel
$A$2
```

Mixed references:

```excel
$A2
```

```excel
A$2
```

Use absolute or mixed references when required by formula copying, filling, or repeated calculations.

## 11. Multi-Sheet Range Identification

If the requested data exists in another worksheet:

1. Identify the correct worksheet.
2. Confirm that the required column exists.
3. Identify the actual data range.
4. Generate a worksheet-qualified reference.

Example pattern:

```excel
=SUM('Sales Data'!B2:B100)
```

The worksheet name and range must come from actual workbook information.

## 12. Data Validation Before Formula Generation

Check:

- Does the required column exist?
- Is the column appropriate for the requested calculation?
- Are the values in a suitable data type?
- Is the range complete?
- Are blank rows or text values present?
- Are the conditions mapped correctly?
- Is the output cell valid?
- Does the formula reference the intended worksheet?

If the required information is unavailable, do not guess.

## 13. Common Range Identification Errors

| Problem | Required Action |
|---|---|
| Incorrect column selected | Recheck the header and data |
| Header included in calculation | Identify the actual data start row |
| Last records excluded | Recheck the used range |
| Blank rows present | Inspect dataset boundaries |
| Wrong worksheet selected | Verify worksheet context |
| Multiple matching columns | Ask for clarification |
| Invalid output cell | Identify a safe target location |
| Different range lengths | Verify conditional and sum ranges |

## 14. Final Range Verification

Before formula execution, confirm:

1. Correct workbook.
2. Correct worksheet.
3. Correct column.
4. Correct first data row.
5. Correct last data row.
6. Correct criteria ranges.
7. Correct output cell.
8. Valid formula references.

Only after validation should the formula be passed to the appropriate available Excel execution tool.

## 15. Important Rules

- Never assume column letters.
- Never assume fixed row numbers.
- Never assume a worksheet name.
- Never include headers in calculations unless explicitly required.
- Never overwrite existing data without authorization.
- Use actual workbook metadata.
- Verify all formula references before execution.