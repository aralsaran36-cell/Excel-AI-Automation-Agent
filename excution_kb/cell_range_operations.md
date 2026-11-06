# Cell Range Operations

## 1. Purpose

This document defines how the Excel AI Agent identifies, interprets, validates, and performs operations on individual cells and cell ranges.

The agent must understand the user's requested operation, identify the correct target cells, and execute the operation using the appropriate available Excel execution tool.

This document focuses on cell-level and range-level operations, not formula selection or formula generation.

## 2. Types of Cell Range References

The agent must recognize different types of cell and range references.

| Reference Type | Example | Description |
|---|---|---|
| Single Cell | A1 | Refers to one cell |
| Horizontal Range | A1:E1 | Multiple cells in one row |
| Vertical Range | A1:A10 | Multiple cells in one column |
| Rectangular Range | A1:E10 | Multiple rows and columns |
| Entire Column | A:A | Refers to an entire column |
| Multiple Columns | A:C | Refers to multiple columns |
| Entire Row | 1:1 | Refers to an entire row |
| Multiple Rows | 1:10 | Refers to multiple rows |
| Non-Contiguous Range | A1:A5,C1:C5 | Separate ranges |
| Sheet-Qualified Range | Sheet1!A1:B10 | Range in a specific worksheet |

The agent must correctly interpret each reference before execution.

## 3. Understanding User Requests

The agent must identify the intended cell range from natural-language instructions.

Examples:

| User Request | Required Interpretation |
|---|---|
| Select A1 to D10 | Select A1:D10 |
| Read column B | Read the required cells in column B |
| Get the values from row 5 | Read the requested cells in row 5 |
| Select the entire sales column | Identify the sales column using headers |
| Read the first 10 records | Identify the data range and read 10 records |
| Select all used cells | Identify the worksheet's actual used range |
| Clear cells from B2 to B20 | Clear B2:B20 |
| Copy A1:C10 to E1 | Copy the source range to E1 |

Do not assume that a user-provided range is the only possible reference. Resolve natural-language references using actual workbook metadata.

## 4. Identifying Target Ranges

Before performing an operation, identify the correct target range.

### 4.1 Explicit Cell References

When the user provides an exact cell reference:

- Use the specified cell or range.
- Verify that the reference is valid.
- Confirm that the referenced worksheet exists.
- Check whether the requested range is within worksheet boundaries.
- Do not modify the requested range without a valid reason.

### 4.2 Header-Based Identification

When the user refers to a column using its header name:

1. Inspect the worksheet headers.
2. Identify the matching header.
3. Determine its actual column position.
4. Identify the relevant data rows.
5. Construct the appropriate range.

Example:

User request:

"Select the Salary column."

Workbook structure:

| A | B | C |
|---|---|---|
| Employee ID | Name | Salary |
| 101 | Arun | 30000 |
| 102 | Kumar | 40000 |

Correct range:

C2:C3

Do not assume that Salary is always located in column C.

### 4.3 Used Range Identification

When the user requests all used cells:

- Inspect the actual worksheet content.
- Identify the first used row and column.
- Identify the last used row and column.
- Exclude completely unused areas.
- Consider whether formatting-only cells are included in the workbook's reported used range.
- Prefer the actual data boundaries when the user means the data table.

### 4.4 Relative Range Identification

When the user describes a range relatively:

Examples:

- First column
- Last five rows
- Second data column
- Entire table
- First 10 records

The agent must resolve the reference using worksheet structure and available metadata.

Do not guess row or column positions.

## 5. Supported Cell Range Operations

### 5.1 Read Range

Purpose: Retrieve values or information from a specified cell range.

Examples:

- Read A1:D10.
- Get all values from the Name column.
- Read the first five records.
- Retrieve the complete data table.

Rules:

- Identify the correct worksheet.
- Validate the target range.
- Retrieve the requested values.
- Preserve row and column relationships.
- Avoid retrieving unrelated cells.
- Handle empty cells correctly.

### 5.2 Write Range

Purpose: Insert or update values in one or more cells.

Examples:

- Enter 50000 in B5.
- Fill A1:A5 with values.
- Write employee details into A2:D2.
- Update the selected range with new values.

Rules:

- Validate the destination range.
- Verify that the supplied values match the target dimensions.
- Preserve the intended row and column order.
- Do not overwrite unrelated data.
- Confirm whether existing content will be replaced.
- Handle merged cells appropriately.

### 5.3 Clear Range

Purpose: Remove information from selected cells.

Examples:

- Clear B2:B10.
- Remove values from the selected range.
- Clear the contents of the Salary column.

Distinguish between:

| Operation | Expected Behaviour |
|---|---|
| Clear Contents | Remove cell values and formulas while preserving formatting |
| Clear Formats | Remove formatting while preserving values |
| Clear All | Remove contents and formatting |
| Delete Cells | Remove cells and shift surrounding cells according to the specified direction |

Important rules:

- Never confuse clearing contents with deleting cells.
- Do not clear an entire worksheet unless explicitly requested.
- Confirm potentially destructive operations when the target is ambiguous.
- Preserve unrelated data and formatting.

### 5.4 Copy Range

Purpose: Duplicate cell content from a source range to a destination range.

Examples:

- Copy A1:B10 to D1.
- Duplicate the first five records.
- Copy the header row to another location.

Rules:

- Identify the source range.
- Identify the destination location.
- Validate both ranges.
- Check whether the destination contains existing data.
- Preserve relative cell relationships.
- Respect the requested copy behaviour.

### 5.5 Move Range

Purpose: Move cell content from one location to another.

Examples:

- Move A1:C10 to E1.
- Move the Salary column after the Department column.

Rules:

- Identify the source and destination.
- Validate both locations.
- Check for overlapping ranges.
- Consider the effect on surrounding cells.
- Avoid accidental data loss.
- Verify the result after execution.

### 5.6 Insert Cells or Ranges

Purpose: Insert new cells into an existing worksheet.

Examples:

- Insert three cells above the selected location.
- Insert a new column between B and C.
- Add a blank row before row 5.

Rules:

- Identify the insertion position.
- Determine the required number of cells, rows, or columns.
- Determine the required shift direction when inserting individual cells.
- Check for possible effects on existing data.
- Preserve worksheet structure.

### 5.7 Delete Cells or Ranges

Purpose: Remove selected cells, rows, or columns.

Examples:

- Delete cells B2:B5.
- Delete the entire third column.
- Remove rows 5 to 10.

Rules:

- Distinguish between deleting cell contents and deleting cells.
- Identify whether the request targets cells, complete rows, or complete columns.
- Determine the required shift behaviour.
- Check whether other data may be affected.
- Verify the remaining worksheet structure.

## 6. Range Dimension Validation

Before writing or copying values, verify the dimensions of the source and destination.

For a rectangular range:

- Row count = Last row - First row + 1
- Column count = Last column - First column + 1

Example:

Range A1:C5:

- Rows = 5
- Columns = 3
- Total cells = 15

If the user provides 15 values, the agent must arrange them according to the intended row and column structure.

For copy operations, validate source and destination dimensions when the operation requires matching sizes.

Do not silently truncate or expand supplied data.

## 7. Handling Multiple Ranges

A user may request operations on multiple separate ranges.

Examples:

- Clear A1:A10 and C1:C10.
- Read A1:B5 and D1:E5.
- Copy A1:A10 to C1:C10 and E1:E10.

Rules:

1. Identify every requested range.
2. Validate each range independently.
3. Preserve the requested execution order.
4. Check for overlapping ranges.
5. Avoid including unintended cells.
6. Execute through an appropriate available tool.

If one range is invalid, do not silently ignore it. Report the affected range and determine whether the remaining operations can safely proceed.

## 8. Handling Merged Cells

Merged cells require special attention.

Rules:

- Detect merged cells when workbook metadata supports it.
- Identify the complete merged area.
- Avoid treating individual cells inside a merged area as independent cells.
- Use the merged area's top-left cell when accessing its stored value.
- Avoid partial operations on merged ranges when they could produce an invalid result.
- Do not unmerge cells unless explicitly requested or required by an authorized operation.

## 9. Handling Empty Cells and Data Boundaries

The agent must distinguish between:

- Empty cells.
- Cells containing spaces.
- Cells containing formulas that return empty strings.
- Cells containing actual values.
- Formatting-only cells.

Rules:

- Do not assume that the first blank cell indicates the end of a data table.
- Check the relevant column or table boundaries.
- Avoid including unrelated notes or summary sections.
- Preserve empty cells when reading rectangular ranges.
- Do not automatically remove blank rows unless requested.

## 10. Range Operation Execution Workflow

Follow this workflow for every cell range operation.

1. Understand the user's requested action.
2. Identify the target workbook.
3. Identify the target worksheet.
4. Determine the requested source and destination ranges.
5. Resolve natural-language references using workbook metadata.
6. Validate cell and range references.
7. Check dimensions and data boundaries.
8. Identify potential overwriting or data-loss risks.
9. Select the appropriate available Excel execution tool.
10. Execute the operation.
11. Verify the execution result.
12. Report the outcome accurately.

## 11. Error Handling

Handle the following situations:

| Situation | Required Behaviour |
|---|---|
| Invalid cell reference | Reject or resolve the invalid reference |
| Worksheet not found | Identify the available worksheets |
| Empty target range | Report that no relevant data was found |
| Invalid range dimensions | Correct the dimensions before execution |
| Source and destination overlap | Evaluate the operation's safety |
| Protected worksheet | Report the protection limitation |
| Merged cell conflict | Resolve the merged range issue |
| Destination contains data | Prevent unintended overwriting |
| Execution failure | Follow the execution recovery rules |
| Partial execution | Report successful and failed ranges separately |

Do not claim that an operation succeeded unless the execution result confirms it.

## 12. Important Rules

- Never assume fixed cell addresses.
- Never assume fixed column letters.
- Never assume a worksheet name.
- Always inspect actual workbook metadata.
- Identify ranges using headers when appropriate.
- Distinguish cell references from row and column references.
- Validate source and destination ranges.
- Preserve the original data structure.
- Avoid unintended overwriting.
- Do not confuse clearing contents with deleting cells.
- Handle merged cells carefully.
- Preserve empty cells when their positions matter.
- Validate multi-range operations independently.
- Verify the result after execution.
- Use only available Excel execution tools.
- Follow execution recovery rules when an operation fails.

## 13. Final Execution Principle

The Excel AI Agent must identify the exact target range before executing any cell-level operation.

Range identification, dimension validation, operation selection, and execution verification must be treated as separate responsibilities.

The agent must prioritize data integrity and execute only the operation requested by the user.