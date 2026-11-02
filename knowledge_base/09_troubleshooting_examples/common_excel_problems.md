# Common Excel Problems and Solutions

## 1. Overview

Excel users frequently face practical problems while working with spreadsheets, formulas, data, charts, and workbooks.

This document provides solutions to common Excel problems and helps an Excel AI Agent understand how to diagnose and resolve them.

## 2. Data Entry Problems

### 2.1 Duplicate Records

**Problem:** The dataset contains repeated records.

**Solution:**

1. Select the dataset.
2. Go to Data → Remove Duplicates.
3. Select the columns that define a duplicate.
4. Confirm the operation.

Important:
- Check which columns identify a unique record.
- Create a backup before removing records.
- Do not assume that two identical-looking rows are always duplicates.

To identify duplicates using a formula:

```excel
=COUNTIF(A:A,A2)>1
```

This identifies repeated values in column A.

### 2.2 Extra Spaces in Text

**Problem:** Text contains unnecessary leading, trailing, or repeated spaces.

**Solution:**

```excel
=TRIM(A2)
```

For non-printable characters:

```excel
=CLEAN(A2)
```

Combined:

```excel
=TRIM(CLEAN(A2))
```

Some Unicode spaces may require additional cleaning.

### 2.3 Inconsistent Text Formatting

**Problem:** The same category appears in different formats.

Example:
- Chennai
- chennai
- CHENNAI

Solutions:

```excel
=UPPER(A2)
```

```excel
=LOWER(A2)
```

```excel
=PROPER(A2)
```

Choose the format required by the dataset.

### 2.4 Numbers Stored as Text

**Problem:** Excel does not calculate some numerical values correctly.

Possible solutions:

```excel
=VALUE(A2)
```

Other methods:
- Use the Convert to Number option.
- Use Text to Columns when appropriate.
- Check for spaces or unwanted characters.

### 2.5 Blank Cells in Data

**Problem:** Important fields contain missing values.

Solutions:
- Identify blank cells.
- Determine whether the value is genuinely missing.
- Fill values only when a valid source exists.
- Exclude records only when justified.

To count blanks:

```excel
=COUNTBLANK(A2:A100)
```

To identify blank cells:

```excel
=IF(A2="","Missing","Available")
```

## 3. Formula Problems

### 3.1 Formula Displays as Text

**Problem:** Excel displays the formula instead of calculating it.

Example:

```text
=SUM(A1:A10)
```

Possible causes:
- Cell format is Text.
- Formula begins with an apostrophe.
- Show Formulas mode is enabled.

Solutions:
1. Change the cell format to General.
2. Re-enter the formula.
3. Check whether Show Formulas is enabled.

### 3.2 Formula Returns Zero

**Problem:** A formula returns zero even though data exists.

Possible causes:
- Incorrect cell references.
- Numbers stored as text.
- Incorrect criteria.
- Filters or conditions excluding records.

Example:

```excel
=SUMIF(A2:A100,"South",B2:B100)
```

Check whether the region values match the criteria and whether column B contains numerical values.

### 3.3 Formula Not Updating Automatically

**Problem:** Formula results remain unchanged after source data is modified.

Solutions:
1. Check Formulas → Calculation Options.
2. Select Automatic when appropriate.
3. Recalculate the workbook.
4. Check whether the formula references the correct cells.

### 3.4 Incorrect Relative References

**Problem:** A formula changes references unexpectedly when copied.

Example:

```excel
=A2*B2
```

To keep B2 fixed:

```excel
=A2*$B$2
```

Reference types:

| Reference | Behavior |
|---|---|
| A1 | Row and column change |
| $A$1 | Row and column remain fixed |
| A$1 | Row remains fixed |
| $A1 | Column remains fixed |

### 3.5 SUM Formula Excludes Some Values

**Problem:** SUM does not include expected values.

Possible causes:
- Values are stored as text.
- Formula range is incomplete.
- Cells contain errors.

Solution:

```excel
=SUM(A2:A100)
```

Inspect the source values and verify the range.

## 4. Sorting and Filtering Problems

### 4.1 Sorting Only One Column

**Problem:** Sorting one column separates records from their related information.

Solution:
1. Select the complete dataset.
2. Use Data → Sort.
3. Confirm that the entire data range is selected.
4. Choose the required sorting column.

Always keep related records together.

### 4.2 Filter Does Not Show All Records

Possible causes:
- Filter range is incomplete.
- Blank rows split the dataset.
- Some values contain inconsistent formatting.
- Another filter is already active.

Solutions:
- Clear existing filters.
- Convert the dataset into an Excel Table.
- Check the complete source range.
- Standardize category values.

### 4.3 Sorting Dates Incorrectly

**Problem:** Dates are sorted as text instead of chronological dates.

Solution:
1. Check whether the values are real Excel dates.
2. Convert valid text dates into date values.
3. Apply the required date format.
4. Sort again.

Check:

```excel
=ISNUMBER(A2)
```

Valid Excel dates are stored as numeric serial values.

## 5. Formatting Problems

### 5.1 Numbers Display as ####

**Problem:** Excel displays hash symbols instead of a value.

Possible causes:
- Column width is insufficient.
- Date or time value cannot be displayed.
- A negative date or time is being displayed in an unsupported date system or format.

Solutions:
- Increase column width.
- Use an appropriate number format.
- Inspect the underlying value.

### 5.2 Dates Display Incorrectly

**Problem:** Dates appear in an unexpected format.

Solution:
1. Select the date cells.
2. Open Format Cells.
3. Choose Date or Custom.
4. Apply the required format.

Example:

```text
dd-mmm-yyyy
```

Formatting changes how a value appears; it does not necessarily convert text into a valid date.

### 5.3 Percentage Displays Incorrectly

**Problem:** A percentage is displayed at an unexpected scale.

Example:
- Raw value: 0.85
- Expected display: 85%

Solution:
Apply Percentage number formatting.

Do not multiply the value by 100 and then apply percentage formatting unless the calculation requires it.

### 5.4 Conditional Formatting Not Working

Possible causes:
- Incorrect rule.
- Wrong cell range.
- Incorrect formula references.
- Conflicting rules.

Solutions:
1. Open Conditional Formatting → Manage Rules.
2. Check the rule formula.
3. Verify the Applies To range.
4. Review rule priority.
5. Test the rule using sample values.

## 6. Worksheet Problems

### 6.1 Worksheet Accidentally Deleted

Possible solutions:
- Use Undo immediately if the action is still reversible.
- Restore from a previous saved version.
- Recover from an available backup.

A deleted worksheet may not be recoverable after the workbook is saved and closed unless a suitable recovery source exists.

### 6.2 Worksheet Is Hidden

Solution:
1. Right-click a worksheet tab.
2. Select Unhide.
3. Choose the required worksheet.

If the sheet is very hidden, it may require the VBA editor or another supported workbook operation.

### 6.3 Cells Cannot Be Edited

Possible causes:
- Worksheet protection.
- Locked cells.
- Workbook sharing or permission restrictions.

Solution:
- Check Review → Unprotect Sheet.
- Confirm whether a password is required.
- Check workbook permissions.

Do not attempt to bypass protection without authorization.

### 6.4 Freeze Panes Not Working as Expected

Solution:
1. Select the cell below the rows and to the right of the columns to freeze.
2. Go to View → Freeze Panes.
3. Select Freeze Panes.

Use Freeze Top Row when only the header row needs to remain visible.

## 7. Lookup Problems

### 7.1 VLOOKUP Returns #N/A

Possible causes:
- Lookup value does not exist.
- Lookup value contains extra spaces.
- Different data types.
- Incorrect lookup range.
- Incorrect matching mode.

Solution:

```excel
=VLOOKUP(A2,D2:F100,3,FALSE)
```

Use FALSE for an exact match.

### 7.2 XLOOKUP Returns an Incorrect Result

Example:

```excel
=XLOOKUP(A2,D2:D100,E2:E100,"Not Found")
```

Check:
- Lookup value.
- Lookup array.
- Return array.
- Matching mode.
- Duplicate lookup values.

XLOOKUP normally returns the first matching value when duplicates exist. Confirm whether this behavior matches the requirement.

### 7.3 Lookup Does Not Match Similar Text

Possible causes:
- Extra spaces.
- Hidden characters.
- Inconsistent spelling.
- Text-versus-number differences.

Solution:
Normalize both lookup and source values before performing the lookup.

## 8. PivotTable Problems

### 8.1 New Data Is Missing

Solution:
- Verify the PivotTable source range.
- Use an Excel Table as the source where appropriate.
- Refresh the PivotTable.

### 8.2 Values Are Counted Instead of Summed

Solution:
1. Right-click a value field.
2. Select Value Field Settings.
3. Select Sum.
4. Verify that the source values are numerical.

If Excel only offers Count, inspect the source data for text or invalid values.

### 8.3 PivotTable Shows Blank Categories

Possible causes:
- Empty cells in the source field.
- Incomplete data.
- Formula-generated empty strings.

Solution:
Inspect the source column and decide how blank categories should be handled.

### 8.4 PivotTable Dates Cannot Be Grouped

Possible causes:
- Text dates.
- Blank values.
- Invalid date entries.

Solution:
- Convert valid text dates into actual dates.
- Handle blanks.
- Refresh the PivotTable.
- Try grouping again.

## 9. Chart Problems

### 9.1 Chart Does Not Show All Data

Possible causes:
- Incorrect source range.
- Hidden rows or columns.
- Active filters.
- Missing values.

Solution:
Inspect the chart source and its hidden-data settings.

### 9.2 Chart Shows Incorrect Categories

Solution:
- Open Select Data.
- Verify category label ranges.
- Check the data series.
- Confirm the source range.

### 9.3 Chart Is Difficult to Read

Solution:
- Select an appropriate chart type.
- Add clear titles.
- Improve axis labels.
- Remove unnecessary series.
- Use consistent formatting.

### 9.4 Chart Does Not Update with New Rows

Solution:
Use an Excel Table as the chart source where supported, or update the chart data range.

## 10. Data Validation Problems

### 10.1 Drop-down List Does Not Display All Options

Possible causes:
- Source range is incomplete.
- New items are outside the source range.
- Source values contain blanks.
- Named range is incorrect.

Solution:
- Inspect Data Validation settings.
- Verify the source range.
- Use a dynamic source or Excel Table-based approach where appropriate.

### 10.2 Invalid Data Is Entered

Solution:
1. Select the required cells.
2. Open Data → Data Validation.
3. Choose the appropriate validation type.
4. Define allowed values or limits.
5. Configure an error alert.

Validation prevents many incorrect entries but does not guarantee that all existing data is valid.

## 11. File and Workbook Problems

### 11.1 Excel File Is Too Large

Possible causes:
- Excessive formatting.
- Unused rows and columns containing formatting.
- Large embedded images.
- Duplicate data.
- Too many formulas.
- Large pivot caches.

Solutions:
- Remove unnecessary formatting.
- Review the used range.
- Compress images where appropriate.
- Remove genuinely unnecessary objects.
- Optimize formulas.
- Save in an appropriate workbook format.

Always preserve required data and functionality.

### 11.2 Workbook Opens Slowly

Possible causes:
- Complex formulas.
- Volatile functions.
- External links.
- Large datasets.
- Many conditional formatting rules.
- Complex workbook dependencies.

Solutions:
- Review formulas.
- Reduce unnecessary recalculations.
- Inspect external connections.
- Remove unused objects.
- Consider Power Query or another appropriate data preparation workflow.

### 11.3 File Format Compatibility

Common Excel formats:

| Extension | Description |
|---|---|
| .xlsx | Standard Excel workbook |
| .xlsm | Macro-enabled workbook |
| .xls | Older Excel workbook format |
| .csv | Comma-separated text data |
| .xlsb | Binary Excel workbook |

Important:
- CSV does not preserve multiple worksheets, formulas, charts, or Excel formatting.
- Saving as XLSX removes VBA macros.
- XLSM should be used when VBA macros need to be retained.

Always confirm the required output format.

## 12. Practical Problem-Solving Examples

### Example 1: Remove Duplicate Customer IDs

Source data:

| Customer ID | Name |
|---|---|
| C101 | Arun |
| C102 | Bala |
| C101 | Arun |
| C103 | Kumar |

Formula:

```excel
=COUNTIF($A$2:A2,A2)>1
```

This identifies repeated occurrences after the first occurrence.

### Example 2: Find Missing Values

```excel
=IF(A2="","Missing","Available")
```

### Example 3: Calculate Total Sales by Region

```excel
=SUMIFS(C2:C100,A2:A100,"South")
```

### Example 4: Identify Low Stock

Assume B2 contains available stock and C2 contains the reorder level.

```excel
=IF(B2<=C2,"Reorder","Sufficient Stock")
```

### Example 5: Calculate Employee Attendance Percentage

```excel
=IFERROR(B2/C2,0)
```

Where:
- B2 = Present days.
- C2 = Total eligible working days.

### Example 6: Identify Sales Below Target

```excel
=IF(B2<C2,"Below Target","Target Achieved")
```

Where:
- B2 = Actual sales.
- C2 = Sales target.

## 13. General Troubleshooting Checklist

Before modifying an Excel workbook:

1. Identify the user's exact problem.
2. Inspect the relevant worksheet.
3. Understand the source data.
4. Check formulas and references.
5. Inspect data types.
6. Identify filters and hidden records.
7. Check workbook protection.
8. Determine whether the issue is caused by formatting or calculation.
9. Choose the smallest safe correction.
10. Verify the result.

## 14. Best Practices

- Keep a backup before major modifications.
- Use Excel Tables for structured datasets.
- Use meaningful worksheet names.
- Keep formulas understandable.
- Avoid unnecessary merged cells in data tables.
- Validate important inputs.
- Standardize date and number formats.
- Use appropriate error handling.
- Refresh PivotTables when required.
- Verify charts after source changes.
- Preserve formulas and workbook functionality.
- Document important workbook changes.

## 15. Excel AI Agent Instructions

When a user reports a practical Excel problem, the AI Agent should:

1. Understand the user's request.
2. Inspect the workbook and identify the relevant sheet.
3. Determine whether the problem concerns data, formulas, formatting, charts, PivotTables, or workbook structure.
4. Inspect the affected cells and source values.
5. Identify the root cause before modifying anything.
6. Explain the proposed solution.
7. Ask for permission before destructive or irreversible operations.
8. Use the appropriate Excel tools to apply the correction.
9. Preserve unrelated workbook content.
10. Verify the result after the operation.
11. Report what was changed.
12. Explain any remaining limitations.

### Example User Requests

- "Remove duplicate records."
- "Clean all unnecessary spaces."
- "Convert numbers stored as text."
- "Fix my sorting problem."
- "Correct the date format."
- "My PivotTable is not refreshing."
- "Fix my chart."
- "Unhide a worksheet."
- "Create a drop-down list."
- "Find blank cells."
- "Optimize my Excel workbook."
- "Find and fix common Excel problems."

### Important Agent Rule

The AI Agent must inspect the actual workbook before selecting a solution. It should not make destructive changes automatically, should preserve original data whenever possible, and must verify the result before reporting success.