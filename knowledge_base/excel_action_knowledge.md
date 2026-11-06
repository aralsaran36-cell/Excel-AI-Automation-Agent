Excel Action Execution Knowledge
Reasoning-Oriented Knowledge for Natural-Language Excel Agents
1. Purpose and Operating Principles
This document provides broad, action-oriented Excel knowledge to help an AI agent reason about and execute user-requested workbook operations. It is not a collection of answers to copy. It is a decision resource: the agent must inspect the actual workbook context, understand the user's intent, select an appropriate method, construct the required operation dynamically, execute it through available tools, and verify the result.
1.1 Core principles
- Treat the user's requested outcome as the goal; do not assume a particular formula or tool before understanding the task.
- Use actual workbook, worksheet, header, data-type, row-count, and range information whenever available.
- Choose the simplest reliable method that satisfies the requested outcome and preserves existing data.
- Distinguish between calculating a result, writing a result, changing source data, formatting cells, and creating a new analytical view.
- Do not invent sheet names, column names, cell addresses, values, criteria, or workbook structure.
- If required information is absent or ambiguous, inspect the workbook first. Ask a concise clarification only when inspection cannot resolve the ambiguity.
- Prefer formulas when the user needs a live, recalculating result. Prefer values when the user explicitly asks for a static result or the task requires a fixed output.
- Consider whether an operation affects one cell, a range, an entire column, a table, a worksheet, or multiple worksheets.
- Protect source data. Avoid destructive changes unless clearly requested and confirmed where appropriate.
- After execution, verify the outcome using available workbook-reading or inspection capabilities.
- Do not claim an operation succeeded unless the tool result or a follow-up inspection supports that claim.
1.2 Reasoning cycle
For each request, reason through these questions as applicable:
1. What final state or result does the user want?
2. Which workbook and worksheet contain the relevant data?
3. What is the actual data layout, including headers, used range, blank rows, tables, and data types?
4. What operation category best fits the request?
5. Which method is most suitable, and why?
6. What exact source cells, ranges, criteria, and destination are required?
7. Could the operation overwrite, delete, reorder, or otherwise alter existing information?
8. What tool capabilities are available to perform the operation?
9. What evidence will establish that the operation completed correctly?
This is a reasoning framework, not a mandatory fixed sequence. Skip irrelevant questions and inspect only what is needed.
2. Understand Workbook Structure Before Acting
2.1 Workbook and worksheet selection
A workbook may contain multiple sheets with similar names or related data. Determine the relevant sheet from the user's wording and workbook context. Use sheet names exactly as they appear. Do not assume that the active sheet is the intended sheet.

When a request refers to:
- "this sheet": resolve the sheet currently being discussed or shown in context.
- "all sheets": determine whether the requested operation applies to every worksheet, including hidden sheets if the tools expose them.
- "sales sheet" or another subject: locate the sheet whose name or contents match that subject.
- a sheet by position ("second sheet"): inspect the workbook's actual sheet order before acting.
If multiple sheets could reasonably match, inspect their headers or ask which one is intended.
2.2 Identify the data region
Do not assume that data starts at A1 or that the first row is always the header. Inspect the used range and sample rows. Identify:
- Header row and header names.
- First and last data rows.
- First and last relevant columns.
- Blank rows or columns inside the data.
- Repeated header rows.
- Totals or notes below the data.
- Excel Tables and their names, if available.
- Merged cells that may affect reading or writing.
A used range can include old formatting beyond the real data. Use actual populated cells and workbook context rather than blindly treating every used-range cell as data.
2.3 Header and column reasoning
Map user language to real headers carefully. For example, "employee name" may correspond to Name, Employee, or Employee Name. Prefer an exact or clearly equivalent header. If two columns are plausible, inspect sample values or ask.

Never silently substitute a similar column when the difference changes the meaning, such as Gross Salary versus Net Salary, Order Date versus Delivery Date, or Revenue versus Profit.

Column positions may change. Prefer header-based identification over hardcoded column letters when the tools support it. Convert to cell addresses only after resolving the actual column and row positions.
2.4 Data type and value inspection
A column's display format does not prove its underlying data type. Inspect representative values and, where possible, the underlying types.

Distinguish:
- Numeric values from numbers stored as text.
- Real Excel dates from date-looking text.
- Empty cells from formulas returning an empty string.
- True/False values from text labels such as "Yes" and "No".
- Percentages stored as decimals from whole-number percentages.
- Currency values from plain numbers.
- Identifiers such as employee IDs and postal codes from quantities.
Do not perform arithmetic on identifiers. Do not convert leading-zero identifiers to numbers if the zeros are meaningful.
2.5 Determine the intended scope
The same request may refer to:
- One selected cell.
- One row or column.
- The current data region.
- An Excel Table.
- All populated rows.
- A filtered subset.
- Multiple sheets or workbooks.
Use the narrowest scope that fulfills the request. If the user says "all employees", include all relevant employee records, not merely visible rows, unless the request explicitly concerns the filtered/visible records.
3. Convert User Intent into an Excel Operation
Interpret the desired outcome, not just isolated keywords.
3.1 Common intent categories
User's intended outcome	Possible operation families
Calculate a value	Formula, aggregation, statistical calculation
Calculate by condition	Conditional formula, grouped aggregation
Find a matching record	Lookup, match, filter, search
Compare records	Difference, comparison formula, conditional logic
Change existing values	Update, replace, transform, fill
Fix inconsistent data	Cleaning, normalization, type conversion
Arrange records	Sort, filter, group
Remove unwanted records	Duplicate removal, row deletion, blank removal
Make data easier to read	Formatting, alignment, number formats, column sizing
Highlight important records	Conditional formatting
Restrict future entries	Data validation
Summarize data	Formula summary, PivotTable, grouped aggregation
Present trends	Chart or dashboard
Reorganize data	Split, combine, transpose, reshape
Explain a result	Read/inspect and interpret; do not modify unless requested


A request may combine multiple categories. Determine dependencies and whether a sequence is needed. Do not force a multi-step workflow when a single operation is sufficient.
3.2 Distinguish verbs precisely
- "Show", "calculate", or "find" may mean return an answer without changing the workbook.
- "Add a formula", "write", or "insert" means modify the workbook.
- "Replace" means overwrite existing content and may require confirmation.
- "Format" means change presentation, not underlying values.
- "Clean" is broad; inspect the issue and identify the intended correction before changing data.
- "Remove" may mean filter from view, clear contents, or delete rows. Determine which is intended.
- "Create a summary" may mean a new summary range, a PivotTable, or a chart. Choose based on the requested result and available capabilities.
3.3 Handle incomplete instructions
If the user says "calculate profit" but the workbook has both Revenue and Sales, and multiple possible cost columns, inspect headers and sample data. If the intended business definition remains unclear, ask whether profit means revenue minus total cost, sales minus product cost, or another definition.

If the user says "highlight high salary", determine whether a threshold is available from context. If not, ask for a threshold or offer a data-driven interpretation only when the user permits it.

Do not invent business rules.
4. Formula Selection and Construction
4.1 Select formulas based on the required logic
Choose a formula based on the operation's meaning and data structure, not because a familiar example exists.
Required logic	Formula families to consider
Add values	SUM, arithmetic operators
Average values	AVERAGE, AVERAGEIF, AVERAGEIFS
Count numeric or non-empty cells	COUNT, COUNTA, COUNTBLANK
Count records matching conditions	COUNTIF, COUNTIFS
Add values matching conditions	SUMIF, SUMIFS
Average values matching conditions	AVERAGEIF, AVERAGEIFS
Find smallest/largest	MIN, MAX, SMALL, LARGE
Return a value based on a condition	IF, IFS, SWITCH
Combine conditions	AND, OR, NOT
Find matching records	XLOOKUP, VLOOKUP, INDEX + MATCH, MATCH
Handle errors	IFERROR, IFNA
Extract or combine text	LEFT, RIGHT, MID, TEXTBEFORE, TEXTAFTER, TEXTJOIN, CONCAT
Clean text	TRIM, CLEAN, SUBSTITUTE, UPPER, LOWER, PROPER
Work with dates	DATE, YEAR, MONTH, DAY, TODAY, EOMONTH, DATEDIF
Dynamic arrays	FILTER, SORT, SORTBY, UNIQUE, SEQUENCE
Round values	ROUND, ROUNDUP, ROUNDDOWN, MROUND
Rank or position	RANK.EQ, MATCH, XMATCH


Formula availability varies by Excel version. If a modern function may not be supported, consider a compatible alternative or check the environment before writing it.
4.2 Build formulas from actual workbook context
Before constructing a formula, resolve:
- The actual source column or cell.
- The first and last relevant data rows.
- The criteria range and criteria value.
- Whether headers or totals are included.
- Whether the formula should be copied down or used once.
- The intended output cell or output column.
- Whether references should change when copied.
Do not copy a sample formula with assumed columns or row numbers. Adapt every reference to the inspected workbook.
4.3 Conditional calculations
For conditional calculations, ensure that criteria ranges and calculation ranges cover corresponding records. In SUMIFS, COUNTIFS, and AVERAGEIFS, mismatched range dimensions can produce incorrect results or errors.

Reason about:
- One condition versus multiple conditions.
- Exact match versus partial text match.
- Numeric thresholds (>, <, >=, <=).
- Date intervals and inclusive/exclusive boundaries.
- Blank or nonblank criteria.
- Text case and extra spaces.
- Whether criteria should be fixed or refer to a cell.
Use a formula that expresses the actual condition. Do not simplify multiple conditions if doing so changes the logic.
4.4 Relative, absolute, and mixed references
Choose references according to how the formula will be reused:
- Relative reference: changes when copied to another location.
- Absolute reference: remains fixed when copied.
- Mixed reference: fixes either the row or the column.
For a formula copied down, determine whether the source row should move while a threshold or lookup range remains fixed. For a formula copied across, determine whether the source column should move while a criteria row remains fixed.

Use absolute references for fixed parameters and lookup ranges where appropriate. Do not make every reference absolute by default.
4.5 Formula placement and fill behavior
When applying a formula to multiple rows:
- Determine the first data row, not the header row.
- Determine the final applicable row from actual data.
- Confirm whether blank rows interrupt the data.
- Check whether a table's calculated-column behavior is available and appropriate.
- Preserve existing formulas and values outside the requested target.
- Avoid filling into totals, notes, or unrelated sections.
- Confirm that relative references adjust correctly.
For a single aggregate result, do not fill the formula down unless the user asks for row-level results.
4.6 Formula alternatives and trade-offs
When multiple methods can produce the same outcome, choose based on maintainability, compatibility, data size, and the user's intended future use.

Examples of reasoning:
- XLOOKUP is often clearer and more flexible than VLOOKUP, but older Excel versions may not support it.
- SUMIFS is suitable for multiple criteria; a PivotTable may be better for a large interactive grouped summary.
- A dynamic-array formula may be concise, but a static output may be more suitable when the user requests a fixed snapshot.
- A helper column may improve transparency when a complex formula becomes difficult to audit.
- A formula is preferable when results must update as source data changes; a value is appropriate for a requested fixed result.
Do not choose a more complex method merely because it is newer.
4.7 Formula correctness checks
Check for:
- Correct operators and parentheses.
- Correct range boundaries.
- Correct criteria syntax.
- Matching range dimensions.
- Text values enclosed in quotation marks where required.
- Correct decimal and date interpretation.
- Correct handling of blanks and errors.
- Correct reference locking.
- No accidental inclusion of headers or totals.
- No unintended circular reference.
A formula returning a number is not proof that it is logically correct. Verify its meaning against the user's request and, where possible, independently inspect a small sample or compare with another calculation.
5. Lookup, Matching, and Record Retrieval
5.1 Understand the lookup relationship
Before choosing a lookup method, identify:
- The lookup key in the user's request.
- The source column containing that key.
- The return column.
- Whether the key should be unique.
- Whether duplicate keys exist.
- Whether exact matching is required.
- What should happen when no match exists.
Do not assume that names are unique. Employee ID, invoice ID, or another stable key is generally safer when available.
5.2 Choose a matching method
- Use XLOOKUP for flexible exact-match retrieval when supported.
- Use VLOOKUP when the return column is to the right of the lookup column and compatibility is important.
- Use INDEX with MATCH when a flexible compatible approach is needed.
- Use FILTER when all matching rows are required rather than one result.
- Use COUNTIF or COUNTIFS to test existence or count matches.
- Use MATCH or XMATCH when the position is needed rather than the returned value.
5.3 Exact versus approximate matching
Use exact matching for identifiers, names, invoice numbers, and most record-retrieval tasks. Approximate matching is appropriate only when the data is designed for ordered ranges, such as tax bands, grading thresholds, or commission slabs. Confirm sorting and boundary behavior before using approximate matching.
5.4 Duplicate keys and missing matches
If multiple source rows share a key, determine whether the user wants:
- The first matching record.
- The last matching record.
- Every matching record.
- A count or summary of matching records.
Do not silently return the first match when duplicates may materially change the result. For missing keys, use an explicit not-found result or report unmatched records; do not fabricate a value.
5.5 Reconciliation and matching across sheets
When comparing two sheets or lists:
- Identify the key columns in both datasets.
- Check data types, whitespace, case, and formatting differences.
- Determine whether matching should be exact or normalized.
- Identify records present in both, only in the first, or only in the second.
- Preserve original identifiers.
- Report unmatched and duplicate records where relevant.
6. Data Cleaning and Quality Improvement
6.1 Diagnose before changing
"Clean this data" is not a precise operation. Inspect the data and identify actual quality issues:
- Duplicate records.
- Missing values.
- Leading/trailing spaces.
- Inconsistent capitalization.
- Inconsistent category labels.
- Numbers stored as text.
- Dates stored as text.
- Mixed units or formats.
- Invalid values and outliers.
- Extra spaces or non-printing characters.
- Inconsistent delimiters or combined fields.
Do not delete or normalize data merely because it looks unusual. Confirm whether the difference is an error or a legitimate business value.
6.2 Duplicate handling
Distinguish exact duplicate rows from duplicate values in one selected column. Determine which columns define a duplicate record. For example, duplicate employee names do not necessarily mean duplicate employees; employee IDs may be the correct key.

Before removing duplicates:
- Identify the comparison columns.
- Determine whether to keep the first or last occurrence.
- Consider whether records contain conflicting values.
- Preserve a backup or create a separate cleaned output when the operation is destructive.
- Verify the number of records removed and the remaining unique records.
If the user only wants to identify duplicates, highlight or report them instead of deleting them.
6.3 Missing values
Distinguish truly empty cells, formulas returning empty strings, placeholders such as N/A, -, Unknown, and zero values. They do not always mean the same thing.

Possible actions include:
- Report missing values.
- Filter rows with missing values.
- Fill with a specified value.
- Use a justified statistical replacement.
- Carry values forward or backward only when the data's meaning supports it.
- Leave blanks unchanged.
Never replace blanks with zero automatically. Zero may mean a real measured value, while blank may mean unknown or not recorded.
6.4 Text normalization
Use trimming, cleaning, case normalization, and replacement only when they preserve meaning. Do not remove punctuation or leading zeros from identifiers unless the user requests it or the data rules clearly establish that they are irrelevant.

For category standardization, inspect distinct values first. Map variants to a canonical label only when their equivalence is clear. Keep a record of transformations where auditability matters.
6.5 Numeric and date conversion
Before converting:
- Inspect representative values.
- Determine decimal and thousands separators.
- Determine date order and locale.
- Identify currency symbols and units.
- Check whether text includes nonnumeric characters.
- Preserve original values when conversion may be lossy.
Do not convert ambiguous dates such as 03/04/2026 without establishing whether they mean March 4 or April 3.
6.6 Outliers and unusual values
An outlier is not automatically an error. Use domain context and statistical evidence. If the user asks to remove outliers, clarify the rule (for example, IQR or a specified threshold) unless it is already established. Prefer flagging unusual records before deletion.
7. Sorting, Filtering, and Selecting Records
7.1 Sorting
Identify:
- The complete record range to sort.
- The sort key or keys.
- Ascending or descending order.
- Whether the first row is a header.
- Whether rows must remain intact.
- Whether the user wants a stable secondary sort.
Sort entire records, not one column alone, unless the user explicitly asks to sort only a standalone list. Sorting one column independently can disconnect values from their related records.

For multi-level sorting, apply the primary key first and secondary keys as tie-breakers. Confirm that dates and numbers are true dates and numbers rather than text.
7.2 Filtering
Determine the required conditions and whether they combine with AND or OR logic. Distinguish:
- Filtering rows for display.
- Copying matching rows to a new location.
- Deleting nonmatching rows.
- Returning matching records in the chat.
Do not delete hidden or filtered-out records when the user only asked to filter. Clarify whether conditions apply to all records or only currently visible records.
7.3 Top, bottom, and threshold selections
For requests such as "top 10 employees by salary":
- Identify the numeric measure.
- Determine whether ties should expand the result.
- Determine whether the user wants a ranking or exactly 10 rows.
- Handle blanks and invalid numeric values.
- Preserve row relationships.
For threshold requests, use the exact threshold and boundary operator. "Above 50000" generally means greater than 50000; "at least 50000" includes 50000.
8. Data Transformation and Reshaping
8.1 Split fields
When splitting names, addresses, codes, or combined values, inspect delimiters and edge cases first. A simple space split may fail for multi-part names or addresses. Use the actual structure and preserve the original field if the transformation could lose information.
8.2 Combine fields
When combining columns:
- Determine the order.
- Choose the delimiter.
- Handle blanks intentionally.
- Avoid unwanted extra spaces or delimiters.
- Preserve original source columns unless replacement is requested.
8.3 Transpose and reshape
Determine whether the user wants:
- A one-time transpose.
- A live linked transpose.
- A wide-to-long transformation.
- A long-to-wide summary.
- A pivoted view.
Preserve headers, identifiers, and relationships. Check whether the destination area is large enough and whether existing content would be overwritten.
8.4 Fill and sequence operations
For filling values, infer a pattern only when the data supports it. Distinguish copying a constant, extending a sequence, repeating a pattern, and filling formulas. Do not infer a trend from a few ambiguous values without checking.
8.5 Insert, delete, move, and clear
Distinguish:
- Clear contents while retaining formatting.
- Clear formats while retaining values.
- Delete cells and shift surrounding cells.
- Delete an entire row or column.
- Move cells or ranges.
- Insert blank rows or columns.
These operations have different consequences. Check formulas, references, tables, and related records before making structural changes. Confirm destructive or broad-scope changes when needed.
9. Formatting and Presentation Operations
9.1 Number formats
Choose a display format that matches the meaning:
- Currency for monetary values.
- Percentage for ratios.
- Date/time formats for dates and times.
- Integer or decimal precision for quantities.
- Text for identifiers that must preserve leading zeros.
Formatting changes how a value is displayed, not necessarily the underlying value. Do not use number formatting as a substitute for correcting an incorrect data type.
9.2 Cell and range formatting
Use formatting to improve readability and communicate hierarchy:
- Headers should be distinguishable from data.
- Totals and subtotals should be visually identifiable.
- Consistent alignment and number formats improve scanning.
- Avoid excessive colors, borders, and font variations.
- Apply formatting to the intended range only.
Do not overwrite values when the user only asks for formatting.
9.3 Conditional formatting
Translate the user's meaning into a condition:
- Greater than or less than a threshold.
- Duplicate or unique values.
- Top/bottom values.
- Dates in a time window.
- Text containing a phrase.
- Formula-based row or cell conditions.
Identify whether the condition applies to one cell, a column, or an entire row. For formula-based rules, use correct relative and absolute references so the rule adjusts as intended. Check rule precedence and overlapping conditions when multiple rules exist.
9.4 Data validation
Determine what entries should be allowed:
- A fixed list.
- Whole numbers or decimals within limits.
- Dates within a range.
- Text length limits.
- A custom formula-based condition.
Identify the target input cells and the source of any dropdown list. Do not create a dropdown with invented choices. Preserve existing validation unless the user asks to replace it.
9.5 Column widths, row heights, freeze panes, and view settings
Use AutoFit or appropriate sizing when readability is the goal. Avoid making every column excessively wide. Freeze rows or columns based on what users need to keep visible while scrolling. Freeze panes relative to the selected cell and verify the correct rows/columns remain visible.

View settings affect presentation and navigation, not source data.
10. Excel Tables and Structured Data
10.1 When a Table is useful
An Excel Table is useful when a dataset needs:
- Structured headers.
- Automatic expansion as rows are added.
- Consistent formatting.
- Filter controls.
- Structured references.
- Calculated columns.
Before creating a Table, inspect the data boundaries and ensure the header row is correct. Avoid including notes, totals outside the dataset, or unrelated blank areas.
10.2 Structured references
When using Table formulas, identify the actual Table name and column names. Structured references can make formulas easier to understand and maintain, but they must match the workbook's real names.
10.3 Table changes
When adding rows or columns, determine whether the operation should extend the existing Table. Preserve calculated columns, totals rows, and filters where possible. Do not create a second Table over an existing Table's range.
11. Summary Analysis and PivotTables
11.1 Choose a summary method
Use formulas when the user needs a small number of explicit calculations or a live result in a known location. Consider a PivotTable when the user needs flexible grouping, multiple measures, category comparisons, or interactive exploration. Use a separate summary range when a simple report is sufficient and PivotTable support is unavailable.
11.2 Grouping logic
Identify:
- Grouping dimensions (department, region, product, month).
- Measures (sum of sales, average salary, count of orders).
- Filters (year, status, region).
- Sorting and ranking requirements.
- Treatment of blanks and unknown categories.
Do not group by a field whose values are inconsistent without first considering normalization.
11.3 Date grouping
For monthly or yearly summaries, ensure the source is a true date field. Distinguish calendar year from financial year and calendar month from rolling periods. Do not assume the financial year definition.
11.4 Totals and subtotals
Check whether totals should be sums, averages, counts, distinct counts, or another measure. Averages of group averages may be mathematically different from the overall average. Use weighted calculations when the business meaning requires them.
12. Charts and Visualization
12.1 Select a chart based on the analytical question
Analytical purpose	Chart types to consider
Compare categories	Bar or column
Show change over time	Line
Show composition of a whole	Stacked bar/column; pie only for a small number of clear categories
Show relationship between numeric variables	Scatter
Show distribution	Histogram or box plot, if available
Compare actual versus target	Bar/column with target reference
Show cumulative contribution	Pareto-style chart, if supported


Choose a chart that makes the comparison clear. Do not use a pie chart for many categories or unrelated measures.
12.2 Chart data reasoning
Identify category labels, numeric series, units, date order, and missing values. Avoid including totals as ordinary data points unless intended. Determine whether the chart should use one series or multiple series and whether axes need meaningful labels.
12.3 Chart integrity
Check that:
- The selected range matches the intended data.
- Headers are interpreted correctly.
- Categories are in the intended order.
- Dates are chronologically ordered.
- Units and number formats are clear.
- The chart title describes the actual measure.
- No unrelated blank or total rows are included.
Do not imply causation from a chart that only shows correlation or comparison.
13. Date and Time Reasoning
Excel dates are stored as serial values in many workbooks, but date-looking text may not behave as a real date. Inspect the source type before calculating.

Common operations include:
- Extracting year, month, or day.
- Calculating elapsed days.
- Finding month-end or year-end.
- Adding days or months.
- Identifying overdue records.
- Grouping by month, quarter, or year.
- Calculating age or service duration.
Reason about boundary conditions:
- Is the end date included?
- Does "within 30 days" include today?
- Is the period calendar-based or a fixed number of days?
- Should weekends or holidays be excluded?
- Is the date in local format or an ambiguous text format?
Use working-day functions only when the required weekend and holiday rules are known. Do not assume a particular financial year or date locale.
14. Text and Identifier Reasoning
For text operations, identify whether the task requires extraction, replacement, normalization, splitting, joining, or comparison.

Consider:
- Leading/trailing spaces.
- Multiple consecutive spaces.
- Non-printing characters.
- Case sensitivity.
- Punctuation.
- Delimiters.
- Leading zeros.
- Unicode characters and language-specific text.
- Text that looks numeric but is an identifier.
Do not convert IDs, phone numbers, postal codes, or account-like identifiers to numeric values if formatting is significant. When comparing text, decide whether case and whitespace differences should count as different values.
15. Statistical and Analytical Reasoning
Select the statistic that answers the user's question:
- Mean for arithmetic average when values are suitable.
- Median when a typical middle value is more representative under skew or outliers.
- Mode for the most frequent value.
- Minimum and maximum for extremes.
- Range for spread between extremes.
- Standard deviation or variance for dispersion.
- Percentile or quartile for distribution position.
- Rank for relative position.
Do not label a metric incorrectly. A count of non-empty cells is not necessarily a count of valid records. A simple average may be inappropriate when records have different weights. Explain or clarify the intended definition when it affects the result.

For comparisons, ensure that the groups use comparable periods, units, and populations. Do not infer causation from descriptive statistics.
16. Cross-Sheet and Cross-Workbook Operations
Before using references across sheets:
- Confirm the source and destination sheet names.
- Confirm that corresponding columns have compatible meanings.
- Confirm whether the sheets share a common key.
- Check whether the source range may expand.
- Consider spaces or special characters in sheet names.
- Avoid broken references when sheets are renamed or moved.
For cross-workbook operations, confirm the source workbook is available and identify whether the result should remain linked or become a static copy. External links can break when files move or are renamed. Do not create external references without understanding the user's intended maintenance behavior.
17. Formula and Workbook Error Diagnosis
17.1 Common formula errors
Error	Possible causes to investigate
#DIV/0!	Division by zero or an empty denominator
#N/A	Lookup value not found or unavailable
#VALUE!	Wrong value type or invalid argument
#REF!	Invalid or deleted cell reference
#NAME?	Misspelled function/name or unsupported function
#NUM!	Invalid numeric argument or out-of-range calculation
#NULL!	Incorrect range intersection syntax
#####	Often insufficient column width or incompatible date/time display


Do not hide errors with IFERROR before identifying their cause. Error handling should preserve meaningful information rather than conceal a broken calculation.
17.2 Debugging method
Inspect the formula, its referenced ranges, argument types, and sample source values. Check whether the issue comes from:
- Incorrect formula logic.
- Incorrect range boundaries.
- Text stored as numbers.
- Invalid dates.
- Missing lookup keys.
- Inconsistent criteria.
- Broken references.
- Unsupported functions.
- Circular references.
- Calculation settings.
Correct the underlying cause when possible. If a fallback is needed, make the fallback behavior explicit.
17.3 Formula consistency
When formulas are filled down or across, compare neighboring formulas for consistent patterns. A single formula that differs may be intentional or may be an error. Do not overwrite exceptions without checking their purpose.
18. Safe Execution and Change Management
18.1 Protect existing workbook content
Before writing, determine whether the destination is empty or already contains values, formulas, formatting, validation, or merged cells. Avoid overwriting unrelated content.

When the destination is uncertain:
- Inspect the target area.
- Choose a clearly separate output area only if that fits the user's request.
- Ask for a destination when placement materially matters.
18.2 Destructive operations
Deletion, clearing, replacing, bulk updates, and broad restructuring can cause data loss. Confirm the exact scope and intended effect when the request is ambiguous or high-impact. Use a backup or separate output when appropriate and supported.
18.3 Preserve formulas and relationships
When moving or deleting rows/columns, consider formulas, named ranges, Tables, charts, validation rules, and cross-sheet references. Do not assume that changing visible cells has no effect elsewhere.
18.4 Idempotence and repeated requests
Consider whether repeating an action would create duplicate sheets, duplicate columns, repeated formatting rules, or repeated records. When possible, inspect the current state and update the existing intended object rather than creating another copy. Do not silently overwrite a similarly named object that may be unrelated.
19. Verification After Execution
Verification is part of the task, not an optional presentation step.

Choose checks appropriate to the operation:
- Formula write: read the destination formula and resulting value.
- Fill-down: inspect the first, middle, and last formulas for reference progression.
- Aggregation: compare with a direct calculation or inspect a sample of included records.
- Lookup: verify matched keys and behavior for missing or duplicate keys.
- Cleaning: compare before/after row counts and inspect changed values.
- Duplicate removal: verify the selected duplicate definition and remaining key uniqueness.
- Sorting: confirm records remain intact and the sort order is correct.
- Filtering: confirm the conditions and whether source rows were preserved.
- Formatting: inspect the target range and ensure values remain unchanged.
- Chart: verify source range, labels, series, and category order.
- Sheet creation: confirm the new sheet exists and contains the intended output.
- Deletion: confirm only the intended object was removed.
Report what was actually completed. If verification is unavailable, state that the operation was performed but could not be independently verified.
20. Communicating Results
After execution, give a concise, factual summary:
- What operation was completed.
- Which workbook/sheet/range was affected, when known.
- The important result or output location.
- Any assumptions that materially affected the result.
- Any errors, skipped records, or unverified parts.
Do not claim perfect accuracy without evidence. Do not bury important warnings. If the task could not be completed, explain the exact blocker and what information or capability is needed.
21. Handling Ambiguity and Uncertainty
Ask a clarification only when the missing information changes the operation or could cause an incorrect or destructive result.

Examples:
- Multiple sheets contain the same type of data.
- Several columns could represent the requested measure.
- A threshold or business rule is missing.
- "Remove duplicates" does not identify the key columns.
- "Delete empty rows" could mean fully empty rows or rows with a blank in one required field.
- "Profit" has multiple possible definitions.
- A date format is ambiguous.
- A destination would overwrite existing data.
Do not ask questions that workbook inspection can answer. Do not interrupt for harmless choices when a conventional, reversible default is clear. State the assumption when it is safe and relevant.
22. Advanced Reasoning Patterns
22.1 Decompose compound requests
A request such as "clean the sales data, calculate monthly revenue, and create a chart" contains distinct outcomes. Identify dependencies:
- Cleaning may affect which records are included.
- Monthly revenue depends on a valid date and revenue field.
- The chart depends on the summary output.
Determine the necessary operations from the actual workbook. Do not apply a predefined sequence if the workbook already contains a cleaned dataset or summary.
22.2 Choose between formula, built-in feature, and transformation
A task may be possible through a formula, a built-in Excel feature, or a data transformation. Compare:
- Whether the result must update dynamically.
- Whether the user needs an interactive result.
- Whether the source data should remain unchanged.
- Whether the workbook supports the method.
- Whether the available execution tools can perform it reliably.
Choose the method that best satisfies the user's outcome and constraints.
22.3 Work with large datasets
For large data ranges:
- Avoid unnecessary full-workbook reads.
- Inspect headers and representative samples first.
- Use targeted ranges and efficient aggregation methods.
- Avoid repeated cell-by-cell operations when bulk operations are available.
- Preserve data types and row alignment.
- Verify counts and representative outputs.
Do not truncate the scope silently. If a tool imposes a limit, disclose it and use a suitable alternative if available.
22.4 Conflicting signals
If the user's wording conflicts with workbook structure, do not blindly follow either. For example, the user may say "monthly sales" while the date column contains text or the workbook has separate invoice and payment dates. Explain the conflict and resolve it through inspection or clarification.
22.5 Confidence and evidence
Distinguish:
- Known facts from workbook inspection.
- Inferences from column names or sample values.
- Assumptions made to proceed.
- Results confirmed by tool output.
- Results independently verified by reading the workbook again.
The agent should be transparent about uncertainty rather than inventing certainty.
23. Knowledge Use Instructions for the Excel Agent
Use this document as reasoning context, not as a command list or a set of fixed workflows.

When relevant knowledge is retrieved:
1. Interpret the user's actual request.
2. Inspect the available workbook context.
3. Apply the relevant decision principles.
4. Select the appropriate Excel operation and method.
5. Construct arguments from actual workbook data and available tool schemas.
6. Execute only the requested scope.
7. Verify the result.
8. Communicate the outcome accurately.
The examples and formula names in this document are illustrative knowledge. They are not mandatory templates. Adapt them to the actual workbook, Excel version, user intent, and available tools.

Tool descriptions and schemas define what operations are executable and what arguments they require. This document provides Excel-domain reasoning for deciding what should be done and how the operation should behave. The agent must combine both sources of information.

Never invent tool names, arguments, workbook values, sheet names, or results. Never claim a tool action occurred when it did not.
24. Final Reasoning Standard
A high-quality Excel agent does not merely recognize a keyword and copy a formula. It understands the requested outcome, inspects the workbook, reasons about the data and constraints, selects an appropriate method, adapts that method to the actual structure, executes safely, and verifies the result.

The goal is independent, context-aware Excel action execution across varied workbooks—not memorization of fixed examples.