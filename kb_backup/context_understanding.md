# Context Understanding

## 1. Purpose

This document defines how the Excel AI Agent understands and interprets user requests before selecting tools or performing Excel operations.

The agent must understand the complete conversational context, identify the user's actual intention, resolve references, recognize relevant workbook information, and determine the expected outcome.

The agent must not rely only on individual keywords. It must interpret the complete request using available conversation context and workbook metadata.

## 2. Understanding Natural Language Requests

Users may express the same Excel operation in different ways.

The agent must recognize different expressions that represent the same intention.

Examples:

| User Request | Intended Meaning |
|---|---|
| Find total sales | Calculate the sum of sales |
| Add all salary values | Calculate the sum of salary |
| Give me the average marks | Calculate average marks |
| Remove duplicate records | Identify and remove duplicate rows |
| Make the heading bold | Apply bold formatting to the heading |
| Arrange sales from highest to lowest | Sort sales in descending order |
| Highlight values above 50000 | Apply conditional formatting or direct highlighting |
| Find employees in Chennai | Filter records based on location |
| Create a new sheet | Insert a new worksheet |

The agent must identify the requested operation even when the user uses informal language, incomplete sentences, or alternative expressions.

## 3. Intent Identification

Before execution, identify the primary intent of the request.

Classify the request into an appropriate operational category.

| Intent Category | Description |
|---|---|
| Data Reading | Retrieve values or information |
| Data Entry | Insert new values |
| Data Modification | Update existing values |
| Data Cleaning | Remove duplicates, blanks, or unwanted data |
| Data Analysis | Calculate or analyze data |
| Formula Operation | Generate or apply Excel formulas |
| Formatting | Change cell appearance |
| Sorting | Arrange data in a specified order |
| Filtering | Display records matching conditions |
| Cell Range Operation | Perform operations on selected cells or ranges |
| Worksheet Operation | Create, rename, delete, or manage worksheets |
| Workbook Operation | Perform workbook-level operations |
| Chart Operation | Create or modify charts |
| Table Operation | Create or modify Excel tables |

A request may contain multiple intents. Identify all requested operations rather than selecting only the first one.

## 4. Workbook Context Understanding

The agent must understand the workbook currently involved in the request.

Identify:

- Available workbook files.
- Currently active workbook.
- Workbook names and paths.
- Workbook format.
- Available worksheets.
- Workbook metadata.
- Relevant data structures.

Rules:

- Use actual workbook information.
- Never assume a workbook name.
- Do not select an unrelated workbook.
- If multiple workbooks are available, resolve the intended workbook from conversation context.
- Ask the user when the target workbook cannot be determined reliably.

## 5. Worksheet Context Understanding

Identify the worksheet relevant to the user's request.

Consider:

- Active worksheet.
- Worksheet names.
- Available worksheet data.
- Column headers.
- Data boundaries.
- Existing tables.
- Relevant cell ranges.

Examples:

User request:

"Calculate total sales."

If multiple worksheets contain sales information, inspect the workbook structure and determine which worksheet is relevant.

User request:

"Calculate total sales in the January sheet."

Use the explicitly specified worksheet after verifying that it exists.

Rules:

- Never assume a worksheet name.
- Do not automatically use the active worksheet when another worksheet is explicitly mentioned.
- Use conversational context to resolve references such as "that sheet" or "previous sheet".
- Ask for clarification if multiple worksheets are equally relevant.

## 6. Column and Header Understanding

Users frequently refer to columns using business terms instead of Excel column letters.

The agent must identify the correct column using actual headers.

Example:

| A | B | C | D |
|---|---|---|---|
| Employee ID | Employee Name | Department | Salary |
| 101 | Arun | Sales | 30000 |
| 102 | Kumar | HR | 40000 |

Possible user expressions:

- Employee names
- Staff names
- Salary column
- Department details
- Employee ID field

The agent must map these expressions to the appropriate workbook columns.

Rules:

- Inspect actual headers.
- Consider meaningful synonyms.
- Use the surrounding request to determine the intended column.
- Avoid matching unrelated columns based only on similar words.
- Ask for clarification when multiple columns could match.

## 7. Conversational Context and Reference Resolution

The agent must maintain relevant context across consecutive user messages.

Users may refer to previously mentioned workbooks, worksheets, columns, ranges, or operations without repeating their names.

Examples:

Conversation:

User: "Calculate total sales in the Sales sheet."

User: "Now find the average."

Interpretation:

Calculate the average of the previously discussed sales data in the same worksheet, unless the new request indicates a different target.

Other reference expressions:

| User Expression | Contextual Interpretation |
|---|---|
| That sheet | Previously identified worksheet |
| Same column | Previously identified column |
| Those records | Previously identified records |
| Do the same | Repeat the previous operation with applicable context |
| Now sort them | Sort the previously identified data |
| Also calculate average | Add an average calculation to the current task |
| Change it to descending | Modify the previous sorting direction |

Rules:

- Preserve relevant context from previous messages.
- Apply only the changes introduced by the latest request.
- Do not unnecessarily ask users to repeat known information.
- Do not carry forward unrelated context.
- Resolve ambiguous references before execution.

## 8. Multi-Intent Request Understanding

A single request may contain multiple operations.

Example:

"Calculate total sales, find average sales, and highlight values above 50000."

Required interpretation:

1. Calculate total sales.
2. Calculate average sales.
3. Highlight sales values above 50000.

Another example:

"Remove duplicates, sort employees by salary, and create a new worksheet."

Required interpretation:

1. Remove duplicate records.
2. Sort employees by salary.
3. Create a new worksheet.

Rules:

- Identify every requested operation.
- Determine dependencies between operations.
- Preserve the intended execution order.
- Avoid unrelated operations.
- Track the execution status of each operation independently.
- Report partial success accurately.

## 9. Conditions and Constraint Understanding

The agent must identify conditions and constraints expressed in the request.

Examples:

| User Request | Identified Condition |
|---|---|
| Employees with salary above 50000 | Salary > 50000 |
| Sales from Chennai | City = Chennai |
| Employees aged between 25 and 35 | Age >= 25 AND Age <= 35 |
| Products with stock below 10 | Stock < 10 |
| Sales from January | Date belongs to January |
| Top five highest salaries | Descending salary ranking, limit 5 |

Understand:

- Greater than and less than conditions.
- Equal and not equal conditions.
- Between conditions.
- Multiple AND conditions.
- Multiple OR conditions.
- Date and time conditions.
- Text matching conditions.
- Top and bottom record requirements.

Do not ignore conditions or change their meaning.

## 10. Output Expectation Understanding

Identify what result the user expects.

Possible output types:

| User Expectation | Required Output |
|---|---|
| Tell me the total | Return the calculated value |
| Show the records | Display matching records |
| Add the result to Excel | Write the result into the workbook |
| Create a summary | Generate an appropriate summary |
| Highlight the cells | Apply formatting to the relevant cells |
| Create a chart | Generate a chart |
| Export the data | Save or export using an available tool |
| Explain the result | Provide a natural-language explanation |

Important distinctions:

- Calculating a value does not automatically mean writing it into Excel.
- Reading data does not automatically mean modifying the workbook.
- Showing a result in the conversation is different from inserting it into a worksheet.
- Do not create additional output unless requested or necessary.

## 11. Ambiguity Detection

The agent must detect requests that have multiple possible interpretations.

Examples:

**Example 1**

User: "Remove the old data."

Possible interpretations:

- Clear previous records.
- Delete an old worksheet.
- Remove records before a specific date.

Action: Ask which data should be removed.

**Example 2**

User: "Calculate the total."

Possible interpretations:

- Total sales.
- Total salary.
- Total quantity.

Action: Inspect context and available columns. Ask if the intended field remains unclear.

**Example 3**

User: "Sort the data."

Possible interpretations:

- Sort by name.
- Sort by salary.
- Sort by date.

Action: Ask which column should determine the sorting order if context does not resolve it.

**Example 4**

User: "Update the sheet."

Possible interpretations:

- Update cell values.
- Update formatting.
- Add new records.

Action: Ask what should be updated.

Rules:

- Do not guess when ambiguity affects correctness.
- Use existing context to resolve minor ambiguity.
- Ask one focused clarification question when necessary.
- Do not ask unnecessary questions when the intended operation is clear.

## 12. Handling Incomplete Requests

Users may provide incomplete instructions.

Examples:

- "Find average."
- "Highlight duplicates."
- "Sort by salary."
- "Calculate percentage."
- "Create chart."

The agent must determine whether the available context is sufficient.

If sufficient:

Proceed using the identified workbook structure and context.

If insufficient:

Ask only for the missing information necessary to perform the operation.

Do not request information already available in workbook metadata or conversation history.

## 13. Understanding Corrections and Follow-up Instructions

The agent must recognize when a user corrects or modifies a previous request.

Examples:

User: "Sort by salary."

User: "Sorry, sort by age instead."

Interpretation:

Replace the sorting column with Age.

User: "Highlight values above 50000."

User: "Make it 60000."

Interpretation:

Update the highlighting condition to values above 60000.

Rules:

- Apply the latest explicit correction.
- Retain unaffected parts of the original request.
- Do not execute outdated instructions after they have been replaced.
- Re-evaluate the operation before execution when the correction changes its meaning.

## 14. Context Understanding Decision Workflow

Follow this sequence before selecting an execution tool.

1. Read the complete user request.
2. Review relevant conversation history.
3. Identify the primary intent.
4. Identify additional requested operations.
5. Resolve workbook references.
6. Resolve worksheet references.
7. Identify relevant columns and data ranges.
8. Extract conditions and constraints.
9. Determine the expected output.
10. Detect missing information or ambiguity.
11. Resolve references using available workbook metadata.
12. Ask for clarification if essential information is unavailable.
13. Pass the structured interpretation to the appropriate next stage.

## 15. Important Rules

- Understand the complete request, not individual keywords.
- Never assume user intent.
- Use conversation history when relevant.
- Use actual workbook and worksheet metadata.
- Resolve column references using real headers.
- Recognize synonyms and alternative expressions.
- Identify all requested operations.
- Preserve conditions and constraints.
- Distinguish reading, calculating, and modifying data.
- Identify the expected output.
- Detect ambiguous instructions.
- Do not ask unnecessary questions.
- Do not execute operations based on unsupported assumptions.
- Apply user corrections immediately.
- Maintain relevant context across follow-up requests.

## 16. Final Interpretation Principle

Before performing any Excel operation, the agent must establish a clear understanding of:

1. What the user wants.
2. Which workbook is involved.
3. Which worksheet is involved.
4. Which data or range is targeted.
5. What conditions must be applied.
6. What result the user expects.

Only after establishing this understanding should the request proceed to the appropriate operation-selection and execution stages.

The agent must prioritize correct interpretation over immediate execution.