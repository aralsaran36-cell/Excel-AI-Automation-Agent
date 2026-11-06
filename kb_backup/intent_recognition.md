# Intent Recognition

## 1. Purpose

This document defines how the Excel AI Agent recognizes and classifies user intentions from natural-language requests.

The agent must identify the actual operation requested by the user, determine the relevant action category, extract operation-specific parameters, and distinguish between similar or overlapping intents.

Intent recognition must consider the complete request rather than relying only on individual keywords.

## 2. Intent Recognition Fundamentals

Every user request should be analyzed using the following elements:

1. Action – What does the user want to do?
2. Target – What workbook, worksheet, data, cell, or range is involved?
3. Condition – Are there any conditions or restrictions?
4. Scope – How much data or how many objects are affected?
5. Output – What result does the user expect?
6. Execution Requirement – Does the request require reading, calculation, modification, or another operation?

Example:

User request:

"Highlight employees whose salary is above 50000."

Recognized intent:

- Action: Highlight
- Target: Employee salary records
- Condition: Salary > 50000
- Scope: Matching employee records
- Output: Highlight matching cells or records
- Operation Category: Formatting

## 3. Primary Intent Categories

The agent must classify requests into the appropriate primary intent category.

| Intent | Description | Example |
|---|---|---|
| Read | Retrieve workbook information or cell values | Read A1:D10 |
| Write | Insert new values | Enter 50000 in B5 |
| Update | Modify existing values | Change Arun's salary to 60000 |
| Delete | Remove cells, rows, columns, or worksheets | Delete row 5 |
| Clear | Remove cell contents or formatting | Clear A1:A10 |
| Calculate | Perform mathematical calculations | Calculate total sales |
| Formula | Generate or apply Excel formulas | Apply SUM formula |
| Lookup | Find related information | Find salary using employee ID |
| Filter | Display matching records | Show Sales department employees |
| Sort | Arrange records | Sort salary descending |
| Format | Change cell appearance | Make headings bold |
| Conditional Formatting | Apply formatting based on conditions | Highlight values above 50000 |
| Data Cleaning | Clean or standardize data | Remove duplicate records |
| Chart | Create or modify charts | Create a sales chart |
| Table | Create or modify Excel tables | Convert data into a table |
| Worksheet | Manage worksheets | Rename a sheet |
| Workbook | Manage workbook-level operations | Save the workbook |
| Export | Export or generate external files | Export data to CSV |
| Explain | Provide information or instructions | Explain the SUM formula |

The agent must select the category that best represents the requested action.

## 4. Action Recognition

Recognize action words and alternative expressions used by users.

| User Expressions | Recognized Action |
|---|---|
| Show, display, get, read, fetch | Read |
| Add, enter, insert, fill | Write |
| Change, modify, replace, edit | Update |
| Remove, delete, eliminate | Delete or Clear, depending on context |
| Calculate, compute, find total | Calculate |
| Apply formula, insert formula | Formula |
| Search, find, look up | Lookup or Search |
| Arrange, order, rank | Sort or Ranking |
| Filter, show only, display matching | Filter |
| Highlight, color, emphasize | Formatting |
| Make bold, italic, underline | Formatting |
| Create, generate, build | Create |
| Rename, change name | Rename |
| Copy, duplicate | Copy |
| Move, transfer | Move |
| Explain, describe, teach | Explain |

Important:

- Action words provide clues, not final decisions.
- Interpret the complete sentence.
- Distinguish similar actions using their intended outcomes.
- Do not select an operation based on keyword matching alone.

## 5. Distinguishing Similar Intents

Some requests contain similar expressions but require different operations.

### 5.1 Delete vs Clear

| User Request | Recognized Intent |
|---|---|
| Clear values in A1:A10 | Clear Contents |
| Remove formatting from A1:A10 | Clear Formats |
| Delete cells A1:A10 | Delete Cells |
| Delete row 5 | Delete Row |
| Remove duplicate records | Data Cleaning |

### 5.2 Calculate vs Formula

| User Request | Recognized Intent |
|---|---|
| What is the total sales? | Calculate and return result |
| Add SUM formula in E2 | Formula insertion |
| Calculate total and show in E2 | Calculate and write result |
| Explain how SUM works | Explain |

### 5.3 Search vs Filter

| User Request | Recognized Intent |
|---|---|
| Find employee Arun | Search |
| Show all employees in Sales | Filter |
| Find the salary of employee 101 | Lookup |
| Display only employees earning above 50000 | Filter |

### 5.4 Sort vs Ranking

| User Request | Recognized Intent |
|---|---|
| Arrange salary from highest to lowest | Sort |
| Find the highest salary | MAX calculation |
| Find the top five salaries | Ranking / LARGE |
| Give each employee a salary rank | Ranking |

### 5.5 Write vs Update

| User Request | Recognized Intent |
|---|---|
| Enter a new employee record | Write |
| Change Arun's existing salary | Update |
| Add a new row at the bottom | Append / Write |
| Replace all existing values | Update or Replace |

## 6. Multi-Intent Recognition

A single request may contain multiple independent or dependent operations.

Example:

"Remove duplicates, calculate total sales, and create a summary chart."

Recognized intents:

1. Data Cleaning – Remove duplicates.
2. Calculation – Calculate total sales.
3. Chart Creation – Create a summary chart.

Another example:

"Filter Chennai employees, sort them by salary, and highlight salaries above 50000."

Recognized intents:

1. Filter records by location.
2. Sort filtered records by salary.
3. Apply conditional formatting.

Rules:

- Recognize every requested action.
- Separate independent operations.
- Identify dependencies between operations.
- Preserve the user's requested execution order.
- Do not omit secondary operations.
- Avoid generating additional operations that were not requested.

## 7. Intent Priority and Conflict Resolution

When a request contains multiple actions, determine their relationship.

Possible relationships:

| Relationship | Description |
|---|---|
| Independent | Operations can be performed separately |
| Sequential | One operation depends on another |
| Conditional | An operation executes only when a condition is satisfied |
| Alternative | User specifies one of multiple possible actions |
| Conflicting | Two instructions cannot be satisfied together |

Example:

"Remove duplicates and then calculate total sales."

Relationship: Sequential.

The duplicate-removal operation must be completed before calculating the total if the intended total excludes duplicates.

Example:

"Sort the data by salary ascending and descending."

Relationship: Conflicting if both directions are requested as the final state of the same range.

Ask which sorting direction should be applied.

Rules:

- Identify operation dependencies.
- Resolve conflicting instructions before execution.
- Do not arbitrarily discard an instruction.
- Respect explicit ordering words such as before, after, first, and then.

## 8. Conditional Intent Recognition

Identify conditions that determine which records or cells are affected.

Recognize:

### Comparison Conditions

- Greater than
- Less than
- Equal to
- Not equal to
- Greater than or equal to
- Less than or equal to

### Range Conditions

- Between two values
- Within a date range
- Above a specified threshold
- Below a specified threshold

### Text Conditions

- Contains specific text
- Begins with specific text
- Ends with specific text
- Matches exact text

### Logical Conditions

- AND
- OR
- NOT

Examples:

| User Request | Recognized Condition |
|---|---|
| Salary above 50000 | Salary > 50000 |
| Age between 20 and 30 | Age >= 20 AND Age <= 30 |
| Department is Sales or HR | Department = Sales OR HR |
| Names starting with A | Name begins with A |
| Sales above 10000 in Chennai | Sales > 10000 AND City = Chennai |

Do not change the logical meaning of user conditions.

## 9. Scope Recognition

Identify the amount of data affected by the requested operation.

Possible scopes:

- Single cell
- Specific range
- Entire column
- Entire row
- Complete data table
- Selected records
- Entire worksheet
- Entire workbook
- Multiple worksheets

Examples:

| User Request | Scope |
|---|---|
| Change A1 | Single Cell |
| Format A1:D10 | Specific Range |
| Highlight the Salary column | Entire Relevant Column |
| Delete row 5 | Entire Row |
| Remove all duplicate employees | Complete Relevant Dataset |
| Rename the January sheet | Single Worksheet |
| Apply formatting to every sheet | Multiple Worksheets |

Rules:

- Identify the smallest scope that satisfies the request.
- Do not expand the operation unnecessarily.
- Verify the scope before destructive operations.
- Do not interpret an entire-workbook request as an active-sheet request.

## 10. Output Intent Recognition

Determine whether the user expects a conversational response or an actual workbook modification.

Classify the expected output as:

| Output Intent | Example |
|---|---|
| Return Value | Tell me total sales |
| Display Data | Show all employee records |
| Modify Workbook | Enter the total in E2 |
| Apply Formatting | Highlight duplicate values |
| Create Object | Create a chart |
| Generate File | Export to CSV |
| Explain | Explain how VLOOKUP works |
| Confirm Status | Tell me whether the workbook was saved |

Important:

- Do not modify the workbook when the user only asks for an explanation.
- Do not return only a conversational answer when the user explicitly requests a workbook modification.
- Identify whether the result must be saved or exported.

## 11. Conversational Intent Recognition

Users may express follow-up requests that modify previously recognized intentions.

Examples:

Conversation:

User: "Calculate total sales."

User: "Also find the average."

Recognized intent:

Add an average calculation while retaining the relevant sales context.

Conversation:

User: "Highlight salaries above 50000."

User: "Actually make it 60000."

Recognized intent:

Update the threshold from 50000 to 60000.

Conversation:

User: "Sort employees by salary."

User: "Descending."

Recognized intent:

Apply descending sorting to the previously identified salary column.

Rules:

- Use relevant previous intent.
- Recognize additions, corrections, and replacements.
- Preserve unaffected parameters.
- Do not execute outdated instructions after a correction.
- Do not treat every follow-up as a completely new task.

## 12. Unsupported Intent Recognition

The agent must identify requests that cannot be fulfilled using currently available capabilities.

Examples:

- Create an unsupported Excel object.
- Execute an unavailable external integration.
- Perform an operation requiring an unavailable tool.
- Access a protected workbook without permission.

Rules:

1. Identify the requested operation.
2. Check whether the required capability exists.
3. Determine whether a supported alternative is available.
4. Do not invent tool capabilities.
5. Explain the limitation clearly.
6. Ask for additional information or authorization when necessary.

Recognizing an intent does not mean the agent is capable of executing it.

## 13. Ambiguous Intent Handling

When the request does not provide enough information to determine the intended operation, ask a clarification question.

Examples:

User: "Remove the data."

Possible intents:

- Clear cell contents.
- Delete selected rows.
- Delete an entire worksheet.

User: "Make it better."

Possible intents:

- Improve formatting.
- Clean the data.
- Improve the chart.
- Modify formulas.

User: "Find the result."

Possible intents:

- Calculate a total.
- Find a maximum value.
- Perform a lookup.

Rules:

- First use available context to resolve ambiguity.
- Ask only when multiple interpretations remain possible.
- Ask one focused question.
- Do not execute destructive or irreversible operations based on assumptions.

## 14. Intent Recognition Output Structure

After analyzing the request, produce a structured internal representation for the next processing stage.

Recommended fields:

| Field | Description |
|---|---|
| Primary Intent | Main requested operation |
| Secondary Intents | Additional requested operations |
| Action | Specific action to perform |
| Target | Workbook, worksheet, data, or object |
| Conditions | Required conditions |
| Scope | Affected data boundaries |
| Output Type | Expected result |
| Dependencies | Operation ordering requirements |
| Constraints | User-specified restrictions |
| Ambiguity | Missing or unclear information |
| Execution Requirement | Read, write, modify, calculate, etc. |
| Confidence | Confidence in the interpretation |

Example:

User request:

"Find employees in Chennai earning above 50000 and sort them by salary."

Recognized representation:

- Primary Intent: Filter
- Secondary Intent: Sort
- Action: Filter and sort
- Target: Employee dataset
- Conditions: City = Chennai AND Salary > 50000
- Scope: Matching employee records
- Output Type: Display filtered and sorted records
- Dependencies: Filtering before sorting
- Ambiguity: None if the relevant columns are identified
- Execution Requirement: Read and organize data

The structured representation is an internal processing result and does not need to be displayed to the user unless requested.

## 15. Intent Recognition Workflow

Follow these steps for every user request.

1. Read the complete user message.
2. Identify the main action.
3. Recognize additional actions.
4. Classify each action.
5. Identify the target object.
6. Extract conditions and constraints.
7. Determine the affected scope.
8. Identify the expected output.
9. Distinguish similar operations.
10. Identify dependencies and conflicts.
11. Consider relevant conversational context.
12. Check whether the intent is sufficiently clear.
13. Produce the structured intent representation.
14. Pass the recognized intent to the appropriate next processing stage.

## 16. Important Rules

- Never classify intent using keywords alone.
- Understand the complete sentence.
- Identify primary and secondary actions.
- Distinguish similar operations.
- Recognize multiple operations in one request.
- Preserve conditions and logical relationships.
- Identify the correct operation scope.
- Recognize conversational corrections.
- Identify expected output.
- Detect unsupported operations.
- Resolve ambiguity before execution.
- Do not invent additional user requirements.
- Do not execute operations during intent recognition.
- Do not confuse intent recognition with tool selection.

## 17. Final Intent Recognition Principle

The Excel AI Agent must accurately determine what action the user expects before selecting an execution method.

Intent recognition is responsible for identifying the requested operation and its requirements.

It must not independently decide which Excel tool to execute or generate the final execution parameters.

The recognized intent should be passed to the appropriate downstream processing stages for operation selection, parameter identification, validation, and execution.