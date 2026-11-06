# Request Variation

## 1. Purpose

This document defines how the Excel AI Agent handles variations in user requests.

Users may express the same intention using different words, sentence structures, abbreviations, informal language, or incomplete instructions.

The agent must recognize these variations and map them to the same intended operation without depending on exact keywords or predefined sentence patterns.

The objective is to provide consistent Excel functionality regardless of how the user phrases a request.

## 2. Types of Request Variations

The agent must recognize the following types of natural-language variations.

| Variation Type | Description |
|---|---|
| Synonyms | Different words expressing the same action |
| Alternative Phrasing | Different sentence structures with the same meaning |
| Informal Language | Casual or conversational expressions |
| Short Requests | Brief instructions with limited words |
| Incomplete Requests | Requests missing some information |
| Polite Requests | Instructions expressed as questions or polite statements |
| Direct Commands | Explicit action-oriented instructions |
| Abbreviations | Shortened words or commonly used abbreviations |
| Typographical Errors | Minor spelling mistakes |
| Rephrased Requests | Same operation expressed differently |
| Multilingual Expressions | Requests containing multiple languages |
| Follow-up Variations | Instructions that modify previous requests |

## 3. Synonym Recognition

The agent must recognize alternative words that represent the same operation.

### Examples

| User Expressions | Intended Operation |
|---|---|
| Show, display, get, retrieve, fetch | Read |
| Add, enter, insert, fill | Write |
| Change, modify, update, edit | Update |
| Remove, eliminate, clear | Delete or Clear, depending on context |
| Calculate, compute, work out | Calculate |
| Find, search, locate, look for | Search or Lookup |
| Arrange, organize, order | Sort |
| Highlight, emphasize, mark | Formatting |
| Create, generate, build | Create |
| Duplicate, copy, replicate | Copy |
| Rename, change the name | Rename |
| Explain, describe, tell me about | Explain |

Important:

- Consider the complete sentence.
- Do not assume that all synonyms always represent identical operations.
- Distinguish actions using their target and expected outcome.

## 4. Alternative Sentence Structures

The same request may be expressed using different grammatical structures.

Example: Calculate total sales.

Possible variations:

- Calculate total sales.
- What is the total sales amount?
- Can you find the sum of sales?
- Give me the overall sales value.
- How much did we sell in total?
- I need the total amount of sales.
- Please add all sales values.

All variations should map to the same calculation intention when the target data is clear.

Another example: Sort salary in descending order.

- Sort salary from highest to lowest.
- Arrange employees by salary in descending order.
- Show the highest-paid employees first.
- Put the largest salary values at the top.
- Order the salary column from maximum to minimum.

The agent must recognize equivalent meanings even when sentence structures differ.

## 5. Informal and Conversational Requests

Users may communicate using casual expressions instead of formal Excel terminology.

Examples:

| User Request | Intended Operation |
|---|---|
| Show me all the names | Read Name column |
| Get rid of duplicates | Remove Duplicates |
| Make this bold | Apply Bold Formatting |
| Put the highest value first | Sort Descending |
| Give me the average | Calculate Average |
| Find out who earns the most | Identify Maximum Salary |
| Clean this data | Identify required data-cleaning operations |
| Put everything in a new sheet | Create a worksheet and copy or move data, depending on context |

Rules:

- Understand conversational meaning.
- Do not require technical Excel terminology.
- Use context to identify the correct target.
- Ask for clarification when the requested action remains unclear.

## 6. Short and Abbreviated Requests

Users may provide very short instructions.

Examples:

| User Request | Interpretation |
|---|---|
| Total sales | Calculate total sales if context is sufficient |
| Avg salary | Calculate average salary |
| Sort A to Z | Sort ascending |
| Top 5 | Retrieve the top five records or values based on context |
| Remove duplicates | Remove duplicate records |
| Bold heading | Apply bold formatting |
| Sheet rename | Rename the relevant worksheet |
| Export CSV | Export the requested data to CSV |

Rules:

- Use conversation history and workbook structure.
- Resolve missing parameters from available context.
- Do not invent missing information.
- Ask only when essential information cannot be determined.

## 7. Polite and Indirect Requests

Users may express instructions as questions or polite requests.

Examples:

| User Request | Intended Operation |
|---|---|
| Can you calculate total sales? | Calculate |
| Could you highlight duplicate values? | Conditional Formatting |
| Would you please sort this data? | Sort |
| I want to see the employee records | Read |
| Is it possible to create a chart? | Create Chart |
| I need this column in uppercase | Text Transformation |

Rules:

- Recognize polite questions as potential action requests.
- Do not interpret every question as a request for explanation.
- Determine whether the user expects an actual workbook operation or only information.

## 8. Typographical and Spelling Variations

Users may make minor spelling mistakes or type quickly.

Examples:

| User Input | Intended Meaning |
|---|---|
| Calcluate total | Calculate total |
| Averge salary | Average salary |
| Hightlight duplicates | Highlight duplicates |
| Remve blank rows | Remove blank rows |
| Sort decending | Sort descending |
| Creat new sheet | Create new worksheet |

Rules:

- Correct obvious spelling errors internally.
- Preserve the intended meaning.
- Do not change meaningful values, identifiers, or cell references.
- Ask for clarification when a typing error creates multiple possible interpretations.

## 9. Multilingual and Mixed-Language Requests

Users may combine English with another language or use transliterated expressions.

Example Tamil-English variations:

| User Request | Intended Meaning |
|---|---|
| Total sales calculate pannu | Calculate total sales |
| Salary highest value sollu | Find highest salary |
| Duplicate rows remove pannu | Remove duplicate rows |
| Name column bold pannu | Make Name column bold |
| New sheet create pannu | Create a new worksheet |
| Chennai employees mattum show pannu | Filter Chennai employees |
| Data descending order-la arrange pannu | Sort descending |
| Average salary enna? | Calculate average salary |

Rules:

- Recognize supported multilingual expressions.
- Preserve the original values and identifiers.
- Interpret the intended operation rather than translating word by word.
- Handle mixed-language instructions consistently.
- Ask for clarification if the meaning cannot be determined reliably.

## 10. Variations in Conditions

Users may express the same logical condition in different ways.

### Greater Than

- Salary above 50000.
- Salary greater than 50000.
- Employees earning more than 50000.
- Salary exceeding 50000.

Interpretation: Salary > 50000.

### Less Than

- Salary below 30000.
- Salary less than 30000.
- Employees earning under 30000.

Interpretation: Salary < 30000.

### Between

- Salary between 30000 and 50000.
- Salary from 30000 to 50000.
- Employees earning within the 30000–50000 range.

Interpretation: Salary >= 30000 AND Salary <= 50000, unless the user specifies different boundary rules.

### Equality

- Department is Sales.
- Show only Sales employees.
- Employees belonging to Sales.
- Department must match Sales.

Interpretation: Department = Sales.

The agent must preserve comparison operators and boundary conditions accurately.

## 11. Variations in Sorting and Ranking Requests

Users may express sorting and ranking instructions differently.

| User Request | Interpretation |
|---|---|
| Highest to lowest | Descending |
| Largest to smallest | Descending |
| Maximum to minimum | Descending |
| Lowest to highest | Ascending |
| Smallest to largest | Ascending |
| A to Z | Text Ascending |
| Z to A | Text Descending |
| Top 10 values | Return 10 largest values |
| Bottom 5 values | Return 5 smallest values |
| Rank employees by salary | Rank based on salary |

Distinguish between sorting an entire dataset and retrieving selected top or bottom values.

## 12. Variations in Data Cleaning Requests

Users may describe data-cleaning operations without using technical terminology.

| User Request | Intended Operation |
|---|---|
| Remove repeated entries | Remove Duplicates |
| Delete duplicate employees | Remove Duplicate Records |
| Get rid of empty rows | Remove Blank Rows |
| Fix inconsistent capitalization | Standardize Text |
| Remove unwanted spaces | Trim Text |
| Clean the names column | Identify the required text-cleaning operation |
| Find missing values | Identify Blank or Missing Values |
| Replace N/A with zero | Replace Values |

Do not assume that a general request such as "clean this data" authorizes every possible cleaning operation.

Identify the specific requested transformation before modifying data.

## 13. Follow-Up Request Variations

Users may modify an existing request using short follow-up instructions.

Examples:

Conversation:

User: "Calculate total sales."

User: "Also average."

Interpretation: Calculate average sales using the same relevant data.

Conversation:

User: "Highlight salary above 50000."

User: "Make it 60000."

Interpretation: Change the threshold to 60000.

Conversation:

User: "Sort by salary."

User: "No, by age."

Interpretation: Replace Salary with Age as the sorting column.

Conversation:

User: "Create a new sheet."

User: "Name it Summary."

Interpretation: Rename the newly created worksheet to Summary.

Rules:

- Retain relevant previous parameters.
- Recognize additions and corrections.
- Replace only the parameters that the user changes.
- Do not repeat already completed operations unnecessarily.
- Do not treat short follow-ups as independent requests when they clearly depend on previous context.

## 14. Ambiguous Variation Handling

Different expressions may represent different operations depending on context.

Example:

User: "Remove the values."

Possible interpretations:

- Clear cell contents.
- Delete cells.
- Remove records containing specific values.

Example:

User: "Give me the top employees."

Possible interpretations:

- Highest salary.
- Highest performance score.
- Longest experience.

Example:

User: "Make it clean."

Possible interpretations:

- Remove duplicates.
- Remove blank rows.
- Standardize formatting.
- Clean text values.

Rules:

- Use available workbook and conversation context.
- Do not force an interpretation based on a single phrase.
- Ask a focused clarification question when necessary.
- Avoid executing ambiguous destructive operations.

## 15. Request Normalization

After recognizing a request variation, convert it into a consistent internal representation.

Normalization should include:

- Standardized action.
- Identified target.
- Extracted conditions.
- Identified scope.
- Expected output.
- Relevant parameters.

Example:

Original request:

"Can you get rid of duplicate employee entries?"

Normalized interpretation:

- Action: Remove Duplicates
- Target: Employee Dataset
- Scope: Relevant Employee Records
- Output: Dataset without duplicate records

Another example:

Original request:

"Ennoda salary column-la highest value sollu."

Normalized interpretation:

- Action: Calculate Maximum
- Target: Salary Column
- Output: Highest Salary Value

Important:

Normalization must preserve the original meaning. It must not introduce additional operations or change user constraints.

## 16. Handling Unknown Expressions

When the agent encounters an unfamiliar expression:

1. Analyze the complete sentence.
2. Identify recognizable action and target references.
3. Consider relevant conversation history.
4. Inspect workbook context when necessary.
5. Determine whether a reliable interpretation is possible.
6. Ask for clarification when the meaning remains uncertain.

Do not map an unknown expression to an unrelated operation simply because it contains a familiar keyword.

## 17. Request Variation Recognition Workflow

Follow this sequence:

1. Receive the original user request.
2. Identify the language and sentence structure.
3. Detect synonyms and alternative expressions.
4. Recognize informal or abbreviated instructions.
5. Identify possible spelling variations.
6. Extract the intended action.
7. Resolve the target and conditions.
8. Compare the request with relevant previous context.
9. Normalize the recognized request.
10. Check for ambiguity.
11. Pass the normalized interpretation to intent recognition and context-processing stages.

## 18. Important Rules

- Support multiple ways of expressing the same intention.
- Do not depend on exact keywords.
- Recognize synonyms and alternative sentence structures.
- Understand informal and conversational language.
- Handle short and incomplete requests.
- Recognize obvious spelling mistakes.
- Support multilingual and mixed-language expressions.
- Preserve numerical values and cell references.
- Preserve logical conditions.
- Distinguish similar operations.
- Use conversation history for follow-up requests.
- Normalize without changing the original meaning.
- Ask for clarification when necessary.
- Never execute an operation solely because a phrase resembles a known command.

## 19. Final Request Variation Principle

The Excel AI Agent must provide consistent operation recognition regardless of how users express their requests.

Different expressions with the same meaning should produce the same normalized interpretation, while expressions with different intended outcomes must remain distinguishable.

The agent must prioritize semantic understanding over exact phrase matching.