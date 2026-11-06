# Worksheet Operations

## 1. Purpose

This document defines how the Excel AI Agent should understand, select, and execute worksheet-related operations.

The agent must identify the correct workbook and worksheet before performing any worksheet modification.

All operations must use the actual workbook metadata and available MCP execution tools.

## 2. Worksheet Operation Categories

Worksheet operations are divided into the following categories:

1. Worksheet identification.
2. Worksheet information retrieval.
3. Worksheet creation.
4. Worksheet deletion.
5. Worksheet renaming.
6. Worksheet selection.
7. Worksheet copying.
8. Worksheet movement.
9. Worksheet visibility management.
10. Worksheet protection.

Select only the operation required by the user's request.

## 3. Worksheet Identification

Before executing an operation, identify the target worksheet.

Determine:

- Target workbook.
- Existing worksheet names.
- Active worksheet.
- Requested worksheet name.
- Whether the worksheet exists.
- Whether the operation requires an existing or new worksheet.

Examples:

| User Request | Required Action |
|---|---|
| Show all sheet names | Retrieve worksheet names |
| Which sheet is active? | Retrieve active worksheet |
| Check whether Sales exists | Verify worksheet existence |
| Show worksheet information | Retrieve worksheet metadata |

Never assume worksheet names or positions.

## 4. Worksheet Creation

Use the appropriate worksheet creation tool when the user requests a new worksheet.

Example:

User request:

"Create a worksheet named Sales Report."

Execution requirements:

1. Identify the target workbook.
2. Extract the requested worksheet name.
3. Check whether the name already exists.
4. Validate the requested name.
5. Select the worksheet creation tool.
6. Map the workbook and worksheet name to the actual tool parameters.
7. Execute the operation.
8. Verify that the new worksheet exists.

Important rules:

- Do not create duplicate worksheets unintentionally.
- Do not rename an existing worksheet as a substitute for creation.
- Preserve existing worksheets.

## 5. Worksheet Deletion

Use the appropriate deletion tool only when the user explicitly requests worksheet removal.

Example:

User request:

"Delete the Old Data worksheet."

Execution requirements:

1. Identify the target workbook.
2. Verify that Old Data exists.
3. Identify the exact worksheet to delete.
4. Check whether the operation is destructive.
5. Confirm with the user when required by the application's safety policy.
6. Execute the deletion.
7. Verify that the worksheet no longer exists.

Important rules:

- Never delete a worksheet based on an assumed name.
- Never delete multiple worksheets when only one is requested.
- Do not delete a worksheet containing important data without authorization.

## 6. Worksheet Renaming

Use the appropriate renaming tool when the user requests a worksheet name change.

Example:

User request:

"Rename Sheet1 to Employee Data."

Execution requirements:

1. Identify the target workbook.
2. Verify that Sheet1 exists.
3. Extract the new worksheet name.
4. Validate the new name.
5. Check for duplicate worksheet names.
6. Execute the renaming operation.
7. Verify the updated worksheet name.

Important rules:

- Preserve all existing worksheet data.
- Do not create a new worksheet.
- Do not modify unrelated worksheet names.

## 7. Worksheet Selection

Use the appropriate worksheet selection tool when the user requests switching to another worksheet.

Example:

User request:

"Open the Sales worksheet."

Execution requirements:

1. Identify the target workbook.
2. Verify that the Sales worksheet exists.
3. Select the appropriate worksheet.
4. Verify the active worksheet.

Do not confuse worksheet selection with worksheet creation.

## 8. Worksheet Copying

Use the appropriate copying tool when the user requests a duplicate worksheet.

Example:

User request:

"Create a copy of the Sales worksheet named Sales Backup."

Execution requirements:

1. Identify the source worksheet.
2. Verify that the source exists.
3. Extract the destination worksheet name.
4. Check whether the destination already exists.
5. Select the worksheet copying tool.
6. Map the source and destination parameters.
7. Execute the operation.
8. Verify the copied worksheet.

Preserve the original worksheet.

## 9. Worksheet Movement

Use the appropriate movement tool when the user requests changing worksheet order.

Example:

User request:

"Move the Sales worksheet to the first position."

Execution requirements:

1. Identify the target workbook.
2. Verify that the Sales worksheet exists.
3. Identify the requested destination position.
4. Check the supported movement parameters.
5. Execute the operation.
6. Verify the new worksheet order.

Do not modify worksheet names or contents during movement.

## 10. Worksheet Visibility

Use the appropriate visibility tool when the user requests hiding or showing worksheets.

Examples:

- Hide the Sales worksheet.
- Unhide the Employee worksheet.
- Show all hidden worksheets.

Execution requirements:

1. Identify the target workbook.
2. Identify the target worksheet.
3. Determine the requested visibility state.
4. Check whether the worksheet exists.
5. Execute the appropriate operation.
6. Verify the visibility state.

Do not confuse hiding a worksheet with deleting it.

## 11. Worksheet Protection

Use the appropriate protection tool when the user requests protecting or unprotecting a worksheet.

Examples:

- Protect the Sales worksheet.
- Remove worksheet protection.

Execution requirements:

1. Identify the target workbook.
2. Identify the target worksheet.
3. Determine the requested protection operation.
4. Identify required password parameters, if supported.
5. Validate the available tool schema.
6. Execute the operation.
7. Verify the protection status when supported.

Never invent or assume worksheet passwords.

## 12. Multiple Worksheet Operations

Some requests require multiple worksheet operations.

Example:

User request:

"Create a Backup worksheet, copy the Sales worksheet into it, and hide the original Sales worksheet."

Execution sequence:

1. Identify the workbook.
2. Verify that Sales exists.
3. Create the Backup worksheet if required by the requested workflow.
4. Copy the required worksheet data or worksheet according to the requested meaning.
5. Verify the backup.
6. Hide the original Sales worksheet.
7. Verify both operations.

Execute dependent operations in the correct order.

Do not continue destructive operations if an essential preceding operation fails.

## 13. Worksheet Validation

Before executing any worksheet operation, verify:

- Correct workbook.
- Correct worksheet.
- Valid worksheet name.
- Valid destination name, when applicable.
- Worksheet existence.
- Duplicate-name conflicts.
- Operation compatibility.
- Required parameters.
- Destructive operation authorization.

## 14. Execution Recovery

If a worksheet operation fails:

1. Capture the actual error.
2. Identify the failed operation.
3. Check the current workbook state.
4. Determine whether the worksheet was partially modified.
5. Identify whether the issue is related to the worksheet name, workbook, tool parameters, or execution.
6. Select an appropriate recovery action.
7. Retry only if safe.
8. Verify the final worksheet state.

Do not repeat creation, deletion, copying, or movement blindly.

## 15. Important Rules

- Never assume a worksheet name.
- Never assume worksheet position.
- Never assume the active workbook.
- Never delete without proper authorization.
- Never overwrite an existing worksheet unintentionally.
- Never confuse hiding with deletion.
- Never confuse copying with renaming.
- Never modify worksheet contents during a name or position change.
- Always inspect actual workbook metadata.
- Always verify worksheet operations after execution.

## 16. Final Execution Principle

Every worksheet operation must follow this sequence:

**Understand Request → Identify Workbook → Identify Worksheet → Select Tool → Map Parameters → Validate → Execute → Verify → Respond**

The final response must accurately describe the operation performed and its actual execution status.