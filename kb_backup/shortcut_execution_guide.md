# Shortcut Execution Guide

## 1. Purpose

This document defines how the Excel AI Agent identifies, selects, validates, and executes Excel keyboard shortcuts.

The agent must understand shortcut-related user requests, identify the intended keyboard action, determine the applicable shortcut based on the operating system and Excel environment, and execute it using an available shortcut or keyboard interaction tool.

The agent must not assume that every keyboard shortcut works identically across different operating systems or Excel environments.

## 2. Shortcut Request Recognition

Users may request shortcuts in different ways.

Examples:

| User Request | Intended Action |
|---|---|
| Save the workbook using shortcut | Save Workbook |
| Give me the shortcut for bold | Identify Bold Shortcut |
| Apply bold using Ctrl+B | Execute Bold Shortcut |
| Select the entire column using keyboard | Select Entire Column |
| Undo the last action | Undo |
| Open Find using shortcut | Open Find |
| Use keyboard shortcut to insert a row | Insert Row |
| Shortcut for creating a new sheet | Identify New Worksheet Shortcut |

Distinguish between:

- Asking for a shortcut explanation.
- Asking to execute an operation using a shortcut.
- Asking to identify an appropriate shortcut.
- Providing a specific shortcut for execution.

Do not execute a shortcut when the user only asks what it means.

## 3. Shortcut Categories

Recognize shortcuts based on their operational purpose.

### 3.1 File Operations

| Operation | Windows Shortcut | macOS Shortcut |
|---|---|---|
| Save Workbook | Ctrl + S | Command + S |
| Save As | F12 | Command + Shift + S |
| Open Workbook | Ctrl + O | Command + O |
| Create New Workbook | Ctrl + N | Command + N |
| Print | Ctrl + P | Command + P |
| Close Workbook | Ctrl + W | Command + W |

### 3.2 Editing Operations

| Operation | Windows Shortcut | macOS Shortcut |
|---|---|---|
| Copy | Ctrl + C | Command + C |
| Cut | Ctrl + X | Command + X |
| Paste | Ctrl + V | Command + V |
| Undo | Ctrl + Z | Command + Z |
| Redo | Ctrl + Y | Command + Y |
| Find | Ctrl + F | Command + F |
| Replace | Ctrl + H | Control + H |
| Select All | Ctrl + A | Command + A |
| Delete Cell Contents | Delete | Delete |
| Insert Cells | Ctrl + Shift + + | Control + Shift + + |

### 3.3 Formatting Operations

| Operation | Windows Shortcut | macOS Shortcut |
|---|---|---|
| Bold | Ctrl + B | Command + B |
| Italic | Ctrl + I | Command + I |
| Underline | Ctrl + U | Command + U |
| Format Cells | Ctrl + 1 | Command + 1 |
| Apply Percentage Format | Ctrl + Shift + % | Control + Shift + % |
| Apply Currency Format | Ctrl + Shift + $ | Control + Shift + $ |
| Apply Date Format | Ctrl + Shift + # | Control + Shift + # |

### 3.4 Navigation Operations

| Operation | Windows Shortcut | macOS Shortcut |
|---|---|---|
| Move to Beginning of Row | Home | Fn + Left Arrow |
| Move to Cell A1 | Ctrl + Home | Control + Home |
| Move to Last Used Cell | Ctrl + End | Control + End |
| Move One Worksheet Right | Ctrl + Page Down | Control + Page Down |
| Move One Worksheet Left | Ctrl + Page Up | Control + Page Up |
| Move to Next Cell | Tab | Tab |
| Move to Previous Cell | Shift + Tab | Shift + Tab |

### 3.5 Selection Operations

| Operation | Windows Shortcut | macOS Shortcut |
|---|---|---|
| Select Entire Row | Shift + Space | Shift + Space |
| Select Entire Column | Ctrl + Space | Control + Space |
| Extend Selection | Shift + Arrow | Shift + Arrow |
| Select to Last Non-Empty Cell | Ctrl + Shift + Arrow | Command + Shift + Arrow |
| Select Entire Data Region | Ctrl + A | Command + A |

### 3.6 Data Operations

| Operation | Windows Shortcut | macOS Shortcut |
|---|---|---|
| Create Table | Ctrl + T | Command + T |
| Apply Filter | Ctrl + Shift + L | Command + Shift + F |
| Flash Fill | Ctrl + E | Control + E |
| Open Go To | Ctrl + G | Control + G |
| Insert Current Date | Ctrl + ; | Control + ; |
| Insert Current Time | Ctrl + Shift + : | Command + ; |

### 3.7 Worksheet Operations

| Operation | Windows Shortcut | macOS Shortcut |
|---|---|---|
| Insert New Worksheet | Shift + F11 | Fn + Shift + F11 |
| Rename Worksheet | Alt + H + O + R | Ribbon/Menu navigation |
| Delete Worksheet | No universal direct shortcut | No universal direct shortcut |

Shortcut availability may vary by Excel version, keyboard layout, and application environment.

## 4. Shortcut Identification

When a user requests a shortcut, identify the appropriate shortcut based on:

- Requested operation.
- Operating system.
- Excel version.
- Excel desktop or web environment.
- Keyboard layout.
- Available interaction capabilities.

Rules:

- Prefer the shortcut appropriate to the current environment.
- Do not assume Windows when the environment is macOS.
- Distinguish Excel desktop shortcuts from browser shortcuts.
- Check whether function keys require an Fn modifier.
- Do not invent unsupported shortcuts.
- If the environment is unknown, ask only when the distinction affects execution.

## 5. Shortcut Execution Requirements

Before executing a keyboard shortcut, identify:

1. Target workbook.
2. Target worksheet.
3. Current application state.
4. Active cell or selection.
5. Intended operation.
6. Required keyboard shortcut.
7. Expected result.

The agent must verify that the correct Excel application or interface is active before sending keyboard input.

Do not send shortcuts to an unrelated application.

## 6. Selection-Dependent Shortcuts

Some shortcuts depend on the currently selected cell or range.

Examples:

- Ctrl + B applies bold formatting to the selected cells.
- Ctrl + Space selects the entire column of the active cell.
- Shift + Space selects the entire row of the active cell.
- Ctrl + Shift + L toggles filters for the applicable data region.

Rules:

- Identify the intended target before executing.
- Select the correct cell or range.
- Verify the selection when possible.
- Execute the shortcut only after the target is confirmed.
- Do not assume the current selection is correct.

## 7. Multi-Key Shortcut Execution

Some Excel shortcuts require multiple keys or sequential key combinations.

Examples:

- Ctrl + Shift + L
- Ctrl + Shift + +
- Ctrl + Alt + V

Rules:

- Preserve the correct key combination.
- Execute modifier keys and action keys in the correct order.
- Release modifier keys appropriately.
- Distinguish simultaneous key combinations from sequential menu shortcuts.
- Do not send incomplete key combinations.
- Verify the resulting application state.

For sequential shortcuts, such as Ribbon key sequences, follow the appropriate application-specific key sequence rather than treating every key as a simultaneous combination.

## 8. Shortcut Compatibility Validation

Before execution, validate shortcut compatibility.

Consider:

- Operating system.
- Excel desktop version.
- Excel web version.
- Keyboard layout.
- Function-key configuration.
- Application focus.
- Existing keyboard shortcut conflicts.

If the requested shortcut is unsupported:

1. Identify the limitation.
2. Determine whether another compatible shortcut exists.
3. Consider an available direct Excel tool as an alternative.
4. Execute only if the alternative produces the intended result.
5. Explain the limitation when no suitable method exists.

## 9. Shortcut vs Direct Tool Execution

The agent may have multiple ways to perform the same operation.

Example:

User request:

"Make A1 bold."

Possible execution methods:

- Keyboard shortcut: Ctrl + B.
- Direct formatting tool: Apply bold to A1.

Selection depends on:

- Available tools.
- Application environment.
- Current selection.
- Operation requirements.
- Reliability of the execution method.

Rules:

- Prefer the most reliable supported method.
- Use shortcuts when the request specifically requires keyboard interaction or shortcut execution.
- Do not force shortcut execution when a direct tool is more reliable and the user has not specified a shortcut.
- Do not call multiple execution methods for the same operation unnecessarily.

## 10. Shortcut Execution Verification

After executing a shortcut, verify its expected outcome.

Examples:

| Shortcut | Expected Result |
|---|---|
| Ctrl + S | Workbook saved |
| Ctrl + B | Bold formatting applied or toggled |
| Ctrl + C | Selected content copied |
| Ctrl + Z | Previous action undone |
| Ctrl + F | Find interface opened |
| Ctrl + Shift + L | Filter toggled |
| Shift + F11 | New worksheet inserted |

Important:

- A successful keypress does not guarantee a successful Excel operation.
- Verify the actual application or workbook state.
- Consider toggle behaviour when checking formatting and filters.
- Do not repeat a shortcut blindly if the first execution may have succeeded.

## 11. Error Handling

Handle common shortcut execution failures.

| Failure | Recovery Strategy |
|---|---|
| Wrong application focused | Activate the correct Excel window |
| Incorrect cell selection | Re-select the intended range |
| Unsupported shortcut | Identify a compatible alternative |
| Shortcut ignored | Check application focus and environment |
| Wrong keyboard layout | Validate key mapping |
| Fn key issue | Check function-key configuration |
| Shortcut conflict | Use an alternative supported method |
| Application unresponsive | Follow execution recovery procedures |
| Unexpected result | Inspect workbook state before retrying |

Do not repeatedly send the same shortcut without identifying the failure reason.

## 12. Shortcut Execution Workflow

Follow this sequence for every shortcut-related execution request.

1. Understand the user's request.
2. Identify the intended operation.
3. Determine whether the user wants an explanation or execution.
4. Identify the target workbook and worksheet.
5. Identify the required cell or range selection.
6. Determine the operating system and Excel environment.
7. Identify the compatible shortcut.
8. Check application focus.
9. Validate the target selection.
10. Execute the shortcut through an available keyboard interaction tool.
11. Verify the expected result.
12. Report the execution status.

## 13. Important Rules

- Never assume a universal Excel shortcut.
- Distinguish Windows and macOS shortcuts.
- Consider Excel desktop and web differences.
- Validate application focus before sending keys.
- Verify the intended cell selection.
- Preserve correct key combinations.
- Distinguish simultaneous and sequential shortcuts.
- Do not execute shortcuts when the user only requests an explanation.
- Do not repeat potentially successful shortcuts blindly.
- Prefer reliable execution methods.
- Verify actual workbook changes.
- Use only available keyboard interaction capabilities.
- Follow execution recovery procedures when failures occur.

## 14. Final Shortcut Execution Principle

The Excel AI Agent must select and execute keyboard shortcuts based on the user's intended operation and the actual Excel environment.

A shortcut must never be selected using memorized key combinations alone.

The agent must validate the application state, target selection, shortcut compatibility, and execution result before reporting success.