# ResetSheet Macro

This VBA macro resets the currently active worksheet to its default state. It is intended for use in Excel environments where worksheets may contain images and custom content that need to be cleared quickly.

## Features

- Prompts the user for confirmation before performing any reset.
- Deletes all **pictures** (shapes of type `msoPicture`) from the active sheet.
- Clears **all cell contents and borders**.
- Displays a completion message upon success.

## How It Works

1. Displays a confirmation message:
   - `"Reset to default settings?"`
   - If the user selects **Cancel**, the macro ends without making changes.

2. If confirmed:
   - Iterates through all shapes in the active sheet.
   - Deletes any shape that is a picture (`msoPicture`).
   - Removes all borders from the worksheet.
   - Clears all cell contents (`ws.Cells.Clear`).

3. Shows a message box:
   - `"Reset complete."`

## Usage

1. Press `Alt + F11` to open the **VBA Editor**.
2. Paste the macro into a module.
3. Run `ResetSheet` while the target sheet is active.

## Warning

- **This action cannot be undone.**
- Make sure to save your work or backup the sheet before running.

## Example

```vba
Sub ResetSheet()
    Dim shp As Shape
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim result As VbMsgBoxResult
    result = MsgBox("Reset to default settings?", vbOKCancel + vbQuestion, "OK")
    If result = vbCancel Then Exit Sub

    For Each shp In ws.Shapes
        If shp.Type = msoPicture Then shp.Delete
    Next shp

    ws.Cells.Borders.LineStyle = xlNone
    ws.Cells.Clear

    MsgBox "Reset complete."
End Sub