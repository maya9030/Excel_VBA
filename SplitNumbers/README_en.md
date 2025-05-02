# SplitNumbers Macro

This VBA macro is designed for use in **memory sports (memorization competitions)**. It splits a long numeric string into 4-digit chunks, then further divides each chunk into two 2-digit numbers. These 2-digit numbers are then used as coordinates to retrieve associated words from a **10×10 lookup matrix**, allowing for rapid visualization and memorization.

---

## Purpose

The macro enables the automation of converting digit sequences into visual representations by referencing a predefined **2-digit number to word matrix**, typically structured as a 10×10 grid.

### Intended Use:

- The matrix is laid out as a **10×10 grid (0–9 along both rows and columns)**.
- Each cell in the matrix corresponds to a 2-digit number (e.g., 23, 87).
- Upon execution, the macro:
  1. Removes whitespace from a numeric string
  2. Splits the string into **4-digit chunks**
  3. Further splits each chunk into **two 2-digit numbers**
  4. Uses these numbers to **lookup corresponding words** in the matrix using dynamically generated `VLOOKUP` formulas
  5. Outputs all related values and formulas to the sheet

### Example Use Cases:

- Practicing number-to-word image conversion for memory competition
- Automatically generating practice material based on numeric sequences
- Enhancing mnemonic strategies through visual anchors

---

## Workflow

1. User is prompted to select:
   - The cell containing the numeric input
   - The cell where output should begin

2. The macro:
   - Removes whitespace using regular expressions
   - Splits the input into 4-digit chunks
   - Extracts the first and last two digits of each chunk
   - Outputs:
     - The full 4-digit chunk
     - Each 2-digit half
     - Two `VLOOKUP` formulas referencing the external matrix

---

## Requirements

- The lookup matrix must exist in range `B2:L12` by default
- The input should be a **pure numeric string without spaces or line breaks**
- Chunks with fewer than 4 digits are ignored

---

## Sample Code

```vba
Option Explicit

Sub SplitNumbers()
    Dim targetCell As Range
    Dim startCell As Range
    Dim outputCell As Range
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")

    With reg
        .Pattern = "[\s]+"
        .Global = True
    End With

    Dim number As String
    Dim i As Long, chunk As String
    Dim matrix As Range
    Set matrix = Range("B2:L12")

    Set targetCell = selectCell("Please select the target cell.")
    Set startCell = selectCell("Please select the starting cell for output.")

    number = reg.Replace(targetCell.Value, "")

    For i = 1 To Len(number) Step 4
        chunk = Mid$(number, i, 4)
        If Len(chunk) < 4 Then Exit For

        Set outputCell = startCell.Offset((i - 1) \ 4, 0)
        With outputCell
            .NumberFormat = "@"
            Union(.Offset(0, 1), .Offset(0, 3)).NumberFormat = "@"
            .Value = Format(chunk, "0000")
            .Offset(0, -1).Value = i + 5
            .Offset(0, 1).Value = Format$(Left(chunk, 2), "00")
            .Offset(0, 3).Value = Format$(Right(chunk, 2), "00")
            .Offset(0, 2).Formula = "=VLOOKUP(VALUE(LEFT(" & _
                .Offset(0, 1).Address & ",1)," & matrix.Address & _
                ",VALUE(RIGHT(" & .Offset(0, 1).Address & ",1)+2),FALSE)"
            .Offset(0, 4).Formula = "=VLOOKUP(VALUE(LEFT(" & _
                .Offset(0, 3).Address & ",1)," & matrix.Address & _
                ",VALUE(RIGHT(" & .Offset(0, 3).Address & ",1)+2),FALSE)"
        End With
    Next i

    MsgBox "Split completed"
End Sub

Function selectCell(text As String) As Range
    Dim returnCell As Range
    On Error Resume Next
    Set returnCell = Application.InputBox(text, Type:=8)
    On Error GoTo 0

    If returnCell Is Nothing Then
        MsgBox "Invalid", vbExclamation
        Exit Function
    End If

    Set selectCell = returnCell
End Function