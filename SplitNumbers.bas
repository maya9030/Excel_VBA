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
        If Len(chunk) < 4 Then
            Exit For
        End If
        
        Set outputCell = startCell.Offset((i - 1) \ 4, 0)
        With outputCell
            
            .NumberFormat = "@"
            
            Union(.Offset(0, 1), .Offset(0, 3)).NumberFormat = "@"
            
            .Value = Format(chunk, "0000")
            
            .Offset(0, -1).Value = i + 5
            
            .Offset(0, 1).Value = Format$(Left(chunk, 2), "00")
            
            .Offset(0, 3).Value = Format$(Right(chunk, 2), "00")
            
            .Offset(0, 2).Formula = "=VLOOKUP(VALUE(LEFT(" & _
                                    .Offset(0, 1).Address & _
                                    ",1))," & _
                                    matrix.Address & _
                                    ",VALUE(RIGHT(" & _
                                    .Offset(0, 1).Address & _
                                    ",1)+2),FALSE)"
                                    
            .Offset(0, 4).Formula = "=VLOOKUP(VALUE(LEFT(" & _
                                    .Offset(0, 3).Address & _
                                    ",1))," & _
                                    matrix.Address & _
                                    ",VALUE(RIGHT(" & _
                                    .Offset(0, 3).Address & _
                                    ",1)+2),FALSE)"
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

