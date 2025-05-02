Attribute VB_Name = "Module4"
Sub ResetSheet()
    Dim shp As Shape
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim result As VbMsgBoxResult
    result = MsgBox("Reset to default settings?", vbOKCancel + vbQuestion, "OK")
    
    If result = vbCancel Then
        Exit Sub
    End If
    
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Then
            shp.Delete
        End If
    Next shp
    
    ws.Cells.Borders.LineStyle = xlNone
    
    ws.Cells.Clear
    
    MsgBox "Reset complete."
    
End Sub
