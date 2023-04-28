Attribute VB_Name = "Module1"
Sub Multiyear()
'define'
    Dim i As Variant
    For i = 2 To 753001
    For Each ws In Worksheets

'lable'
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Value"
    ws.Cells(2, 15).Value = "Greatest Increase%"
    ws.Cells(3, 15).Value = "Greatest Decrease&"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
'tinker'
    ws.Range("A2:A753001").Copy
    ws.Range("I2:753001").PasteSpecial
    
'Yearly Change'
    ws.Cells(i, 10).Value = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
    If ws.Cells(i, 10).Value < 0 Then
        Cells(i, 10).Value.Interior.Color = vbRed
    Else
        Cells(i, 10).Value.Interior.Color = vbGreen
    End If
'Percent Change'
    ws.Cells(i, 11).Value = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value / ws.Cells(i, 3).Value
'Total Stock Value'
    ws.Cells(i, 12).Value = WorksheetFunction.Sum(ws.Cells(2, 7), ws.Cells(i, 7))
'Greatest Increase'
    ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Cells(i, 11))
'Greatest Decrease'
    ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Cells(i, 11))
'Greatst Total Volume"
    ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Cells(i, 7))
    
    Next i
    Next ws
    
End Sub
