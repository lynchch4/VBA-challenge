VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub Clear()

For Each ws In Worksheets

ws.Range("I:Z") = ""
ws.Range("I:Z").Interior.ColorIndex = 2

Next ws

End Sub

Sub Stock()

For Each ws In Worksheets

Dim Counter As Integer
Dim Counter2 As Integer
Dim VolTotal As String
Dim LastRow As Long

LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row

Counter = 1
Counter2 = 1
VolTotal = 0


    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
    
        Counter2 = Counter2 + 1
        
        VolTotal = VolTotal + ws.Cells(i, 7)
        
    
    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ws.Cells(Counter + 1, 9).Value = ws.Cells(i, 1).Value
    
    'close value - Cells(counter, 10).Value = Cells(i, 6).Value
    'open value - Cells(counter, 11).Value = Cells(i - Counter2 + 1, 3).Value
    
        ws.Cells(Counter + 1, 10).Value = ws.Cells(i, 6).Value - ws.Cells(i - Counter2 + 1, 3).Value
        
        ws.Cells(Counter + 1, 11).Value = ws.Cells(i, 6).Value / ws.Cells(i - Counter2 + 1, 3).Value - 1
        
        ws.Cells(Counter + 1, 11).NumberFormat = "0.00%"
    
        ws.Cells(Counter + 1, 12).Value = VolTotal + ws.Cells(i, 7).Value
    
            Counter = 1 + Counter
            
            Counter2 = 1
    
            VolTotal = 0
    
    End If
        
    Next i
    
    ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Range("K2:K10000"))
        ws.Cells(2, 17).NumberFormat = "0.00%"

    ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Range("K2:K10000"))
        ws.Cells(3, 17).NumberFormat = "0.00%"

    ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("L2:L10000"))
          

Call TickerValue


Next ws

End Sub



Sub TickerValue()
 
For Each ws In Worksheets

 Dim LastRow2 As Integer
 
 LastRow2 = ws.Range("I" & Rows.Count).End(xlUp).Row
 
 
 
 For i = 2 To LastRow2
        
        If ws.Cells(i, 11).Value = ws.Cells(2, 17).Value Then
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        
        ElseIf ws.Cells(i, 11).Value = ws.Cells(3, 17).Value Then
            ws.Cells(3, 16).Value = Cells(i, 9).Value
        
        ElseIf ws.Cells(i, 12).Value = ws.Cells(4, 17).Value Then
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        
        
        End If
        
        If ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
            ws.Cells(i, 11).Interior.ColorIndex = 3
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 4
            ws.Cells(i, 11).Interior.ColorIndex = 4
        End If
    
    
    Next i
Next ws

End Sub


