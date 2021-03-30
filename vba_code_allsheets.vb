'Attribute VB_Name = "Module1"
Sub Stock_Market_Analysis()

Dim ws As Worksheet
Dim WorksheetName As String
Dim ticker As String
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim vol As Double
Dim Row As Long
Dim start As Long
Dim i As Long

'run through each worksheet
For Each ws In Worksheets
    WorksheetName = ws.Name
    MsgBox WorksheetName
    'ws.Activate
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly change"
    ws.Cells(1, 11).Value = "percentage change"
    ws.Cells(1, 12).Value = "Total stock volume"
     
    vol = 0
    Row = 2
    start = 2
    'Setup initial Open Price
    Open_Price = ws.Cells(start, 3).Value
    
    'loop through all ticker symbol
    For i = 2 To LastRow
       
    'check if its still within the same ticker symbol, if not ...
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     
    'set Ticker Name
     Ticker_Name = ws.Cells(i, 1).Value
     ws.Cells(Row, 9).Value = Ticker_Name
     
    'Set close Price
     Close_Price = ws.Cells(i, 6).Value
     
    'Add Yearly Change
     yearly_change = Close_Price - Open_Price
     ws.Cells(Row, 10).Value = yearly_change
     
    'Add Percent Change
     If (Open_Price = 0 And Close_Price = 0) Then
      percent_change = 0
     ElseIf (Open_Price = 0 And Close_Price <> 0) Then
         percent_change = 1
     Else
         percent_change = yearly_change / Open_Price
         ws.Cells(Row, 11).Value = percent_change
         ws.Cells(Row, 11).NumberFormat = "0.00%"
     End If
           
    'Conditional formatting for yearly_change
     If ws.Cells(Row, 10).Value > 0 Then
        ws.Cells(Row, 10).Interior.ColorIndex = 4
     ElseIf ws.Cells(Row, 10).Value < 0 Then
           ws.Cells(Row, 10).Interior.ColorIndex = 3
     ElseIf ws.Cells(Row, 10).Value = 0 Then
           ws.Cells(Row, 10).Interior.ColorIndex = 0
          
     End If
           
     'Add Total volumn
        vol = vol + ws.Cells(i, 7).Value
        ws.Cells(Row, 12).Value = vol
    
     'Add one to the sumary table row
        Row = Row + 1
           
     'reset the open Price
        start = i + 1
           
     'reset the Volumn Total
        vol = 0
           
     'If cells are the same ticker
    Else
        vol = vol + ws.Cells(i, 7).Value
             
        End If
    Next i
    
     
    Next ws
    
End Sub



























