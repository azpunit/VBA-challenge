Sub Column_Creation():

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    Dim Ticker_Symbol As String
    Dim LastRow As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim Summary_Table_Row As Variant
    Dim Ticker_Change As Boolean
  
Summary_Table_Row = 2
LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
Ticker_Change = True
ws.Columns("J:L").AutoFit

For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      Close_Price = ws.Cells(i, 6).Value
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
      Yearly_Change = Close_Price - Open_Price
      If Open_Price = 0 Then
      Percent_Change = 0
      Else
      Percent_Change = Yearly_Change / Open_Price
      End If
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
      ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
      ws.Range("J" & summary_table_Row).NumberFormat = "0.00"
      ws.Range("K" & Summary_Table_Row).Value = Percent_Change
      ws.Range("K" & summary_table_Row).NumberFormat = "0.00%"
      ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
      Summary_Table_Row = Summary_Table_Row + 1
      Ticker_Change = True
    Else
    If Ticker_Change = True Then
      Ticker_Symbol = ws.Cells(i, 1).Value
      Open_Price = ws.Cells(i, 3).Value
      Ticker_Change = False
      Total_Stock_Volume = 0
  End If
End If
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
Next i
Next ws

End Sub

Sub Conditional_Formating():

Dim LastRow1 As Double
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

LastRow1 = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

For i = 2 To LastRow1
     If ws.Cells(i, 10).Value > 0 Then
   ws.Cells(i, 10).Interior.ColorIndex = 4
     ElseIf ws.Cells(i, 10).Value < 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 3
     ElseIf ws.Cells(i, 10).Value = 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 0
     End If
   Next i
Next ws

End Sub

Sub Values_Calculation():

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

Dim LastRow2 As Double
Dim Max_Percent As Double
Dim Min_Percent As Double
Dim Max_Total_Volume As Variant

Max_Total_Volume = 0
Max_Percent = 0
Min_Percent = 0
LastRow2 = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
ws.Columns("O").AutoFit
ws.Columns("Q").AutoFit

For i = 2 To LastRow2
    If Max_Total_Volume <= ws.Cells(i, 12) Then
      Max_Total_Volume = ws.Cells(i, 12).Value
      ws.Range("P4").Value = ws.Cells(i, 9).Value
      ws.Range("Q4").Value = ws.Cells(i, 12).Value
      ws.Range("Q4").NumberFormat = "0.0000E+0"
    ElseIf Max_Percent <= ws.Cells(i, 11).Value Then
      Max_Percent = ws.Cells(i, 11).Value
      ws.Range("P2").Value = ws.Cells(i, 9).Value
      ws.Range("Q2").Value = ws.Cells(i, 11).Value
      ws.Range("Q2").NumberFormat = "0.00%"
    ElseIf Min_Percent >= ws.Cells(i, 11).Value Then
      Min_Percent = ws.Cells(i, 11).Value
      ws.Range("P3").Value = ws.Cells(i, 9).Value
      ws.Range("Q3").Value = ws.Cells(i, 11).Value
      ws.Range("Q3").NumberFormat = "0.00%"
    End If   
  Next i
Next ws

End Sub

