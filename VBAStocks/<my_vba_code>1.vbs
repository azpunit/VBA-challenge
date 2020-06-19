Sub Column_Creation():

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
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
LastRow = Cells(Rows.Count, "A").End(xlUp).Row
Ticker_Change = True
Columns("J:L").AutoFit

For i = 2 To LastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      Close_Price = Cells(i, 6).Value
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
      Yearly_Change = Close_Price - Open_Price
      If Open_Price = 0 Then
      Percent_Change = 0
      Else
      Percent_Change = Yearly_Change / Open_Price
      End If
      Range("I" & Summary_Table_Row).Value = Ticker_Symbol
      Range("J" & Summary_Table_Row).Value = Yearly_Change
      Range("J" & summary_table_Row).NumberFormat = "0.00"
      Range("K" & Summary_Table_Row).Value = Percent_Change
      Range("K" & summary_table_Row).NumberFormat = "0.00%"
      Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
      Summary_Table_Row = Summary_Table_Row + 1
      Ticker_Change = True
    Else
    If Ticker_Change = True Then
      Ticker_Symbol = Cells(i, 1).Value
      Open_Price = Cells(i, 3).Value
      Ticker_Change = False
      Total_Stock_Volume = 0
  End If
End If
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
Next i

End Sub

Sub Conditional_Formating():

Dim LastRow1 As Double

LastRow1 = Cells(Rows.Count, "I").End(xlUp).Row

For i = 2 To LastRow1
     If Cells(i, 10).Value > 0 Then
    Cells(i, 10).Interior.ColorIndex = 4
     ElseIf Cells(i, 10).Value < 0 Then
    Cells(i, 10).Interior.ColorIndex = 3
     ElseIf Cells(i, 10).Value = 0 Then
    Cells(i, 10).Interior.ColorIndex = 0
     End If
   Next i

End Sub

Sub Values_Calculation():

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

Dim LastRow2 As Double
Dim Max_Percent As Double
Dim Min_Percent As Double
Dim Max_Total_Volume As Variant

Max_Total_Volume = 0
Max_Percent = 0
Min_Percent = 0
LastRow2 = Cells(Rows.Count, "I").End(xlUp).Row
Columns("O").AutoFit
Columns("Q").AutoFit

For i = 2 To LastRow2
    If Max_Total_Volume <= Cells(i, 12) Then
      Max_Total_Volume = Cells(i, 12).Value
      Range("P4").Value = Cells(i, 9).Value
      Range("Q4").Value = Cells(i, 12).Value
      Range("Q4").NumberFormat = "0.0000E+0"
    ElseIf Max_Percent <= Cells(i, 11).Value Then
      Max_Percent = Cells(i, 11).Value
      Range("P2").Value = Cells(i, 9).Value
      Range("Q2").Value = Cells(i, 11).Value
      Range("Q2").NumberFormat = "0.00%"
    ElseIf Min_Percent >= Cells(i, 11).Value Then
      Min_Percent = Cells(i, 11).Value
      Range("P3").Value = Cells(i, 9).Value
      Range("Q3").Value = Cells(i, 11).Value
      Range("Q3").NumberFormat = "0.00%"
    End If   
  Next i

End Sub

