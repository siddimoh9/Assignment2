Attribute VB_Name = "Module1"
Sub assignment_2_retry():
'have to configure each ws to run sepratelt as variable; start of first and last layer
Dim ws As Worksheet

For Each ws In Worksheets

'Declare main variables whilst also starting values and setup, figure out lastrow setup and variables
Dim Ticker_Name As String
   Ticker_Name = " "
Dim Total_T As Double
   Total_T = 0
Dim Lastrow As Long
Dim i As Long
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
' Naming and adding columns
ws.Range("I1").EntireColumn.Insert
ws.Cells(1, 9).Value = "Ticker"
ws.Range("J1").EntireColumn.Insert
ws.Cells(1, 10).Value = "Yearly_Change"
ws.Range("K1").EntireColumn.Insert
ws.Cells(1, 11).Value = "Percent Change"
ws.Range("L1").EntireColumn.Insert
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker_Name"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest_Percentage_Increase"
ws.Range("O3").Value = "Greatest_Percentage_Decrease"
ws.Range("O4").Value = "Greatest_Total_Volume"
     
' Set Variables for the prices, changes, percentages and range of values. Also set variables for the bonus assessment.
Dim opening As Double
    opening = 0
Dim closing As Double
    closing = 0
Dim change As Double
    change = 0
Dim perchange As Double
    perchange = 0
Dim Ticker_Increase As String
    Ticker_Increase = ""
Dim Value_Increase As Double
    Value_Increase = 0
Dim Ticker_Decrease As String
    Ticker_Decrease = ""
Dim Value_Decrease As Double
    Value_Decrease = 0
Dim GreatestVolume_T As String
    GreatestVolume_T = ""
Dim GreatestVolume_TV As Double
    GreatestVolume_TV = 0
Dim Formatchange As Double
Dim Ticker_Row As Long: Ticker_Row = 1
'first loop
For i = 2 To Lastrow
    If opening = 0 Then
        opening = ws.Cells(i, 3).Value
    End If
      
' second loop for when we are contained in same ticker name; <> means not equal to; name the ticker and out the value; use i function; use opening for start and close for ending values; and make formula, establish what exactly the price change is
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      Ticker_Row = Ticker_Row + 1
      Ticker_Name = ws.Cells(i, 1).Value
      ws.Cells(Ticker_Row, "I").Value = Ticker_Name
      closing = ws.Cells(i, 6).Value
      change = closing - opening
      ws.Cells(Ticker_Row, "J").Value = change
'create a color index for the cells
    If change < 0 Then
            ws.Cells(Ticker_Row, "J").Interior.ColorIndex = 3
        ElseIf change > 0 Then
            ws.Cells(Ticker_Row, "J").Interior.ColorIndex = 4
    End If
'format .##%, and do percentage chnge, fill in the values for both conditions of increase and decrease
perchange = (change / opening)
ws.Cells(Ticker_Row, "K").Value = perchange
ws.Cells(Ticker_Row, "K").NumberFormat = ".##%"
If Value_Increase < perchange Then
      Ticker_Increase = Ticker_Name
      Value_Increase = perchange
End If
If Value_Decrease > perchange Then
      Ticker_Decrease = Ticker_Name
      Value_Decrease = perchange
End If
      
opening = 0
' Add to the Ticker Total
Total_T = Total_T + ws.Cells(i, 7).Value
ws.Cells(Ticker_Row, "L").Value = Total_T
ws.Cells(Ticker_Row, "L").NumberFormat = "0"
If GreatestVolume_TV < Total_T Then
   GreatestVolume_TV = Total_T
   GreatestVolume_T = Ticker_Name
End If
Total_T = 0
    
Else

Total_T = Total_T + ws.Cells(i, 7).Value

End If
     
'close second last loop
Next i
    
'summary table values
ws.Cells(2, "P") = Ticker_Increase
ws.Cells(2, "Q") = Value_Increase
ws.Cells(3, "P") = Ticker_Decrease
ws.Cells(3, "Q") = Value_Decrease
ws.Cells(4, "P") = GreatestVolume_T
ws.Cells(4, "Q") = GreatestVolume_TV
ws.Range("Q1:Q4").NumberFormat = ".##%"
'clos elast loop
    Next ws



End Sub
