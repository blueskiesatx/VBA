Sub StockAnalyst_3()
For Each ws In Worksheets

' Created a Variable to Hold File Name, Last Row, Last Column, and Year
        Dim WorksheetName As String

      ' Grabbed the WorksheetName
        WorksheetName = ws.Name
'set variables
Dim ticker As Variant
Dim volume As Double
Dim lastrow As Long
Dim summary_number As Integer
Dim openP As Double
Dim closeP As Double

'You will also need to display the ticker symbol to coincide with the total volume.
'Set values in column 9 to "Ticker" and values in column 10 to "Yearly Price Change"

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Price Change"
'Set values in column 11 to "Percent Change" and values in column 12 to "Total Stock Volume"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"


'Create a script that will loop through all the stocks and take the following info.
lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
volume = 0
summary_number = 2
openP = ws.Cells(2, 3)
'get last row in the sheet using column A (size of the table of parts)
For i = 2 To lastrow
If (ws.Cells(i + 1, 1).Value) <> (ws.Cells(i, 1).Value) Then

'Yearly change from what the stock opened the year at to what the closing price was.
ticker = ws.Cells(i, 1).Value
closeP = ws.Cells(i, 6).Value
volume = volume + ws.Cells(i, 7).Value
ws.Range("I" & summary_number).Value = ticker
ws.Range("J" & summary_number).Value = closeP - openP
'You should also have conditional formatting that will highlight positive change in green and negative change in red.
If (closeP - openP) < 0 Then
ws.Range("J" & summary_number).Interior.ColorIndex = 3
Else
ws.Range("J" & summary_number).Interior.ColorIndex = 4
End If
'The percent change from the what it opened the year at to what it closed. What if have zeroes, or no change?
If openP = 0 Then
ws.Range("K" & summary_number).Value = 0
Else
ws.Range("K" & summary_number) = FormatPercent(((closeP - openP) / openP))
End If
ws.Range("L" & summary_number).Value = volume
summary_number = summary_number + 1
volume = 0
openP = ws.Cells(i + 1, 3).Value
closeP = 0
Else
volume = volume + ws.Cells(i, 7).Value
End If
Next i
'Your solution will also be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
Dim bestest As Double
Dim worstest As Double
Dim greatestV As Double
Dim lastcolumnrow As Integer
bestest = ws.Cells(2, 10).Value
worstest = ws.Cells(2, 10).Value
greatestV = ws.Cells(2, 12).Value
lastcolumnrow = ws.Range("L" & Rows.Count).End(xlUp).Row
For i = 2 To lastcolumnrow
If ws.Cells(i + 1, 12) > bestest Then
bestest = ws.Cells(i + 1, 12).Value
ws.Range("O2") = "Greatest % Increase"
ws.Range("P2") = ws.Cells(i + 1, 10).Value
ws.Range("Q2") = ws.Cells(i + 1, 11).Value
End If
If ws.Cells(i + 1, 12) < worstest Then
worstest = ws.Cells(i + 1, 12).Value
ws.Range("O3") = "Greatest % Decrease"
ws.Range("P3") = ws.Cells(i + 1, 10).Value
ws.Range("Q3") = ws.Cells(i + 1, 11).Value
End If
If ws.Cells(i + 1, 12) > greatestV Then
greatestV = ws.Cells(i + 1, 12).Value
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P4") = ws.Cells(i + 1, 11).Value
ws.Range("Q4") = ws.Cells(i + 1, 12).Value
Next ws

End Sub
