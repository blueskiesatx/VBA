Sub StockAnalyst()
For Each ws In Worksheets

' Created a Variable to Hold File Name, Last Row, Last Column, and Year
        Dim WorksheetName As String

        ' Determine the Last Row
        'lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
'set variables
Dim ticker As Variant
Dim volume As Double
Dim lastrow As Long
Dim summary_number As Integer

'You will also need to display the ticker symbol to coincide with the total volume.
'Set values in column 9 to "Ticker" and values in column 10 to "Total Stick Volume"
volume = 0
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Total Stock Volume"
summary_number = 2

'get last row in the sheet using column A (size of the table of parts)
lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
If (Cells(i + 1, 1).Value) <> (Cells(i, 1).Value) Then
ticker = Cells(i, 1).Value
volume = volume + Cells(i, 7).Value
'output values for ticker and volume sum in columns I and J
Range("I" & summary_number).Value = ticker
Range("J" & summary_number).Value = volume
summary_number = summary_number + 1
'reset volume sum
volume = 0

Else
volume = volume + Cells(i, 7).Value
End If
Next i
Next ws
End Sub