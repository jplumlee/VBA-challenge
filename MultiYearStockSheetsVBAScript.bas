Attribute VB_Name = "Module1"
Sub MultiYearStockSheets()

Dim ws As Worksheet
For Each ws In Worksheets

Dim ticker As String
Dim volume As Double
volume = 0
Dim ychange As Double
ychange = 0
Dim percentchange As Double
percentchange = 0
Dim yopen As Double
yopen = 0
Dim yclose As Double
yclose = 0
Dim gincreae As Double
gincrease = 0
Dim gincreaseticker As String
Dim gdecrease As Double
gdecrease = 0
Dim gdecreaseticker As String
Dim gtotalvolume As Double
gtotalvolume = 0
Dim gtotalvolumeticker As String

Dim Lastrow As Long
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Dim tablerow As Long
tablerow = 2
Dim i As Long

'Add column/row titles
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"


'Capture year open for first ticker of current worksheet
yopen = ws.Cells(2, 3).Value

'Loop to add data to table
For i = 2 To Lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Find Values
ticker = ws.Cells(i, 1).Value
yclose = ws.Cells(i, 6).Value

'Calc ychange
ychange = yclose - yopen

'Prevent divisible by zero error
If yopen > 0 Then
'Calc percentchange
percentchange = (ychange / yopen) * 100
End If

'Calc volume
volume = volume + ws.Cells(i, 7).Value

'Insert Values
ws.Cells(tablerow, 9).Value = ticker
ws.Cells(tablerow, 10).Value = ychange
'Covert number to string and format as percent - https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/type-conversion-functions
ws.Cells(tablerow, 11).Value = CStr(percentchange) & "%"
ws.Cells(tablerow, 12).Value = volume
ws.Range("P2").Value = gincreaseticker
ws.Range("Q2").Value = CStr(gincrease) & "%"
ws.Range("P3").Value = gdecreaseticker
ws.Range("Q3").Value = CStr(gdecrease) & "%"
ws.Range("P4").Value = gtotalvolumeticker
ws.Range("Q4").Value = gtotalvolume

'Change ychange cell color
If ws.Cells(tablerow, 10) <= 0 Then
ws.Cells(tablerow, 10).Interior.ColorIndex = 3
ElseIf ws.Cells(tablerow, 10) > 0 Then
ws.Cells(tablerow, 10).Interior.ColorIndex = 4
End If

'Concept for finding min/max - https://stackoverflow.com/questions/45422688/vba-for-loop-to-find-maximum-value-in-a-column
'Calc gincrease by looping through percentchange for the max
If percentchange > gincrease Then
gincrease = percentchange
gincreaseticker = ticker
End If

'Calc gdecrease by looping through percentchange for the min
If percentchange < gdecrease Then
gdecrease = percentchange
gdecreaseticker = ticker
End If

'Calc gtotalvolume by looping through volume for the max
If volume > gtotalvolume Then
gtotalvolume = volume
gtotalvolumeticker = ticker
End If

'Go to next row in table and clear values from previous ticker calcs
tablerow = tablerow + 1
yclose = 0
yopen = 0
ychange = 0
percentchange = 0
volume = 0

'Capture the next ticker's yopen
yopen = ws.Cells(i + 1, 3).Value

'Else increase ticker volume
Else
volume = volume + ws.Cells(i, 7).Value
End If

Next i
'Make columns autofit to contents
ws.Columns("A:Q").AutoFit
Next ws

End Sub
