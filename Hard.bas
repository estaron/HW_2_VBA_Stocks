Attribute VB_Name = "Module1"
Sub Hard()
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate

Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"
Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Cells(2, 15) = "Greatest % increase"
Cells(3, 15) = "Greatest % decrease"
Cells(4, 15) = "Greatest Total Volume"
Cells(2, 9) = Cells(2, 1)
j = 3
For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
If Cells(i, 1) <> Cells(i + 1, 1) Then
Cells(j, 9) = Cells(i + 1, 1)
j = j + 1
End If
Next i

Dim sum As Double
sum = Cells(2, 7)
j = 2
For i = 3 To Cells(Rows.Count, 1).End(xlUp).Row
If Cells(i, 1) = Cells(i + 1, 1) Then
    sum = sum + Cells(i + 1, 7)
    Else: Cells(j, 12) = sum
    sum = 0
    j = j + 1
End If
Next i

Dim begin As Double
Dim finish As Double
Dim begincell As String
k = 2
begincell = Cells(2, 1)
begin = Cells(2, 3)
For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
 If begincell <> Cells(i + 1, 1) Then
 Cells(k, 10) = Cells(i, 6) - begin
 If begin <> 0 Then
 Cells(k, 11) = Cells(k, 10) / begin
 Else: Cells(k, 11) = 0
 End If
 If Cells(k, 10) > 0 Then
 Cells(k, 10).Interior.ColorIndex = 4
 ElseIf Cells(k, 10) < 0 Then
 Cells(k, 10).Interior.ColorIndex = 3
 End If
 k = k + 1
 begin = Cells(i + 1, 3)
 begincell = Cells(i + 1, 1)
 End If
    Next i

Dim max As Double
Dim maxline As Integer
max = Cells(2, 11)
maxline = 2
Dim min As Double
Dim minline As Integer
min = Cells(2, 11)
minline = 2
Dim maxV As Double
Dim maxlineV As Integer
maxV = Cells(2, 12)
maxlineV = 2

For i = 3 To Cells(Rows.Count, 11).End(xlUp).Row
If Cells(i, 11) > max Then
max = Cells(i, 11)
maxline = i
ElseIf Cells(i, 11) < min Then
min = Cells(i, 11)
minline = i
End If
If Cells(i, 12) > maxV Then
maxV = Cells(i, 12)
maxlineV = i
End If
Next i


Cells(2, 17) = max
Cells(2, 16) = Cells(maxline, 9)
Cells(3, 17) = min
Cells(3, 16) = Cells(minline, 9)
Cells(4, 17) = maxV
Cells(4, 16) = Cells(maxlineV, 9)

For k = 2 To Cells(Rows.Count, 1).End(xlUp).Row
Cells(k, 11).NumberFormat = "0.00%"
Next k
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"
Next ws


End Sub
