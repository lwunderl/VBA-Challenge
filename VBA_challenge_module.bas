Attribute VB_Name = "VBA_challenge_module"
Sub header()
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Columns("K").NumberFormat = "0.00%"
Range("L1").Value = "Total Stock Volume"
Columns("I:L").AutoFit
End Sub
Sub greatest_header()
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("Q2").NumberFormat = "0.00%"
Range("O3").Value = "Greatest % Decrease"
Range("Q3").NumberFormat = "0.00%"
Range("O4").Value = "Greatest Total Volume"
Columns("O").AutoFit
Columns("P:Q").ColumnWidth = 10
End Sub
Sub ticker()
Dim ticker As String
Dim r As Integer
Dim last_row As Long

r = 2
last_row = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
ticker = Cells(r, 1).Value

For i = 2 To last_row
    If Cells(i, 1).Value = ticker Then
    Cells(r, 9).Value = ticker
    Else
    r = r + 1
    ticker = Cells(i, 1).Value
    Cells(r, 9).Value = ticker
    End If
Next i

End Sub
Sub yearly_change()
Dim year_open As Double
Dim year_close As Double
Dim r As Integer
Dim last_row As Long
Dim ticker As String
Dim first_day As Long
Dim last_day As Long

r = 2
ticker = Cells(r, 9).Value
last_row = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
year_open = 0
year_close = 0
first_day = Cells(2, 2).Value
last_day = Cells(2, 2).Value

For i = 2 To last_row
    If Cells(i, 1).Value = ticker And Cells(i, 2).Value <= first_day Then
        first_day = Cells(i, 2).Value
        year_open = Cells(i, 3).Value
    
    ElseIf Cells(i, 1).Value = ticker And Cells(i, 2).Value >= last_day Then
        last_day = Cells(i, 2).Value
        year_close = Cells(i, 6).Value
        Cells(r, 10).Value = year_close - year_open
        Cells(r, 11).Value = (year_close / year_open) - 1
        
    ElseIf Cells(i, 1).Value <> ticker Then
        r = r + 1
        ticker = Cells(r, 9).Value
        year_open = Cells(i, 3).Value
        year_close = Cells(i, 3).Value
        first_day = Cells(i, 2).Value
        last_day = Cells(i, 2).Value
    End If
Next i

End Sub
Sub total_volume()
Dim r As Integer
Dim last_row As Long
Dim ticker As String
Dim first_day As Long
Dim last_day As Long

r = 2
ticker = Cells(r, 9).Value
last_row = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

For i = 2 To last_row
    If Cells(i, 1).Value = ticker Then
        Cells(r, 12).Value = Cells(r, 12).Value + Cells(i, 7).Value
        
    ElseIf Cells(i, 1).Value <> ticker Then
        r = r + 1
        ticker = Cells(r, 9).Value
        Cells(r, 12).Value = Cells(r, 12).Value + Cells(i, 7).Value
    End If
Next i

End Sub
Sub color_format()
Dim last_row As Long

last_row = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

For i = 2 To last_row
    If Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
        Cells(i, 11).Interior.ColorIndex = 3
    ElseIf Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
        Cells(i, 11).Interior.ColorIndex = 4
End If
Next i

End Sub
Sub greatest_increase()
Dim percentage As Double
Dim last_row As Long
Dim ticker As String

last_row = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
percentage = 0
ticker = "string"

For i = 2 To last_row
    If Cells(i, 11).Value > percentage Then
        percentage = Cells(i, 11).Value
        ticker = Cells(i, 9).Value
        Range("P2").Value = ticker
        Range("Q2").Value = percentage
    End If
Next i
End Sub
Sub greatest_decrease()
Dim percentage As Double
Dim last_row As Long
Dim ticker As String

last_row = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
percentage = 0
ticker = "string"

For i = 2 To last_row
    If Cells(i, 11).Value < percentage Then
        percentage = Cells(i, 11).Value
        ticker = Cells(i, 9).Value
        Range("P3").Value = ticker
        Range("Q3").Value = percentage
    End If
Next i
End Sub
Sub greatest_volume()
Dim volume As LongLong
Dim last_row As Long
Dim ticker As String

last_row = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
volume = 0
ticker = "string"

For i = 2 To last_row
    If Cells(i, 12).Value > volume Then
        volume = Cells(i, 12).Value
        ticker = Cells(i, 9).Value
        Range("P4").Value = ticker
        Range("Q4").Value = volume
    End If
Next i
End Sub
Sub main()
Dim ws_count As Integer
ws_count = ActiveWorkbook.Worksheets.Count
For i = 1 To ws_count
Worksheets(i).Activate
    header
    greatest_header
    ticker
    yearly_change
    total_volume
    color_format
    greatest_increase
    greatest_decrease
    greatest_volume
Next i

End Sub
