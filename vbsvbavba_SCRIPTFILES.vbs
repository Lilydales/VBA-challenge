Option Explicit
Public Const StartColumn = 10

Sub Run_Analysis()
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    If IsNumeric(ws.Name) Then Module1.CheckThroughTicket (ws.Name)
Next ws

End Sub

Sub BonusAssignment(strSheetName)
Dim ws As Worksheet

Set ws = ThisWorkbook.Sheets(strSheetName)

ws.Cells(3, 17) = "Greatest % Increase"
ws.Cells(3, 19) = Format(WorksheetFunction.Max(ws.Range("L:L")), "0.00%")
ws.Cells(3, 18) = WorksheetFunction.XLookup(ws.Cells(3, 19), ws.Range("L:L"), ws.Range("J:J"))

ws.Cells(4, 17) = "Greatest % Decrease"
ws.Cells(4, 19) = Format(WorksheetFunction.Min(ws.Range("L:L")), "0.00%")
ws.Cells(4, 18) = WorksheetFunction.XLookup(ws.Cells(4, 19), ws.Range("L:L"), ws.Range("J:J"))

ws.Cells(5, 17) = "Greatest Total Volume"
ws.Cells(5, 19) = WorksheetFunction.Max(ws.Range("M:M"))
ws.Cells(5, 18) = WorksheetFunction.XLookup(ws.Cells(5, 19), ws.Range("M:M"), ws.Range("J:J"))

ws.Cells(2, 18) = "Ticker"
ws.Cells(2, 19) = "Value"

End Sub

Sub AddFormatCondition(strSheetName)
Dim ws As Worksheet
Dim TotalRow As Long

Set ws = ThisWorkbook.Sheets(strSheetName)
TotalRow = WorksheetFunction.CountA(ws.Range("L:L")) + 1
On Error Resume Next

With ws.Range("L3:L" & TotalRow)
.FormatConditions.Delete
.FormatConditions.Add xlCellValue, xlLess, 0
.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
End With

With ws.Range("L3:L" & TotalRow)
.FormatConditions.Add xlCellValue, xlGreater, 0
.FormatConditions(2).Interior.Color = RGB(0, 255, 0)
End With

End Sub

Sub CheckThroughTicket(strSheetName)
Dim TotalRow, i, NextRow, k As Long
Dim strDate As String
Dim ws As Worksheet


Set ws = ThisWorkbook.Sheets(strSheetName)
TotalRow = WorksheetFunction.CountA(ws.Range("A:A"))

Module1.CreateNewColumns (strSheetName)

For i = 2 To TotalRow
    'If k = 350 Then Exit For
    If WorksheetFunction.CountIf(ws.Range("J:J"), ws.Cells(i, 1)) = 0 Then
        NextRow = WorksheetFunction.CountA(ws.Range("J:J")) + 2
        ws.Cells(NextRow, 10) = ws.Cells(i, 1)
        'ws.Cells(NextRow, 11) = ws.Cells(NextRow, 11) + ws.Cells(i, 6) - ws.Cells(i, 3)
        'ws.Cells(NextRow, 12) = ws.Cells(NextRow, 12) + ws.Cells(i, 2)
        ws.Cells(NextRow, 13) = ws.Cells(i, 7)
        
        strDate = ws.Name & "0102"
        If ws.Cells(i, 2) = strDate Then ws.Cells(NextRow, 14) = ws.Cells(i, 3)
        strDate = ws.Name & "1231"
        If ws.Cells(i, 2) = strDate Then
        ws.Cells(NextRow, 15) = ws.Cells(i, 6)
        strDate = "YearEnd"
        End If
        k = k + 1
    Else
        NextRow = WorksheetFunction.Match(ws.Cells(i, 1), ws.Range("J1:J1000000"))
        'ws.Cells(NextRow, 11) = ws.Cells(NextRow, 11) + ws.Cells(i, 6) - ws.Cells(i, 3)
        'ws.Cells(NextRow, 12) = ws.Cells(NextRow, 12) + ws.Cells(i, 2)
        ws.Cells(NextRow, 13) = ws.Cells(NextRow, 13) + ws.Cells(i, 7)
        
        strDate = ws.Name & "0102"
        If ws.Cells(i, 2) = strDate Then ws.Cells(NextRow, 14) = ws.Cells(i, 3)
        strDate = ws.Name & "1231"
        If ws.Cells(i, 2) = strDate Then
        ws.Cells(NextRow, 15) = ws.Cells(i, 6)
        strDate = "YearEnd"
        End If
        k = k + 1
    End If
    If strDate = "YearEnd" Then
    ws.Cells(NextRow, 11) = ws.Cells(NextRow, 15) - ws.Cells(NextRow, 14)
    ws.Cells(NextRow, 12) = Format(ws.Cells(NextRow, 11) / ws.Cells(NextRow, 14), "0.00%")
    End If
Next i

Call Module1.AddFormatCondition(strSheetName)
Call Module1.BonusAssignment(strSheetName)

ws.Range("J1:S" & NextRow).EntireColumn.AutoFit
End Sub


Sub CreateNewColumns(strSheetName)
Dim ws As Worksheet
Dim rgHeader(5) As String
Dim i, j As Integer

Set ws = ThisWorkbook.Sheets(strSheetName)
'Set Header Name
rgHeader(0) = "Ticker"
rgHeader(1) = "Yearly Change"
rgHeader(2) = "Percent Change"
rgHeader(3) = "Total Stock Volume"
rgHeader(4) = "Starting Price"
rgHeader(5) = "Ending Price"

If ws.Cells(2, StartColumn) = "" Then
    For i = StartColumn To 15
        ws.Cells(2, i) = rgHeader(j)
        j = j + 1
    Next i
End If
End Sub
