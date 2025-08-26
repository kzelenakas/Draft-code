Attribute VB_Name = "Module1"
Sub ImportExcelToStatsSheet()
    Dim wsStats As Worksheet
    Dim wbImport As Workbook
    Dim filePath As String
    Dim sheetExists As Boolean
    Dim sht As Worksheet

    ' Check if "Stats" sheet exists
    sheetExists = False
    For Each sht In ThisWorkbook.Sheets
        If sht.Name = "Stats" Then
            sheetExists = True
            Set wsStats = sht
            Exit For
        End If
    Next sht

    ' If "Stats" sheet doesn't exist, create it
    If Not sheetExists Then
        Set wsStats = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsStats.Name = "Stats"
    End If

    ' Prompt user to select an Excel file
    filePath = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx), *.xls; *.xlsx", , "Select Excel File")

    ' Exit if no file is selected
    If filePath = "False" Then Exit Sub

    ' Open the selected workbook
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set wbImport = Workbooks.Open(filePath)

    ' Copy data from the first sheet of the imported workbook
    With wbImport.Sheets(1).UsedRange
        wsStats.Cells.Clear
        .Copy Destination:=wsStats.Range("A1")
    End With

    ' Close the imported workbook
    wbImport.Close SaveChanges:=False
    Application.DisplayAlerts = True

    ' Run header consolidation
    ConsolidateHeaderRows wsStats

    ' Autofit all columns
    wsStats.Columns.AutoFit

    ' Delete rows where column A contains "TERMINATED"
    DeleteTerminatedRows wsStats

    ' Delete the last row
    DeleteLastRow wsStats

    Application.ScreenUpdating = True

    MsgBox "Import complete: headers consolidated, columns autofit, 'TERMINATED' rows and last row removed!", vbInformation
End Sub

Sub ConsolidateHeaderRows(ws As Worksheet)
    Dim lastCol As Long
    Dim i As Long
    Dim headerValue As String

    ' Unmerge all cells in the first two rows
    ws.Range("1:2").UnMerge

    ' Find the last column with data in row 1 or 2
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Combine row 1 and row 2 into row 1
    For i = 1 To lastCol
        headerValue = Trim(ws.Cells(1, i).Value) & " " & Trim(ws.Cells(2, i).Value)
        ws.Cells(1, i).Value = Trim(headerValue)
    Next i

    ' Delete row 2
    ws.Rows(2).Delete
End Sub

Sub DeleteTerminatedRows(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop from bottom to top to avoid skipping rows after deletion
    For i = lastRow To 2 Step -1 ' Start from row 2 to preserve headers
        If InStr(1, ws.Cells(i, 1).Value, "TERMINATED", vbTextCompare) > 0 Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

Sub DeleteLastRow(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then
        ws.Rows(lastRow).Delete
    End If
End Sub
