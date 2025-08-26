Attribute VB_Name = "Module3"
Sub ImportExcelToDataAMC()
    Dim ws As Worksheet
    Dim wbImport As Workbook
    Dim filePath As String
    Dim sheetExists As Boolean
    Dim sht As Worksheet

    ' Check if "DataAMC" sheet exists
    sheetExists = False
    For Each sht In ThisWorkbook.Sheets
        If sht.Name = "DataAMC" Then
            sheetExists = True
            Set ws = sht
            Exit For
        End If
    Next sht

    ' If "DataAMC" sheet doesn't exist, create it
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "DataAMC"
    End If

    ' Prompt user to select an Excel file
    filePath = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx), *.xls; *.xlsx", , "Select Excel File")

    ' Exit if no file is selected
    If filePath = "False" Then Exit Sub

    ' Clear existing data in the "DataAMC" sheet
    ws.Cells.Clear

    ' Open the selected workbook
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set wbImport = Workbooks.Open(filePath)

    ' Copy data from the first sheet of the imported workbook
    wbImport.Sheets(1).UsedRange.Copy Destination:=ws.Range("A1")

    ' Close the imported workbook
    wbImport.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ' Run cleanup after import
    CleanUpHtmlEntities ws

    MsgBox "Excel file imported and cleaned successfully to the 'DataAMC' sheet!", vbInformation
End Sub

Sub CleanUpHtmlEntities(ws As Worksheet)
    Dim cell As Range
    Dim replacements As Variant
    Dim i As Long

    replacements = Array( _
        "&lt;*&gt;", "", _
        "&nbsp;", "", _
        "&quot;", "", _
        "&rsquo;", "", _
        "&rdquo;", "", _
        "&#39;", "", _
        "&gt;", "", _
        "&ldquo;", "", _
        "bull;", "", _
        "ndash;", "", _
        "amp;", "", _
        "&frac12;", "", _
        "&lsquo;", "", _
        "=-", "", _
        " ", "" _
    )

    Application.ScreenUpdating = False
    For Each cell In ws.UsedRange
        If Not IsError(cell.Value) Then
            If VarType(cell.Value) = vbString Then
                For i = 0 To UBound(replacements) Step 2
                    cell.Value = Replace(cell.Value, replacements(i), replacements(i + 1), , , vbTextCompare)
                Next i
            End If
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub
