Attribute VB_Name = "XLStoStaff"
Sub ImportExcelToDataStaff()
    Dim ws As Worksheet
    Dim wbImport As Workbook
    Dim filePath As String
    Dim sheetExists As Boolean
    Dim sht As Worksheet

    ' Check if "DataStaff" sheet exists
    sheetExists = False
    For Each sht In ThisWorkbook.Sheets
        If sht.Name = "DataStaff" Then
            sheetExists = True
            Set ws = sht
            Exit For
        End If
    Next sht

    ' If "DataStaff" sheet doesn't exist, create it
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "DataStaff"
    End If

    ' Prompt user to select an Excel file
    filePath = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx), *.xls; *.xlsx", , "Select Excel File")

    ' Exit if no file is selected
    If filePath = "False" Then Exit Sub

    ' Clear existing data in the "DataStaff" sheet
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

    ' Run cleanup after import
    CleanUpHtmlEntities ws
    RemoveHtmlTags ws
    FormatColumnsAndRows ws

    Application.ScreenUpdating = True

    MsgBox "Excel file imported, cleaned, and formatted successfully to the 'DataStaff' sheet!", vbInformation
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

    For Each cell In ws.UsedRange
        If Not IsError(cell.Value) Then
            If VarType(cell.Value) = vbString Then
                For i = 0 To UBound(replacements) Step 2
                    cell.Value = Replace(cell.Value, replacements(i), replacements(i + 1), , , vbTextCompare)
                Next i
            End If
        End If
    Next cell
End Sub

Sub RemoveHtmlTags(ws As Worksheet)
    Dim cell As Range
    Dim regEx As Object

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.Pattern = "<[^>]+>"

    For Each cell In ws.UsedRange
        If Not IsError(cell.Value) Then
            If VarType(cell.Value) = vbString Then
                If regEx.Test(cell.Value) Then
                    cell.Value = regEx.Replace(cell.Value, "")
                End If
            End If
        End If
    Next cell
End Sub

Sub FormatColumnsAndRows(ws As Worksheet)
    Dim colG As Range

    ' Autofit all columns
    ws.Columns.AutoFit

    ' Format column G: set width and wrap text
    Set colG = ws.Columns("G")
    colG.ColumnWidth = 50
    colG.WrapText = True

    ' Autofit row heights
    ws.Rows.AutoFit

    ' Hide columns C, D, E, H, I, J, K
    ws.Columns("C").Hidden = True
    ws.Columns("D").Hidden = True
    ws.Columns("E").Hidden = True
    ws.Columns("H").Hidden = True
    ws.Columns("I").Hidden = True
    ws.Columns("J").Hidden = True
    ws.Columns("K").Hidden = True
End Sub

