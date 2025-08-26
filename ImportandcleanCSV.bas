Attribute VB_Name = "ImportCleanCSV"
Sub ImportCSVToDataSheet()
    Dim ws As Worksheet
    Dim filePath As String

    ' Check if the "Data" sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Data")
    On Error GoTo 0

    ' If "Data" sheet doesn't exist, create it
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Data"
    End If

    ' Prompt user to select a CSV file
    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select CSV File")

    ' Exit if no file is selected
    If filePath = "False" Then Exit Sub

    ' Clear existing data in the "Data" sheet
    ws.Cells.Clear
    ' Import the CSV file
    With ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A1"))
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFilePlatform = xlWindows
        .TextFileParseType = xlDelimited
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

    MsgBox "CSV file imported successfully to the 'Data' sheet!", vbInformation
End Sub

Sub CleanUpHtmlEntities()
    Dim ws As Worksheet
    Dim cell As Range
    Dim replacements As Variant
    Dim i As Long

    Set ws = ActiveSheet

    replacements = Array( _
        "<*>", "", _
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
        "?", "" _
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
