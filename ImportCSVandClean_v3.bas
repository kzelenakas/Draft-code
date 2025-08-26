Attribute VB_Name = "ImportCSVandClean"

Sub ImportCSVandClean()
    Dim ws As Worksheet
    Dim filePath As String
    Dim sheetExists As Boolean
    Dim sht As Worksheet

    ' Check if the "Data" sheet exists
    sheetExists = False
    For Each sht In ThisWorkbook.Sheets
        If sht.Name = "Data" Then
            sheetExists = True
            Set ws = sht
            Exit For
        End If
    Next sht

    ' If "Data" sheet doesn't exist, create it
    If Not sheetExists Then
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

    ' Run cleanup after import
    CleanUpHtmlEntities ws

    MsgBox "CSV file imported and cleaned successfully!", vbInformation
End Sub

Sub CleanUpHtmlEntities(ws As Worksheet)
    Dim cell As Range
    Dim replacements As Variant
    Dim i As Long

    replacements = Array( _
        "<*>", "", _
        " ", "", _
        """", "", _
        "’", "", _
        "”", "", _
        "&#39;", "", _
        ">", "", _
        "“", "", _
        "bull;", "", _
        "ndash;", "", _
        "amp;", "", _
        "½", "", _
        "‘", "", _
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
