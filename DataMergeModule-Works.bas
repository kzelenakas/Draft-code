Attribute VB_Name = "Module11"
Sub CreateDataMergedSheet()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim sheetName As String
    sheetName = "DataMerged"
    
    ' Check if the sheet already exists
    sheetExists = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            sheetExists = True
            Exit For
        End If
    Next ws

    ' Create or activate the sheet
    If sheetExists Then
        Set ws = ThisWorkbook.Sheets(sheetName)
        ws.Activate
    Else
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
        ws.Activate
    End If

    ' Move to the next macro
    MergeToDataMerged
End Sub

Sub MergeToDataMerged()
    Dim wsAMC As Worksheet, wsStaff As Worksheet, wsMerged As Worksheet
    Dim lastRowAMC As Long, lastRowStaff As Long, lastRowMerged As Long
    Dim lastCol As Long

    ' Safely reference source sheets
    On Error Resume Next
    Set wsAMC = ThisWorkbook.Sheets("DataAMC")
    Set wsStaff = ThisWorkbook.Sheets("DataStaff")
    On Error GoTo 0
    
    If wsAMC Is Nothing Or wsStaff Is Nothing Then
        MsgBox "Source sheets 'DataAMC' or 'DataStaff' not found.", vbExclamation
        Exit Sub
    End If

    ' Create or clear DataMerged sheet
    On Error Resume Next
    Set wsMerged = ThisWorkbook.Sheets("DataMerged")
    On Error GoTo 0

    If wsMerged Is Nothing Then
        Set wsMerged = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsMerged.Name = "DataMerged"
       ' MsgBox "'DataMerged' sheet was created.", vbInformation
    Else
        wsMerged.Cells.Clear
      '  MsgBox "'DataMerged' sheet already exists and has been cleared.", vbInformation
    End If

    ' Find last column in DataAMC
    lastCol = wsAMC.Cells(1, wsAMC.Columns.Count).End(xlToLeft).Column

    ' Copy headers from DataAMC with formatting
    wsAMC.Range(wsAMC.Cells(1, 1), wsAMC.Cells(1, lastCol)).Copy
    wsMerged.Cells(1, 1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    Application.CutCopyMode = False

    ' Copy data from DataAMC
    lastRowAMC = wsAMC.Cells(wsAMC.Rows.Count, 1).End(xlUp).Row
    If lastRowAMC > 1 Then
        wsAMC.Range(wsAMC.Cells(2, 1), wsAMC.Cells(lastRowAMC, lastCol)).Copy
        wsMerged.Cells(2, 1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        Application.CutCopyMode = False
    End If

    ' Copy data from DataStaff (excluding header)
    lastRowStaff = wsStaff.Cells(wsStaff.Rows.Count, 1).End(xlUp).Row
    If lastRowStaff > 1 Then
        lastRowMerged = wsMerged.Cells(wsMerged.Rows.Count, 1).End(xlUp).Row
        wsStaff.Range(wsStaff.Cells(2, 1), wsStaff.Cells(lastRowStaff, lastCol)).Copy
        wsMerged.Cells(lastRowMerged + 1, 1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        Application.CutCopyMode = False
    End If

  '  MsgBox "Data merged successfully into 'DataMerged' with formatting preserved!", vbInformation
    
    FormatColumnsAndRows wsMerged
End Sub

Sub FormatColumnsAndRows(ws As Worksheet)
    Dim colA As Range, colB As Range, colF As Range, colG As Range

    ' Format columns
    Set colA = ws.Columns("A")
    colA.ColumnWidth = 30
    colA.WrapText = True

    Set colB = ws.Columns("B")
    colB.ColumnWidth = 20
    colB.WrapText = True

    Set colF = ws.Columns("F")
    colF.ColumnWidth = 15
    colF.WrapText = True

    Set colG = ws.Columns("G")
    colG.ColumnWidth = 75
    colG.WrapText = True

    ' Autofit row heights
    ws.Rows.AutoFit

    ' Hide columns
    ws.Columns("C").Hidden = True
    ws.Columns("E").Hidden = True
    ws.Columns("H").Hidden = True
    ws.Columns("I").Hidden = True

    ' Vertical alignment
    With ws
        .Columns("A").VerticalAlignment = xlCenter
        .Columns("B").VerticalAlignment = xlCenter
        .Columns("F").VerticalAlignment = xlCenter
    End With

  '  MsgBox "Vertical alignment set to center for columns A, B, and F.", vbInformation
End Sub
