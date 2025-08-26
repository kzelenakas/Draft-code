Attribute VB_Name = "Categorize"
Option Explicit

Private dict As Object
Private sortedKeys() As Variant
Private regExAlpha As Object
Private regExSpaces As Object

Public Sub CategorizeThemes()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim theme As String
    Dim uncategorizedCount As Long
    Dim totalCount As Long
    Dim i As Long, j As Long

    Set ws = ThisWorkbook.Worksheets("Pivot")
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

    ws.Range("L1").Value = "Theme"
    If lastRow >= 2 Then ws.Range("L2:L" & lastRow).ClearContents

    uncategorizedCount = 0
    totalCount = 0

    ' Create dictionary object for keywords and categories
    Set dict = CreateObject("Scripting.Dictionary")

    ' Populate keywords dictionary by calling sub from ModuleKeywords
    Call PopulateKeywordDictionary(dict) ' This procedure must exist in ModuleKeywords

    ' Normalize keys in dictionary: all should be lowercase & trimmed already in ModuleKeywords,
    ' but if not, normalization can be ensured before sorting here.

    ' Sort keys by descending length to prioritize longer/multi-word keywords
    If dict.Count > 0 Then
        ReDim sortedKeys(dict.Count - 1)
        i = 0
        Dim key As Variant
        For Each key In dict.Keys
            sortedKeys(i) = key
            i = i + 1
        Next key
        For i = LBound(sortedKeys) To UBound(sortedKeys) - 1
            For j = i + 1 To UBound(sortedKeys)
                If Len(CStr(sortedKeys(i))) < Len(CStr(sortedKeys(j))) Then
                    Dim tmp As Variant
                    tmp = sortedKeys(i)
                    sortedKeys(i) = sortedKeys(j)
                    sortedKeys(j) = tmp
                End If
            Next j
        Next i
    Else
        MsgBox "Keyword/category dictionary is empty. Cannot categorize.", vbExclamation
        Exit Sub
    End If

    ' Create RegExp objects for cleaning input text
    Set regExAlpha = CreateObject("VBScript.RegExp")
    regExAlpha.Pattern = "[^a-z\s]"
    regExAlpha.Global = True

    Set regExSpaces = CreateObject("VBScript.RegExp")
    regExSpaces.Pattern = "\s+"
    regExSpaces.Global = True

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Loop through input data in column G
    For Each cell In ws.Range("G2:G" & lastRow)
        If Not IsEmpty(cell) Then
            totalCount = totalCount + 1
            If Not IsError(cell.Value) And VarType(cell.Value) = vbString Then
                Dim cellValue As String
                cellValue = RemoveNonAlphabetic(LCase(Trim(cell.Value)), regExAlpha, regExSpaces)
                theme = FindBestTheme(cellValue)
                If theme = "" Then
                    theme = "No Primary noted"
                    uncategorizedCount = uncategorizedCount + 1
                End If
                cell.Offset(0, 5).Value = theme ' Output to column L (G + 5)
            Else
                cell.Offset(0, 5).Value = "No Primary noted"
                uncategorizedCount = uncategorizedCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    Set regExAlpha = Nothing
    Set regExSpaces = Nothing

    ' Debug report — optional: print some dictionary info in Immediate window
    Debug.Print "Total keywords loaded: " & dict.Count
    Debug.Print "Uncategorized count: " & uncategorizedCount & " out of " & totalCount

    If totalCount > 0 And (uncategorizedCount / totalCount) > 0.1 Then
        MsgBox "Warning: More than 10% of cells are uncategorized. Please refine the categorization logic.", vbExclamation
    End If

    MsgBox "Categorization complete! " & uncategorizedCount & " out of " & totalCount & " entries were uncategorized.", vbInformation

    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Set regExAlpha = Nothing
    Set regExSpaces = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "Line: " & Erl, vbCritical
End Sub

'===========================================
' Find best matching theme for the cleaned input string.
' Matching order:
' 1) Full phrase whole-word match (multi-word keys)
' 2) Whole word match (including optional plural "s")
' 3) Substring match fallback (least precise)
'===========================================
Private Function FindBestTheme(cellValue As String) As String
    Dim currentKey As Variant
    Dim re As Object
    Dim pattern As String
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = False

    ' Exact phrase match first (handle multiword keywords)
    For Each currentKey In sortedKeys
        If Len(currentKey) > 0 Then
            pattern = "\b" & Replace(currentKey, " ", "\s+") & "\b"
            re.Pattern = pattern
            If re.Test(cellValue) Then
                FindBestTheme = dict(currentKey)
                Exit Function
            End If
        End If
    Next currentKey

    ' Whole word match with optional trailing plural 's'
    For Each currentKey In sortedKeys
        If Len(currentKey) > 0 Then
            pattern = "\b" & currentKey & "s?\b"
            re.Pattern = pattern
            If re.Test(cellValue) Then
                FindBestTheme = dict(currentKey)
                Exit Function
            End If
        End If
    Next currentKey

    ' Substring fallback
    For Each currentKey In sortedKeys
        If Len(currentKey) > 1 Then
            If InStr(cellValue, currentKey) > 0 Then
                FindBestTheme = dict(currentKey)
                Exit Function
            End If
        End If
    Next currentKey

    FindBestTheme = ""
End Function

'===========================================
' Remove non-alphabetic characters and normalize whitespace in input
'===========================================
Private Function RemoveNonAlphabetic(ByVal inputString As String, ByVal alphaRegEx As Object, ByVal spacesRegEx As Object) As String
    Dim cleanedString As String
    cleanedString = alphaRegEx.Replace(inputString, " ")
    cleanedString = spacesRegEx.Replace(cleanedString, " ")
    RemoveNonAlphabetic = Trim(cleanedString)
End Function
