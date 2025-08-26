Attribute VB_Name = "ClaudeModule1"
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
    Dim totalCount As Long
    Dim uncategorizedCount As Long
    Dim i As Long, j As Long
    
    ' Updated to use column G as input and column L as output
    Set ws = ThisWorkbook.Worksheets("Pivot")
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ws.Range("L1").Value = "Theme"
    If lastRow >= 2 Then ws.Range("L2:L" & lastRow).ClearContents
    
    totalCount = 0
    uncategorizedCount = 0
    
    ' Create and populate dictionary using ModuleKeywords
    Set dict = CreateObject("Scripting.Dictionary")
    Call PopulateKeywordDictionary(dict)
    
    ' Validate dictionary was populated
    If dict.Count = 0 Then
        MsgBox "Keyword dictionary is empty. Cannot categorize.", vbCritical
        Exit Sub
    End If
    
    ' Sort keys by descending length for better matching priority
    ReDim sortedKeys(dict.Count - 1)
    i = 0
    Dim key As Variant
    For Each key In dict.Keys
        sortedKeys(i) = key
        i = i + 1
    Next key
    
    ' Bubble sort by length (longest first)
    For i = LBound(sortedKeys) To UBound(sortedKeys) - 1
        For j = i + 1 To UBound(sortedKeys)
            If Len(sortedKeys(i)) < Len(sortedKeys(j)) Then
                Dim temp As Variant
                temp = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = temp
            End If
        Next j
    Next i
    
    ' Create RegExp objects for text cleaning
    Set regExAlpha = CreateObject("VBScript.RegExp")
    regExAlpha.Pattern = "[^a-z\s]"
    regExAlpha.Global = True
    
    Set regExSpaces = CreateObject("VBScript.RegExp")
    regExSpaces.Pattern = "\s+"
    regExSpaces.Global = True
    
    ' Optimize Excel performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Process each cell in column G (updated from D)
    For Each cell In ws.Range("G2:G" & lastRow)
        If Not IsEmpty(cell) Then
            totalCount = totalCount + 1
            If VarType(cell.Value) = vbString And Not IsError(cell.Value) Then
                Dim cellText As String
                cellText = RemoveNonAlphabetic(LCase(Trim(cell.Value)), regExAlpha, regExSpaces)
                theme = FindBestTheme(cellText)
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
    
    ' Restore Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Cleanup objects
    Set regExAlpha = Nothing
    Set regExSpaces = Nothing
    
    ' Show completion message
    If totalCount > 0 And (uncategorizedCount / totalCount) > 0.1 Then
        MsgBox "Warning: More than 10% uncategorized. Consider refining keywords.", vbExclamation
    End If
    
    MsgBox "Categorization complete! Uncategorized: " & uncategorizedCount & " of " & totalCount, vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Set regExAlpha = Nothing
    Set regExSpaces = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description & " at line " & Erl, vbCritical
End Sub

Private Function FindBestTheme(text As String) As String
    Dim key As Variant
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = False
    
    ' 1. Exact phrase match (multiword keywords)
    For Each key In sortedKeys
        If Len(key) > 0 Then
            re.Pattern = "\b" & Replace(key, " ", "\s+") & "\b"
            If re.Test(text) Then
                FindBestTheme = dict(key)
                Exit Function
            End If
        End If
    Next key
    
    ' 2. Whole word match with optional plural 's'
    For Each key In sortedKeys
        If Len(key) > 0 Then
            re.Pattern = "\b" & key & "s?\b"
            If re.Test(text) Then
                FindBestTheme = dict(key)
                Exit Function
            End If
        End If
    Next key
    
    ' 3. Substring match fallback
    For Each key In sortedKeys
        If Len(key) > 1 Then
            If InStr(text, key) > 0 Then
                FindBestTheme = dict(key)
                Exit Function
            End If
        End If
    Next key
    
    FindBestTheme = ""
End Function

Private Function RemoveNonAlphabetic(inputStr As String, alphaRegEx As Object, spacesRegEx As Object) As String
    Dim cleanedStr As String
    cleanedStr = alphaRegEx.Replace(inputStr, " ")
    cleanedStr = spacesRegEx.Replace(cleanedStr, " ")
    RemoveNonAlphabetic = Trim(cleanedStr)
End Function
