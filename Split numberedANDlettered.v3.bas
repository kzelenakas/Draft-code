Attribute VB_Name = "SplitNumberedOrLetteredItemsInColumnG"
Sub SplitNumberedOrLetteredItemsInColumnG()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

    Dim i As Long
    For i = lastRow To 1 Step -1
        Dim cell As Range
        Set cell = ws.Cells(i, "G")

        Dim text As String
        text = Trim(cell.Value)

        If text <> "" Then
            Dim splitItems As Variant
            splitItems = SplitByNumberOrLetter(text)

            If UBound(splitItems) > 0 Then
                Dim j As Long
                For j = UBound(splitItems) To 1 Step -1
                    ws.Rows(i + 1).Insert Shift:=xlDown
                    ws.Cells(i + 1, "G").Value = RemoveNumberingOrLettering(Trim(splitItems(j)))
                Next j

                ' Replace original row with first item
                ws.Cells(i, "G").Value = RemoveNumberingOrLettering(Trim(splitItems(0)))
            Else
                ws.Cells(i, "G").Value = RemoveNumberingOrLettering(Trim(splitItems(0)))
            End If
        End If
    Next i
End Sub

Function SplitByNumberOrLetter(text As String) As Variant
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    ' Matches: 1. 2. 1) 2) a. b. a) b) and so on
    regex.Pattern = "(\d+\)|\d+\.)|([a-zA-Z]\)|[a-zA-Z]\.)"

    Dim matches As Object
    Set matches = regex.Execute(text)

    If matches.Count = 0 Then
        SplitByNumberOrLetter = Array(text)
        Exit Function
    End If

    Dim result() As String
    ReDim result(matches.Count - 1)

    Dim i As Long, startPos As Long, endPos As Long
    For i = 0 To matches.Count - 1
        startPos = matches(i).FirstIndex + Len(matches(i).Value)
        If i < matches.Count - 1 Then
            endPos = matches(i + 1).FirstIndex
        Else
            endPos = Len(text)
        End If
        result(i) = Mid(text, startPos + 1, endPos - startPos)
    Next i

    SplitByNumberOrLetter = result
End Function

Function RemoveNumberingOrLettering(itemText As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    ' Remove leading numbering/lettering patterns with optional space
    regex.Pattern = "^\s*(\d+\)|\d+\.)|^\s*([a-zA-Z]\)|[a-zA-Z]\.)"

    RemoveNumberingOrLettering = Trim(regex.Replace(itemText, ""))
End Function
