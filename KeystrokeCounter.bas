Attribute VB_Name = "Tastschreiben"
' select in Sub 'addKeystrokeCountToSecondTable' whether the behavior for line breaks should be queried or not
Public queryIgnoreLineBreak As Boolean
Public ignoreLineBreak As Boolean

Sub countKeystrokes()
    queryIgnoreLineBreak = True ' select in Sub 'addKeystrokeCountToSecondTable' whether the behavior for line breaks should be queried or not
    If queryIgnoreLineBreak Then
        ignoreLineBreak = checkUpIgnoreLineBreak
    Else
        ignoreLineBreak = False ' default value of line break behaviour. Select if line breaks should be ignored or not
    End If
    Dim line As String
    Dim keystrokes, keystrokesOld As Integer
    Dim keystrokesList As New Collection
    
    goToTable 2, 2, True
    goToTable 2, 1, False
    Do While Selection.Information(wdWithInTable)
        line = getActualLine
        keystrokesOld = keystrokes
        keystrokes = keystrokes + countKeystrokesFromLine(line)
        
        If ignoreLineBreak And keystrokes = keystrokesOld Then
            keystrokesList.Add ""
        Else
            keystrokesList.Add keystrokes
        End If
        Selection.MoveDown wdLine, 1
    Loop
    writeKeystrokesToTable keystrokesList
    Selection.HomeKey Unit:=wdStory
End Sub
Private Function checkUpIgnoreLineBreak() As Boolean
    Dim antwort As VbMsgBoxResult
    antwort = MsgBox("Möchten Sie Zeilenumbrüche (Absätze) ignorieren?", vbYesNo, "Bestätigung")
    
    If antwort = vbYes Then
        checkUpIgnoreLineBreak = True
    Else
        checkUpIgnoreLineBreak = False
    End If
End Function
Private Function isDoubleKeystroke(character As String) As Boolean
    ' add all characters with double keystroke in  list
    doubleKeystrokes = Array("€", "\", "{", "[", "]", "}", "²", "³", "°", "!", """", "§", "$", "%", "&", "/", "(", ")", "=", "?", "*", ">", ";", ":", "_", "@", "|", "'")
    For i = LBound(doubleKeystrokes) To UBound(doubleKeystrokes)
        If doubleKeystrokes(i) = character Then
            isDoubleKeystroke = True
            Exit Function
        End If
    Next i
    isDoubleKeystroke = False
End Function
Private Function goToTable(table As Integer, column As Integer, delete As Boolean)
    Dim doc As Document
    Dim tbl As table
    Dim rng As Range
    
    Set doc = ActiveDocument
    Selection.HomeKey Unit:=wdStory
    If doc.Tables.Count < table Then
        MsgBox "Das Dokument hat nicht gen gend Tabellen.", vbExclamation
        Exit Function
    End If
    Set tbl = doc.Tables(table)
    Set rng = tbl.Cell(1, column).Range
    If delete Then
        rng.delete
    End If
    rng.Collapse Direction:=wdCollapseStart
    rng.Select
End Function
Private Function getActualLine() As String
    Dim startRange As Range
    Dim endRange As Range
    Dim markRange As Range
    
    Selection.HomeKey Unit:=wdLine
    Set startRange = Selection.Range
    Selection.EndKey Unit:=wdLine
    Set endRange = Selection.Range
    
    If Not ignoreLineBreak And Mid(Selection.Text, currentPosition + 1, 1) = vbCr Then
        Set markRange = ActiveDocument.Range(startRange.End, endRange.Start + 1)
    Else
        Set markRange = ActiveDocument.Range(startRange.End, endRange.Start)
    End If
    getActualLine = markRange.Text
End Function
Private Function countKeystrokesFromLine(inputString As String) As Integer
    Dim keystrokes As Integer
    Dim character As String
    keystrokes = 0
    For i = 1 To Len(inputString)
        keystrokes = keystrokes + getKeystrokeFromCharacter(Mid(inputString, i, 1))
    Next i
    countKeystrokesFromLine = keystrokes
End Function
Private Sub writeKeystrokesToTable(keystrokesList As Collection)
    goToTable 2, 2, True
    Selection.Range.InsertAfter keystrokesList(keystrokesList.Count)
    For i = keystrokesList.Count - 1 To 1 Step -1
        Selection.Range.InsertAfter keystrokesList(i) & vbLf
    Next i
End Sub
Private Function getKeystrokeFromCharacter(character As String)
    If character = "…" Then
       getKeystrokeFromCharacter = 3
    ElseIf character Like "[A-Z]" Or isDoubleKeystroke(character) Then
        getKeystrokeFromCharacter = 2
    Else
        getKeystrokeFromCharacter = 1
    End If
End Function
