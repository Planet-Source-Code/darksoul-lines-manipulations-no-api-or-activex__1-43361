Attribute VB_Name = "modStrLine"
'Hi! i wanna tell you that
'everything of this project
'was created by [[.DarkSouL.]]
' -----------------------
'| MSN: dark_soul@123.cl |
' -----------------------
'
'Thanks for voting...! :P
'Lol!

Function ReadLines(Text As String, StartLine, EndLine) As String

    On Error GoTo MsgAlert
    Dim Lines As Variant
    
    Lines = Split(Text, vbNewLine)
    ActualLine = StartLine

    Do Until ActualLine > EndLine
        ReadLines = ReadLines & Lines(ActualLine - 1) & vbNewLine
        ActualLine = ActualLine + 1
    Loop

    Exit Function
    
MsgAlert:
    MsgBox "Error reading some lines from a string.", vbCritical, "Read Lines"
End Function
Function GetTotalLines(Text As String) As Long

    Dim LineNumber%, tmpString$
    LineNumber = 1
    tmpString = Text

    Do While InStr(tmpString, Chr(13))
        tmpString = Mid$(tmpString, InStr(tmpString, Chr(13)) + 1)
        LineNumber = LineNumber + 1
    Loop

    GetTotalLines = CLng(LineNumber)
    
End Function

Function DetermineLine(Text As TextBox) As Long

    On Error Resume Next

    Dim strTemp$, Line%, CurrentPos%

    CurrentPos = Text.SelStart
    strTemp = Left(Text, CurrentPos)

    If Not strTemp Like "*" & vbNewLine Then strTemp = strTemp & vbNewLine Else: Line = Line + 1

    DoEvents

    Do While InStr(strTemp, Chr$(13))
        Line = Line + 1
        strTemp = Mid$(strTemp, InStr(strTemp, vbNewLine) + 2)
    Loop

    DetermineLine = CLng(Line)

End Function

Function DetermineColumn(Text As TextBox) As Long

    On Error Resume Next
    Dim strTemp$, CurrentPos%

    CurrentPos = Text.SelStart
    strTemp = StrReverse(Left(Text, CurrentPos))
    
    If InStr(strTemp, Chr(10)) > 0 Then
        strTemp = Left(strTemp, InStr(strTemp, Chr(10)) - 1)
        If strTemp = "" Then DetermineColumn = 1 Else: DetermineColumn = Len(strTemp) + 1
            Else
        DetermineColumn = Len(strTemp) + 1
    End If

End Function

Sub SetLineAndColumn(Text As TextBox, Optional LineNumber = 1, Optional Column = 1)

    On Error Resume Next
    Dim strTemp$, strText$, tmpLine%
    
    Do Until tmpLine + 1 >= LineNumber
        strTemp2 = Mid(Text, Len(strText) + 1, Len(Text))
        strTemp = Left$(strTemp2, InStr(strTemp2, Chr(10)))
        strText = strText & strTemp
        tmpLine = tmpLine + 1
    Loop
    
    strTemp = Left$(strTemp2, InStr(strTemp2, Chr(10)) - 2)
    If Column - 1 > Len(strTemp) Then Column = Len(strTemp) + 1
    
    Text.SelStart = Len(strText) + (Column - 1)
    Text.SetFocus

End Sub

