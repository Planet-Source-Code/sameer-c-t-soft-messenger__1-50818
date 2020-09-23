Attribute VB_Name = "modEncryption"
Option Explicit

Public Function gfnDecodeMsg(strCodedMsg As String, strPassword As String) As String
    Dim intAsk As Integer
    Dim strStart As String, strEnd As String, strTmpPass As String
    Dim intCount As Integer, itstops As Variant
    Dim lngAccessNum As Long, lngCharac As Long
    Dim strNumber As String, strEdit  As String, strBeforeStart As String
    
    strBeforeStart = strCodedMsg
    strStart = strCodedMsg
    
    On Error GoTo errhandler


    Do
        strTmpPass = strPassword


        For intCount = 1 To Len(strPassword)
            strNumber = Mid(strTmpPass, 1, 1)
            lngAccessNum = Asc(strNumber)
            strTmpPass = Right(strTmpPass, Len(strTmpPass) - 1)
            strEdit = Mid(strStart, 1, 1)
            strStart = Right(strStart, Len(strStart) - 1)
            lngCharac = Asc(strEdit) - lngAccessNum
            strEnd = strEnd + Chr(Asc(strEdit) - lngAccessNum)
        Next intCount

    Loop Until itstops

errhandler:
    intAsk = InStr(strEnd, strPassword)


    If intAsk Then
        'a = StrComp(intAsk, strPassword)
        'If a = 1 Then MsgBox "Access Denied", 12, "Cipher": Exit Function
        strEnd = Right(strEnd, Len(strEnd) - Len(strPassword))
        'MsgBox "Text DeCiphered", 12, "Cipher"
        gfnDecodeMsg = strEnd
    Else
        'MsgBox "Access Denied", 12, "Cipher"
        gfnDecodeMsg = strBeforeStart
    End If

End Function

Public Function gfnEncodeMsg(strNormalMsg As String, strPassword As String)
    Dim strStart As String, strEnd As String, strBegin As String, strTmpPass As String
    Dim intCount As Integer, itstops As Variant
    Dim strNumber As String, strEdit  As String, strBeforeStart As String
    Dim lngAccessNum As Long, lngCharac As Long

    'If strPassword = "" Then MsgBox "Please put in a password", 12, "Cipher": Exit Function
    strStart = strPassword + strNormalMsg
    
    On Error GoTo errhandler


    Do
        strTmpPass = strPassword


        For intCount = 1 To Len(strPassword)
            strNumber = Mid(strTmpPass, 1, 1)
            lngAccessNum = Asc(strNumber)
            strTmpPass = Right(strTmpPass, Len(strTmpPass) - 1)
            strEdit = Mid(strStart, 1, 1)
            strStart = Right(strStart, Len(strStart) - 1)
            lngCharac = Asc(strEdit) + lngAccessNum
            strEnd = strEnd + Chr(Asc(strEdit) + lngAccessNum)
        Next intCount

    Loop Until itstops

errhandler:
    'MsgBox "Text Ciphered", 12, "Cipher"
    gfnEncodeMsg = strEnd
End Function


