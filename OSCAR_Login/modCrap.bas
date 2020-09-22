Attribute VB_Name = "modCRAP"
Option Explicit
Public Function Qt(strString As String) As String
    'encircle a string in quotes
    Qt = Chr(34) & strString & Chr(34)
End Function
Public Function StringBlock(strString As String) As String
    'string block
    StringBlock = Chr(Len(strString)) & strString
End Function
Public Function GetStringBlock(strString As String, lngStart As Long) As String
    'get string block
    GetStringBlock = Mid$(strString, lngStart, Asc(Mid$(strString, lngStart - 1, 1)))
End Function
Public Function ChrH(strString As String) As String
    'hex to ascii
    Dim strHex() As String
    strHex = Split(strString$, " ")
    Dim i As Integer
    For i = 0 To UBound(strHex)
        ChrH = ChrH & Chr("&H" & strHex(i))
        DoEvents
    Next i
End Function
Public Function ChrA(strString As String) As String
    'decimal to ascii
    Dim strDec() As String
    strDec = Split(strString$, " ")
    Dim i As Integer
    For i = 0 To UBound(strDec)
        ChrA = ChrA & Chr(strDec(i))
        DoEvents
    Next i
End Function
Public Function AscToDec(strString As String) As String
    'ascii to decimal
    Dim i As Integer
    For i = 1 To Len(strString)
        AscToDec = AscToDec & Asc(Mid$(strString, i, 1))
        If i <> Len(strString) Then AscToDec = AscToDec & " "
        DoEvents
    Next i
End Function
Public Function AscToHex(strString As String) As String
    'ascii to hex
    Dim i As Integer
    For i = 1 To Len(strString)
        AscToHex = AscToHex & Format$(Hex$(Asc%(Mid$(strString, i, 1))), "00")
        If i <> Len(strString) Then AscToHex = AscToHex & " "
        DoEvents
    Next i
End Function
Public Function ChrUUID(strUUID As String) As String
    'cap uuid to ascii
    strUUID = Replace$(strUUID, "-", vbNullString)
    Dim i As Integer
    For i = 1 To Len(strUUID) Step 2
        ChrUUID = ChrUUID & Chr$("&H" & Mid$(strUUID, i, 2))
    Next i
End Function
Public Function GetTLV(strData As String, intType As Integer) As String
    'get tlv
    Dim strType As String, lngLength As Long, strValue As String
    Dim L1 As Long
    strType$ = Word(intType)
    L1& = InStr(1, strData, strType)
    If L1 = 0 Then GetTLV = "": Exit Function
    lngLength& = GetWord(Mid$(strData, L1& + 2, 2))
    strValue$ = Mid$(strData, L1& + 4, lngLength)
    GetTLV = strValue
End Function
Public Function GetAllTLV(strData As String, intType As Integer) As String
    'get all tlv's of a certain type
    Dim strTempData As String
    strTempData = strData 'so we dont mess with the actual data
    Dim i As Integer
    For i = 0 To UBound(Split(strTempData, Word(intType)))
        If GetTLV(strTempData, intType) <> "" Then
            Dim lngTemp As Long, strFullTLV As String
            GetAllTLV = GetAllTLV & GetTLV(strTempData, intType) & vbCrLf 'get the value out of the tlv
            strFullTLV = Word(intType) & WordS(GetTLV(strTempData, intType)) 'put the tlv back together
            lngTemp& = InStr(1, strTempData, strFullTLV) 'find it
            strTempData$ = Replace$(strTempData, Mid$(strTempData, 1, lngTemp& + Len(strFullTLV)), "") 'remove it
        End If
        DoEvents
    Next i
End Function
Public Function TLV(intType As Integer, strValue As String) As String
    'generate a tlv
    TLV = Word(intType) & Word(Len(strValue)) & strValue
End Function
Public Function WordS(strString As String) As String
    'two byte integer of the length and then the string
    WordS = Word(Len(strString)) & strString
End Function
Public Function EncryptPW(ByRef strPass As String) As String
    Dim arrTable() As Variant
    Dim strEncrypted As String
    Dim lngX As Long
    Dim strHex As String
    
    arrTable = Array(243, 179, 108, 153, 149, 63, 172, 182, 197, 250, 107, 99, 105, 108, 195, 154)
    
    For lngX = 0 To Len(strPass$) - 1
        strHex = Chr(Asc(Mid(strPass, lngX + 1, 1)) Xor CLng(arrTable((lngX Mod 16))))
        strEncrypted = strEncrypted & strHex
    Next
    
    EncryptPW = strEncrypted
End Function
'Put a value up to 65535 into this, and get a 2 byte integer
Public Function Word(ByVal lngVal As Long) As String
    Word = Chr(lngVal \ 256) & Chr(lngVal Mod 256)
End Function
'Input a 2 byte integer into this, and get a value out
Public Function GetWord(ByVal strVal As String) As Long
    GetWord = (Asc(strVal) * 256) + Asc(Mid$(strVal, 2, 1))
End Function
