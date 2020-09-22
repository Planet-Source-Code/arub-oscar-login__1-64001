Attribute VB_Name = "modOSCAR"
Option Explicit
Private SNAC As String
Public Function Auth_Login(strUserName As String, strPassword As String) As String 'login authorization
    SNAC = ChrA("0 0 0 1")
    Auth_Login = SNAC & TLV(1, strUserName) & TLV(2, EncryptPW(strPassword)) & _
    TLV(3, "AOL Instant Messenger, version 5.2.3292/WIN32") & _
    TLV(22, ChrA("1 9")) & _
    TLV(23, ChrA("0 5")) & _
    TLV(24, ChrA("0 2")) & _
    TLV(25, ChrA("0 0")) & _
    TLV(26, ChrA("12 220")) & _
    TLV(20, ChrA("0 0 0 238")) & _
    TLV(14, "en") & TLV(15, "us")
End Function
Public Function Set_ICBM() As String 'set icbm parameter's like max message size
    SNAC = ChrA("0 4 0 2 0 0 0 0 0 0")
    Set_ICBM = SNAC & ChrA("0 0 0 0 0 3 31 64 3 231 3 231 0 0 0 0")
End Function
Public Function Auth_Ready() As String 'client ready
    SNAC = ChrH("00 01 00 02 00 00 00 00 00 02")
    Auth_Ready = SNAC & _
    ChrH("00 01 00 04 01 10 08 E5 00 13 00 03 01 10 08 E5 00 02 00 01 01 10 08 E5 00 03 00 01 01 10 08 E5 00 04 00 01 01 10 08 E5 00 06 00 01 01 10 08 E5 00 08 00 01 01 04 00 01 00 09 00 01 01 10 08 E5 00 0A 00 01 01 10 08 E5 00 0B 00 01 01 10 08 E5")
End Function
Public Function Auth_Cookie(strCookie As String) As String 'authorization cookie
    Auth_Cookie = ChrH("00 00 00 01") & TLV(6, strCookie)
End Function
Public Function Request_Limit(lngType As Long) As String 'Request limitation's for specific service
    SNAC = Word(lngType) & ChrH("00 02 00 00 00 00 00 02")
    Request_Limit = SNAC   'just the snac
End Function
Public Function Request_Versions() As String 'request version numbers
    SNAC = ChrH("00 01 00 17 00 00 00 00 00 17")
    Request_Versions = SNAC & ChrH("00 01 00 03 00 13 00 03 00 02 00 01 00 03 00 01 00 04 00 01 00 06 00 01 00 08 00 01 00 09 00 01 00 0A 00 01 00 0B 00 01")
End Function
Public Function Request_Rate_Info() As String 'Ask server for rate info
    SNAC = ChrH("00 01 00 06 00 00 00 00 00 06")
    Request_Rate_Info = SNAC 'just the snac
End Function
Public Function Client_Capabilities() As String 'Capabilities
    Client_Capabilities = ChrUUID("09461341-4C7F-11D1-8222-444553540000") & _
    ChrUUID("09461342-4C7F-11D1-8222-444553540000") & _
    ChrUUID("09461343-4C7F-11D1-8222-444553540000") & _
    ChrUUID("09461345-4C7F-11D1-8222-444553540000") & _
    ChrUUID("09461346-4C7F-11D1-8222-444553540000") & _
    ChrUUID("09461347-4C7F-11D1-8222-444553540000") & _
    ChrUUID("09461348-4C7F-11D1-8222-444553540000") & _
    ChrUUID("0946134A-4C7F-11D1-8222-444553540000") & _
    ChrUUID("0946134B-4C7F-11D1-8222-444553540000") & _
    ChrUUID("0946134D-4C7F-11D1-8222-444553540000") & _
    ChrUUID("0946134E-4C7F-11D1-8222-444553540000") & _
    ChrUUID("748F2420-6287-11D1-8222-444553540000")
End Function
Public Function IDCookie() As String 'generate a random id cookie
    Dim i As Integer
    For i = 1 To 6
        IDCookie = IDCookie & Chr(Int(256 * Rnd))
    Next i
    IDCookie = IDCookie & Chr(0) & Chr(0)
End Function
Public Function Set_Info(strProfile As String) As String 'set profile and capabilities
    SNAC = ChrH("00 02 00 04 00 00 00 00 00 04")
    Set_Info = SNAC & _
    ChrH("00 01 00 21 74 65 78 74 2F 78 2D 61 6F 6C 72 74 66 3B 20 63 68 61 72 73 65 74 3D 22 75 73 2D 61 73 63 69 69 22") & _
    TLV(2, strProfile) & TLV(5, Client_Capabilities)
End Function
Public Function Send_Message(strUserName As String, strMessage As String) 'send icbm message
    SNAC = ChrH("00 04 00 06 00 00 00 00 00 06")
    Send_Message = SNAC & IDCookie & ChrH("00 01") & _
    StringBlock(strUserName) & TLV(2, ChrH("05 01 00 03 01 01 02 01 01") & WordS(ChrH("00 00 00 00") & strMessage))
End Function
Public Function Request_NewService(lngServ As Long) As String 'request a new service
    SNAC = ChrH("00 01 00 04 00 00 00 00 00 00")
    Request_NewService = SNAC & Word(lngServ)
End Function
Public Function Rate_Ack() As String 'ackknowledge rate
    SNAC = ChrH("00 01 00 08 00 00 00 00 00 08")
    Rate_Ack = SNAC & _
    ChrH("00 01 00 02 00 03 00 04 00 05")
End Function
Public Function Change_Ready() As String 'client ready
    SNAC = ChrH("00 01 00 02 00 00 00 00 00 02")
    Change_Ready = SNAC & ChrH("00 01 00 04 00 10 08 E5 00 07 00 01 00 10 08 E5")
End Function
