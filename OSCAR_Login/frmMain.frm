VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OSCAR_Login - Arub"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   3870
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstEmails 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdFindBy 
      Caption         =   "FindBy Email"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSendICBM 
      Caption         =   "Send ICBM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   1
      Left            =   2760
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   3120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtPW 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Password"
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtUIN 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Username"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2970
      Left            =   2520
      Picture         =   "frmMain.frx":0000
      Top             =   120
      Width           =   1305
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
'Description: OSCAR Protocol example that shows the basics of logging into the AIM service, _
              connecting to a service, sending/recieving messages, and _
              finding a user by their email address
'Date: 12/16/05
'Author: Arub
'Site: www.arubs.net
'-----------------------------------------------------------
Option Explicit
Private LocalSeq(1000) As Long
Private StoredData(1000)
'--------------used these to help explain packets ---------------------
'----------------------------------------------------------------------
Private Const SK_AUTH As Integer = 0
Private Const SK_BOS As Integer = 1
Private Enum OSCAR_Channel
    C_AUTH = 1
    C_DATA = 2
    C_DISCONNECT = 4
    C_KEEPALIVE = 5
End Enum
'----------------------------------------------------------------------
'----------------------------------------------------------------------

Private Sub cmdConfirm_Click()
    SendPacket SK_BOS, C_DATA, Request_NewService(7)
End Sub

Private Sub cmdFindBy_Click()
    Dim strEmail As String
    strEmail = InputBox("What's the email?")
    SendPacket SK_BOS, C_DATA, ChrH("00 0A 00 02 00 00 00 00 00 02") & strEmail
End Sub

Private Sub cmdLogin_Click()
    Sock(SK_AUTH).Close
    Sock(SK_AUTH).Connect "login.oscar.aol.com", 5190
    sData 2, 0
End Sub
Private Sub ParseData(intIndex As Integer, strData As String)
    On Error Resume Next
    Debug.Print "{" & AscToDec(Mid$(strData, 7, 4)) & "}" & AscToHex(Mid$(strData, 11))
    Select Case intIndex
        Case SK_AUTH 'auth
            Select Case Asc(Mid$(strData, 2, 1)) 'chan
                Case C_AUTH 'send authorisation for login
                    SendPacket SK_AUTH, C_AUTH, Auth_Login(txtUIN.Text, txtPW.Text)
                Case 4 'disconnect/connect to bos server
                    If InStr(1, GetTLV(strData, 5), ":") = 0 Then    'error
                        MsgBox "Login Error:" & GetWord(GetTLV(strData, 8))
                    Else
                        Dim strBOS() As String
                        strBOS = Split(GetTLV(strData, 5), ":", 2) 'bos server:port
                        txtUIN.Text = GetTLV(strData, 1) 'displa y formatted uin
                        sData 0, GetTLV(strData, 6) 'authorisation cookie
                        Call KillConnection(SK_AUTH) 'close authoriser connection
                        Sock(SK_BOS).Close
                        Sock(SK_BOS).Connect strBOS(0), strBOS(1) 'connect to bos server
                    End If
            End Select
        Case SK_BOS 'BOS
            If Asc(Mid$(strData, 2, 1)) = C_AUTH Then
                SendPacket SK_BOS, C_AUTH, Auth_Cookie(CStr(gData(0))) 'send back the sotred cooke from the authoriser
            End If
            strData = Mid$(strData, 7) 'get rid of data before the snac header
            Select Case AscToDec(Mid$(strData, 1, 4)) 'seperate by snac headers

                Case "0 1 0 3" 'supported services list
                    SendPacket SK_BOS, C_DATA, Request_Versions
                Case "0 1 0 5" 'server redirect
                    Dim intSock As Integer, strServ As String
                    sData 1, GetTLV(Mid$(strData, 15), 6) 'store service authorisation cookie
                    strServ = GetTLV(Mid$(strData, 4), 5) 'server to connect to
                    intSock = Sock.UBound + 1
                    Load Sock(intSock)
                    Sock(intSock).Close
                    Sock(intSock).Connect strServ, 5190 'create new connection to server
                Case "0 1 0 7"  'requested rate info
                    SendPacket SK_BOS, C_DATA, Rate_Ack 'acknowledge rate info
                    SendPacket SK_BOS, C_DATA, ChrH("00 01 00 0E 00 00 00 00 00 0E") 'privacy settings
                    SendPacket SK_BOS, C_DATA, ChrH("00 13 00 02 00 00 00 00 00 02") 'request ssi parameters
                    SendPacket SK_BOS, C_DATA, ChrH("00 13 00 04 00 00 00 00 00 00") 'request buddylist
                    SendPacket SK_BOS, C_DATA, Request_Limit(2)  'request location limitations
                    SendPacket SK_BOS, C_DATA, Request_Limit(3) 'request buddylist management limitations
                    SendPacket SK_BOS, C_DATA, ChrH("00 04 00 04 00 00 00 00 00 00") 'request icbm limitations
                    SendPacket SK_BOS, C_DATA, Request_Limit(9) 'request privacy limitationns
                Case "0 1 0 19" 'message of the day
                    SendPacket SK_BOS, C_DATA, Request_Rate_Info 'request rate limit info
                Case "0 1 0 24"
                    'version numbers
                Case "0 2 0 3" 'server replied with profile param's, finish login
                    SendPacket SK_BOS, C_DATA, Set_ICBM 'set icbm parameters like max message size
                    SendPacket SK_BOS, C_DATA, Set_Info("Hi there, %n!") 'send profile along with capabilities
              '      SendPacket SK_BOS, C_DATA, ChrH("00 02 00 0B 00 00 20 20 00 0B") & StringBlock(txtUIN.Text)  'some weird packet
                    SendPacket SK_BOS, C_DATA, Auth_Ready 'client ready
                Case "0 4 0 7" 'incoming message
                    If Mid$(strData, 19, 2) = ChrH("00 01") Then 'regular im
                        Dim A1, A2, A3, A4
                        A1 = GetStringBlock(strData, 22) 'aim screenname
                        A2 = InStr(1, strData, A1)
                        A3 = InStr(A2 + 1, strData, ChrA("0 0 0 0"))
                        A4 = Mid(strData, A3 + 4) 'message
                        A4 = Left$(A4, Len(A4) - 4) 'get rid of 4 non message bytes at the end
                        MsgBox A1 & " says " & A4, vbInformation, "Message"
                    End If
                Case "0 10 0 1" 'email error
                    MsgBox "There are no accounts on that email or they are all on privacy!", vbInformation
                Case "0 10 0 3" 'email lookup returned
                    Dim strEmails() As String
                    strEmails = Split(GetAllTLV(strData, 1), vbCrLf)
                    Dim i As Integer
                    For i = 0 To UBound(strEmails) - 1
                        lstEmails.AddItem strEmails(i)
                    Next i
                Case "0 19 0 6" 'buddy list
                    If gData(2) = 0 Then 'only sending this once
                        sData 2, 1
                        SendPacket SK_BOS, C_DATA, ChrH("00 13 00 07 00 00 00 00 00 00") 'request buddylist notifications
                        MsgBox "Logged In.", vbInformation + vbOKOnly
                    End If
            End Select
        Case Else 'service
            If Asc(Mid$(strData, 2, 1)) = C_AUTH Then
                SendPacket intIndex, C_AUTH, Auth_Cookie(CStr(gData(1))) 'send back the stored service cookie
            End If
            strData = Mid$(strData, 7)
            Select Case AscToDec(Mid$(strData, 1, 4))
                Case "0 1 0 3"
                    SendPacket intIndex, C_DATA, ChrH("00 01 00 17 00 00 00 00 00 17 00 01 00 03 00 07 00 01") 'versions request
                Case "0 1 0 7"
                    SendPacket intIndex, C_DATA, Rate_Ack
                    SendPacket intIndex, C_DATA, Change_Ready
                    SendPacket intIndex, C_DATA, ChrH("00 07 00 06 00 00 66 CA 00 06")
                Case "0 7 0 7"
                    MsgBox "Check your email for verification steps", vbInformation
                    Call KillConnection(intIndex)
            End Select
    End Select
End Sub
Private Sub SendPacket(intIndex As Integer, intChannel As OSCAR_Channel, strData As String)
    If Sock(intIndex).State <> sckConnected Then Exit Sub
    LocalSeq(intIndex) = LocalSeq(intIndex) + 1
    If LocalSeq(intIndex) >= 65535 Then LocalSeq(intIndex) = 0
    Sock(intIndex).SendData _
    "*" & _
    Chr(intChannel) & _
    Word(LocalSeq(intIndex)) & _
    WordS(strData)
End Sub
Private Sub KillConnection(intIndex As Integer)
    SendPacket intIndex, C_DISCONNECT, vbNullString
    Sock(intIndex).Close
    If intIndex > 1 Then Unload Sock(intIndex) ' we dont want to remove the bos or auth connection
End Sub
Private Sub cmdSendICBM_Click()
    Dim strMessage As String, strUIN As String
    strUIN = InputBox("Who do you want to send to?")
    strMessage = InputBox("What's the message?")
    SendPacket SK_BOS, C_DATA, Send_Message(strUIN, strMessage)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    For i = 0 To Sock.UBound
        Sock(i).Close
        If i > 1 Then Unload Sock(i) 'unload only socket's created for services
    Next i
    End
End Sub

Private Sub Sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'By Xeon
    On Error Resume Next
    Dim strDatas As String
    Dim lngLength As Long
Split:
    Sock(Index).PeekData strDatas, vbString
    lngLength = GetWord(Mid(strDatas, 5, 2))
    If bytesTotal >= lngLength + 6 Then
        Sock(Index).GetData strDatas, vbString, lngLength + 6
        Call ParseData(Index, Mid(strDatas, 1, Len(strDatas)))
        bytesTotal = bytesTotal - (lngLength + 6)
        If bytesTotal > 0 Then GoTo Split
    End If
End Sub
Private Function gData(intIndex As Integer)
    gData = StoredData(intIndex)
End Function
Private Function sData(intIndex As Integer, strData)
    StoredData(intIndex) = strData
End Function
