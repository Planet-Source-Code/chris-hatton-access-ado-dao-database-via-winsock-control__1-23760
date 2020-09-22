VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmServer 
   Caption         =   "Office Server"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8250
   Icon            =   "FrmServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   8250
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2520
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   2040
      Top             =   720
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmServer.frx":12FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmServer.frx":174E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmServer.frx":1BA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3080
      Left            =   3720
      TabIndex        =   7
      Top             =   3360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5424
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"FrmServer.frx":1FFE
   End
   Begin MSWinsockLib.Winsock sockMail 
      Left            =   3000
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LVMsgs 
      Height          =   1215
      Left            =   3720
      TabIndex        =   6
      Top             =   1680
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.ListBox Userlist 
      Appearance      =   0  'Flat
      Height          =   4710
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2040
      Top             =   240
   End
   Begin MSWinsockLib.Winsock ServerSck 
      Index           =   0
      Left            =   2520
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "visit www.chris.hatton.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Database"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5160
      Picture         =   "FrmServer.frx":20B8
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mail Delivery Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   3000
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Outgoing Messages"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Connected Users"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Office Server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   8295
   End
   Begin VB.Menu Mnu1 
      Caption         =   "&File"
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Mnu2 
      Caption         =   "&View"
      Begin VB.Menu MnuProfile 
         Caption         =   "&View User Profiles"
      End
      Begin VB.Menu MnuDelivery 
         Caption         =   "&Delivery Options"
      End
   End
   Begin VB.Menu Mnu3 
      Caption         =   "&Mail"
      Begin VB.Menu mnu3OutGo 
         Caption         =   "&Process Outgoing Mail"
      End
      Begin VB.Menu Mnu3SndRecv 
         Caption         =   "&Send Mail"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCompress 
         Caption         =   "&Compress Database"
      End
   End
   Begin VB.Menu mnu4 
      Caption         =   "&Help"
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu4About 
         Caption         =   "A&bout"
      End
   End
End
Attribute VB_Name = "FrmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sckmax As Integer 'Winsock multi connections
Public SndUserList As Boolean ' new user logged in? if TRUE then send online list
Dim List As String 'Actual list of users
Public ActiveCon As Long 'How many connections are allow to the server
Public CheckMailTmr As Long
Dim FileInit As Long
Public Mtimer As New MailTimer
Private Sub Form_Load()
Me.Hide
FrmStart.Label1.Caption = "Loading Server......"
DoEvents
FrmStart.Show
        ADOConnect
If ODBC.DBConnect = True Then
        Call CreateHeaders
        openDAO
   Call LdEMail
End If

ServerSck(0).LocalPort = 9456           'Sets the local port for the first sock
ServerSck(0).Listen
Me.Label5.Caption = "Server IP: " & ServerSck(0).LocalIP
Me.Label6.Caption = ""

On Error Resume Next
If GetRegKey(HKEY_LOCAL_MACHINE, "Office Server", "CheckState", "") = "True" Then Timer3.Enabled = True
CheckMailTmr = GetRegKey(HKEY_LOCAL_MACHINE, "Office Server", "SndRec", "")

'FrmStart.Hide


End Sub

Private Sub Form_Resize()
On Error Resume Next
RichTextBox1.Width = Me.Width - 3900
LVMsgs.Width = Me.Width - 3900
Label3.Width = Me.Width - 3900
Label4.Width = Me.Width - 3900
LVMsgs.ColumnHeaders.Item(2).Width = Me.Width - 6200
Shape1.Width = Me.Width
LVMsgs.Height = Me.Height - 6000
Label4.Top = Me.Height - 4250
RichTextBox1.Top = Me.Height - 3900
Userlist.Height = Me.Height - 2400
If Label4.Top - 1000 < LVMsgs.Top Then RichTextBox1.Top = 3000: LVMsgs.Height = 930: Label4.Top = 2680: RichTextBox1.Height = Me.Height - 3700


End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
DB.Close
Set DB = Nothing
Set cn = Nothing
End
End Sub
Private Sub SendMailQue()
Dim i As Integer
Dim emID As Long
 If InternetGetConnectedState(0&, 0&) = 1 Then
    For i = 1 To LVMsgs.ListItems.Count
        emID = LVMsgs.ListItems.Item(i).Text
            Call EmOutBox(emID)
            Call EmRemove(emID)
    Next i
    LVMsgs.ListItems.Clear
End If

Unload FrmConnector

End Sub
Private Sub mnu3OutGo_Click()

    If LVMsgs.ListItems.Count = 0 Then
        MsgBox "No mail to send", vbExclamation + vbOKOnly, "Checking Outgoing Mail"
            Exit Sub
    End If

Call SendMailQue


End Sub

Public Sub SendMail(POPAddy, RecvMail, FromAddy, strFrom, ToAddy, Subject, Message As String)
On Error GoTo MailError
Message = Split(Message, "[~N10~]")(0)
sockMail.Close
sockMail.Connect POPAddy, "25"

Do While sockMail.State <> sckConnected

If sockMail.State = sckClosed Then
repsonse = MsgBox("Error Can't Establish Connection " & vbCrLf & "  Retry Connecting?", vbInformation + vbYesNo, "Can't Find Connection")
If repsonse = vbYes Then

Else

GoTo unloadit
End If
End If
DoEvents

Loop
RichTextBox1.Text = RichTextBox1.Text & vbCrLf & vbCrLf & "Session Open:" & vbNewLine
sockMail.SendData "MAIL FROM: " & "chatton1@hotmail.com" & Chr$(13) & Chr$(10) 'leave this incase of error
DoEvents

sockMail.SendData "RCPT TO: " & RecvMail & Chr$(13) & Chr$(10) ' recievers email address"
DoEvents
RichTextBox1.Text = RichTextBox1.Text & Time & ":  " & "Sending Message to " & ToAddy & vbNewLine
RichTextBox1.Text = RichTextBox1.Text & Time & ":  " & "Subject:  " & Subject & vbNewLine

sockMail.SendData "DATA" & Chr$(13) & Chr$(10)
DoEvents
RichTextBox1.Text = RichTextBox1.Text & Time & ":  " & "Communicating to " & POPAddy & vbNewLine

sockMail.SendData "FROM: " & FromAddy & " <" & FromAddy & ">" & Chr$(13) & Chr$(10)
sockMail.SendData "TO: " & strFrom & " <" & ToAddy & ">" & Chr$(13) & Chr$(10)
sockMail.SendData "SUBJECT: " & Subject & Chr$(13) & Chr$(10)
sockMail.SendData Data & Message
sockMail.SendData Chr$(13) & Chr$(10) & "." & Chr$(13) & Chr$(10)
DoEvents

sockMail.SendData "QUIT" & Chr$(13) & Chr$(10)

RichTextBox1.Text = RichTextBox1.Text & "Message Sent!       " & Time & "    " & _
Format(Now, "short Date") & vbNewLine & "Closing Session:" & vbCrLf & "***********************************************" & vbCrLf
RichTextBox1.SelStart = Len(RichTextBox1.Text)
sockMail.Close

unloadit:

Exit Sub
    On Error Resume Next
MailError:
    
  If IsNull(POPAddy) = True Then POPAddy = "POP Server ERROR!"
    RichTextBox1.Text = RichTextBox1.Text & vbCrLf & cbcrlf & "Error Posting Message" & _
    vbCrLf & "POP SERVER: = " & POPAddy & vbCrLf & "Mail To: = " & ToAddy & vbCrLf & _
    "Mail From: = " & FromAddy & vbCrLf & "Subject: = " & Subject & vbCrLf & "Message Dump" & vbCrLf & Message & vbCrLf & _
    vbCrLf & "Deleting Message from que" & vbCrLf & "Closing Session:" & vbCrLf & "***********************************************" & vbCrLf
RichTextBox1.SelStart = Len(RichTextBox1.Text)

sockMail.Close


End Sub

Private Sub Mnu3SndRecv_Click()
        
On Error GoTo cancheck

With FrmConnector
    .Show vbOLEDisplayContent, Me
        .Label1.Caption = "Connecting to Mail Servers"
        FrmConnector.Caption = "Connecting to Mail Servers"
     InternetAutodial INTERNET_AUTODIAL_FORCE_UNATTENDED, 0

        .Label1.Caption = "Sending..."

    If LVMsgs.ListItems.Count = 0 Then
        Unload FrmConnector
    Exit Sub
        Else
        Call SendMailQue
    End If


  
  End With

cancheck:
End Sub

Private Sub mnu4About_Click()
'FrmStart.Label4.Caption = "Office Server" & vbCrLf & vbCrLf & "Beta Version 1.0"
'FrmStart.Label3 = "Author: Chris Hatton": FrmStart.Label3.ToolTipText = "Email: Chris@Hatton.com"
FrmStart.Timer1.Enabled = False
FrmStart.Label1.Caption = ODBC.MSDatabase
FrmStart.Show 1
End Sub

Private Sub mnuCompress_Click()
Call CompactJetDatabase(App.Path & "\OSDB.mdb")
End Sub

Private Sub MnuDelivery_Click()
FrmMailOpt.Show 1
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub MnuProfile_Click()
ClearForm
RegUsers    'display user accounts
FrmProfile.Show 1
End Sub

Private Sub ServerSck_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
        sckmax = sckmax + 1                                     'Increases the user count
        Load ServerSck(sckmax)                                 'Loads a new socket
        ServerSck(sckmax).LocalPort = 0                        'Sets a random port to listen to
        ServerSck(sckmax).Accept requestID                     'Accept the user
            If Userlist.ListCount + 1 > ActiveCon Then
                ServerSck(sckmax).SendData "NoConnection"
               
                    Exit Sub
            End If

        ServerSck(sckmax).SendData "welcome" & Chr(10)         'Tell the user that they are connected
   
   
End Sub
Private Sub ServerSck_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim GetRecv As String

    Dim Command      As String
    Dim NewArrival   As String
    Dim Data         As String
    Static DataCnt   As Long
   ServerSck(Index).GetData GetRecv
Debug.Print GetRecv


On Error GoTo ExitRoutine

If Mid$(GetRecv, 1, 9) = "VUserName" Then
    Userlist.AddItem Split(Mid(GetRecv, 10, Len(GetRecv)), Chr(10))(1) & "/" & Index
    DelLstDup Userlist
    
    Call SckRecordset.sckUserName(Mid(GetRecv, 10, Len(GetRecv)))
End If

If Mid$(GetRecv, 1, 9) = "VPassword" Then
    Call SckRecordset.sckPassword(Mid(GetRecv, 10, Len(GetRecv)))
End If

If Mid$(GetRecv, 1, 11) = "GetUserList" Then
    Dim FindPort As MultiSck
    Set FindPort = New MultiSck
        FindPort.GetSck Mid$(GetRecv, 12, 12)   'User is requesting list, get there port number
        
        Call UsrList    'gets the all the users in a single string
        
        Call usrPorts   'get all the active user ports to send userlist to.
    Set FindPort = Nothing
    
    End If

If Mid$(GetRecv, 1, 7) = "SignOff" Then
    Dim SignoffPort As MultiSck
    Set SignoffPort = New MultiSck
        SignoffPort.GetSck Mid$(GetRecv, 8, Len(GetRecv))
        Call UsrRemove(Mid$(GetRecv, 8, Len(GetRecv)), SignoffPort.Sck)
    Set SignoffPort = Nothing
    End If
    
If Mid$(GetRecv, 1, 7) = "OffList" Then '*
    Dim LstOffline As MultiSck
    Set LstOffline = New MultiSck
        LstOffline.GetSck (Mid$(GetRecv, 8, Len(GetRecv)))
        LstOffline.LstOffline
      
    Set LstOffline = Nothing
End If


If Mid$(GetRecv, 1, 7) = "GetInfo" Then '*
    Dim Iuser As String
    Dim Iport As String
    Dim GetUsrInfo As MultiSck
    Set GetUsrInfo = New MultiSck
        Iport = Split(Mid$(GetRecv, 8, Len(GetRecv)), Chr(10))(0)
        GetUsrInfo.GetSck Iport
        Iuser = Split(Mid$(GetRecv, 8, Len(GetRecv)), Chr(10))(1)
        GetUsrInfo.GetUserI Iuser
    Set GetUsrInfo = Nothing
    
End If


If Mid$(GetRecv, 1, 12) = "CreateFolder" Then
    
    Call NewFolder(Mid$(GetRecv, 14, Len(GetRecv)))
    
End If

If Mid$(GetRecv, 1, 13) = "CustomFolders" Then Call GetFolders(Mid$(GetRecv, 14, Len(GetRecv)))

If Mid$(GetRecv, 1, 12) = "DeleteFolder" Then
    GetRecv = Split(GetRecv, "GetMessages")(0) 'incase of error
    Call DelFolder(Mid$(GetRecv, 14, Len(GetRecv)))
End If

If Mid$(GetRecv, 1, 11) = "GetMessages" Then '*
    Dim GetMessages As MultiSck
    Set GetMessages = New MultiSck
        GetMessages.GetSck2 (Mid$(GetRecv, 13, Len(GetRecv)))
        GetMessages.SendMsgs (Mid$(GetRecv, 13, Len(GetRecv)))
     
    Set GetMessages = Nothing

End If

If Mid$(GetRecv, 1, 14) = "ExportMessages" Then '*
    Dim ExportM As MultiSck
    Set ExportM = New MultiSck
        
        ExportM.GetSck2 (Mid$(GetRecv, 15, Len(GetRecv)))
        ExportM.SendExport (Mid$(GetRecv, 15, Len(GetRecv)))
     
    Set ExportM = Nothing

End If


If Mid$(GetRecv, 1, 11) = "DragMessage" Then
On Error Resume Next
    Dim User, Folder, From, Subj, Discript, Rdate, Msgid, StrUser As String
        GetRecv = Mid$(GetRecv, 12, Len(GetRecv))
        User = Split(GetRecv, Chr(10))(0)
        Folder = Split(GetRecv, Chr(10))(1): Folder = Split(Folder, "~F~")(0)
        From = Split(GetRecv, "~F~")(1): From = Split(From, "~~")(0)
        Subj = Split(GetRecv, "~~")(1): Subj = Split(Subj, "~~")(0)
        Rdate = Split(GetRecv, "~~")(2) ': Rdate = Split(Rdate, "~~")(1)
        Msgid = Split(GetRecv, "~~")(3)
        Discript = Split(GetRecv, "~~")(4)
        StrUser = Split(GetRecv, "~~")(5)
        StrUser = Split(StrUser, "Edit")(0) 'avoiding multi key select error
        
        SckRecordset.MoveRecord User, Folder, From, Subj, Discript, Rdate, Msgid, StrUser
End If

If Mid$(GetRecv, 1, 12) = "DeleteRecord" Then
        Call SckRecordset.DelRecord(Mid$(GetRecv, 14, Len(GetRecv)))
End If

If Mid$(GetRecv, 1, 10) = "DelMessage" Then

    Dim GetUser As String
    Dim DelRecord As Long
    GetRecv = Split(GetRecv, "EditMessage")(0) 'incase of error
    GetRecv = Split(GetRecv, "DelMessage")(0) 'incase of error
    GetUser = Split(Mid$(GetRecv, 11, Len(GetRecv)), Chr(10))(0)
    DelRecord = Split(Mid$(GetRecv, 11, Len(GetRecv)), Chr(10))(1)
    Call DelMessage(GetUser, DelRecord)
End If

If Mid$(GetRecv, 1, 10) = "NewMessage" Then
    Dim strWho, StrSub, StrMsg, SvMsg, Tdate As String
    Dim MsgCounter As Long
    Dim Notification As MultiSck
    Set Notification = New MultiSck
    
    SvMsg = Mid$(GetRecv, 11, Len(GetRecv))
        strWho = Split(SvMsg, "~~")(0)
        StrUser = Split(SvMsg, "~~")(1)
        StrSub = Split(SvMsg, "~~")(2)
        SvMsg = Split(SvMsg, "~~")(3)
        Tdate = Date
        
            Call NewMessage(StrUser, strWho, StrSub, SvMsg, Tdate)
            Call SentMessage(StrUser, strWho, StrSub, SvMsg, Tdate)

            Call MessageCount(StrUser)  'get recordset count
            MsgCounter = SckRecordset.MessageCounter    'new message id for listview
            Notification.GetSck4 (StrUser)  'get socket number
            Notification.Notifiy strWho, StrSub, SvMsg, Tdate, MsgCounter
            
        Set Notification = Nothing
            If InStr(1, GetRecv, "OpenFile,") > 0 Then _
            GetRecv = Split(GetRecv, Chr(10))(1) 'send attachment

        
End If

If Mid$(GetRecv, 1, 5) = "Email" Then
    Dim EmMsg, EmDate As String
    Dim EmUser As String
    Dim EmWho As String
    Dim EmSub As String
    EmMsg = Mid$(GetRecv, 6, Len(GetRecv))
        EmWho = Split(EmMsg, "~~")(0)
        EmUser = Split(EmMsg, "~~")(1)
        EmSub = Split(EmMsg, "~~")(2)
        EmMsg = Split(EmMsg, "~~")(3)
        EmDate = Date
          
        MailIDcheck EmUser
        MailTOcheck EmWho
          
         
        EmInbox Recordsets.usrEmail, Recordsets.UsrToEmail, EmMsg, EmDate, EmSub
        If GetRegKey(HKEY_LOCAL_MACHINE, "Office Server", "SendDirect", "") = "True" Then Mnu3SndRecv_Click
    
            
End If


If Mid$(GetRecv, 1, 11) = "EditMessage" Then
    Dim strEditUsr, strEditRec, strEditFld, strEditMsg, strSvMsg As String
    strSvMsg = Mid$(GetRecv, 12, Len(GetRecv))
        strEditUsr = Split(strSvMsg, "~~")(0)
        strEditRec = Split(strSvMsg, "~~")(1)
        strEditFld = Split(strSvMsg, "~~")(2)
        strEditMsg = Split(strSvMsg, "~~")(3)
        
        Call EditMessage(strEditUsr, strEditRec, strEditFld, strEditMsg)
End If

If Mid$(GetRecv, 1, 8) = "EmptyBin" Then
        Call SckRecordset.EmptyBin(Mid$(GetRecv, 9, Len(GetRecv)))
        
End If

If Mid$(GetRecv, 1, 7) = "ComData" Then
    CompressData.User = Mid$(GetRecv, 8, Len(GetRecv))
    mnuCompress_Click
End If

If Mid$(GetRecv, 1, 11) = "RegisterNew" Then
    Dim RDetails As String
    Dim RName, RAddy, RAddy1, RCountry, RPhone, RFax, Rcom, _
    REmail, RWebsite, RPass As String
        RDetails = Mid$(GetRecv, 12, Len(GetRecv))
        
        RName = Split(RDetails, "~~~")(0)
        RAddy = Split(RDetails, "~~~")(1)
        RAddy1 = Split(RDetails, "~~~")(2)
        RCountry = Split(RDetails, "~~~")(3)
        RPhone = Split(RDetails, "~~~")(4)
        RFax = Split(RDetails, "~~~")(5)
        Rcom = Split(RDetails, "~~~")(6)
        REmail = Split(RDetails, "~~~")(7)
        RWebsite = Split(RDetails, "~~~")(8)
        RPass = Split(RDetails, "~~~")(9)
        
    Call NewAccount(RName, RAddy, RAddy1, RCountry, RPhone, RFax, _
    Rcom, REmail, RWebsite, RPass)
'name, address, address1, country, phone, fax, company,
'email, website, password
        
        rego.MSGUsrSuccess

End If





If Mid$(GetRecv, 1, 10) = "GetMailAcc" Then
    Dim SendAccount As MultiSck
    Set SendAccount = New MultiSck
        Dim AccName As String
            AccName = Mid$(GetRecv, 11, Len(GetRecv))
                SendAccount.GetSck2 AccName
                SendAccount.GetMailAcc AccName
    Set SendAccount = Nothing
    
End If

If Mid$(GetRecv, 1, 8) = "SaveMail" Then
    Dim SvUser, SvPop, SvSmtp, SvAccount, SvPass As String
    SvUser = Split(Mid$(GetRecv, 9, Len(GetRecv)), Chr(10))(0)
    SvPop = Split(Mid$(GetRecv, 9, Len(GetRecv)), Chr(10))(1)
    SvSmtp = Split(Mid$(GetRecv, 9, Len(GetRecv)), Chr(10))(2)

   
    Recordsets.SvEmailAcc SvUser, SvPop, SvSmtp

End If



ExitRoutine: 'dont write code pass this point



End Sub


Private Sub Timer1_Timer()
Dim DBStat As String
If ODBC.DBConnect = True Then DBStat = "Open" Else DBStat = "Closed"
Label7.Caption = "Current Connections "
Label8.Caption = "(" & Userlist.ListCount & ")" & "  " & "(" & ActiveCon & ")"
Label2.Caption = "Database = " & DBStat & " (OSDB.mdb) "
End Sub

Public Sub DelLstDup(listBox As listBox)
On Error GoTo exitdel
' *** Removes any dupes incase server makes a mistake
    Dim a%, b%
    For a% = 0 To listBox.ListCount - 1
        For b% = 0 To listBox.ListCount - 1
            If b% <> a% Then
                If Split(listBox.List(a%), "/")(0) = Split(listBox.List(b%), "/")(0) Then
                    listBox.RemoveItem a%
                    'listBox.RemoveItem "" ' if the listbox finds a "" entry remove it!
                    b% = b% - 1
                End If
            End If
        Next b%
    Next a%
    Exit Sub
exitdel:
    Exit Sub
End Sub
Public Sub usrPorts()
On Error Resume Next
Dim i As Integer
Dim sckPorts As Long
    For i = 0 To Userlist.ListCount - 1
        sckPorts = Split(Userlist.List(i), "/")(1)
        ServerSck(sckPorts).SendData "UserList" _
        & Chr(10) & List
    DoEvents
    Next i
End Sub

Public Sub UsrList()
Dim User As Integer
    List = ""
    For User = 0 To Userlist.ListCount - 1
        List = List & Split(Userlist.List _
        (User), "/")(0) & "_"
    Next User

End Sub
Public Sub UsrRemove(UserName As String, Port As Long)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim sckPorts As Long
For i = 0 To Userlist.ListCount - 1
    If Userlist.List(i) = UserName & "/" & Port Then _
    Userlist.RemoveItem ((i))
    Call UsrList
Next i

For j = 0 To Userlist.ListCount - 1
    sckPorts = Split(Userlist.List(j), "/")(1)
        ServerSck(sckPorts).SendData "UserList" _
        & Chr(10) & List
        DoEvents
Call UsrList
Next j

End Sub
Private Sub CreateHeaders()
FrmStart.Label1.Caption = "Creating Headers"
LVMsgs.ColumnHeaders.Clear

With LVMsgs.ColumnHeaders
    .Add , , , 270
    .Add , , "To", 2200
    .Add , , "From", 1650
    

End With

End Sub
Private Sub Timer2_Timer()
Dim LiveUser As String
Dim LiveIP, i As Long
On Error Resume Next
    For i = 0 To Userlist.ListCount
        LiveUser = Userlist.List(i)
        LiveIP = Split(LiveUser, "/")(1)
            If ServerSck.Item(LiveIP).State <> 7 Then
                Userlist.RemoveItem (i)
                Call UsrList
                Call Broadcast("UserList" & Chr(10) & List)
            End If
    Next i



End Sub
Private Sub Broadcast(BrMessage As String)
Dim CastUser As String
Dim CastIP, i As Long
    For i = 0 To Userlist.ListCount - 1
        CastUser = Userlist.List(i)
        CastIP = Split(CastUser, "/")(1)
            ServerSck.Item(CastIP).SendData BrMessage
    Next i

End Sub

Private Sub Timer3_Timer()
    If Mtimer.Elapsed > CheckMailTmr * 60000 Then
    Mnu3SndRecv_Click
    Mtimer.Reset
    End If
End Sub

Private Sub Userlist_DblClick()
Dim Person As String
Dim IP As Long
    Dim GetInfo As FrmUser
    Set GetInfo = New FrmUser
        
        Person = Split(Userlist.Text, "/")(0)
        IP = Split(Userlist.Text, "/")(1)
        Call GetUserInfo(Person)
    With GetInfo
            .Caption = "User Info: " & Person
            .Label15.Caption = Me.ServerSck(IP).RemoteHostIP
            .Label8.Caption = Recordsets.usrCom
            .Label9.Caption = Recordsets.usrName
            .Label10.Caption = Recordsets.usrAddy
            .Label11.Caption = Recordsets.usrAddy1
            .Label12.Caption = Recordsets.usrPhone
            .Label13.Caption = Recordsets.usrFax
            .Label14.Caption = Recordsets.usrEmail
    End With
    GetInfo.Show 1
    
    Set GetInfo = Nothing

End Sub
