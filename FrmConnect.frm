VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect to Server"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "FrmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4905
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   1200
      Top             =   3600
   End
   Begin VB.CommandButton cmdnew 
      Cancel          =   -1  'True
      Caption         =   "Create User"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Create a New User Account"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   3600
   End
   Begin VB.CheckBox Check2 
      Caption         =   "&Connect automatically"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2280
      Top             =   3600
   End
   Begin MSWinsockLib.Winsock Usersock 
      Left            =   2760
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Save Password"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdcon 
      Caption         =   "Connect"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   4200
      Picture         =   "FrmConnect.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      Caption         =   "Enter your user name and password and connect to the Server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label3 
      Caption         =   "Ser&ver IP"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "&Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "&User name:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "FrmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ServerIP As String
Public AutoConnect As Boolean
Dim Counter As Long
Public strUserName As String    'current username


Private Sub Check1_Click()
If Check1.Value = 1 Then Check2.Enabled = True Else Check2.Enabled = False
End Sub

Private Sub CmdCancel_Click()
On Error Resume Next
Dim canAuto As PwSettings
Set canAuto = New PwSettings

If Not Usersock.State = sckClosed Then _
Usersock.SendData "SignOff" & strUserName
DoEvents

If Check2.Value = 0 Then
        canAuto.Autocon = False
End If

If Usersock.State = sckClosed Then
       Unload Me
       Unload FrmMain
       End
        Else
        
        Usersock.Close
        DataRecieve.Status = 2
        cmdcon.Enabled = True
End If

 

FrmMain.DisControls True

FrmMain.TVcontact.Nodes.Clear
FrmMain.TVdir.Nodes.Clear
FrmMain.LvMail.ListItems.Clear
FrmMain.WindowState = vbMinimized
Timer2.Enabled = False
cmdnew.Enabled = True

Set canAuto = Nothing
Me.MousePointer = 0
    
End Sub

Private Sub cmdcon_Click()

Dim PassEvent As PwSettings
Set PassEvent = New PwSettings
    
    
    Timer2.Enabled = True
    PassEvent.UserName = Text1(0).Text 'update to the
                                       'latest user
    

    strUserName = PassEvent.UserName
If Check2.Value = 1 Then
    PassEvent.Autocon = True 'if Auto connect is check save it
    Else
    PassEvent.Autocon = False
End If

If Check1.Value = 1 Then
    PassEvent.SavePass = True
    PassEvent.Password = Text1(1).Text
    PassEvent.ServerIP = Text1(2).Text
    
        Else
    
    PassEvent.Password = Text1(1).Text
    PassEvent.SavePass = False
    
End If

If Not Text1(2).Text = "" Then         'Begin Transmission.
    ServerIP = PassEvent.ServerIP
    DataRecieve.Status = 1
    Call Connect
    
End If

Set PassEvent = Nothing




End Sub

Public Sub Connect()

On Error GoTo again

again:

Usersock.Connect ServerIP, 9456
Do Until Usersock.State = sckConnected
    cmdcon.Enabled = False
    cmdnew.Enabled = False
    Check1.Enabled = False
    Check2.Enabled = False
    DoEvents: DoEvents: DoEvents: DoEvents

Loop


End Sub


Private Sub cmdnew_Click()
FrmNewUser.Show 1

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
If Check1.Value = 1 Then Check2.Enabled = True Else Check2.Enabled = False


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim reponse As String
If Not Usersock.State = sckClosed Then Usersock.Close
FrmMain.DisControls True
FrmMain.LvMail.ListItems.Clear
FrmMain.TVcontact.Nodes.Clear
FrmMain.TVdir.Nodes.Clear

reponse = MsgBox("Would you like to see more winsock projects?" & vbNewLine & vbNewLine & _
"Feel free to contact me at: Chris@Hatton.com" & _
vbNewLine & "Please send me a vote on what you would rate this project.", vbInformation + vbYesNo, "Would you like to see more winsock projects?")

If reponse = vbYes Then

Shell "start http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria=Chris+Hatton&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&lngWId=1&B1=Quick+Search&optSort=Alphabetical"
Shell "start http://www.chris.hatton.com"

Else

    End
End If


End

End Sub


Private Sub Timer1_Timer()
Label5.Caption = DataRecieve.Status             'Keeps winsock status up to date.
FrmMain.StatusBar1.Panels(1).Text = DataRecieve.Statusbar
    
  
    If DataRecieve.SckStat = 6 Then Me.WindowState = vbMinimized
    If DataRecieve.SckStat = 2 Then cmdcon.Enabled = True
    If Usersock.State = sckConnected Then
        If DataRecieve.SckStat = 3 Or DataRecieve.SckStat = 5 Then GoTo skip
        
        DataRecieve.Status = 3
        FrmMain.DisControls False           'enables all the GUI Controls
        Usersock.SendData "GetUserList" & strUserName
        Statusbar = 3
        
        
skip:
    
    End If
    
    
    
End Sub
Private Sub Timer2_Timer()
 
Counter = Counter + 1
'Me.MousePointer = 11

If Counter = 20 Then Usersock.Close: Call Connect: Label5.Caption = Label5.Caption & vbCrLf & "Reconnecting..."
If Counter = 40 Then Usersock.Close: Call Connect: Label5.Caption = Label5.Caption & vbCrLf & "Reconnecting..."
If Counter = 60 Then Usersock.Close: Call Connect: Label5.Caption = Label5.Caption & vbCrLf & "Reconnecting..."


If Counter > 80 Then
    MsgBox "Connection Time Out", vbCritical + vbOKOnly, "Connection Error"
    Timer2.Enabled = False
    CmdCancel = True
 '   Me.MousePointer = 0
End If



End Sub

Private Sub Timer3_Timer()
Dim Response As Variant
If Usersock.State = 8 Then Response = MsgBox("Network Disconnect Detected" & vbCrLf & vbCrLf & _
"Click OK to Reconnect" & vbCrLf & "Cancel to Work Offline", vbCritical + vbOKCancel)

If Response = 1 Then
    FrmMain.DisControls True
    FrmMain.LvMail.ListItems.Clear
    FrmMain.TVcontact.Nodes.Clear
    FrmMain.TVdir.Nodes.Clear
    FrmConnect.Usersock.Close
    FrmConnect.cmdcon = True
End If

If Response = 2 Then
    FrmMain.TVcontact.Nodes.Clear
    Timer3.Enabled = False
End If

End Sub

Private Sub Usersock_DataArrival(ByVal bytesTotal As Long)
Dim DataArrval As String
 Usersock.GetData DataArrval
  'DataRecieve.Status = 5              'verify user and pass
 If DataArrval = "" Then MsgBox "Communication Error"
 If Len(DataArrval) = 0 Then
   Exit Sub
    Else
    
    Call DataRecieve.ParseData(DataArrval)
 End If
End Sub
Private Sub Usersock_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'FrmDownload.Label3.Caption = bytesSent & "KB "
End Sub
