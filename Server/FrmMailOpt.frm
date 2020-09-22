VERSION 5.00
Begin VB.Form FrmMailOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mail Delivery"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   Icon            =   "FrmMailOpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcan 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "&Hang up after sending and recieving"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   2160
      Width           =   4215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "&Send Messages Immediately"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Text            =   "10"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Send Messages every"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "   Delivery Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5415
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   1680
      X2              =   5280
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label2 
      Caption         =   "Dial-up Options"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "minutes"
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   1680
      X2              =   5280
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label3 
      Caption         =   "Mail Account Options"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "FrmMailOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcan_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
If Check3.Value = 1 Then
    SaveRegKey HKEY_LOCAL_MACHINE, "Office Server", "Hangup", "True"
    Else
    SaveRegKey HKEY_LOCAL_MACHINE, "Office Server", "Hangup", "False"
End If

If Check2.Value = 1 Then
    SaveRegKey HKEY_LOCAL_MACHINE, "Office Server", "SendDirect", "True"
    Else
    SaveRegKey HKEY_LOCAL_MACHINE, "Office Server", "SendDirect", "False"
End If

If Check1.Value = 1 Then
    If Text1.Text <= 0 Then Text1.Text = "1"
    SaveRegKey HKEY_LOCAL_MACHINE, "Office Server", "SndRec", Text1.Text
    SaveRegKey HKEY_LOCAL_MACHINE, "Office Server", "CheckState", "True"
    FrmServer.CheckMailTmr = Text1.Text
    FrmServer.Mtimer.Reset
    FrmServer.Timer3 = True
    Else
    SaveRegKey HKEY_LOCAL_MACHINE, "Office Server", "CheckState", "False"
    'SaveRegKey HKEY_LOCAL_MACHINE, "Office Server", "SndRec", ""
    FrmServer.Timer3 = False

End If


Unload Me
End Sub

Private Sub Form_Load()
If GetRegKey(HKEY_LOCAL_MACHINE, "Office Server", "Hangup", "") = "True" Then Check3.Value = 1
If GetRegKey(HKEY_LOCAL_MACHINE, "Office Server", "SendDirect", "") = "True" Then Check2.Value = 1
If GetRegKey(HKEY_LOCAL_MACHINE, "Office Server", "CheckState", "") = "True" Then Check1.Value = 1
Text1.Text = GetRegKey(HKEY_LOCAL_MACHINE, "Office Server", "SndRec", "")
End Sub
