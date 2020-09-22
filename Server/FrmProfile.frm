VERSION 5.00
Begin VB.Form FrmProfile 
   Caption         =   "User Profiles"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   8655
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   15
      TabIndex        =   13
      Top             =   480
      Width           =   4680
      Begin VB.TextBox Text1 
         Height          =   1485
         Index           =   5
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "FrmProfile.frx":0000
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   960
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   960
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   960
         Width           =   3615
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Details"
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
         Left            =   3240
         TabIndex        =   8
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1320
         Width           =   3615
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Form"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Delete Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "Create Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Discription"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Email:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Name:"
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
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Password:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Company:"
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
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Phone:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdMail 
      Caption         =   "Mail Properties"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   5040
      Width           =   1335
   End
   Begin VB.ListBox lusers 
      Height          =   3765
      Left            =   4800
      TabIndex        =   10
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "New / Existing User Form"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15
      TabIndex        =   18
      Top             =   0
      Width           =   4680
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registered User Accounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   19
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "FrmProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClear_Click()
ClearForm
End Sub

Private Sub CmdMail_Click()

FrmMailCFG.Label6.Caption = "Mail Properties for    (" & lusers.Text & ")"
FrmMailCFG.Show 1
End Sub

Private Sub CmdNew_Click()
AddAccount
RefreshALL
End Sub

Private Sub cmdSave_Click()
UpdateAcc
If UsrExist = False Then Exit Sub
RefreshALL      'Refreshes all text fields and listboxes.
End Sub
Private Sub Command1_Click()
DelAcc
RefreshALL
End Sub

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub lusers_Click()
Recordsets.RegUser = lusers.Text
ItemClick
CmdMail.Enabled = True
End Sub
Public Sub RefreshALL()
lusers.Clear
RegUsers
ClearForm
ItemClick

End Sub

Private Sub Timer1_Timer()


End Sub
