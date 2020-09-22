VERSION 5.00
Begin VB.Form FrmMailCFG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Setup"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCan 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server Information     "
      Height          =   3255
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   5535
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "O&utgoing mail (SMTP):"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "&Incoming mail (POP3):"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "FrmMailCFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcan_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
Recordsets.EmailAcc (FrmProfile.lusers.Text)
Unload Me
End Sub

Private Sub Form_Load()
Call ClearForm
Recordsets.MailInfo (FrmProfile.lusers.Text)
End Sub

Public Sub ClearForm()
Dim j As Integer
With Me.Text1
    For j = .LBound To .UBound
    .Item(j).Text = ""
    Next j
End With

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index))

End Sub
