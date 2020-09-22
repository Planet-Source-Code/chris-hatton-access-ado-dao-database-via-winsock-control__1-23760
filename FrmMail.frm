VERSION 5.00
Begin VB.Form FrmMail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Email Setup"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   Icon            =   "FrmMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcan 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server Information     "
      Height          =   3255
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   5535
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "&Incoming mail (POP3):"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "O&utgoing mail (SMTP):"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
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
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "FrmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcan_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
FrmConnect.Usersock.SendData "SaveMail" & FrmConnect.strUserName & Chr(10) & _
Text1(0).Text & Chr(10) & Text1(1).Text & Chr(10)

Unload Me
End Sub

Private Sub Form_Load()
Call ClearForm
End Sub
Public Sub ClearForm()
Dim j As Integer
With Me.Text1
    For j = .LBound To .UBound
    .Item(j).Text = ""
    Next j
End With

End Sub


