VERSION 5.00
Begin VB.Form FrmStart 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   LinkTopic       =   "Form2"
   ScaleHeight     =   3615
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5520
      Top             =   3840
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label3"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   3105
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      TabIndex        =   1
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2265
      Left            =   120
      Picture         =   "FrmStart.frx":0000
      Top             =   120
      Width           =   5760
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3360
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3615
      Left            =   0
      Top             =   0
      Width           =   6150
   End
End
Attribute VB_Name = "FrmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Long

Private Sub Form_Click()
Unload Me
End Sub
Private Sub Form_Load()
Label2.Caption = "": Label3.Caption = "visit www.chris.hatton.com"
FrmServer.ActiveCon = 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmServer.Show
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ToolTipText = Label1.Caption
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Label4_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
MousePointer = 11
counter = counter + 1
If counter >= 25 Then
    Unload Me
   
    Timer1.Enabled = False
    MousePointer = 0
End If

End Sub
