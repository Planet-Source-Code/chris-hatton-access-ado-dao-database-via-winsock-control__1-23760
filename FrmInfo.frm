VERSION 5.00
Begin VB.Form FrmInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Information"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "FrmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label18 
      Caption         =   "Website"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label17"
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Company:"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Address:"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Phone"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Fax"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Email:"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "IP Address:"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label8"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label10"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label11"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label12"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label13"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label14"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label15"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3600
      Picture         =   "FrmInfo.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label16"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "FrmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label16.Caption = vbNewLine & "                             User Information"

End Sub
