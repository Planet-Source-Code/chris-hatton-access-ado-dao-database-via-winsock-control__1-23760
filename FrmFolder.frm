VERSION 5.00
Begin VB.Form FrmFolder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Folder"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "FrmFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Create"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Folder Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FrmFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCan_Click()
FrmMain.MousePointer = 0
        cmdOK.Enabled = True
        Text1.Enabled = True
Unload Me
End Sub

Private Sub cmdOK_Click()
If Text1.Text = "" Then Unload Me
With FrmConnect.Usersock
    If Not .State = sckClosed Then _
    .SendData "CreateFolder" & Chr(10) & Text1.Text
        MousePointer = 11
        cmdOK.Enabled = False
        Text1.Enabled = False
End With



End Sub

