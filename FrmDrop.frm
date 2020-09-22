VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDrop 
   Caption         =   "Select Folder"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   4290
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCan 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   5040
      Width           =   1335
   End
   Begin MSComctlLib.TreeView TVdir 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6800
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrop.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrop.frx":02F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrop.frx":05B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrop.frx":08AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrop.frx":0AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrop.frx":0E7C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Double click on Folder to Move Message to"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "FrmDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Message As String
Private Sub cmdCan_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()

FrmMain.LvMail.ListItems.Remove (FrmMain.LvMail.SelectedItem.Index)
TVdir_DblClick

End Sub

Private Sub TVdir_Click()
Dim i As Integer
If TVdir.SelectedItem.Text = "[Personal Folders]" Then Exit Sub
    For i = 1 To TVdir.Nodes.Count
    TVdir.Nodes.Item(i).Bold = False
    Next i

TVdir.SelectedItem.Bold = True
cmdOK.Enabled = True
End Sub

Private Sub TVdir_DblClick()
Dim i As Integer
With FrmConnect
.Usersock.SendData "DragMessage" & .strUserName & Chr(10) & TVdir.SelectedItem.Text & "~F~" & Message
  '  For i = 1 To FrmMain.TVdir.Nodes.Count
       ' FrmMain.TVdir.Nodes.Item(i).Bold = False
        'If FrmMain.TVdir.Nodes.Item(i).Text = TVdir.SelectedItem.Text Then
       '     FrmMain.TVdir.Nodes.Item(i).Selected = True
        '    FrmMain.TVdir.SelectedItem.Bold = True
        'End If
'    Next i

FrmMain.LvMail.ListItems.Remove (FrmMain.LvMail.SelectedItem.Index)

Unload Me
End With

End Sub

