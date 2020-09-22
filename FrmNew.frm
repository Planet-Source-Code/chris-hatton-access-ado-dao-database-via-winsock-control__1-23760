VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form FrmNew 
   Caption         =   "New Message"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7995
   Icon            =   "FrmNew.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7200
      Top             =   1080
   End
   Begin MSComctlLib.ListView LVFiles 
      Height          =   1455
      Left            =   0
      TabIndex        =   9
      Top             =   5640
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2566
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNew.frx":068A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNew.frx":0ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNew.frx":3292
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNew.frx":36E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNew.frx":3B3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNew.frx":3E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNew.frx":42AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   525
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   926
      BandCount       =   2
      BandBorders     =   0   'False
      VariantHeight   =   0   'False
      _CBWidth        =   7575
      _CBHeight       =   525
      _Version        =   "6.0.8450"
      Child1          =   "Toolbar1"
      MinHeight1      =   465
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      MinHeight2      =   360
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   465
         Left            =   165
         TabIndex        =   8
         Top             =   30
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   820
         ButtonWidth     =   2381
         ButtonHeight    =   820
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Send    "
               Description     =   "MnuNew"
               Object.ToolTipText     =   "Send"
               ImageIndex      =   1
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Send as an Email"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Print        "
               Description     =   "MnuDel"
               Object.ToolTipText     =   "Print"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Reply       "
               Object.ToolTipText     =   "Reply"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1320
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   960
      Width           =   6015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   6015
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6588
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"FrmNew.frx":45C6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7320
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNew.frx":46C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNew.frx":49DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1365
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   975
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   615
      Width           =   855
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         WindowList      =   -1  'True
         Begin VB.Menu MnuSNew 
            Caption         =   "Mail Message"
         End
      End
      Begin VB.Menu mnuSendF 
         Caption         =   "&Send"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "S&end via Email"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Sa&ve As"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuMove 
         Caption         =   "Move to Folder"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuClose 
         Caption         =   "Close"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "FrmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim ATFlist As String
Public MsgItm As Long

Private Sub Form_Load()
'LVFiles.ColumnHeaders.Add , , "File Attachments", 1000
'Me.Height = FrmMain.Height - 2000

End Sub

Private Sub Form_Resize()
On Error Resume Next
RichTextBox1.Width = Me.Width - 100
RichTextBox1.Height = Me.Height - 2600
Text1(0).Width = Me.Width - 1260
Text1(1).Width = Me.Width - 1260
Combo1.Width = Me.Width - 1260
CoolBar1.Width = Me.Width - 100
LVFiles.Width = Me.Width - 100
LVFiles.Height = Me.Height - 6000
If FrmMain.ATFiles = True Then RichTextBox1.Height = Me.Height - 3900
If FrmMain.ATFiles = True Then LVFiles.Visible = True
If FrmMain.ATFiles = True Then LVFiles.Top = Me.Height - 2000
If LVFiles.Height < 2000 Then LVFiles.Height = 1595
Label4.Left = Me.Width - 2500
LVFiles.ColumnHeaders.Item(1).Width = FrmMain.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub LVFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
LVFiles.ToolTipText = Split(LVFiles.SelectedItem.Text, "\")(UBound(Split(LVFiles.SelectedItem.Text, "\")))
End Sub

Private Sub MnuClose_Click()
Unload Me
End Sub

Private Sub mnuCopy_Click()
SendKeys "^c"

End Sub

Private Sub mnuCut_Click()
SendKeys "^x"
End Sub

Private Sub MnuDelete_Click()
Call FrmMain.ItmDelete
Unload Me
End Sub

Private Sub mnuEmail_Click()
    FrmConnect.Usersock.SendData "Email" & FrmConnect.strUserName & Chr(10) & _
    Text1(0).Text & "~~" & Combo1.Text & "~~" & Text1(1).Text & "~~" & _
    RichTextBox1.Text & "[~N10~]" & "~~" & Date
    
Unload Me

End Sub

Private Sub mnuMove_Click()
'from ' subject 'date ' id ' message ' whofrom
FrmMain.DragMessage = Text1(0).Text & "~~" & Text1(1) & "~~" & Label4.Caption & "~~" & MsgItm & "~~" & Split(RichTextBox1.Text, "[~N10~]")(0) & "~~" & FrmConnect.strUserName
Call FrmMain.DragFolder

End Sub

Private Sub mnuPaste_Click()
SendKeys "^v"

End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.Index = 1 Then
    
    FrmConnect.Usersock.SendData "Email" & FrmConnect.strUserName & Chr(10) & _
    Text1(0).Text & "~~" & Combo1.Text & "~~" & Text1(1).Text & "~~" & _
    RichTextBox1.Text & "[~N10~]" & "~~" & Date
    End If
Unload Me
End Sub

Private Sub mnuPrint_Click()
Call PrintMsg
End Sub

Private Sub mnuSendF_Click()
Call SendMessage
End Sub

Private Sub MnuSNew_Click()
FrmMain.CreateNew
End Sub
Private Sub mnuUndo_Click()
SendKeys "^z"
End Sub

Private Sub RichTextBox1_Click()
If Toolbar1.Buttons(4).Enabled = True Then Unload Me: Call FrmMain.Reply
End Sub



Private Sub Timer1_Timer()

On Error Resume Next
If FrmMain.ATFiles = True Then RichTextBox1.Height = Me.Height - 3900
If FrmMain.ATFiles = True Then LVFiles.Visible = True
If FrmMain.ATFiles = True Then LVFiles.Top = Me.Height - 2000
If LVFiles.Height < 2000 Then LVFiles.Height = 1595

If FrmMain.ATFiles = True Then LVFiles.Visible = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index

Case 1: Call SendMessage
Case 3: Call PrintMsg
Case 4: Unload Me: Call FrmMain.Reply

End Select


End Sub
Sub SendMessage()
'newmessageFROM~~TO~~SUBJECT~~MESSAGE
Dim j As Integer

If Combo1.Text = "" Then
    MsgBox "There must be a name in the TO: field", vbExclamation + vbOKOnly, "Send Message"
Exit Sub
End If

If Text1(0).Text = "" Then
    MsgBox "There must be a name in the From: field", vbExclamation + vbOKOnly, "Send Message"
Exit Sub
End If

'If RichTextBox1.Text = "" Then
 '   MsgBox "No blank messages allowed.", vbExclamation + vbOKOnly, "Send Message"
'Exit Sub
'End If

FrmConnect.Usersock.SendData "NewMessage" & Text1(0).Text & _
"~~" & Combo1.Text & "~~" & Text1(1).Text & "~~" & RichTextBox1.Text & "[~N10~]" & "~~" & Date & Chr(10)

Unload Me
End Sub

Sub PrintMsg()
Dim FrmPrinter As FrmPrint
Set FrmPrinter = New FrmPrint
    
        FrmPrinter.Label1.Caption = Text1(0).Text
        FrmPrinter.Label2.Caption = "Subject: " & Text1(1).Text
        FrmPrinter.RichTextBox1.Text = RichTextBox1.Text
        Const ErrCancel = 32755
            FrmMain.PrintDiag.CancelError = True
            
            On Error GoTo errorPrinter
                FrmMain.PrintDiag.Flags = 64
                FrmMain.PrintDiag.ShowPrinter
                FrmPrinter.PrintForm

        Set FrmPrinter = Nothing
       

errorPrinter:
            If Err = ErrCancel Then
        
        
    Set FrmPrinter = Nothing

        Exit Sub
            End If
End Sub


