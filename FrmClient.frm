VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Office Messenger"
   ClientHeight    =   7320
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   9870
   Icon            =   "FrmClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   9870
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame HSplit 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000B&
      Height          =   75
      Left            =   2880
      MousePointer    =   7  'Size N S
      TabIndex        =   18
      Top             =   3480
      Width           =   6885
   End
   Begin VB.Frame ctlSplitter 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000B&
      Height          =   6105
      Left            =   2880
      MousePointer    =   9  'Size W E
      TabIndex        =   17
      Top             =   960
      Width           =   45
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   9240
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":07F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":0BAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   7065
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5115
            MinWidth        =   5115
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5680
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LvMail 
      Height          =   2085
      Left            =   2970
      TabIndex        =   1
      Top             =   1395
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3678
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList3"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog PrintDiag 
      Left            =   9240
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9240
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":0FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":131A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":176E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":1BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":2016
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":47CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":4AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":4F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":538E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":56AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9240
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   14
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":59CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":5CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":5F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":6276
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":64C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":6846
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":6BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":6EE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":969A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":9AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":9F42
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":A396
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":A6B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":AB06
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":AF5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":B276
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":B59A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":B8B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":BD0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":C162
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":C85A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":CCAE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   840
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   1482
      BandCount       =   1
      _CBWidth        =   9855
      _CBHeight       =   840
      _Version        =   "6.0.8450"
      MinHeight1      =   780
      Width1          =   6360
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   780
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   1376
         ButtonWidth     =   2223
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "New Message"
               Description     =   "MnuNew"
               Object.ToolTipText     =   "Create a New Message"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete Message"
               Description     =   "MnuDel"
               Object.ToolTipText     =   "Delete"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Reply"
               Description     =   "MnuReply"
               Object.ToolTipText     =   "Reply to Message"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Connect Server"
               Description     =   "MnuConnect"
               Object.ToolTipText     =   "Connect to the Server"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               Object.ToolTipText     =   "Print"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Object.ToolTipText     =   "Refresh List"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Options"
               Object.ToolTipText     =   "Options"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView TVdir 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4471
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   0
      OLEDropMode     =   1
   End
   Begin MSComctlLib.TreeView TVcontact 
      Height          =   3280
      Left            =   0
      TabIndex        =   10
      Top             =   3500
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5794
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   2970
      TabIndex        =   2
      Top             =   3480
      Width           =   6255
      Begin VB.Label Label6 
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Recieved:"
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
         Left            =   3960
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
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
         Left            =   960
         TabIndex        =   7
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label3 
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
         Left            =   720
         TabIndex        =   6
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label2 
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
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
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
         TabIndex        =   4
         Top             =   120
         Width           =   615
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2775
      Left            =   3000
      TabIndex        =   13
      Top             =   4200
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   4895
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"FrmClient.frx":CFD2
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
   Begin VB.Label Label9 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   282
      Left            =   3000
      TabIndex        =   16
      Top             =   1030
      Width           =   3495
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   6435
      TabIndex        =   15
      Top             =   1080
      Width           =   2820
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   960
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   2055
      Left            =   3000
      Top             =   1400
      Width           =   6255
   End
   Begin VB.Menu mnu0File 
      Caption         =   "&File"
      Begin VB.Menu Mnu0New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu0Open 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnu0SaveAs 
         Caption         =   "S&ave As"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnu0Print 
         Caption         =   "Print..."
      End
      Begin VB.Menu mnu0Rubbish 
         Caption         =   "Empty Rubbish Bin"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu0Close 
         Caption         =   "Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnu0Edit 
      Caption         =   "&Edit"
      Begin VB.Menu mnu0Del 
         Caption         =   "Delete"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu0NewFold 
         Caption         =   "&New Folder"
      End
      Begin VB.Menu Mnu0MvFold 
         Caption         =   "&Move to Folder"
      End
   End
   Begin VB.Menu mnu0View 
      Caption         =   "&View"
      Begin VB.Menu Mnu0Preview 
         Caption         =   "Preview Pa&ne"
      End
      Begin VB.Menu mnu0Opt 
         Caption         =   "O&ptions                 "
      End
   End
   Begin VB.Menu mnu4Tools 
      Caption         =   "&Tools"
      Begin VB.Menu mnu4Reindex 
         Caption         =   "Compress Database"
      End
      Begin VB.Menu mnu4EmailAcc 
         Caption         =   "Email Properties"
      End
   End
   Begin VB.Menu mnu0Help 
      Caption         =   "&Help"
      Begin VB.Menu mnu0About 
         Caption         =   "A&bout Office Messenger"
      End
   End
   Begin VB.Menu Menu1 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuNew 
         Caption         =   "&New Message"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh List:"
      End
      Begin VB.Menu MnuInfo 
         Caption         =   "User Information"
      End
      Begin VB.Menu Menu2 
         Caption         =   "Mnu2"
         Visible         =   0   'False
         Begin VB.Menu Mnu2Open 
            Caption         =   "&Open"
         End
         Begin VB.Menu MnuNewFold 
            Caption         =   "&New Folder"
         End
         Begin VB.Menu sep3 
            Caption         =   "-"
         End
         Begin VB.Menu MnuDelFolder 
            Caption         =   "D&elete Folder"
         End
         Begin VB.Menu Mnu2Rubbish 
            Caption         =   "Empty Rubbish Bin"
         End
      End
   End
   Begin VB.Menu Menu3 
      Caption         =   "mnu3"
      Visible         =   0   'False
      Begin VB.Menu mnu3Open 
         Caption         =   "Open"
      End
      Begin VB.Menu mnu3Print 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuReply 
         Caption         =   "Reply"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu3Delete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnu3Move 
         Caption         =   "&Move to Folder..."
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const GWL_STYLE = (-16)
Private Const LVM_FIRST = &H1000
Private Const LVM_GETHEADER = (LVM_FIRST + 31)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Private Const LVS_EX_FULLROWSELECT = &H20
Private Const HDS_BUTTONS = &H2
  
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Offline As Boolean ' work offline detection if true then read cache
Public Logged, MailIcon As Boolean
Public strMessage As String
Public AllUsersList As String 'get all the users in a variable, will come in handy
Public DragMessage As String 'Get the information when dragging a message
Dim TC As NOTIFYICONDATA
Public Selected As String
Public HiddenPreview As Boolean
Public ColumnSet As String
Public ATFiles As Boolean
'Public Lvstore As MsgLayout
Public OFMSGER As MsgLayout



Private Sub Frame1_Click()
RichTextBox1_GotFocus
End Sub

Private Sub HSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Res As Long
HSplit.BackColor = vbBlack

ReleaseCapture
On Error Resume Next
        Res = SendMessage(HSplit.hwnd, WM_SYSCOMMAND, 61458, 0)
         HSplit.BackColor = vbButtonFace

         If HSplit.Top < 1500 Then HSplit.Top = 2400
         If HSplit.Top > FrmMain.Height - 2000 Then HSplit.Top = FrmMain.Height - 3000
         LvMail.Height = HSplit.Top - 1410
         Shape1.Height = HSplit.Top - 1410
         
         Frame1.Top = HSplit.Top - 20
         HSplit.Width = Frame1.Width
         HSplit.Left = Frame1.Left
         RichTextBox1.Height = FrmMain.Height - LvMail.Height - 3080
         RichTextBox1.Top = HSplit.Top + 720

End Sub

Private Sub ctlSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Res As Long
ctlSplitter.BackColor = vbBlack
DoEvents
ReleaseCapture

On Error Resume Next
        Res = SendMessage(ctlSplitter.hwnd, WM_SYSCOMMAND, 61458, 0)
         ctlSplitter.BackColor = vbButtonFace
         If ctlSplitter.Left > 9180 Then ctlSplitter.Left = 9000
         If ctlSplitter.Left < 2675 Then ctlSplitter.Left = 2700
         If HSplit.Top < 1500 Then HSplit.Top = 2500
         LvMail.Left = ctlSplitter.Left + 60
         Shape1.Left = ctlSplitter.Left + 60
         Frame1.Left = ctlSplitter.Left + 60
         Label7.Left = ctlSplitter.Left + 60
         Label9.Left = ctlSplitter.Left + 60
         
         Label6.Width = FrmMain.Width - ctlSplitter.Left - 5300
         Label3.Width = FrmMain.Width - ctlSplitter.Left - 1100
         Label4.Width = FrmMain.Width - ctlSplitter.Left - 1300
         Label1.Width = FrmMain.Width - ctlSplitter.Left - 450
         Label2.Width = FrmMain.Width - ctlSplitter.Left - 450
         Label5.Width = FrmMain.Width - ctlSplitter.Left - 4300
         Label7.Width = FrmMain.Width - ctlSplitter.Left - 300
         
         TVdir.Width = ctlSplitter.Left
         TVcontact.Width = ctlSplitter.Left
         HSplit.Left = ctlSplitter.Left + 60
         ctlSplitter.Top = TVdir.Top
         RichTextBox1.Left = ctlSplitter.Left + 60
         LvMail.Width = FrmMain.Width - ctlSplitter.Left - 260
         Shape1.Width = FrmMain.Width - ctlSplitter.Left - 260
         
         RichTextBox1.Width = FrmMain.Width - ctlSplitter.Left - 285
         Frame1.Width = FrmMain.Width - ctlSplitter.Left - 275
         LvMail.ColumnHeaders.Item(3).Width = FrmMain.Width - ctlSplitter.Left - 4340
       
End Sub


Private Sub Form_Load()

'Remove these few lines if you deside to have your own personal name saved.

DoEvents
SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "UserName", "Chris Hatton"
DoEvents
SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "Password", "Password"
DoEvents
SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "SavePassword", "True"
'remove


If GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "AutoLogon", "") = "True" Then Call GetConnection
If GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "HidePreview", "") = "True" Then Call FrmStyle: Mnu0Preview.Checked = False Else:  Mnu0Preview.Checked = True


Call GetConnection
Me.WindowState = vbMinimized
DisControls True
Label8.Caption = Format(Now, "long Date")
Label9.Caption = "  Inbox Folder"
Me.Caption = "  Inbox Folder - " & "Office Messenger"
'SendMessage HSplit.hWnd, &HF4&, &H8&, 0&
'SendMessage ctlSplitter.hWnd, &HF4&, &H8&, 0&
'Set Lvstore = New MsgLayout
End Sub

Sub lvcolumns()
LvMail.View = lvwReport

With LvMail.ColumnHeaders
    
    .Add , , , 280
    .Add , , "From", 1500
    .Add , , "Subject", 8250
    .Add , , "Received", 1830
    .Add , , , 180  'blank '5
   
    .Item(5).Position = 1
    
End With
435
Dim hHeader As Long
    hHeader = SendMessage(LvMail.hwnd, LVM_GETHEADER, 0, ByVal 0&)
    SetWindowLong hHeader, GWL_STYLE, GetWindowLong(hHeader, GWL_STYLE) Xor HDS_BUTTONS

End Sub

Private Sub Form_Resize()
On Error Resume Next

'If HSplit.Top < 1800 Then Me.Height = 6500

         
Label7.Left = ctlSplitter.Left + 60
LvMail.Width = FrmMain.Width - 3150
Shape1.Width = FrmMain.Width - 3150
Label7.Width = FrmMain.Width - 3150
'LvMail.ColumnHeaders.Item(3).Width = FrmMain.Width - 6980
CoolBar1.Width = FrmMain.Width - 160
Frame1.Width = FrmMain.Width - 3100
DoEvents
If Not HiddenPreview = True Then RichTextBox1.Height = FrmMain.Height - 600
'If Not HiddenPreview = True Then LvMail.Height = HSplit.Top - 1430
'If Not HiddenPreview = True Then Shape1.Height = HSplit.Top - 1430
If HiddenPreview = True Then LvMail.Height = FrmMain.Height - 2340
If HiddenPreview = True Then Shape1.Height = FrmMain.Height - 2340

DoEvents
Label6.Width = FrmMain.Width - 8100
DoEvents
'RichTextBox1.Width = FrmMain.Width - 2500
DoEvents
TVcontact.Height = FrmMain.Height - 4430
DoEvents
HSplit.Width = Frame1.Width
DoEvents
HSplit.Left = Frame1.Left
DoEvents
ctlSplitter.Height = FrmMain.Height - 1900
DoEvents
If Not HSplit.Top <= 2000 Then HSplit.Top = FrmMain.Height - 4500
If Me.Height < 6800 Then HSplit.Top = 2100
'Frame1.Top = HSplit.Top + 100

        If Not HiddenPreview = True Then LvMail.Height = HSplit.Top - 1430
         If Not HiddenPreview = True Then Shape1.Height = HSplit.Top - 1430

         Frame1.Top = HSplit.Top - 20
         DoEvents
         RichTextBox1.Height = FrmMain.Height - LvMail.Height - 3090
         RichTextBox1.Top = HSplit.Top + 720
         ctlSplitter.Top = TVdir.Top
         Label3.Width = FrmMain.Width - ctlSplitter.Left - 1100
         Label4.Width = FrmMain.Width - ctlSplitter.Left - 1300
         Label1.Width = FrmMain.Width - ctlSplitter.Left - 450
         Label2.Width = FrmMain.Width - ctlSplitter.Left - 450
         Label5.Width = FrmMain.Width - ctlSplitter.Left - 4300
         Label7.Width = FrmMain.Width - ctlSplitter.Left - 300
         Label6.Width = FrmMain.Width - ctlSplitter.Left - 5300
         LvMail.Left = ctlSplitter.Left + 60
         Shape1.Left = ctlSplitter.Left + 60
         HSplit.Left = ctlSplitter.Left + 60
         Frame1.Left = ctlSplitter.Left + 60
         RichTextBox1.Left = ctlSplitter.Left + 60
         HSplit.Left = Frame1.Left
         TVdir.Width = ctlSplitter.Left
         TVcontact.Width = ctlSplitter.Left
         DoEvents
         Label8.Left = Me.Width - 2400
         DoEvents
         RichTextBox1.Left = ctlSplitter.Left + 60
         DoEvents
         LvMail.Width = FrmMain.Width - ctlSplitter.Left - 260
         DoEvents
         Shape1.Width = FrmMain.Width - ctlSplitter.Left - 260
             
         DoEvents
         RichTextBox1.Width = FrmMain.Width - ctlSplitter.Left - 275
         DoEvents
         Frame1.Width = FrmMain.Width - ctlSplitter.Left - 260
         DoEvents
         LvMail.ColumnHeaders.Item(3).Width = FrmMain.Width - ctlSplitter.Left - 4340


         

         
         
         If Label8.Left < 6550 Then Label8.Visible = False Else Label8.Visible = True
         If Label9.Width > Label7.Width Then Label9.Visible = False Else Label9.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
RemoveIcon
With FrmConnect.Usersock
    If FrmConnect.WindowState = vbMinimized Then FrmConnect.WindowState = vbNormal
    On Error Resume Next
    If Not .State = sckClosed Then .SendData "SignOff" & FrmConnect.strUserName
DoEvents
        If Not .State = sckClosed Then .Close
End With
  FrmConnect.Timer1.Enabled = False
  Exit Sub
'Set Lvstore = Nothing
Set OFMSGER = Nothing
End Sub
Public Sub AddIcon()
Dim rc As Long
If MailIcon = True Then Exit Sub
TC.cbSize = Len(TC)
                TC.hwnd = FrmMain.hwnd
                TC.uID = vbNull
                TC.uFlags = NIF_DOALL
                TC.uCallbackMessage = WM_MOUSEMOVE
                TC.hIcon = FrmMain.Icon
                TC.sTip = "New Office Mail" & vbNullChar
                rc = Shell_NotifyIcon(NIM_ADD, TC)
                MailIcon = True
                Beep
End Sub
Public Sub RemoveIcon()
Dim rc As Long
rc = Shell_NotifyIcon(NIM_DELETE, TC)
MailIcon = False
End Sub

Private Sub Label1_Click()
RichTextBox1_GotFocus
End Sub

Private Sub Label2_Click()
RichTextBox1_GotFocus
End Sub

Private Sub Label3_Click()
RichTextBox1_GotFocus
End Sub

Private Sub Label4_Click()
RichTextBox1_GotFocus
End Sub

Private Sub Label5_Click()
RichTextBox1_GotFocus
End Sub

Private Sub Label6_Click()
RichTextBox1_GotFocus
End Sub

Public Sub LvMail_Click()
Dim itm As ListItem
Dim Folder As String
Dim i As Integer
'Dim Lvstore As MsgLayout
'Set Lvstore = New MsgLayout
Dim rs As Long
Dim rc As Long
On Error Resume Next
If HiddenPreview = True And LvMail.SelectedItem.ListSubItems(2).Bold = False Then GoTo SkipPre      'skip the preview if not avialiable
RichTextBox1_LostFocus
    With LvMail
        For i = 1 To LvMail.ListItems.Count
        LvMail.ListItems.Item(i).Ghosted = False
        Next i
        
        Set itm = .ListItems.Item(.SelectedItem.Index)
            Lvstore.GetMsgStore (.SelectedItem.Text)
            Label3.Caption = itm.SubItems(1)
            Label4.Caption = itm.SubItems(2)
            Label6.Caption = itm.SubItems(3)
            .SelectedItem.Bold = False
            itm.ListSubItems.Item(1).Bold = False
            itm.ListSubItems.Item(2).Bold = False
            itm.ListSubItems.Item(3).Bold = False
            FrmMain.StatusBar1.Panels.Item(3).ToolTipText = ""
            RemoveIcon
            'FrmMain.StatusBar1.Panels.Item(3).Picture = Nothing
            LvMail.ListItems.Item(.SelectedItem.Index).SmallIcon = 2
            'LvMail.ListItems.Item(.SelectedItem.Index).Ghosted = True
            
    End With
    
    With FrmConnect
        Folder = TVdir.SelectedItem.Text 'been read
        If TVdir.SelectedItem.Text = "Inbox" Then Folder = "Discription"
            .Usersock.SendData "EditMessage" & .strUserName & "~~" & _
            LvMail.SelectedItem.Text & "~~" & Folder & _
            "~~" & RichTextBox1.Text
    End With

    
SkipPre:
        LvMail.Refresh
        Set itm = Nothing
       ' Set Lvstore = Nothing
        
End Sub

Private Sub LvMail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

EnhListView_SortColumns LvMail, ColumnHeader.Index, False

End Sub

Private Sub LvMail_DblClick()
If LvMail.ListItems.Count = 0 Then Exit Sub
Call MsOpen
End Sub

Private Sub LvMail_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDelete Then mnu3Delete_Click
If KeyCode = vbKeyF5 Then mnuRefresh_Click
If KeyCode = 40 Then Call PreItmUp
If KeyCode = 38 Then Call PreItmDown

End Sub
Private Sub PreItmDown()
Dim itm As ListItem
Dim Folder As String
Dim i As Integer
'Dim Lvstore As MsgLayout
'Set Lvstore = New MsgLayout
Dim rs As Long
Dim rc As Long
On Error Resume Next
If HiddenPreview = True And LvMail.SelectedItem.ListSubItems(2).Bold = False Then GoTo SkipPre      'skip the preview if not avialiable

    With LvMail
        For i = 1 To LvMail.ListItems.Count
        LvMail.ListItems.Item(i).Ghosted = False
        Next i
        
        Set itm = .ListItems.Item(.SelectedItem.Index - 1)
            Lvstore.GetMsgStore (.ListItems.Item(.SelectedItem.Index - 1).Text)
            Label3.Caption = itm.SubItems(1)
            Label4.Caption = itm.SubItems(2)
            Label6.Caption = itm.SubItems(3)
            .SelectedItem.Bold = False
            itm.ListSubItems.Item(1).Bold = False
            itm.ListSubItems.Item(2).Bold = False
            itm.ListSubItems.Item(3).Bold = False
            FrmMain.StatusBar1.Panels.Item(3).ToolTipText = ""
            RemoveIcon
            'FrmMain.StatusBar1.Panels.Item(3).Picture = Nothing
            LvMail.ListItems.Item(.SelectedItem.Index).SmallIcon = 2
            'LvMail.ListItems.Item(.SelectedItem.Index).Ghosted = True
            
    End With
    
    With FrmConnect
        Folder = TVdir.SelectedItem.Text
        If TVdir.SelectedItem.Text = "Inbox" Then Folder = "Discription"
            .Usersock.SendData "EditMessage" & .strUserName & "~~" & _
            LvMail.SelectedItem.Text & "~~" & Folder & _
            "~~" & RichTextBox1.Text
    End With
SkipPre:
        LvMail.Refresh
        Set itm = Nothing
       ' Set Lvstore = Nothing
End Sub
Private Sub PreItmUp()
Dim itm As ListItem
Dim Folder As String
Dim i As Integer
'Dim Lvstore As MsgLayout
'Set Lvstore = New MsgLayout
Dim rs As Long
Dim rc As Long
On Error Resume Next
If HiddenPreview = True And LvMail.SelectedItem.ListSubItems(2).Bold = False Then GoTo SkipPre      'skip the preview if not avialiable

    With LvMail
        For i = 1 To LvMail.ListItems.Count
        LvMail.ListItems.Item(i).Ghosted = False
        Next i
        
        Set itm = .ListItems.Item(.SelectedItem.Index + 1)
            Lvstore.GetMsgStore (.ListItems.Item(.SelectedItem.Index + 1).Text)
            Label3.Caption = itm.SubItems(1)
            Label4.Caption = itm.SubItems(2)
            Label6.Caption = itm.SubItems(3)
            .SelectedItem.Bold = False
            itm.ListSubItems.Item(1).Bold = False
            itm.ListSubItems.Item(2).Bold = False
            itm.ListSubItems.Item(3).Bold = False
            FrmMain.StatusBar1.Panels.Item(3).ToolTipText = ""
            RemoveIcon
            'FrmMain.StatusBar1.Panels.Item(3).Picture = Nothing
            LvMail.ListItems.Item(.SelectedItem.Index).SmallIcon = 2
            'LvMail.ListItems.Item(.SelectedItem.Index).Ghosted = True
            
    End With
    
    With FrmConnect
        Folder = TVdir.SelectedItem.Text
        If TVdir.SelectedItem.Text = "Inbox" Then Folder = "Discription"
            .Usersock.SendData "EditMessage" & .strUserName & "~~" & _
            LvMail.SelectedItem.Text & "~~" & Folder & _
            "~~" & RichTextBox1.Text
    End With
SkipPre:
        LvMail.Refresh
        Set itm = Nothing
       ' Set Lvstore = Nothing
End Sub

Private Sub LvMail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call MsOpen

End Sub

Private Sub LvMail_KeyUp(KeyCode As Integer, Shift As Integer)
LvMail.Refresh
End Sub

Private Sub LvMail_LostFocus()
LvMail.Refresh
End Sub

Private Sub LvMail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If LvMail.ListItems.Count = 0 Then Exit Sub
If Button = vbRightButton Then
PopupMenu Menu3
End If

End Sub

Private Sub LvMail_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
On Error Resume Next
Dim itm As ListItem
Set itm = LvMail.ListItems.Item(LvMail.SelectedItem.Index)
    Data.SetData itm.SubItems(1) & " " & itm.SubItems(2) & " " & itm.SubItems(3) & "~~" & LvMail.SelectedItem.Text & " " & Split(RichTextBox1.Text, "[~N10~]")(0) 'copy to clipboard
    DragMessage = itm.SubItems(1) & "~~" & itm.SubItems(2) & "~~" & itm.SubItems(3) & "~~" & LvMail.SelectedItem.Text & "~~" & Split(RichTextBox1.Text, "[~N10~]")(0) & "~~" & FrmConnect.strUserName
Set itm = Nothing
End Sub
Private Sub Mnu0Close_Click()
Unload Me
End Sub

Private Sub mnu0Del_Click()

Call ItmDelete
End Sub

Private Sub Mnu0MvFold_Click()
Call MainMove
Call DragFolder
End Sub

Private Sub Mnu0New_Click()
Call CreateNew
End Sub

Private Sub mnu0Open_Click()
Call MsOpen
End Sub

Private Sub mnu0Opt_Click()
Call GetOptions
End Sub

Private Sub Mnu0Preview_Click()
If Mnu0Preview.Checked = True Then
    SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "HidePreview", "True"
    Call FrmStyle
    Exit Sub
End If

If Mnu0Preview.Checked = False Then
    SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "HidePreview", "False"
    Mnu0Preview.Checked = True
    HiddenPreview = False
    Frame1.Visible = True
    RichTextBox1.Visible = True
    HSplit.Visible = True
    LvMail.Height = FrmMain.HSplit.Top - 1410
    Shape1.Height = FrmMain.HSplit.Top - 1410
End If

End Sub

Private Sub mnu0Print_Click()
Call PrintFrm
End Sub

Private Sub mnu0Rubbish_Click()
Call EmptyRubbish
End Sub

Private Sub mnu2Cache_Click()
FrmConnect.Usersock.SendData "GetCache~~" & FrmConnect.strUserName & "~~" & _
TVdir.SelectedItem.Text

End Sub

Private Sub Mnu2Open_Click()
Selected = ""
TVdir_Click
End Sub
Private Sub Mnu2Rubbish_Click()
Call EmptyRubbish
End Sub

Private Sub EmptyRubbish()
Dim Response As String
Dim i As Integer
Response = MsgBox("Are you Sure you want to Delete all items?", vbExclamation + vbOKCancel, "Empty Rubbish Bin?")
    If Response = 1 Then
       FrmConnect.Usersock.SendData "EmptyBin" & FrmConnect.strUserName
       LvMail.ListItems.Clear
    End If

End Sub

Public Sub mnu3Delete_Click()
Call ItmDelete
End Sub
Public Sub ItmDelete()
On Error Resume Next
Dim itm As ListItem

If Not TVdir.SelectedItem.Text = "Rubbish Bin" Then
    Set itm = LvMail.ListItems.Item(LvMail.SelectedItem.Index)
        DragMessage = itm.SubItems(1) & "~~" & itm.SubItems(2) & "~~" & itm.SubItems(3) & "~~" & LvMail.SelectedItem.Text & "~~" & Split(RichTextBox1.Text, "[~N10~]")(0) & "~~" & FrmConnect.strUserName
        LvMail.ListItems.Remove (LvMail.SelectedItem.Index)
        LvMail.Refresh

        FrmConnect.Usersock.SendData "DragMessage" & _
        FrmConnect.strUserName & Chr(10) & "Rubbish Bin" & "~F~" & DragMessage
    
    Set itm = Nothing
Else

With FrmConnect

    .Usersock.SendData "DelMessage" & .strUserName & Chr(10) & FrmMain.LvMail.SelectedItem.Text

End With

With LvMail
If .SelectedItem.Text = "" Then
    MsgBox "No Message to Delete", vbOKOnly, "Delete Error"
    Exit Sub
Else
    .ListItems.Remove (.SelectedItem.Index)
    .ListItems.Item(.ListItems.Count).Selected = True
    RichTextBox1.Text = ""
    Label3.Caption = ""
    Label4.Caption = ""
    Label6.Caption = ""
End If
End With

End If

End Sub

Private Sub mnu3Move_Click()
Call MainMove
Call DragFolder
End Sub
Public Sub MainMove()
Dim itm As ListItem
Set itm = LvMail.ListItems.Item(LvMail.SelectedItem.Index)
    DragMessage = itm.SubItems(1) & "~~" & itm.SubItems(2) & "~~" & itm.SubItems(3) & "~~" & LvMail.SelectedItem.Text & "~~" & Split(RichTextBox1.Text, "[~N10~]")(0) & "~~" & FrmConnect.strUserName
Set itm = Nothing

End Sub

Public Sub MsOpen()
Dim itm As ListItem
Dim i As Integer
Dim frmView As FrmNew
Set frmView = New FrmNew
'Dim LvMessage As MsgLayout
'Set LvMessage = New MsgLayout
On Error Resume Next
With LvMail
If Not .SelectedItem.Text = "" Then
Set itm = .ListItems.Item(.SelectedItem.Index)
  Lvstore.GetMsgStore (.SelectedItem.Text)
        
        With FrmNew
            For i = 0 To UBound(Split(AllUsersList, "_")) - 1
            .Combo1.AddItem Split(AllUsersList, "_")(i)
                Next i
            .mnuEmail.Enabled = False
            .mnuSendF.Enabled = False
'            .mnuEdit.Enabled = False
            '.mnuInsert.Enabled = False
            .MsgItm = LvMail.SelectedItem.Text
            .Combo1.AddItem FrmConnect.strUserName
            .Combo1.Text = FrmConnect.strUserName
            .Combo1.BackColor = Me.BackColor
            .Combo1.Appearance = vbFlat
            .Combo1.Enabled = False
            .Toolbar1.Buttons.Item(1).Enabled = False
            .Text1(0).Text = itm.SubItems(1)
            .Text1(1).Text = itm.SubItems(2)
            .Label4.Caption = itm.SubItems(3)
            .RichTextBox1.Text = Split(Lvstore.SendMessageID, "[~N10~]")(0)
            .Caption = itm.SubItems(2)
            .Text1(0).Locked = True
            .Text1(0).BackColor = Me.BackColor
            .Text1(0).Appearance = vbFlat
            .Text1(0).BorderStyle = 0
            .Text1(1).Locked = True
            .Text1(1).BackColor = Me.BackColor
            .Text1(1).Appearance = vbFlat
            .Text1(1).BorderStyle = 0
            .Label1.FontBold = False
            .Label2.FontBold = False
            .Label3.FontBold = False
            .Show
        End With
Else
MsgBox "Error With Message Content", vbExclamation + vbOKOnly
End If
End With

Set itm = Nothing
'Set LvMessage = Nothing
Set FrmNew = Nothing

End Sub

Private Sub mnu3Open_Click()
Call MsOpen
End Sub

Private Sub mnu3Print_Click()
Call PrintFrm
End Sub



Private Sub mnu4EmailAcc_Click()
FrmConnect.Usersock.SendData "GetMailAcc" & FrmConnect.strUserName
End Sub

Private Sub mnu4Reindex_Click()
FrmConnect.Usersock.SendData "ComData" & FrmConnect.strUserName
MsgBox "Database Compressed", vbInformation

End Sub

Private Sub MnuDelFolder_Click()
With FrmConnect.Usersock
    If Not .State = sckClosed Then _
    .SendData "DeleteFolder" & Chr(10) & FrmMain.TVdir.SelectedItem.Text
    MousePointer = 11
End With

End Sub

Private Sub MnuInfo_Click()
FrmConnect.Usersock.SendData "GetInfo" & FrmConnect.strUserName & Chr(10) & TVcontact.SelectedItem.Text
End Sub

Private Sub mnuNew_Click()
TVcontact_DblClick
End Sub

Private Sub MnuNewFold_Click()
Call newFolder
End Sub

Public Sub mnuRefresh_Click()
On Error Resume Next
LvMail.ListItems.Clear
TVcontact.Nodes.Clear
TVdir.Nodes.Clear
FrmConnect.Usersock.SendData "GetUserList" & FrmConnect.strUserName
Statusbar = 4
End Sub

Private Sub mnuReply_Click()
Call Reply
End Sub
Private Sub RichTextBox1_Click()
'If RichTextBox1.Text = "" Then Exit Sub
'Call Reply
End Sub




Private Sub RichTextBox1_GotFocus()
Frame1.BackColor = &H700000
Label4.BackColor = &H700000
Label3.BackColor = &H700000
Label6.BackColor = &H700000
Label1.BackColor = &H700000
Label2.BackColor = &H700000
Label5.BackColor = &H700000
Frame1.BorderStyle = 0
Label1.ForeColor = vbWhite
Label2.ForeColor = vbWhite
Label5.ForeColor = vbWhite

Label4.ForeColor = vbWhite
Label3.ForeColor = vbWhite
Label6.ForeColor = vbWhite

End Sub

Private Sub RichTextBox1_LostFocus()
Frame1.BackColor = vbButtonFace
Label4.BackColor = vbButtonFace
Label3.BackColor = vbButtonFace
Label6.BackColor = vbButtonFace
Label1.BackColor = vbButtonFace
Label2.BackColor = vbButtonFace
Label5.BackColor = vbButtonFace
Frame1.BorderStyle = 1
Label1.ForeColor = &H0&
Label2.ForeColor = &H0&
Label5.ForeColor = &H0&

Label4.ForeColor = &H0&
Label3.ForeColor = &H0&
Label6.ForeColor = &H0&

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index

Case 1: Call CreateNew
Case 3: Call ItmDelete
Case 5: Call Reply
Case 7: FrmConnect.WindowState = vbNormal
Case 9: Call PrintFrm
Case 11: mnuRefresh_Click
Case 13: Call GetOptions
End Select
End Sub

Private Sub GetConnection()
Dim strAuthen As PwSettings
Set strAuthen = New PwSettings
Dim NewCon As FrmConnect
Set NewCon = FrmConnect
If FrmConnect.WindowState = vbMinimized Then _
FrmConnect.WindowState = vbNormal

With NewCon
        If strAuthen.SavePass = True Then
        .Check1.Value = 1
        Else
        .Check1.Value = 0
    End If
    
    If strAuthen.Autocon = True Then
        .Check2.Value = 1
        .AutoConnect = True
        Else
        .Check2.Value = 0
        .AutoConnect = False
    End If
    
    DataRecieve.Status = 0               'Clear the status description
    .Text1(0).Text = strAuthen.UserName
    .Text1(1).Text = strAuthen.Password
    .Text1(1).SelStart = 0
    .Text1(1).SelLength = Len(.Text1(1).Text)
    .Text1(2).Text = strAuthen.ServerIP
    .Show
    If .AutoConnect = True Then .cmdcon = True
End With

Set NewCon = Nothing
Set strAuthen = Nothing

End Sub

Public Sub CreateNew()
Dim FormMsg As FrmNew
Set FormMsg = New FrmNew
Dim i As Integer
    With FormMsg
         For i = 0 To UBound(Split(AllUsersList, "_")) - 1
        .Combo1.AddItem Split(AllUsersList, "_")(i)
         Next i
            .mnuNew.Enabled = False
            .mnuSave.Enabled = False
            .MnuDelete.Enabled = False
            .mnuMove.Enabled = False
            .Combo1.AddItem ""
            .Text1(0).Text = FrmConnect.strUserName
            .Text1(1).Text = ""
            .RichTextBox1.Text = ""
            .Toolbar1.Buttons(4).Enabled = False
           
        .Show
    End With
Set FormMsg = Nothing
End Sub

Public Sub DisControls(Disable As Boolean)
With Toolbar1.Buttons
If Disable = True Then
FrmMain.LvMail.ColumnHeaders.Clear
    .Item(1).Enabled = False
    .Item(2).Enabled = False
    .Item(3).Enabled = False
    .Item(4).Enabled = False
    .Item(5).Enabled = False
    .Item(6).Enabled = False
    .Item(8).Enabled = False
    .Item(9).Enabled = False
    .Item(11).Enabled = False
    Label3.Caption = ""
    Label4.Caption = ""
    Label6.Caption = ""
    RichTextBox1.Text = ""
    TVdir.Enabled = False
    TVcontact.Enabled = False
    RichTextBox1.Enabled = False
    LvMail.Enabled = False
    LvMail.View = lvwList
    .Item(6).Image = 10

Else
Call lvcolumns          'create columns headers

    .Item(1).Enabled = True
    .Item(2).Enabled = True
    .Item(3).Enabled = True
    .Item(4).Enabled = True
    .Item(5).Enabled = True
    .Item(6).Enabled = True
    .Item(8).Enabled = True
    .Item(9).Enabled = True
    .Item(11).Enabled = True
    .Item(6).Image = 4

    TVdir.Enabled = True
    TVcontact.Enabled = True
    RichTextBox1.Enabled = True
    LvMail.Enabled = True
    StatusBar1.Panels.Item(2).Text = "Logged in: " & FrmConnect.strUserName

End If
End With
End Sub

Private Sub TVcontact_DblClick()
If InStr(1, TVcontact.SelectedItem.Text, "[") = 1 Then _
Exit Sub
Dim i As Integer
Dim NewMessage As FrmNew
Set NewMessage = New FrmNew

With NewMessage
    
On Error Resume Next
    For i = 0 To UBound(Split(AllUsersList, "_")) - 1
        .Combo1.AddItem Split(AllUsersList, "_")(i)
    Next i
        .mnuNew.Enabled = False
        .mnuSave.Enabled = False
        .MnuDelete.Enabled = False
        .mnuMove.Enabled = False
        .Combo1.Text = TVcontact.SelectedItem.Text
        .Combo1.AddItem ""
        .Text1(0).Text = FrmConnect.strUserName
        .Text1(1).Text = ""
        .RichTextBox1.Text = ""
        .Text1(1).TabIndex = 0
        .RichTextBox1.TabIndex = 1
        .Toolbar1.Buttons(4).Enabled = False
    
        .Show
    
End With

Set NewMessage = Nothing



End Sub

Private Sub TVcontact_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then mnuRefresh_Click

End Sub

Private Sub TVcontact_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu Menu1
End If
End Sub

Private Sub TVdir_Click()
On Error Resume Next
Dim i As Integer
   
    With TVdir
        If Selected = .Nodes.Item(.SelectedItem.Index).Text Then
            Exit Sub
        Else
            Selected = .SelectedItem.Text
        End If
    End With
    
    For i = 0 To TVdir.Nodes.Count
        TVdir.Nodes(i).Bold = False
    Next i
   
If Selected = "[Personal Folders]" Then Exit Sub
If Selected = "Inbox" Then Me.Caption = "  Inbox Folder - " & "Office Messenger" Else Me.Caption = "  " & Selected & " - Office Messenger"
If Selected = "Inbox" Then Label9 = "  Inbox Folder" Else Label9 = "  " & Selected

With TVdir.SelectedItem
       .Bold = True
    If .Text = "Inbox" Then
        Call GetUserMessages("Discription")
         FrmMain.strMessage = ""
        Exit Sub
    End If


Call FrmMain.RemoveIcon
Call GetUserMessages(.Text)
End With
End Sub

Private Sub TVdir_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then mnuRefresh_Click

End Sub

Private Sub TVdir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    With TVdir.SelectedItem
        If .Text = "Rubbish Bin" Then Mnu2Rubbish.Visible = True Else Mnu2Rubbish.Visible = False
        If Not .Text = "Rubbish Bin" Or .Text = "[Personal Folders]" Or .Text = "Inbox" Then MnuDelFolder.Caption = "&Delete " & "'" & .Text & "'"
        If .Text = "Rubbish Bin" Or .Text = "[Personal Folders]" Or .Text = "Inbox" Or .Text = "Sent Items" Then MnuDelFolder.Visible = False Else MnuDelFolder.Visible = True
        
    End With
PopupMenu Menu2

End If

End Sub
Private Sub TVdir_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call DragFolder
End Sub
Public Sub DragFolder()
Dim GetFolders As FrmDrop
Set GetFolders = New FrmDrop
Dim TvFolder As Node
Dim i As Integer
    With GetFolders.TVdir
        .Nodes.Clear
        Set TvFolder = .Nodes.Add(, tvwFirst, , "[Personal Folders]", 3)
        Set TvFolder = .Nodes.Add(TvFolder, tvwChild, , "Inbox", 2)
            .Nodes.Item(1).Expanded = True
    

        For i = 3 To TVdir.Nodes.Count
            If TVdir.Nodes.Item(i).Text = "Rubbish Bin" Then
                Set TvFolder = .Nodes.Add(TvFolder, tvwNext, , TVdir.Nodes.Item(i).Text, 4)
                     GoTo nextI
            End If
           
            Set TvFolder = .Nodes.Add(TvFolder, tvwNext, , TVdir.Nodes.Item(i).Text, 2)
nextI:

        Next i
    End With
GetFolders.Message = DragMessage 'copy the message to the frmdrop form
GetFolders.Show 1


Set GetFolders = Nothing

End Sub
Public Sub Reply()
Dim replyMessage As FrmNew
Dim LvMessage As MsgLayout
Dim itm As ListItem
Dim i As Integer
On Error GoTo skipload
Set replyMessage = New FrmNew
'Set LvMessage = New MsgLayout

With replyMessage
    For i = 0 To UBound(Split(AllUsersList, "_")) - 1
            .Combo1.AddItem Split(AllUsersList, "_")(i)
                Next i
    Set itm = LvMail.ListItems.Item(LvMail.SelectedItem.Index)
    Lvstore.GetMsgStore (LvMail.SelectedItem.Text)
    .Combo1.AddItem itm.SubItems(1)
    .Combo1.Text = itm.SubItems(1)
    .Label4.Caption = itm.SubItems(3)
    .Text1(0).Text = FrmConnect.strUserName
    If InStr(1, itm.SubItems(2), "RE:", vbTextCompare) = 1 Then
        .Text1(1).Text = itm.SubItems(2)
            Else
        .Text1(1).Text = "RE: " & itm.SubItems(2)
    End If
    .Toolbar1.Buttons(4).Enabled = False
    .RichTextBox1.Text = Split(Lvstore.SendMessageID, "[~N10~]")(0)
    .RichTextBox1.Text = vbNewLine & vbNewLine & "------Original Message------ " & _
    vbNewLine & "From: " & itm.SubItems(1) & vbNewLine & "Sent: " & Format(itm.SubItems(3), "Long Date") & _
    vbNewLine & "To: " & itm.SubItems(1) & vbNewLine & "Subject: " & itm.SubItems(2) & _
    vbNewLine & vbNewLine & RichTextBox1.Text
    .RichTextBox1.TabIndex = 0
    .Show


End With
'Set LvMessage = Nothing
Set replyMessage = Nothing
Exit Sub
skipload:
MsgBox "Error with message, it may not exist, or is corrupted", vbExclamation + vbOKOnly, "Get Message Error"

'Set LvMessage = Nothing
Set replyMessage = Nothing
End Sub

Public Sub PrintFrm()
Dim FrmPrinter As FrmPrint
Set FrmPrinter = New FrmPrint
'Dim LvMessage As MsgLayout
'Set LvMessage = New MsgLayout
Dim itm As ListItem

Set itm = LvMail.ListItems.Item(LvMail.SelectedItem.Index)
    Lvstore.GetMsgStore (LvMail.SelectedItem.Text)
        FrmPrinter.Label1.Caption = itm.SubItems(1)
        FrmPrinter.Label2.Caption = "Subject: " & itm.SubItems(2)
        FrmPrinter.RichTextBox1.Text = Split(Lvstore.SendMessageID, "[~N10~]")(0)
        Const ErrCancel = 32755
            PrintDiag.CancelError = True
            
            On Error GoTo errorPrinter
                PrintDiag.Flags = 64
                PrintDiag.ShowPrinter
                FrmPrinter.PrintForm

        Set FrmPrinter = Nothing
       ' Set Lvstore = Nothing

errorPrinter:
            If Err = ErrCancel Then
        
        
    Set FrmPrinter = Nothing
    'Set Lvstore = Nothing
        Exit Sub
            End If
End Sub

Private Sub GetOptions()
FrmOptions.Show 1

End Sub

Public Sub FrmStyle()
Mnu0Preview.Checked = False
HiddenPreview = True
Frame1.Visible = False
HSplit.Visible = False
RichTextBox1.Visible = False
LvMail.Height = FrmMain.Height - 2340
Shape1.Height = FrmMain.Height - 2340

End Sub

