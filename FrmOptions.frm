VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "FrmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6240
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6015
      Begin VB.CheckBox Check3 
         Caption         =   "Display Content While Downloading"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   3135
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1660
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Export"
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Top             =   2160
         Width           =   3135
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Hide Message Preview"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Hide Download Status"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1420
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "Exporting Messages are only exported to a text file as commar delimited."
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Export Messages Folder To:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   6  'Inside Solid
         X1              =   240
         X2              =   5880
         Y1              =   1920
         Y2              =   1920
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5400
      Picture         =   "FrmOptions.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Control Panel"
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
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Check1.Value = 1 Then
    SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "HideStatus", "True"
    Else
    SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "HideStatus", "False"
End If

If Check3.Value = 1 Then
    SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "DisplayCont", "True"
    Else
    SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "DisplayCont", "False"
End If


If Check2.Value = 1 Then
    SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "HidePreview", "True"
    Call FrmMain.FrmStyle
    Else
    SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "HidePreview", "False"
    FrmMain.HiddenPreview = False
    FrmMain.Frame1.Visible = True
    FrmMain.RichTextBox1.Visible = True
    FrmMain.HSplit.Visible = True
    FrmMain.LvMail.Height = FrmMain.HSplit.Top - 1410
    FrmMain.Shape1.Height = FrmMain.HSplit.Top - 1410
End If

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

FrmConnect.Usersock.SendData "ExportMessages" & FrmConnect.strUserName

End Sub

Public Sub FileExport(sptmessage As String, recordset As String)
Dim ExMessage(6)
Dim j As Long
Dim FileName As String
Dim Message As String
    
    If Text1.Text = "" Then
        MsgBox "Enter a file name"
        Exit Sub
    End If
    
    If InStr(Text1.Text, ".") Then
        Text1.Text = Split(Text1.Text, ".")(0)
    
    End If
    
Label4.Caption = "Opening file for Export"
ProgressBar1.Max = recordset
FileName = Text1.Text & ".csv"

Open FileName For Output As #1
    For j = 0 To recordset
    
Label4.Caption = "Exporting Messages... " & j & " of " & recordset


ProgressBar1.Value = j

On Error Resume Next
Message = Split(sptmessage, "~}~")(j)
If Message = "" Then GoTo NextMessage

ExMessage(0) = Split(Message, "~%~")(0) 'Folder
ExMessage(2) = Split(Message, "~!~")(1): ExMessage(2) = Split(ExMessage(2), "~#~")(0) 'From
ExMessage(3) = Split(Message, "~#~")(1): ExMessage(3) = Split(ExMessage(3), "~@~")(0) 'Subject
ExMessage(4) = Split(Message, "~@~")(1): ExMessage(4) = Split(ExMessage(4), "~^~")(0): ExMessage(4) = Split(ExMessage(4), "[~N10~]")(0) 'Discription
ExMessage(5) = Split(Message, "~^~")(1): ExMessage(5) = Split(ExMessage(5), "ó")(0) 'Date



Write #1, j & "," & """" & ExMessage(2) & """" & "," & """" & ExMessage(3) & """" & "," & """" & ExMessage(4) & """" & "," & """" & ExMessage(5) & """" & ","



NextMessage:
    Next j
  
Close #1
MsgBox "Done!" & Chr(10) & "FileName = " & UCase(FileName)
ProgressBar1.Value = 0

Label4.Caption = "Exporting Messages Completed!"


End Sub

Private Sub Form_Load()
If GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "HidePreview", "") = "True" Then Check2.Value = 1
If GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "HideStatus", "") = "True" Then Check1.Value = 1
If GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "DisplayCont", "") = "True" Then Check3.Value = 1


End Sub

