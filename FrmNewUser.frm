VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmNewUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registration"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "FrmNewUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSWinsockLib.Winsock RegSock 
      Left            =   6360
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1440
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   4680
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1440
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4200
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1320
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "Register"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   4320
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1440
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3720
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3240
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2760
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label Label15 
      Caption         =   "Town/City"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Required"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   25
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "Required"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   24
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Website:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Email:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "To join please enter in your details and click on the register button."
      Height          =   495
      Left            =   1560
      TabIndex        =   20
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registration Form"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   19
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label7 
      Caption         =   "Company:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Fax:"
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Phone:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Country:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   5160
      Picture         =   "FrmNewUser.frx":0442
      Top             =   240
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   -120
      TabIndex        =   12
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "FrmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReg_Click()
Call RInfoCheck
RegSock.SendData "RegisterNew" & Chr(10) & Text1(0).Text & "~~~" & _
Text1(1).Text & "~~~" & Text1(2).Text & "~~~" & Text1(3).Text & "~~~" & _
Text1(4).Text & "~~~" & Text1(5).Text & "~~~" & Text1(6).Text & "~~~" & _
Text1(7).Text & "~~~" & Text1(8).Text & "~~~" & Text1(9).Text & "~~~"


'name, address, address1, country, phone, fax, company, email, website, password

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call ClearForm
Call RConnect

End Sub

Private Sub ClearForm()
Dim i As Integer


    For i = 0 To Text1.UBound
        Text1(i).Text = ""
    Next i



End Sub

Private Sub RConnect()

     RegSock.Connect FrmConnect.Text1(2).Text, 9456

End Sub
Private Sub RInfoCheck()

If Text1(0).Text = "" Then MsgBox "Name Field is Required": Exit Sub
If Text1(9).Text = "" Then MsgBox "Enter a Password for this Account": Exit Sub


End Sub

Private Sub Form_Unload(Cancel As Integer)
RegSock.Close
End Sub

Private Sub RegSock_DataArrival(ByVal bytesTotal As Long)
Dim DataIN As String
RegSock.GetData DataIN

If Mid(DataIN, 1, 7) = "UsrExst" Then MsgBox Mid(DataIN, 8, Len(DataIN)), vbCritical + vbOKOnly, "User Name Already Taken"

If Mid(DataIN, 1, 10) = "UsrSuccess" Then
    MsgBox Mid(DataIN, 11, Len(DataIN)), vbInformation + vbOKOnly, "Successful!"
    FrmConnect.Text1(0).Text = Text1(0).Text
    FrmConnect.Text1(1).Text = Text1(9).Text
    Unload Me
End If

End Sub

