VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public RSck As Long

Public Sub MSGUsrExst()

FrmServer.ServerSck(RSck).SendData "UsrExst" & "User Name has already been taken" & Chr(10) & _
"Please choose another name"

End Sub

Public Sub MSGUsrSuccess()

FrmServer.ServerSck(RSck).SendData "UsrSuccess" & "Your user name has been Registered Successfully" & Chr(10) & _
"Your details have been saved and click connect when ready."

End Sub


