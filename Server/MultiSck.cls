VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiSck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public uPort As Long
Public Sck As Long
Dim sck2 As Long 'handle the messaging
Dim sck4 As Long
Dim sck5 As Long
Dim sck6 As Long

Public Sub parseuser(UserName As String)
Dim Port As Long
Port = Split(UserName, "/")(1)
If Port > 0 Then uPort = Port

If Port <= 0 Then
        MsgBox "Winsock Port Error" & vbNewLine & _
        "Port Number = " & uPort, vbCritical + vbOKOnly, "Winsock Error"
    Else
        Call sendPass
End If

End Sub

Private Sub sendPass()

FrmServer.ServerSck(sck5).SendData "PasswordRequest"

End Sub
'password verify
Public Sub GetSck6(UserName As String)       'updates server to find out sock number
Dim i As Integer                            'also the main role of folders and logging in
On Error Resume Next                        'users
With FrmServer.Userlist
For i = 0 To .ListCount - 1
        If UserName = Split(.List(i), "/")(0) Then
            sck6 = Split(.List(i), "/")(1) 'make socket public
        End If                                             'so the data goes to
    Next i                                                 'the right person.
    Call sendAccept
End With

End Sub

Public Sub LogonSuccess(UserName As String)
Dim Port As Long
Port = Split(UserName, "/")(1)
If Port > 0 Then uPort = Port

If Port <= 0 Then
        MsgBox "Winsock Port Error" & vbNewLine & _
        "Port Number = " & uPort, vbCritical + vbOKOnly, "Winsock Error"
    Else
        Call sendAccept
End If

End Sub

Private Sub sendAccept()
FrmServer.ServerSck(sck6).SendData "RequestAccepted"   'UserName and Password OK.
FrmServer.SndUserList = True    'Send the Online userlist to the user
End Sub

Public Sub UsrErr(UserName As String, User As Boolean)
Dim users As Integer
With FrmServer
For users = 0 To .Userlist.ListCount - 1
If Split(.Userlist.List(users), "/")(1) = UserName Then 'remove person from list
    .Userlist.RemoveItem (users)        'thats not a registered user
    End If
Next users

.ServerSck(UserName).SendData "UserNameFailed"
End With


End Sub

Public Sub PassErr(UserName As String, Password As Boolean)
Dim users As Integer
With FrmServer
For users = 0 To .Userlist.ListCount - 1
If Split(.Userlist.List(users), "/")(1) = UserName Then
    .Userlist.RemoveItem (users)
    End If
Next users

.ServerSck(UserName).SendData "PasswordFailed"

End With


End Sub
'username verify
Public Sub GetSck5(UserName As String)       'updates server to find out sock number
Dim i As Integer                            'also the main role of folders and logging in
On Error Resume Next                        'users
With FrmServer.Userlist
For i = 0 To .ListCount - 1
        DoEvents
        If UserName = Split(.List(i), "/")(0) Then
            sck5 = Split(.List(i), "/")(1) 'make socket public
        End If                                             'so the data goes to
    Next i                                                 'the right person.
    Call sendPass
End With

End Sub

Public Sub GetSck(UserName As String)       'updates server to find out sock number
Dim i As Integer                            'also the main role of folders and logging in
On Error Resume Next                        'users
With FrmServer.Userlist
For i = 0 To .ListCount - 1
        DoEvents
        If UserName = Split(.List(i), "/")(0) Then
            Sck = Split(.List(i), "/")(1) 'make socket public
        End If                                             'so the data goes to
    Next i                                                 'the right person.
    
End With
End Sub
Public Sub LstOffline()
Dim User As Integer
             
    SckRecordset.OfflneUsers
    FrmServer.ServerSck(Sck).SendData "OfflineList" & Allusers
    Allusers = ""

End Sub
Public Sub GetUserI(UserName As String)
 

    SckRecordset.UsrInfo (UserName)
    
    FrmServer.ServerSck(Sck).SendData "UserInfo" & Chr(10) & SckRecordset.IusrCom & "~~~" & _
    SckRecordset.IusrName & "~~~" & SckRecordset.IusrAddy & "~~~" & SckRecordset.IusrAddy1 & "~~~" & SckRecordset.IusrPhone & "~~~" & _
    SckRecordset.IusrFax & "~~~" & SckRecordset.IusrEmail & "~~~" & SckRecordset.IusrIP & "~~~" & SckRecordset.IusrWeb & "~~~"

End Sub


Public Sub sendFolders(Folder As String)

With FrmServer
    .ServerSck(Sck).SendData "CustomFolders" & Chr(10) & Folder
End With
End Sub
Public Sub DelFolder(Folder As String)
With FrmServer
    .ServerSck(Sck).SendData "DelFolder" & Chr(10) & Folder
End With
End Sub

Public Sub AddFolder(Folder As String)

With FrmServer
    .ServerSck(Sck).SendData "AddFolder" & Chr(10) & Folder
    
End With
End Sub
Public Sub ErrFolder(Folder As String)

With FrmServer
    .ServerSck(Sck).SendData "ErrFolder" & Chr(10) & Folder
End With

End Sub
Public Sub SendMsgs(UserName As String)
With FrmServer

        On Error Resume Next
        Call DBMessages(UserName)       'get personnal messages
        .ServerSck(sck2).SendData "Messages" & Chr(10) & SckRecordset.Folders
        
        SckRecordset.Folders = ""
        
End With

End Sub
Public Sub SendExport(UserName As String)
With FrmServer
       UserName = UserName & Chr(10) & "Discription" 'inbox messages folder
        On Error Resume Next
        Call DBMessages(UserName)       'get personnal messages
        
        .ServerSck(sck2).SendData "Exported" & Chr(10) & SckRecordset.Folders
        
        SckRecordset.Folders = ""
        
End With

End Sub

Public Sub GetSck2(UserName As String)  'This handles all the messaging side of things
Dim i As Integer
On Error Resume Next
With FrmServer.Userlist
UserName = Split(UserName, Chr(10))(0)
For i = 0 To .ListCount - 1
        DoEvents
        If UserName = Split(.List(i), "/")(0) Then
            sck2 = Split(.List(i), "/")(1)
        End If
    Next i
   
End With
End Sub

Public Sub MoveOK()

With FrmServer
    .ServerSck(sck2).SendData "MoveOK" 'tells the client that the move was good
End With

End Sub
Public Sub MoveErr()

With FrmServer
    .ServerSck(sck2).SendData "MoveError" & Chr(10) & "There was a Problem Moving the Message"
End With

End Sub
Public Sub GetSck4(UserName As String)  'This handles new message notificatino
Dim i As Integer
On Error Resume Next
With FrmServer.Userlist
UserName = Split(UserName, Chr(10))(0)
For i = 0 To .ListCount - 1
        
       If UCase(UserName) = Split(UCase(.List(i)), "/")(0) Then
            sck4 = Split(.List(i), "/")(1)
        End If
    Next i
  
End With
End Sub
Public Sub RefreshList() 'refreshes all the users
On Error Resume Next
With FrmServer
        .ServerSck(sck2).SendData "RefreshList"
     
End With

End Sub
Public Sub Notifiy(strWho, StrSub, StrMsg, Tdate As String, counter As Long) 'refreshes just the current user.
On Error Resume Next
With FrmServer
        .ServerSck(sck4).SendData "GetNewMail" & strWho & "~~" & StrSub & _
        "~~" & StrMsg & "~~" & Tdate & "~~" & counter
End With

End Sub

Public Sub SendCache(UserName, Folder As String)
With FrmServer
        On Error Resume Next
       
        Call CacheMessages(UserName, Folder)      'get personnal messages
        
        .ServerSck(sck2).SendData "CacheImport" & Chr(10) & Folder & Chr(10) & SckRecordset.CacheFolder
        
        SckRecordset.CacheFolder = "" 'clear the variable
        
End With

End Sub
Public Sub GetMailAcc(UserName As String)
With FrmServer
        On Error Resume Next
       
        Call MailAccount(UserName)
        .ServerSck(sck2).SendData "MailDetails" & Chr(10) & SckRecordset.AcPOP & _
        Chr(10) & SckRecordset.AcSmtp & Chr(10) & SckRecordset.AcAccount & _
        Chr(10) & SckRecordset.AcPass
        
      
        
End With
End Sub

