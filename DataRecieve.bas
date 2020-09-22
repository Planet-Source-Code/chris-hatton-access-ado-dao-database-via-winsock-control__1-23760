Attribute VB_Name = "DataRecieve"
Public Const WM_SYSCOMMAND = &H112
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Dim NetStatus As String
Dim strStatus
Public SckStat As Integer
Public ATList As Variant

Public Property Get Status() As String
NetStatus = NetStatus
Status = NetStatus
End Property
Public Sub newFolder()
    FrmFolder.Show 1

End Sub


Public Property Let Status(ByVal vNewValue As String)
If vNewValue = 0 Then NetStatus = ""
If vNewValue = 1 Then NetStatus = "Connecting to Ado Server" & vbNewLine

If vNewValue = 2 Then
    NetStatus = NetStatus & "Connection Closed" & vbNewLine
    FrmConnect.CmdCancel.Caption = "Ca&ncel"
End If

If vNewValue = 3 Then
    NetStatus = "Connection Established" & vbNewLine
    FrmConnect.Caption = "Connected to Office Server"
    FrmConnect.CmdCancel.Caption = "D&isconnect"
End If
If vNewValue = 4 Then NetStatus = NetStatus & "Connection Error" & vbNewLine
If vNewValue = 5 Then NetStatus = NetStatus & "Verifing Username and Password" & vbNewLine
If vNewValue = 6 Then NetStatus = "Logging onto Network" & vbNewLine
If vNewValue = 7 Then NetStatus = NetStatus & "User Name Not Registered" & vbNewLine
If vNewValue = 8 Then NetStatus = NetStatus & "Incorrect Password" & vbNewLine

SckStat = vNewValue
End Property
Public Property Get Statusbar() As String
strStatus = strStatus
Statusbar = strStatus
End Property

Public Property Let Statusbar(ByVal vNewValue As String)
If vNewValue = 0 Then strStatus = ""
If vNewValue = 1 Then strStatus = "Downloading Folder Info"
If vNewValue = 2 Then strStatus = "Downloading Messages"
If vNewValue = 3 Then strStatus = "Downloading Online Users List"
If vNewValue = 4 Then strStatus = "Downloading Offline Users List"

End Property


Public Sub ParseData(GetRecv As String)
    Dim i As Integer
    Dim itm As ListItem
    Dim AddCustom As String
    Dim VD As PwSettings
    Set VD = New PwSettings

    Set OFMSGER = New MsgLayout
If GetRecv = "welcome" & Chr(10) Then
    Status = 5
    FrmConnect.Usersock.SendData "VUserName" & Chr(10) & VD.UserName & Chr(10)
    FrmMain.WindowState = vbNormal
    FrmConnect.Timer2.Enabled = False
End If

If GetRecv = "PasswordRequest" Then
    
    VD.SvPassword = True
    FrmConnect.Usersock.SendData "VPassword" & Chr(10) & FrmConnect.Text1(1).Text
    Set VD = Nothing
    End If
    
If Mid(GetRecv, 1, 15) = "RequestAccepted" Then
    Status = 6                      'username and password accepted Log onto network.
    FrmMain.Show
    FrmMain.Logged = True           'all logged in now get online userlist
    If Not FrmConnect.WindowState = vbMinimized Then FrmConnect.WindowState = vbMinimized
    FrmMain.WindowState = vbMaximized
End If
         
If GetRecv = "NoConnection" Then
    MsgBox "No more Available Connections" & vbCrLf & "See help for more details", vbExclamation + vbOKOnly
    FrmConnect.CmdCancel = True
End If
    
If GetRecv = "UserNameFailed" Then
    Status = 0: Status = 5: Status = 7
    FrmConnect.Usersock.Close
    Status = 2
    FrmMain.WindowState = vbMinimized
    MsgBox "User Name is not registered", vbExclamation + vbOKOnly, "Authentication Error"
    FrmConnect.cmdcon.Enabled = True
End If

If GetRecv = "PasswordFailed" Then
    Status = 0: Status = 5: Status = 8
    FrmConnect.Usersock.Close
    'Status = 2
    FrmConnect.cmdcon.Enabled = True
    FrmMain.WindowState = vbMinimized
    MsgBox "Incorrect Password", vbCritical + vbOKOnly, "Authentication Error"
End If

If Mid(GetRecv, 1, 8) = "UserList" Then
    DataRecieve.Status = 3

    Call ModUserList.tvtree(GetRecv)
    FrmConnect.Usersock.SendData "OffList" & FrmConnect.strUserName  'Ask server for the offline list
    Statusbar = 4
    End If

If Mid(GetRecv, 1, 11) = "OfflineList" Then 'Got the Offline List now add it to the treeview control
    Call ModUserList.TvOffline(GetRecv)
    
    End If
Set VD = Nothing
 
If Mid(GetRecv, 1, 9) = "AddFolder" Then    'creating a new custom folder
    AddFolder (Mid(GetRecv, 11, Len(GetRecv)))
   Unload FrmFolder
End If
 
If Mid(GetRecv, 1, 9) = "ErrFolder" Then
    MsgBox "Error Creating New Folder " & vbNewLine & _
    "Make sure the folder " & Mid(GetRecv, 11, Len(GetRecv)) & " doesn't already exist", vbCritical, "New Folder Error"
    Unload FrmFolder
End If

If Mid(GetRecv, 1, 13) = "CustomFolders" Then
    For i = 5 To UBound(Split(Mid(GetRecv, 14, Len(GetRecv)), "-"))
    AddCustom = Split(Mid(GetRecv, 14, Len(GetRecv)), "-")(i)
        Call AddFolder(AddCustom)
    Next i
    Call GetUserMessages("Discription") 'Default folder
End If

If Mid(GetRecv, 1, 9) = "DelFolder" Then             'Delete Selected Folder
FrmMain.TVdir.Nodes.Remove (FrmMain.TVdir.SelectedItem.Index)
FrmMain.MousePointer = 1
End If

If Mid(GetRecv, 1, 8) = "Messages" Then
With FrmMain.LvMail

Dim sptmessage As String
Dim recordset As String
On Error Resume Next
.ListItems.Clear
    sptmessage = Mid(GetRecv, 10, Len(GetRecv))
    recordset = Split(sptmessage, "ó")(1): recordset = Split(recordset, "~'~")(0)
    OFMSGER.MsgLyt sptmessage, recordset
    'FrmConnect.Usersock.SendData "Clear"


End With
End If

If Mid(GetRecv, 1, 11) = "RefreshList" Then
FrmMain.LvMail.ListItems.Clear
    Call GetUserMessages("Discription")
    Statusbar = 4
End If

If Mid(GetRecv, 1, 10) = "GetNewMail" Then
  Dim NewMail(3)
  Dim SDate As String
  Dim Counter As Long
 ' Dim AddMaiLst As MsgLayout
 ' Set AddMaiLst = New MsgLayout
    NewMail(1) = Split(Mid(GetRecv, 11, Len(GetRecv)), "~~")(0)
    NewMail(2) = Split(Mid(GetRecv, 11, Len(GetRecv)), "~~")(1)
    NewMail(3) = Split(Mid(GetRecv, 11, Len(GetRecv)), "~~")(2)
         SDate = Split(Mid(GetRecv, 11, Len(GetRecv)), "~~")(3)
       Counter = Split(Mid(GetRecv, 11, Len(GetRecv)), "~~")(4)
       
    If Not FrmMain.TVdir.SelectedItem.Text = "Inbox" Then
       ' Set AddMaiLst = Nothing
        FrmMain.AddIcon
        Exit Sub
    End If
  DoEvents
    OFMSGER.SingleMessage NewMail(1), NewMail(2), NewMail(3), SDate, Counter
  DoEvents
  'Set AddMaiLst = Nothing

End If


If Mid(GetRecv, 1, 6) = "MoveOK" Then
    FrmConnect.Usersock.SendData "DeleteRecord" & Chr(10) & FrmConnect.strUserName
End If

If Mid(GetRecv, 1, 8) = "MoveError" Then
    MsgBox Mid(GetRecv, 9, Len(GetRecv)), vbCritical + vbOKOnly, "Error"
End If

If Mid(GetRecv, 1, 8) = "Exported" Then
    Dim SplitMessage As String
    Dim RecordCount As String
        On Error Resume Next
        SplitMessage = Mid(GetRecv, 8, Len(GetRecv))
        RecordCount = Split(SplitMessage, "ó")(1): RecordCount = Split(RecordCount, "~'~")(0)
        

        Call FrmOptions.FileExport(SplitMessage, RecordCount)

End If

If Mid(GetRecv, 1, 8) = "UserInfo" Then
    Dim ShowUser As String
    ShowUser = Mid(GetRecv, 10, Len(GetRecv))
    FrmInfo.Label8 = "" & Split(ShowUser, "~~~")(0)
    FrmInfo.Label9 = "" & Split(ShowUser, "~~~")(1)
    FrmInfo.Label10 = "" & Split(ShowUser, "~~~")(2)
    FrmInfo.Label11 = "" & Split(ShowUser, "~~~")(3)
    FrmInfo.Label12 = "" & Split(ShowUser, "~~~")(4)
    FrmInfo.Label13 = "" & Split(ShowUser, "~~~")(5)
    FrmInfo.Label14 = "" & Split(ShowUser, "~~~")(6)
    FrmInfo.Label15 = "" & Split(ShowUser, "~~~")(7)
    FrmInfo.Label17 = "" & Split(ShowUser, "~~~")(8)
    FrmInfo.Caption = "User Information  (" & FrmMain.TVcontact.SelectedItem.Text & ")"
    FrmInfo.Show 1
  
End If

If Mid(GetRecv, 1, 11) = "CacheImport" Then
    Dim GetLength As String
    Dim CaFolder As String
    Dim CaRecord As String
        On Error Resume Next
        
        GetLength = Mid(GetRecv, 12, Len(GetRecv))
        CaFolder = Split(GetRecv, Chr(10))(1)
        CaRecord = Split(GetLength, "ó")(1): CaRecord = Split(CaRecord, "~'~")(0)
       
               
        'data, recordcount, foldername
        Call OFSCache.CacheTemp(GetLength, CaRecord, CaFolder)

End If


If Mid(GetRecv, 1, 11) = "MailDetails" Then
 
    FrmMail.Text1(0).Text = Split(Mid$(GetRecv, 12, Len(GetRecv)), Chr(10))(1)
    FrmMail.Text1(1).Text = Split(Mid$(GetRecv, 12, Len(GetRecv)), Chr(10))(2)

    FrmMail.Label6.Caption = vbCrLf & "    Email Account Setup"
    FrmMail.Show 1

End If


End Sub
Public Sub GetUserMessages(GetFolder As String)
   Dim i As Integer
    
    If GetFolder = "" Then GetFolder = "Discription" 'Discription is the default vaule for 'Inbox'
    'On Error Resume Next
    For i = 1 To FrmMain.TVdir.Nodes.Count
        FrmMain.TVdir.Nodes(i).Bold = False
    Next i
  
    If GetFolder = "Discription" Then
        FrmMain.TVdir.Nodes(2).Selected = True
        FrmMain.TVdir.Nodes.Item(2).Bold = True
    Else
    FrmMain.TVdir.SelectedItem.Bold = True
    End If
    
    FrmConnect.Usersock.SendData "GetMessages" & Chr(10) & FrmConnect.strUserName & Chr(10) & GetFolder
    Statusbar = 2           'Get the messages for the 'Inbox', (defaults to this folder)

End Sub
