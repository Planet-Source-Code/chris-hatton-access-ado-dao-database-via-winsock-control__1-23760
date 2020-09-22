Attribute VB_Name = "SckRecordset"
Dim StrUser As String 'Without "/" port #
Dim OldUser As String 'with "/" Port #
Public Folders As String
Public CacheFolder As String
Public CurrentMsg As Long 'tells this module what current record we are dealing with for deleting records
Dim StrMessage As Variant
Public MessageCounter As Long 'tells the client what message count we are up to (for the single messages function)
Public Allusers As String 'allows the client to see all the registred users.
Public IusrCom, IusrName, IusrAddy, IusrAddy1, IusrPhone, IusrFax, IusrEmail, IusrIP, IusrWeb As String
Public AcPOP, AcSmtp, AcAccount, AcPass As String

Public Sub sckUserName(UserName As String)
On Error GoTo UsrErr
Dim rs As ADODB.Recordset                   'This subroutine checks the database
Set rs = New ADODB.Recordset                'to see if the person logging in actually
Dim ChkLogon As MultiSck
Set ChkLogon = New MultiSck
Dim Sql As String                           'has an account here.
OldUser = UserName & "/" & FrmServer.sckmax 'orginal winsock port
UserName = Split(UserName, Chr(10))(1)
Sql = "Select UserName from Users where UserName = " _
& Chr(34) & UserName & Chr(34)

rs.Open Sql, cn, adOpenForwardOnly, adLockReadOnly

If UCase(UserName) = UCase(rs!UserName) Then
StrUser = rs!UserName   'make this public for this module and then verify it with the password
'Call ChkLogon.parseuser(UserName & "/" & FrmServer.sckmax)
                                    'now that the user is in the database, we must
                                    'find out the exact port his is on or we will
                                    'get criscross passwords.
Call ChkLogon.GetSck5(UserName)
End If

Set ChkLogon = Nothing
rs.Close
Set rs = Nothing
Exit Sub

UsrErr:

    Dim UserErr As MultiSck
    Set UserErr = New MultiSck
        OldUser = Split(OldUser, "/")(1)
        UserErr.UsrErr OldUser, True
    Set UserErr = Nothing

End Sub

Public Sub sckPassword(Password As String)
On Error GoTo PassErr
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim Sql As String
Dim ChkPass As MultiSck
Set ChkPass = New MultiSck

Password = Split(Password, Chr(10))(1)

Sql = "Select * from Users where UserName = " _
& Chr(34) & StrUser & Chr(34)

rs.Open Sql, cn, adOpenForwardOnly, adLockReadOnly
'Debug.Print Password & rs!Password

    If UCase(rs!Password) = UCase(Password) Then
        ChkPass.GetSck6 (StrUser)
       
    Else
        GoTo PassErr
    End If
    
rs.Close
Set rs = Nothing
Set ChkPass = Nothing
Exit Sub
PassErr:
    Dim PasswordErr As MultiSck
    Set PasswordErr = New MultiSck
        OldUser = Split(OldUser, "/")(1)
        PasswordErr.PassErr OldUser, True
    Set PasswordErr = Nothing

End Sub
Public Sub DelFolder(Folder As String)
On Error Resume Next

Set Table = DB.TableDefs(StrUser)

Table.Fields.Delete Folder


If Err.Description = "" Then
    Dim FolderUpdate As MultiSck
    Set FolderUpdate = New MultiSck
        With FolderUpdate
            DoEvents
            .GetSck (StrUser)
            DoEvents
            .DelFolder (Folder)
        End With
    Set FolderUpdate = Nothing
Else
    Dim ErrUpdate As MultiSck
    Set ErrUpdate = New MultiSck
        With ErrUpdate
            .GetSck (StrUser)
            .ErrFolder (Folder)
        End With
    Set ErrUpdate = Nothing
End If
End Sub
Public Sub NewFolder(Folder As String)
Set Table = DB.TableDefs(StrUser)
Set FL = Table.CreateField(Folder, dbMemo)
On Error Resume Next
Table.Fields.Append FL

If Err.Description = "" Then '*
    Dim FolderUpdate As MultiSck
    Set FolderUpdate = New MultiSck     'sends user the new folder if
        With FolderUpdate               'it was successfully created
            DoEvents
            .GetSck (StrUser)           'in database.
            .AddFolder (Folder)
        End With
    Set FolderUpdate = Nothing
Else
    Dim ErrUpdate As MultiSck           'if theres an error tell user
    Set ErrUpdate = New MultiSck        'that the folder has'nt been
        With ErrUpdate                  'created.
            .GetSck (StrUser)
            .ErrFolder (Folder)
        End With
    Set ErrUpdate = Nothing
End If
End Sub

Public Sub DBMessages(UserName As String)
Dim i As Integer
Dim strFields As String
Dim strInbox, strID, StrSub, strFrom, strDate As String
Dim GetFolder As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim NewUser As Long
Dim Read As String
GetFolder = Split(UserName, Chr(10))(1)
UserName = Split(UserName, Chr(10))(0)

rs.Open "select * from [" & UserName & "]", cn, adOpenStatic, adLockReadOnly

        For i = 3 To 3
            For j = 1 To rs.RecordCount
            If IsNull(rs.Fields(GetFolder)) = True Then GoTo MoveNext
                strID = "~*~" & rs!Msgid
                strFrom = "~!~" & rs!From
                StrSub = "~#~" & rs!Subject
                If InStr(1, rs.Fields(GetFolder), "[~N10~]", vbTextCompare) = 0 Then Read = "Y" Else Read = "N"
                strInbox = "~@~" & rs.Fields(GetFolder)
                strDate = "~^~" & rs!Rdate & "ó" & rs.RecordCount & "~'~" & Read & "~}~"
                strFields = GetFolder
                Folders = Folders & strFields & "~%~" & _
                strID & strFrom & StrSub & strInbox & strDate
'Debug.Print Folders
MoveNext:       'if the current record = null then skip the message
                rs.MoveNext
            
            Next j
        Next i
        
If Folders = "" Then
End If
'Debug.Print Folders
rs.Close
Set rs = Nothing

End Sub
Public Sub MoveRecord(UserName, Folder, From, Subject, Message, Rdate, Msgid, StrUser As String)
If Folder = "Messages Folder" Then Folder = "Discription"
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim i As Integer

rs.Open "select * from [" & UserName & "]", cn, adOpenKeyset, adLockOptimistic

rs.AddNew
rs.Fields(Folder) = Message
rs!From = From
rs!Subject = Subject
rs!Rdate = Rdate
rs.Update
rs.Close
Set rs = Nothing
CurrentMsg = Msgid
If Err.Description = "" Then
    Dim MoveOK As MultiSck
    Set MoveOK = New MultiSck
        With MoveOK
            .GetSck2 (StrUser)
            .MoveOK
       
        End With
    Set MoveOK = Nothing
End If
If Not Err.Description = "" Then
    Dim MoveError As MultiSck
    Set MoveError = New MultiSck
        With MoveError
             .GetSck2 (StrUser)
             .MoveErr
        End With
    Set MoveError = Nothing
End If
End Sub
Public Sub DelRecord(UserName As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim Sql As String
Sql = "Select * from [" & UserName & "] where msgid = " & CurrentMsg

rs.Open Sql, cn, adOpenKeyset, adLockOptimistic
On Error Resume Next
If rs!Msgid = CurrentMsg Then
    rs.Delete
    Else
    MsgBox "Delete field " & CurrentMsg
End If
rs.Update

rs.Close
Set rs = Nothing

End Sub
'This sub deletes any message that the user selects in there message list
Public Sub DelMessage(UserName As String, MessageNumber As Long)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim Sql As String
Sql = "Select * from [" & UserName & "] where msgid = " & MessageNumber

rs.Open Sql, cn, adOpenKeyset, adLockOptimistic
On Error Resume Next
If rs!Msgid = MessageNumber Then
    rs.Delete
    Else
    MsgBox "can't Delete field " & MessageNumber
End If
rs.Update

rs.Close
Set rs = Nothing

End Sub
Public Sub NewMessage(UserName As String, WhoFrom, Subject, Message, Rdate As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim Sql As String
StrMessage = Message
    DoEvents
    
    Sql = "Select * from [" & UserName & "] "
    rs.Open Sql, cn, adOpenKeyset, adLockOptimistic
        
       On Error Resume Next
    rs.AddNew

        rs!From = WhoFrom
        rs!Subject = Subject
        rs!Discription = StrMessage
        rs!Rdate = Rdate & " " & Format(Now, "short Time")

    rs.Update
    rs.Close
Set rs = Nothing
End Sub
Public Sub SentMessage(UserName As String, WhoFrom, Subject, Message, Rdate As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim Sql As String
StrMessage = Message
    DoEvents
    
    Sql = "Select * from [" & WhoFrom & "] "
    rs.Open Sql, cn, adOpenKeyset, adLockOptimistic
        
       On Error Resume Next
    rs.AddNew

        rs!From = WhoFrom
        rs!Subject = Subject
        rs![Sent Items] = StrMessage
        rs!Rdate = Rdate & " " & Format(Now, "short Time")

    rs.Update
    rs.Close
Set rs = Nothing
End Sub

Public Sub EditMessage(UserName, Record, Folder, Message As Variant)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim Sql As String
    On Error Resume Next
    Sql = "Select * from [" & UserName & "] where msgid = " & Record
    rs.Open Sql, cn, adOpenKeyset, adLockOptimistic
        
        rs.Fields(Folder) = "" & Message
        
    rs.Update
    rs.Close
Set rs = Nothing
End Sub
Public Sub MessageCount(UserName As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim Sql As String
    On Error Resume Next
    Sql = "Select * from [" & UserName & "] "
    rs.Open Sql, cn, adOpenKeyset, adLockOptimistic
        rs.MoveLast
        MessageCounter = rs!Msgid



rs.Close
Set rs = Nothing
End Sub
Public Sub OfflneUsers()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim i As Integer
Dim Sql As String
    On Error Resume Next
    Sql = "Select * from users"
    rs.Open Sql, cn, adOpenKeyset, adLockOptimistic
        
    For i = 1 To rs.RecordCount
        Allusers = Allusers & rs!UserName & "_"
        
        rs.MoveNext
                Next i
      
    rs.Update
    rs.Close
Set rs = Nothing
End Sub

Public Sub EmptyBin(UserName As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim i As Integer
Dim Sql As String
    On Error Resume Next
    Sql = "Select * from [" & UserName & "] "
    rs.Open Sql, cn, adOpenKeyset, adLockOptimistic
        
    For i = 1 To rs.RecordCount
        If IsNull(rs![Rubbish Bin]) = False Then rs.Delete
       
        rs.MoveNext
                Next i
      
   rs.Update
    rs.Close
Set rs = Nothing

End Sub

Public Sub UsrInfo(UserName As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim Iuser As String
Dim i As Integer
Dim Sql As String
   Sql = "Select * from Users where UserName = " & Chr(34) & UserName & Chr(34)
    rs.Open Sql, cn, adOpenKeyset, adLockOptimistic
On Error Resume Next
        IusrCom = "" & rs!Company
        IusrName = "" & rs!UserName
        IusrAddy = "" & rs!address
        IusrAddy1 = "" & rs!address1
        IusrPhone = "" & rs!Phone
        IusrFax = "" & rs!Fax
        IusrEmail = "" & rs!Email
        IusrWeb = "" & rs!Website
    rs.Close
    
    IusrIP = "N/A"
    
    'On Error Resume Next
        For i = 0 To FrmServer.Userlist.ListCount - 1
            If IusrName = Split(FrmServer.Userlist.List(i), "/")(0) Then
                Iuser = Split(FrmServer.Userlist.List(i), "/")(1)
                IusrIP = FrmServer.ServerSck(Iuser).RemoteHostIP
               
               Else
               
            
            End If
    
        Next i




Set rs = Nothing


End Sub
Public Sub CacheMessages(UserName, GetFolder As String)
Dim i As Integer
Dim strFields As String
Dim strInbox, strID, StrSub, strFrom, strDate As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim Read As String

rs.Open "select * from [" & UserName & "]", cn, adOpenStatic, adLockReadOnly

        For i = 3 To 3
            For j = 1 To rs.RecordCount
            If IsNull(rs.Fields(GetFolder)) = True Then GoTo MoveNext
                strID = "~*~" & rs!Msgid
                strFrom = "~!~" & rs!From
                StrSub = "~#~" & rs!Subject
                If InStr(1, rs.Fields(GetFolder), "[~N10~]", vbTextCompare) = 0 Then Read = "Y" Else Read = "N"
                strInbox = "~@~" & rs.Fields(GetFolder)
                strDate = "~^~" & rs!Rdate & "ó" & rs.RecordCount & "~'~" & Read & "~}~"
                strFields = GetFolder
                CacheFolder = CacheFolder & strFields & "~%~" & _
                strID & strFrom & StrSub & strInbox & strDate
'Debug.Print CacheFolder
MoveNext:       'if the current record = null then skip the message
                rs.MoveNext
            
            Next j
        Next i
        

rs.Close
Set rs = Nothing

End Sub
Public Sub MailAccount(UserName As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    Dim Sql As String
    Sql = "select * from users where username = " & Chr(34) & UserName & Chr(34)
    
    rs.Open Sql, cn, adOpenStatic, adLockReadOnly
    
    AcPOP = "" & rs!pop
    AcSmtp = "" & rs!smtp
    AcAccount = "" & rs!account
    AcPass = "" & rs!mailpassword
    
    rs.Close
    Set rs = Nothing

End Sub


