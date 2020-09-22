Attribute VB_Name = "Recordsets"
Option Explicit
Public RegUser As String
Public UsrExist As Boolean 'If the user is not in database this stops the form from clearing.
Public usrEmail ' make users email addy public to the world
Public UsrToEmail ' make users TO: email addy public to the world
Public usrCom, usrName, usrAddy, usrAddy1, usrPhone, usrFax
Public rego As Registration

Public Sub RegUsers()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim i As Integer
rs.Open "Select Username from Users", cn, adOpenForwardOnly, adLockReadOnly

For i = 1 To rs.RecordCount

FrmProfile.lusers.AddItem rs!UserName 'add users to the list box

rs.MoveNext

Next i

rs.Close

End Sub
Public Sub ClearForm()
Dim j As Integer
With FrmProfile.Text1
    For j = .LBound To .UBound
    .Item(j).Text = ""
    Next j
End With

End Sub

Public Sub AddAccount()             'add the user account to the database
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim i As Integer
Dim Sql As String
rs.Open "Select * from Users", cn, adOpenKeyset, adLockOptimistic

For i = 1 To rs.RecordCount

If FrmProfile.Text1(1).Text = rs!UserName Then
    MsgBox "User Already Exists", vbInformation
    Else
    rs.MoveNext
End If

Next i
With FrmProfile
On Error Resume Next
If rs.EOF = True Then
    rs.AddNew
    rs!Company = .Text1(0).Text
    rs!UserName = .Text1(1).Text
    rs!Password = .Text1(2).Text
    rs!Phone = .Text1(3).Text
    rs!Email = .Text1(4).Text
    rs!Discription = .Text1(5).Text
    
    rs.Update
    rs.Close
    
    CreateUserTable .Text1(1).Text
    ClearForm
    Set rs = Nothing
    End If
End With




End Sub
Public Function GetFolders(UserName As String)  'get all the custom folders
Dim i As Integer                                'created by the user
'Dim DB As Database
Dim FieldsList As String
UserName = Split(UserName, Chr(10))(1)
'Set DB = OpenDatabase(App.Path & "\OSDB.mdb")
    For i = 1 To 30
    On Error GoTo sndFolders
        FieldsList = FieldsList & "-" & DB.TableDefs(UserName).Fields(i).Name
    Next i
sndFolders:
'Debug.Print FieldsList & DB.TableDefs.Count

Dim sendFolders As MultiSck
Set sendFolders = New MultiSck
    With sendFolders
        .GetSck (UserName)
        .sendFolders (FieldsList)
    End With


'Set DB = Nothing


End Function
Public Sub CreateUserTable(UserName As String)
DB.Execute "CREATE TABLE [" & UserName & "] (MsgID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, From Text (15), Subject Memo , Discription Memo, RDate Text (20), [Sent Items] Memo, [Rubbish Bin] Memo );"
DB.Close
Call openDAO
  
End Sub

Public Sub ItemClick()                      'Read users details &
Dim rs As ADODB.Recordset                   'insert them to the text fields
Set rs = New ADODB.Recordset
Dim Sql As String
Dim i As Integer

Sql = "Select * from Users where UserName = " & Chr(34) & RegUser & Chr(34)
rs.Open Sql, cn, adOpenKeyset, adLockOptimistic

For i = 1 To rs.RecordCount

With FrmProfile.Text1
    .Item(0).Text = "" & rs!Company
    .Item(1).Text = "" & rs!UserName
    .Item(2).Text = "" & rs!Password
    .Item(3).Text = "" & rs!Phone
    .Item(4).Text = "" & rs!Email
    .Item(5).Text = "" & rs!Discription
End With


Next i

rs.Close
Set rs = Nothing

End Sub

Public Sub UpdateAcc()              'Update the users details
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim i As Integer
Dim Sql As String
    On Error GoTo UpdateError
Sql = "Select * from Users where UserName = " & Chr(34) & RegUser & Chr(34)

rs.Open Sql, cn, adOpenKeyset, adLockOptimistic

With FrmProfile
  If rs.EOF = True Then
        MsgBox "User doesn't exist in Database", vbExclamation
        UsrExist = False
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If Len(.Text1(0).Text) = 0 Then rs!Company = Null Else rs!Company = .Text1(0).Text
    If Len(.Text1(1).Text) = 0 Then rs!UserName = Null Else rs!UserName = .Text1(1).Text
    If Len(.Text1(2).Text) = 0 Then rs!Password = Null Else rs!Password = .Text1(2).Text
    If Len(.Text1(3).Text) = 0 Then rs!Phone = Null Else rs!Phone = .Text1(3).Text
    If Len(.Text1(4).Text) = 0 Then rs!Email = Null Else rs!Email = .Text1(4).Text
    If Len(.Text1(5).Text) = 0 Then rs!Discription = Null Else rs!Discription = .Text1(5).Text
    
    rs.Update
    rs.Close
    
    ClearForm
    Set rs = Nothing
Exit Sub

UpdateError:

rs.Close
Set rs = Nothing
Exit Sub
End With

End Sub
Public Sub NewUser(UserName As String)
Dim rs As ADODB.Recordset
Dim Sql As String
Set rs = New ADODB.Recordset
Sql = "Select * from [" & UserName & "]"
rs.Open Sql, cn, adOpenKeyset, adLockOptimistic

rs.AddNew

rs!From = "Office Messenger v1.1"
rs!Subject = "Welcome to Office Messenger"

rs!Discription = "Thank you for interest in Office Messenger" & vbNewLine & _
"For more information please visit www.chris.hatton.com"

rs!Rdate = Date
rs.Update
rs.Close
Set rs = Nothing


End Sub

Public Sub DelAcc()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim i As Integer
Dim DB As Database
Dim Sql As String
Dim result As String
Sql = "Select * from Users where UserName = " & Chr(34) & RegUser & Chr(34)

rs.Open Sql, cn, adOpenKeyset, adLockOptimistic

If rs.EOF = True Then
    MsgBox "Can't Delete User Account" & vbNewLine _
    & "User May Not Exist In Database", vbCritical + vbOKOnly
    Exit Sub
    Else
        result = MsgBox("Are You Sure?", vbQuestion + vbYesNo, "Delete Account?")
End If

If result = vbYes Then
    rs.Delete
    rs.Close
    Set rs = Nothing
    On Error Resume Next
    Set DB = OpenDatabase(App.Path & "\OSDB.mdb")
    DB.TableDefs.Delete RegUser
    DB.Close

    Exit Sub
Else

        MsgBox "User Cancelled", vbInformation, "Cancelled"
        rs.Close
        Set rs = Nothing
    End If
    
End Sub

Public Sub EmailAcc(UserName As String)               'Update the users Email Account
                                                      'Mail cfg page, saves info
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim i As Integer
Dim Sql As String
    On Error GoTo UpdateError
Sql = "Select * from Users where UserName = " & Chr(34) & UserName & Chr(34)

rs.Open Sql, cn, adOpenKeyset, adLockOptimistic

With FrmMailCFG
  If rs.EOF = True Then
        MsgBox "Error finding User in Database", vbExclamation
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If Len(.Text1(0).Text) = 0 Then rs!pop = Null Else rs!pop = .Text1(0).Text
    If Len(.Text1(1).Text) = 0 Then rs!smtp = Null Else rs!smtp = .Text1(1).Text
    
    rs.Update
    rs.Close
    
    
    Set rs = Nothing
Exit Sub

UpdateError:

rs.Close
Set rs = Nothing
MsgBox "Error Saving Mail Settings", vbExclamation
Exit Sub
End With

End Sub

Public Sub MailInfo(UserName As String)
Dim rs As ADODB.Recordset                   'insert them to the text fields
Set rs = New ADODB.Recordset
Dim Sql As String
Dim i As Integer

Sql = "Select * from Users where UserName = " & Chr(34) & UserName & Chr(34)
rs.Open Sql, cn, adOpenKeyset, adLockOptimistic



With FrmMailCFG.Text1
    .Item(0).Text = "" & rs!pop
    .Item(1).Text = "" & rs!smtp
End With




rs.Close
Set rs = Nothing


End Sub
Public Sub MailTOcheck(UserName As String)

Dim rs As ADODB.Recordset       'gets email addy
Set rs = New ADODB.Recordset
Dim itm As ListItem
Dim Sql  As String
Dim i As Integer
UserName = Split(UserName, Chr(10))(0)

Sql = "Select * from Users where UserName = " & Chr(34) & UserName & Chr(34)
rs.Open Sql, cn, adOpenKeyset, adLockOptimistic


UsrToEmail = "" & rs!Email
        

rs.Close


Set itm = Nothing
Set rs = Nothing
End Sub

Public Sub MailIDcheck(UserName As String)

Dim rs As ADODB.Recordset       'gets email addy
Set rs = New ADODB.Recordset
Dim itm As ListItem
Dim Sql  As String
Dim i As Integer
UserName = Split(UserName, Chr(10))(0)

Sql = "Select * from Users where UserName = " & Chr(34) & UserName & Chr(34)
rs.Open Sql, cn, adOpenKeyset, adLockOptimistic


usrEmail = "" & rs!Email
        

rs.Close


Set itm = Nothing
Set rs = Nothing
End Sub

Public Sub EmInbox(WhoTo, WhoFrom, Message, EmailDate, Subject As String)
Dim rs As ADODB.Recordset       'take record of outgoing emails
Set rs = New ADODB.Recordset
Dim itm As ListItem
Dim Sql As String
Dim i As Integer
WhoFrom = Split(WhoFrom, Chr(10))(0)

Sql = "Select * from Email "
rs.Open Sql, cn, adOpenKeyset, adLockOptimistic



rs.AddNew

rs!To = "" & WhoTo
rs!From = "" & WhoFrom
rs!Message = "" & Message
rs!EmDate = "" & EmailDate
rs!Subject = "" & Subject
        
rs.Update

rs.MoveLast
    Set itm = FrmServer.LVMsgs.ListItems.Add(, , rs!emailid, , 1)
        itm.SubItems(1) = WhoTo
        itm.SubItems(2) = WhoFrom

rs.Close



Set rs = Nothing


End Sub

Public Sub EmOutBox(emID As Long)
Dim rs As ADODB.Recordset       'take record of outgoing emails
Set rs = New ADODB.Recordset
Dim POPAddy, RecvMail, FromAddy, strFrom, ToAddy, Subject, Message As String
Dim Sql, Sql1 As String

Sql = "Select * from Email where emailid = " & emID
    rs.Open Sql, cn, adOpenKeyset, adLockOptimistic

        
        
        ToAddy = rs!To      'who the message is going to
        FromAddy = rs!From  'who its from
        Message = rs!Message    'The actual Message
        Subject = rs!Subject

    rs.Close

Set rs = Nothing

Set rs = New ADODB.Recordset

Sql1 = "Select * from Users where Email = " & Chr(34) & ToAddy & Chr(34)
        rs.Open Sql1, cn, adOpenKeyset, adLockOptimistic
            
            POPAddy = rs!pop    'get the pop3 server
            strFrom = rs!UserName   'person name on email
            RecvMail = rs!Email     'and there email address


        rs.Close

Set rs = Nothing

FrmServer.SendMail POPAddy, RecvMail, FromAddy, strFrom, ToAddy, Subject, Message


End Sub

Public Sub EmRemove(emID As Long)
Dim rs As ADODB.Recordset     'Delete the emails in the outbox.
Set rs = New ADODB.Recordset
Dim Sql As String

Sql = "Select * from email where emailid = " & emID
 
   rs.Open Sql, cn, adOpenKeyset, adLockOptimistic
   
   rs.Delete

rs.Close

Set rs = Nothing

End Sub
Public Sub LdEMail()
Dim rs As ADODB.Recordset     'loads any emails into the lvControl
Set rs = New ADODB.Recordset
Dim Sql As String
Dim itm As ListItem
Dim i As Integer

Sql = "Select * from email"
 
   rs.Open Sql, cn, adOpenKeyset, adLockOptimistic
   
   For i = 1 To rs.RecordCount
   
      
      
   Set itm = FrmServer.LVMsgs.ListItems.Add(, , rs!emailid, , 1)
        itm.SubItems(1) = rs!To
        itm.SubItems(2) = rs!From

   
   rs.MoveNext
   Next

rs.Close

Set rs = Nothing
End Sub

Public Sub GetUserInfo(UserName As String)
Dim rs As ADODB.Recordset     'get the users information
Set rs = New ADODB.Recordset
Dim Sql As String

    Sql = "Select * from Users where UserName = " & Chr(34) & UserName & Chr(34)
    rs.Open Sql, cn, adOpenKeyset, adLockOptimistic

        usrCom = "" & rs!Company
        usrName = "" & rs!UserName
        usrAddy = "" & rs!address
        usrAddy1 = "" & rs!address1
        usrPhone = "" & rs!Phone
        usrFax = "" & rs!Fax
        usrEmail = "" & rs!Email

    rs.Close

Set rs = Nothing

End Sub


Public Sub NewAccount(Name, Addy, Addy1, Country, Phone, Fax, Company, _
Email, Website, Password As String)     'add the user account to the database



Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Set rego = New Registration

Dim i As Integer
Dim Sql As String

Name = Split(Name, Chr(10))(1)
rs.Open "Select * from Users", cn, adOpenKeyset, adLockOptimistic

rego.RSck = FrmServer.ServerSck.UBound

For i = 1 To rs.RecordCount

If LCase(Name) = LCase(rs!UserName) Then
    
    
    rego.MSGUsrExst
    GoTo stopaccount
    
    Else
    rs.MoveNext
End If


Next i

If rs.EOF = True Then
    
    rs.AddNew
       If Not Len(Name) = 0 Then rs!UserName = Name
       If Not Len(Company) = 0 Then rs!Company = Company
       If Not Len(Password) = 0 Then rs!Password = Password
       If Not Len(Phone) = 0 Then rs!Phone = Phone
       If Not Len(Email) = 0 Then rs!Email = Email
       If Not Len(Fax) = 0 Then rs!Fax = Fax
       If Not Len(Addy) = 0 Then rs!address = Addy
       If Not Len(Addy1) = 0 Then rs!address1 = Addy1
       If Not Len(Website) = 0 Then rs!Website = Website
       If Not Len(Country) = 0 Then rs!Country = Country
     
    rs.Update
    rs.Close
    DoEvents
    CreateUserTable (Name)
    
stopaccount:
    Set rs = Nothing
    End If

End Sub
'updates the users email account remotely.
Public Sub SvEmailAcc(UserName, AcPOP, AcSmtp)            'Update the users Email Account
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim i As Integer
Dim Sql As String

Sql = "Select * from Users where UserName = " & Chr(34) & UserName & Chr(34)

rs.Open Sql, cn, adOpenKeyset, adLockOptimistic
    
    If Len(AcPOP) = 0 Then rs!pop = Null Else rs!pop = AcPOP
    If Len(AcSmtp) = 0 Then rs!smtp = Null Else rs!smtp = AcSmtp
    
    rs.Update
    rs.Close
    
    Set rs = Nothing
    



End Sub
