Attribute VB_Name = "ModUserList"
Dim TvMsgs As Node

Sub tvtree(List As String)
Dim i As Integer
Dim Adduser As Variant
Dim TvOnline As Node
    List = Split(List, Chr(10))(1)
    
    
    With FrmMain.TVdir
        .Nodes.Clear
        Set TvMsgs = .Nodes.Add(, tvwFirst, , "[Personal Folders]", 22)
        Set TvMsgs = .Nodes.Add(TvMsgs, tvwChild, , "Inbox", 2)
            .Nodes.Item(2).Bold = True
            .Nodes.Item(1).Expanded = True
            .Nodes(2).Selected = True
    End With
     
    Debug.Print List
     With FrmMain.TVcontact
        .Nodes.Clear
        Set TvOnline = .Nodes.Add(, tvwFirst, , "[Online Users]", 20)
        For i = 0 To UBound(Split(List, "_")) - 1
            Adduser = Split(List, "_")
                    If i = 0 Then Set TvOnline = .Nodes.Add(TvOnline, tvwChild, , Adduser(i), 15)
                    If Not i = 0 Then Set TvOnline = .Nodes.Add(TvOnline, tvwNext, , Adduser(i), 15)
        Next i
                    Set TvOnline = .Nodes.Add(, tvwChild, , "")
            .Nodes.Item(1).Expanded = True
                 
                
     End With
        
End Sub
Public Sub TvOffline(List As String)
Dim i, j As Integer
Dim TvOffline As Node
Dim Adduser As String
Debug.Print List
List = Split(List, "OfflineList")(1)
With FrmMain.TVcontact
    Set TvOffline = .Nodes.Add(, tvwFirst, , "[All Users]", 20)

For i = 0 To UBound(Split(List, "_"))
    FrmMain.AllUsersList = List
    Adduser = Split(List, "_")(i)
    If Not Adduser = "" Then Set TvOffline = .Nodes.Add _
    (TvOffline, tvwNext, , Adduser)
  
Next i
 End With
 
 With FrmConnect            'get our custom folders.
    .Usersock.SendData "CustomFolders" & Chr(10) & .strUserName
     Statusbar = 1
 End With

End Sub

Public Sub AddFolder(Folder As String)
Dim newFolder As Node
On Error Resume Next
With FrmMain.TVdir
    If Folder = "Rubbish Bin" Then
        Set TvFolder = .Nodes.Add(TvMsgs, tvwNext, , Folder, 4)
            GoTo nextI
    End If
        
        Set newFolder = .Nodes.Add(TvMsgs, tvwNext, , Folder, 2)
            
            .Nodes.Item(1).Expanded = True
        'FrmFolder.Text1.Enabled = True
        'FrmFolder.cmdOK.Enabled = True
        FrmMain.MousePointer = 0
       'FrmFolder.Hide
    End With
nextI:

End Sub



