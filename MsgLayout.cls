VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MsgLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim UnRead As Long

Public Sub MsgLyt(sptmessage As String, recordset As String)
On Error Resume Next
Dim j As Long
Dim Counter As Long
Dim PerMessage(6)
Dim Statuse As Long
Dim itm As ListItem
Dim Message As String
With FrmMain.LvMail
FrmMain.MousePointer = 11
FrmMain.TVdir.Enabled = False
FrmDownload.MousePointer = 11
Statuse = UBound(Split(sptmessage, "~}~")) + 1
Call FrmMain.FrmStyle
FrmDownload.ProgressBar1.Max = Statuse
If Not GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "DisplayCont", "") = "True" Then .Visible = False
If Not GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "HideStatus", "") = "True" Then FrmDownload.Show vbOLEDisplayContent
    Counter = recordset - Statuse
    Statuse = Statuse - 1
    For j = 0 To UBound(Split(sptmessage, "~}~"))

Message = Split(sptmessage, "~}~")(j)
If Message = "" Then GoTo NextMessage

FrmDownload.Label2.Caption = "Recieving Message.." & j + 1 & " of " & Statuse
'FrmDownload.ProgressBar1.ToolTipText = "Recieving Message.." & j + 1 & " of " & Statuse
FrmDownload.ProgressBar1.Value = FrmDownload.ProgressBar1.Value + 1

If j - 3 = UBound(Split(sptmessage, "~}~")) - 4 Then
    FrmDownload.ProgressBar1.Max = j
    FrmDownload.ProgressBar1.Value = j
End If


PerMessage(0) = Split(Message, "~%~")(0) 'Folder
PerMessage(1) = Split(Message, "~*~")(1): PerMessage(1) = Split(PerMessage(1), "~!~")(0) 'ID
PerMessage(2) = Split(Message, "~!~")(1): PerMessage(2) = Split(PerMessage(2), "~#~")(0) 'From
PerMessage(3) = Split(Message, "~#~")(1): PerMessage(3) = Split(PerMessage(3), "~@~")(0) 'Subject
PerMessage(4) = Split(Message, "~@~")(1): PerMessage(4) = Split(PerMessage(4), "~^~")(0) 'Discription
PerMessage(5) = Split(Message, "~^~")(1): PerMessage(5) = Split(PerMessage(5), "ó")(0) 'Date
PerMessage(6) = Split(Message, "~'~")(1):
Debug.Print PerMessage(1) & " " & Message

If PerMessage(6) = "N" Then Set itm = .ListItems.Add(, , PerMessage(1), , 1) Else Set itm = .ListItems.Add(, , PerMessage(1), , 2)  'ID
        itm.SubItems(1) = PerMessage(2): FrmMain.Label3.Caption = PerMessage(2) 'From
        itm.SubItems(2) = PerMessage(3): FrmMain.Label4.Caption = PerMessage(3) 'Subject
        itm.SubItems(3) = PerMessage(5): FrmMain.Label6.Caption = PerMessage(5) 'Date
       
         'DoEvents
        'If FrmMain.TVdir.SelectedItem.Text = "Sent Items" Then FrmMain.LvMail.ColumnHeaders.Item(2).Text = "To" Else FrmMain.LvMail.ColumnHeaders.Item(2).Text = "From"
            FrmMain.strMessage = FrmMain.strMessage & PerMessage(1) & "/" & PerMessage(4) & "~!^~"
            FrmMain.RichTextBox1.Text = Split(PerMessage(4), "[~N10~]")(0)
            .ListItems.Item(4).Selected = True
            DoEvents
            If PerMessage(6) = "N" Then
                UnRead = UnRead + 1
                itm.ListSubItems.Item(j - 1).Bold = True
                itm.ListSubItems.Item(j).Bold = True
                itm.ListSubItems.Item(1).Bold = True
                itm.ListSubItems.Item(2).Bold = True
                itm.ListSubItems.Item(3).Bold = True
                FrmMain.StatusBar1.Panels.Item(3).ToolTipText = "New Message"
                DoEvents
                                
            End If
            
            
NextMessage:
Next j

If Not FrmMain.ColumnSet = "Set" Then EnhListView_SortColumns FrmMain.LvMail, 1, False
    FrmMain.ColumnSet = "Set"
    FrmMain.StatusBar1.Panels(3).Text = FrmMain.LvMail.ListItems.Count & " Items, " & UnRead & " Unread"
    Statusbar = 0
    'FrmDownload.ProgressBar1.Value = 0
    Call LetMsgStore(0)
    FrmMain.MousePointer = 0
    FrmDownload.MousePointer = 0
    FrmMain.Label4 = ""
    FrmMain.Label3 = ""
    FrmMain.Label6 = ""
    FrmMain.RichTextBox1 = ""
    FrmMain.TVdir.Enabled = True
    Unload FrmDownload
   '.Refresh
   
   If GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "HidePreview", "") = "False" Then
    FrmMain.HiddenPreview = False
    FrmMain.Frame1.Visible = True
    FrmMain.RichTextBox1.Visible = True
    FrmMain.HSplit.Visible = True
    FrmMain.LvMail.Height = FrmMain.HSplit.Top - 1410
    FrmMain.Shape1.Height = FrmMain.HSplit.Top - 1410
   End If
   
   If FrmMain.LvMail.ListItems.Count = 0 Then FrmMain.Label4 = "": FrmMain.Label3 = "": FrmMain.Label6 = "": FrmMain.RichTextBox1 = ""
   If Not GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "DisplayCont", "") = "True" Then .Visible = True
   FrmMain.LvMail.ListItems.Item(1).Selected = True
   FrmMain.LvMail.SetFocus
   FrmMain.LvMail_Click
End With
End Sub




Public Sub SingleMessage(strWho, StrSub, StrMsg, TDate As String, Counter As Long)
Dim itm As ListItem
Dim X As Long
With FrmMain.LvMail
FrmMain.strMessage = FrmMain.strMessage & Counter & "/" & StrMsg & "~!^~"

Set itm = .ListItems.Add(, , Counter, , 1)
        itm.SubItems(1) = strWho 'From
        itm.SubItems(2) = StrSub 'Subject
        itm.SubItems(3) = TDate & " " & Format(Now, "short Time")  'Date
        FrmMain.RichTextBox1.Text = StrMsg
    
    For X = 1 To .ListItems.Count
         If .ListItems.Item(X).Text = Counter Then
                
                itm.ListSubItems.Item(1).Bold = True
                itm.ListSubItems.Item(2).Bold = True
                itm.ListSubItems.Item(3).Bold = True
         End If
    Next X
                FrmMain.StatusBar1.Panels.Item(3).ToolTipText = "New Message"

                Call FrmMain.AddIcon
                .Refresh
        

End With
End Sub
