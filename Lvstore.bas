Attribute VB_Name = "Lvstore"

Public SendMessageID As String 'view the message field
Public MessageID As Long


Public Sub LetMsgStore(letID As Long)
'If Not FrmMain.Selected = FrmMain.TVdir.SelectedItem.Text Then FrmMain.strMessage = FrmMain.strMessage & "~!^~"
FrmMain.MousePointer = 11
On Error Resume Next
Dim i As Integer
Dim SortMessage As String
 
    For i = 0 To UBound(Split(FrmMain.strMessage, "~!^~"))
         'DoEvents
         SortMessage = Split(FrmMain.strMessage, "~!^~")(i)
         DoEvents
         MessageID = Split(SortMessage, "/")(0)
           'DoEvents
           If letID = MessageID Then
                SendMessageID = Split(SortMessage, "/")(1)
                FrmMain.RichTextBox1.Text = "" & Split(SendMessageID, "[~N10~]")(0)
          
           Else
           MessageID = letID
            End If
    Next i

FrmMain.MousePointer = 0


End Sub
Public Sub GetMsgStore(ID As Long)
  Call LetMsgStore(ID)
End Sub

