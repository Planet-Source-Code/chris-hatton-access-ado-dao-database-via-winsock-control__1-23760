Attribute VB_Name = "CompressData"
Public Declare Function GetTempPath Lib "kernel32" Alias _
    "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer _
    As String) As Long

Public Const MAX_PATH = 260
Public User As String

 
Public Sub CompactJetDatabase(Location As String, _
    Optional BackupOriginal As Boolean = True)
If User = "" Then User = "ADMIN"

On Error GoTo CompactErr
  cn.Close
  DB.Close

Dim strBackupFile As String
Dim strTempFile As String

'Check the database exists
If Len(Dir(Location)) Then

    ' If a backup is required, do it!
    If BackupOriginal = True Then
        strBackupFile = GetTemporaryPath & "backup.mdb"
        If Len(Dir(strBackupFile)) Then Kill strBackupFile
        FileCopy Location, strBackupFile
    End If

    ' Create temporary filename
    strTempFile = GetTemporaryPath & "temp.mdb"
    If Len(Dir(strTempFile)) Then Kill strTempFile

    ' Do the compacting via DBEngine
    DBEngine.CompactDatabase Location, strTempFile

    ' Remove the original database file
    Kill Location

    ' Copy the temporary now-compressed
    ' database file back to the original
    ' location
    FileCopy strTempFile, Location

    ' Delete the temporary file
    Kill strTempFile

Else

End If
ADOConnect
openDAO

FrmServer.Label6.Caption = "Database Compressed" & vbCrLf & Format(Now, "long date") & vbCrLf & _
"By " & User & "  " & Format(Now, "long time")
FrmServer.Image1.Visible = True
User = ""
    Exit Sub
CompactErr:

  
    Exit Sub

End Sub

Public Function GetTemporaryPath()

Dim strFolder As String
Dim lngResult As Long

strFolder = String(MAX_PATH, 0)
lngResult = GetTempPath(MAX_PATH, strFolder)

If lngResult <> 0 Then
  GetTemporaryPath = Left(strFolder, InStr(strFolder, _
    Chr(0)) - 1)
Else
  GetTemporaryPath = ""
End If

End Function

