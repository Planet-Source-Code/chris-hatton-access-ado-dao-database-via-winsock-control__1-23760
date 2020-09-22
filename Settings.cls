VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PwSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim strName As String       'UserName string
Dim strPass As String       'Password string
Public SvPassword As Boolean   'if the save password is true or false
Public Property Get UserName() As String
strName = GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "UserName", "")
UserName = strName           'Get the current password saved in Registry
End Property
Public Property Let UserName(ByVal vNewValue As String)
SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "UserName", vNewValue
                            'Saves the current user in the connection form.
   End Property
Public Property Get Password() As String
strPass = GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "Password", "")
If SvPassword = True Then   'if the save password box is true then allow user
    Password = strPass      'to see ***** password on screen.
    End If
End Property

Public Property Let Password(ByVal vNewValue As String)
SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "Password", vNewValue
strPass = vNewValue         'saves the new password to the registry
End Property

Public Property Get SavePass() As Boolean
If GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "SavePassword", "") = "" Then
    SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "SavePassword", "False"
    SvPassword = False
    SavePass = SvPassword   'if theres no password then don't save password.
    
    Exit Sub
        Else
        
    SvPassword = GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "SavePassword", "")
    SavePass = SvPassword   'user wants to save password save it to "true" in the registry
End If

End Property

Public Property Let SavePass(ByVal vNewValue As Boolean)
If vNewValue = True Then
    Password = strPass  'if save password is ticked, then save password.
        SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "SavePassword", "True"
        SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "Password", strPass

    Else                'if not save password is ticked then delete password.
        SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "SavePassword", "False"
        DeleteValue HKEY_CURRENT_USER, "OfficeMessenger", "Password"
End If
End Property

Public Property Get ServerIP() As String
If GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "ServerIP", "") = "" Then
   SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "ServerIP", FrmConnect.Usersock.LocalIP
   ServerIP = FrmConnect.Usersock.LocalIP   'if serverip = "" then insert local ip
Else

ServerIP = GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "ServerIP", "")
End If                                      'Display ServerIP to connection form

End Property

Public Property Let ServerIP(ByVal vNewValue As String)
SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "ServerIP", vNewValue
End Property                                'Save serverIP in the registry
Public Property Get Autocon() As Boolean
If GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "AutoLogon", "") = "" Then
    SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "AutoLogon", "False"
End If
    Autocon = GetRegKey(HKEY_CURRENT_USER, "OfficeMessenger", "AutoLogon", "")
End Property

Public Property Let Autocon(ByVal vNewValue As Boolean)
If vNewValue = True Then
    SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "AutoLogon", "True"
Else
    SaveRegKey HKEY_CURRENT_USER, "OfficeMessenger", "AutoLogon", "False"
End If


End Property
