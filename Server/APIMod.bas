Attribute VB_Name = "APIMod"
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
Public Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Public Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2


Public Declare Function InternetAutodial Lib "wininet.dll" _
    (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long


Public Declare Function InternetAutodialHangup Lib "wininet.dll" _
    (ByVal dwReserved As Long) As Long

    Public Const ModConnect As Long = &H1

