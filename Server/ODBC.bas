Attribute VB_Name = "ODBC"
Public cn As ADODB.Connection
Public DBConnect As Boolean
Public MSDatabase
Public Function ADOConnect() As Boolean

On Error GoTo OpenErr




Set cn = New ADODB.Connection

MSDatabase = App.Path & "\" & "OSDB.mdb"
    cn.CursorLocation = adUseClient
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    FrmStart.Label1.Caption = "Connecting to Database"
    cn.Open MSDatabase, Admin
    
    DBConnect = True
    
Exit Function

OpenErr:

    MsgBox "Error Opening " & MSDatabase & vbNewLine & Err.Description, vbCritical, "Open Database Error"
    DBConnect = False


End Function
