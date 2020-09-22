Attribute VB_Name = "cnDAO"
Public DB As Database
Public Table As TableDef
Public FL As DAO.Field
Public Sub openDAO()
Set DB = OpenDatabase(App.Path & "\OSDB.mdb")
FrmStart.Label1.Caption = "Connected to Database"
End Sub
