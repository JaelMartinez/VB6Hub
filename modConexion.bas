Attribute VB_Name = "modConexion"
Public conn As ADODB.Connection

Public Sub ConnectToDatabase()
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=DESKTOP-S6JV37M\MSSQLSERVER01;Initial Catalog=HubDeLecturaDB;User ID=Mega;Password=Mega123;"
    conn.Open
End Sub

