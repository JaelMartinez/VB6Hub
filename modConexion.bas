Attribute VB_Name = "modConexion"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public conn As ADODB.Connection

Public Sub ConnectToDatabase()
    Dim iniFilePath As String
    iniFilePath = App.Path & "\config.ini"
    
    Dim connectionString As String
    connectionString = GetIniValue("Database", "ConnectionString", iniFilePath)
    
    Set conn = New ADODB.Connection
    conn.connectionString = connectionString
    conn.Open
End Sub

Private Function GetIniValue(section As String, key As String, iniFilePath As String) As String
    Dim retVal As String * 255
    Dim strLen As Long
    
    strLen = GetPrivateProfileString(section, key, "", retVal, 255, iniFilePath)
    GetIniValue = Left(retVal, strLen)
End Function


Public Sub CloseDatabaseConnection()
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            conn.Close
        End If
        Set conn = Nothing
    End If
End Sub

