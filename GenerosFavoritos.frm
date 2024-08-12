VERSION 5.00
Begin VB.Form GenerosFavoritos 
   Caption         =   "GenerosFavoritos"
   ClientHeight    =   8100
   ClientLeft      =   4815
   ClientTop       =   3705
   ClientWidth     =   19230
   LinkTopic       =   "Form2"
   ScaleHeight     =   8100
   ScaleWidth      =   19230
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11520
      TabIndex        =   3
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   2
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label Texto 
      BackStyle       =   0  'Transparent
      Caption         =   "Estos son tus generos favoritos basado en tus libros leidos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   1
      Top             =   1320
      Width           =   14175
   End
End
Attribute VB_Name = "GenerosFavoritos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVolver_Click()
    frmHubDeLectura.Show
    Me.Hide
End Sub
Private Sub Form_Activate()
    Call ConnectToDatabase

    ' Consulta para obtener los dos géneros más leídos
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "SELECT TOP 2 L.Genero, COUNT(L.LibroID) AS TotalLeidos FROM LibrosLeidos LL INNER JOIN Libros L ON LL.LibroID = L.LibroID GROUP BY L.Genero ORDER BY TotalLeidos DESC;", conn, adOpenStatic, adLockReadOnly

    ' Verificar si hay resultados
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If Not rs.EOF Then
            Label1.Caption = "1ero: " & rs!Genero
            rs.MoveNext
        End If
        
        If Not rs.EOF Then
            Label2.Caption = "2do: " & rs!Genero
        End If
    Else
        MsgBox "No se encontraron géneros favoritos.", vbInformation, "Información"
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Form_Load()
    Call ConnectToDatabase

    ' Consulta para obtener los dos géneros más leídos
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "SELECT TOP 2 L.Genero, COUNT(L.LibroID) AS TotalLeidos FROM LibrosLeidos LL INNER JOIN Libros L ON LL.LibroID = L.LibroID GROUP BY L.Genero ORDER BY TotalLeidos DESC;", conn, adOpenStatic, adLockReadOnly

    ' Verificar si hay resultados
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If Not rs.EOF Then
            Label1.Caption = "Primer Género Favorito: " & rs!Genero
            rs.MoveNext
        End If
        
        If Not rs.EOF Then
            Label2.Caption = "Segundo Género Favorito: " & rs!Genero
        End If
    Else
        MsgBox "No se encontraron géneros favoritos.", vbInformation, "Información"
    End If

    rs.Close
    Set rs = Nothing
End Sub


