VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form LibrosLeidos 
   Caption         =   "LibrosLeidos"
   ClientHeight    =   8025
   ClientLeft      =   4815
   ClientTop       =   3510
   ClientWidth     =   19200
   LinkTopic       =   "Form5"
   ScaleHeight     =   8025
   ScaleWidth      =   19200
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7695
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   13573
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "LibrosLeidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEliminar_Click()
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Por favor, seleccione un libro para eliminar.", vbExclamation, "Eliminar Libro"
        Exit Sub
    End If

    Dim TituloLibro As String
    TituloLibro = ListView1.SelectedItem.Text

    ' Confirmar eliminación
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Está seguro de que desea eliminar el libro '" & TituloLibro & "' de la lista de libros leídos?", vbYesNo + vbQuestion, "Confirmar Eliminación")
    
    If respuesta = vbYes Then
        ' Eliminar el libro de la base de datos
        Dim sql As String
        sql = "DELETE FROM LibrosLeidos WHERE LibroID IN (SELECT LibroID FROM Libros WHERE Titulo = '" & TituloLibro & "')"
        conn.Execute sql

        ' Eliminar el libro del ListView
        ListView1.ListItems.Remove ListView1.SelectedItem.Index

        MsgBox "Libro eliminado correctamente.", vbInformation, "Eliminar Libro"
    End If
End Sub


Private Sub cmdVolver_Click()
    frmHubDeLectura.Show
    Me.Hide
End Sub

Private Sub Form_Load()
    Call ConnectToDatabase

    ' Configurar ListView
    With ListView1
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .Font.Size = 10
        
        ' Configurar encabezados y ancho de las columnas
        .ColumnHeaders.Add , , "Título", 4000
        .ColumnHeaders.Add , , "Autor", 3000
        .ColumnHeaders.Add , , "Género", 2000
        .ColumnHeaders.Add , , "Calificación", 1500
        .ColumnHeaders.Add , , "Sinopsis", 14000
    End With

    ' Cargar los libros leídos en el ListView
    Call LoadLeidosIntoListView
End Sub

Private Sub LoadLeidosIntoListView()
    ' Conectar a la base de datos y llenar el ListView con los libros leídos
    Set rs = New ADODB.Recordset
    rs.Open "SELECT L.Titulo, L.Autor, L.Genero, L.Calificacion, L.Sinopsis FROM Libros L INNER JOIN LibrosLeidos LL ON L.LibroID = LL.LibroID", conn, adOpenStatic, adLockReadOnly

    Do While Not rs.EOF
        Dim itmX As ListItem
        Set itmX = ListView1.ListItems.Add(, , rs!titulo)
        itmX.SubItems(1) = rs!Autor
        itmX.SubItems(2) = rs!Genero
        itmX.SubItems(3) = rs!Calificacion
        
        ' Dividir la sinopsis en dos líneas si es muy larga
        Dim sinopsis As String
        sinopsis = rs!sinopsis
        
        If Len(sinopsis) > 100 Then
            itmX.SubItems(4) = Left(sinopsis, 100) & vbCrLf & Mid(sinopsis, 101)
        Else
            itmX.SubItems(4) = sinopsis
        End If

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

