VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form LibrosQueQuieresLeer 
   Caption         =   "LibrrosQueQuieresLeer"
   ClientHeight    =   8055
   ClientLeft      =   5220
   ClientTop       =   3105
   ClientWidth     =   19230
   LinkTopic       =   "Form4"
   ScaleHeight     =   8055
   ScaleWidth      =   19230
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7695
      Left            =   4200
      TabIndex        =   1
      Top             =   240
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
End
Attribute VB_Name = "LibrosQueQuieresLeer"
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
    respuesta = MsgBox("¿Está seguro de que desea eliminar el libro '" & TituloLibro & "' de la lista de libros que quieres leer?", vbYesNo + vbQuestion, "Confirmar Eliminación")
    
    If respuesta = vbYes Then
        ' Eliminar el libro de la base de datos
        Dim sql As String
        sql = "DELETE FROM LibrosQueQuieresLeer WHERE LibroID IN (SELECT LibroID FROM Libros WHERE Titulo = '" & TituloLibro & "')"
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

    ' Cargar los libros que quieres leer en el ListView
    Call LoadQuieresLeerIntoListView
End Sub

Private Sub LoadQuieresLeerIntoListView()
    ' Conectar a la base de datos y llenar el ListView con los libros que quieres leer
    Set rs = New ADODB.Recordset
    rs.Open "SELECT L.Titulo, L.Autor, L.Genero, L.Calificacion, L.Sinopsis FROM Libros L INNER JOIN LibrosQueQuieresLeer LQQL ON L.LibroID = LQQL.LibroID", conn, adOpenStatic, adLockReadOnly

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

