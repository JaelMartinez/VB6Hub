VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form LibrosMega 
   Caption         =   "LibrosMega"
   ClientHeight    =   9825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22335
   LinkTopic       =   "Form1"
   ScaleHeight     =   9825
   ScaleWidth      =   22335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuieroLeer 
      Caption         =   "Quiero Leer"
      Height          =   615
      Left            =   840
      TabIndex        =   5
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmdNoMeGusta 
      Caption         =   "No Me Gusta"
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdRecomendar 
      Caption         =   "Recomendar"
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdMarcarLeido 
      Caption         =   "Marcar como Leído"
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   9735
      Left            =   3360
      TabIndex        =   1
      Top             =   -120
      Width           =   18975
      _ExtentX        =   33470
      _ExtentY        =   17171
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "LibrosMega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMarcarLeido_Click()
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Por favor, selecciona un libro primero.", vbExclamation, "Error"
        Exit Sub
    End If

    Dim libroID As Integer
    libroID = GetLibroIDByTitle(ListView1.SelectedItem.Text)

    If libroID = -1 Then Exit Sub

    ' Verificar si ya está en la tabla de libros leídos
    Dim rsCheck As ADODB.Recordset
    Set rsCheck = New ADODB.Recordset
    rsCheck.Open "SELECT * FROM LibrosLeidos WHERE LibroID = " & libroID, conn, adOpenStatic, adLockReadOnly

    If Not rsCheck.EOF Then
        MsgBox "Este libro ya está marcado como leído.", vbInformation, "Información"
    Else
        ' Insertar en la tabla de libros leídos
        conn.Execute "INSERT INTO LibrosLeidos (LibroID, FechaLectura) VALUES (" & libroID & ", GETDATE())"
        MsgBox "Libro marcado como leído.", vbInformation, "Éxito"
    End If

    rsCheck.Close
    Set rsCheck = Nothing
End Sub

Private Sub cmdNoMeGusta_Click()
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Por favor, selecciona un libro primero.", vbExclamation, "Error"
        Exit Sub
    End If

    Dim libroID As Integer
    libroID = GetLibroIDByTitle(ListView1.SelectedItem.Text)

    If libroID = -1 Then Exit Sub

    ' Verificar si ya está en la tabla de libros que no te gustan
    Dim rsCheck As ADODB.Recordset
    Set rsCheck = New ADODB.Recordset
    rsCheck.Open "SELECT * FROM LibrosQueNoTeGustan WHERE LibroID = " & libroID, conn, adOpenStatic, adLockReadOnly

    If Not rsCheck.EOF Then
        MsgBox "Este libro ya está marcado como que no te gusta.", vbInformation, "Información"
    Else
        ' Insertar en la tabla de libros que no te gustan
        conn.Execute "INSERT INTO LibrosQueNoTeGustan (LibroID, Motivo) VALUES (" & libroID & ", 'No especificado')"
        MsgBox "Libro marcado como que no te gusta.", vbInformation, "Éxito"
    End If

    rsCheck.Close
    Set rsCheck = Nothing
End Sub

Private Sub cmdQuieroLeer_Click()
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Por favor, selecciona un libro primero.", vbExclamation, "Error"
        Exit Sub
    End If

    Dim libroID As Integer
    libroID = GetLibroIDByTitle(ListView1.SelectedItem.Text)

    If libroID = -1 Then Exit Sub

    ' Verificar si ya está en la tabla de libros que quieres leer
    Dim rsCheck As ADODB.Recordset
    Set rsCheck = New ADODB.Recordset
    rsCheck.Open "SELECT * FROM LibrosQueQuieresLeer WHERE LibroID = " & libroID, conn, adOpenStatic, adLockReadOnly

    If Not rsCheck.EOF Then
        MsgBox "Este libro ya está marcado como que quieres leer.", vbInformation, "Información"
    Else
        ' Insertar en la tabla de libros que quieres leer
        conn.Execute "INSERT INTO LibrosQueQuieresLeer (LibroID, Prioridad) VALUES (" & libroID & ", 1)"
        MsgBox "Libro marcado como que quieres leer.", vbInformation, "Éxito"
    End If

    rsCheck.Close
    Set rsCheck = Nothing
End Sub

Private Sub cmdRecomendar_Click()
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Por favor, selecciona un libro primero.", vbExclamation, "Error"
        Exit Sub
    End If

    Dim libroID As Integer
    libroID = GetLibroIDByTitle(ListView1.SelectedItem.Text)

    If libroID = -1 Then Exit Sub

    ' Verificar si ya está en la tabla de libros recomendados
    Dim rsCheck As ADODB.Recordset
    Set rsCheck = New ADODB.Recordset
    rsCheck.Open "SELECT * FROM LibrosRecomendados WHERE LibroID = " & libroID, conn, adOpenStatic, adLockReadOnly

    If Not rsCheck.EOF Then
        MsgBox "Este libro ya está marcado como recomendado.", vbInformation, "Información"
    Else
        ' Insertar en la tabla de libros recomendados
        conn.Execute "INSERT INTO LibrosRecomendados (LibroID, FechaRecomendacion) VALUES (" & libroID & ", GETDATE())"
        MsgBox "Libro marcado como recomendado.", vbInformation, "Éxito"
    End If

    rsCheck.Close
    Set rsCheck = Nothing
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
        .Font.Size = 10 ' Cambiar tamaño de la fuente
        
        ' Configurar encabezados y ancho de las columnas
        .ColumnHeaders.Add , , "Título", 4000 ' Ancho ajustado
        .ColumnHeaders.Add , , "Autor", 3000  ' Ancho ajustado
        .ColumnHeaders.Add , , "Género", 2000  ' Ancho ajustado
        .ColumnHeaders.Add , , "Calificación", 1500  ' Ancho ajustado
        .ColumnHeaders.Add , , "Sinopsis", 14000  ' Ancho ajustado
    End With

    ' Cargar los libros en el ListView
    Call LoadBooksIntoListView
End Sub


Private Sub LoadBooksIntoListView()
    ' Conectar a la base de datos y llenar el ListView con los libros
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM Libros", conn, adOpenStatic, adLockReadOnly

    Do While Not rs.EOF
        Dim itmX As ListItem
        Set itmX = ListView1.ListItems.Add(, , rs!titulo)
        itmX.SubItems(1) = rs!Autor
        itmX.SubItems(2) = rs!Genero
        itmX.SubItems(3) = rs!Calificacion
        
        ' Dividir la sinopsis en dos líneas si es muy larga
        Dim sinopsis As String
        sinopsis = rs!sinopsis
        
        If Len(sinopsis) > 100 Then ' Ajusta el número 100 según lo que se vea mejor
            itmX.SubItems(4) = Left(sinopsis, 100) & vbCrLf & Mid(sinopsis, 101)
        Else
            itmX.SubItems(4) = sinopsis
        End If

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub
Private Function GetLibroIDByTitle(titulo As String) As Integer
    Dim rsLibroID As ADODB.Recordset
    Set rsLibroID = New ADODB.Recordset
    rsLibroID.Open "SELECT LibroID FROM Libros WHERE Titulo = '" & Replace(titulo, "'", "''") & "'", conn, adOpenStatic, adLockReadOnly

    If rsLibroID.EOF Then
        MsgBox "No se encontró el LibroID para este libro.", vbExclamation, "Error"
        rsLibroID.Close
        Set rsLibroID = Nothing
        GetLibroIDByTitle = -1
    Else
        GetLibroIDByTitle = rsLibroID!libroID
        rsLibroID.Close
        Set rsLibroID = Nothing
    End If
End Function

