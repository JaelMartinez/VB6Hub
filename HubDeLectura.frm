VERSION 5.00
Begin VB.Form frmHubDeLectura 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Hub de lectura"
   ClientHeight    =   8190
   ClientLeft      =   4200
   ClientTop       =   3915
   ClientWidth     =   19305
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   19305
   Begin VB.CommandButton cmdVerLibrosMega 
      BackColor       =   &H00FFFF00&
      Caption         =   "Ver libros"
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdVerLibrosLeidos 
      Caption         =   "Ver libros"
      Height          =   615
      Left            =   3960
      TabIndex        =   4
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton cmdVerLibrosQueQuieresLeer 
      Caption         =   "Ver libros"
      Height          =   615
      Left            =   9960
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdVerLibrosQueNoTeGustan 
      Caption         =   "Ver libros"
      Height          =   615
      Left            =   9960
      TabIndex        =   2
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton cmdVerGenerosFavoritos 
      Caption         =   "Ver generos"
      Height          =   615
      Left            =   15840
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdVerLibrosRecomendados 
      Caption         =   "Ver libros"
      Height          =   615
      Left            =   15840
      TabIndex        =   0
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Image imgtogglemode 
      Height          =   615
      Left            =   18240
      Top             =   240
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   1440
      Top             =   600
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   2655
      Left            =   1440
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Image Image3 
      Height          =   2655
      Left            =   7440
      Top             =   720
      Width           =   1815
   End
   Begin VB.Image Image4 
      Height          =   2655
      Left            =   7440
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Image Image5 
      Height          =   2655
      Left            =   13320
      Top             =   720
      Width           =   1815
   End
   Begin VB.Image Image6 
      Height          =   2655
      Left            =   13320
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Catalogo de libros de Mega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   11
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Libros que ya leiste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   10
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Libros que quieres leer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9720
      TabIndex        =   9
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Libros que no te gustaron"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9720
      TabIndex        =   8
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Generos favoritos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15600
      TabIndex        =   7
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Libros recomendados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15480
      TabIndex        =   6
      Top             =   4800
      Width           =   1935
   End
End
Attribute VB_Name = "frmHubDeLectura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmHubDeLectura.frm
Dim rs As ADODB.Recordset
Dim isDarkMode As Boolean

Private Sub cmdVerGenerosFavoritos_Click()
    GenerosFavoritos.Show
    Me.Hide
End Sub

Private Sub cmdVerLibrosLeidos_Click()
    LibrosLeidos.Show
    Me.Hide
End Sub

Private Sub cmdVerLibrosMega_Click()
    LibrosMega.Show
    Me.Hide
End Sub

Private Sub cmdVerLibrosQueNoTeGustan_Click()
    LibrosQueNoTeGustan.Show
    Me.Hide
End Sub

Private Sub cmdVerLibrosQueQuieresLeer_Click()
    LibrosQueQuieresLeer.Show
    Me.Hide
End Sub

Private Sub cmdVerLibrosRecomendados_Click()
    LibrosRecomendados.Show
    Me.Hide
End Sub

Private Sub Form_Load()
    ' Conectar a la base de datos utilizando el módulo
    Call ConnectToDatabase
    
    If conn.State = adStateOpen Then
        MsgBox "Conexión exitosa a la base de datos.", vbInformation, "Conexión"
    Else
        MsgBox "Error al conectar a la base de datos.", vbCritical, "Error"
    End If
    
    ' Resto de la inicialización del formulario
    Image1.Stretch = True
    Image2.Stretch = True
    Image3.Stretch = True
    Image4.Stretch = True
    Image5.Stretch = True
    Image6.Stretch = True

    ' Cargar imágenes en los controles Image
    Image1.Picture = LoadPicture("C:\Users\Pobrito\Desktop\MegaHubLibros\Imagenes\2.jpg")
    Image2.Picture = LoadPicture("C:\Users\Pobrito\Desktop\MegaHubLibros\Imagenes\2.jpg")
    Image3.Picture = LoadPicture("C:\Users\Pobrito\Desktop\MegaHubLibros\Imagenes\2.jpg")
    Image4.Picture = LoadPicture("C:\Users\Pobrito\Desktop\MegaHubLibros\Imagenes\2.jpg")
    Image5.Picture = LoadPicture("C:\Users\Pobrito\Desktop\MegaHubLibros\Imagenes\2.jpg")
    Image6.Picture = LoadPicture("C:\Users\Pobrito\Desktop\MegaHubLibros\Imagenes\2.jpg")

    ' Inicializar el modo a claro
    isDarkMode = False
    imgtogglemode.Stretch = True
    imgtogglemode.Picture = LoadPicture("C:\Users\Pobrito\Desktop\MegaHubLibros\Imagenes\sol.jpg")
End Sub

Private Sub imgToggleMode_Click()
    ' Alternar entre los modos
    isDarkMode = Not isDarkMode
    
    Dim newBackColor As Long
    Dim newForeColor As Long
    Dim togglePicture As String

    If isDarkMode Then
        newBackColor = RGB(43, 43, 43)
        newForeColor = RGB(255, 255, 255)
        togglePicture = "C:\Users\Pobrito\Desktop\MegaHubLibros\Imagenes\luna.jpg"
    Else
        newBackColor = RGB(255, 255, 255)
        newForeColor = RGB(0, 0, 0)
        togglePicture = "C:\Users\Pobrito\Desktop\MegaHubLibros\Imagenes\sol.jpg"
    End If
    
    imgtogglemode.Picture = LoadPicture(togglePicture)
    
    Load frmHubDeLectura
    Load GenerosFavoritos
    Load LibrosLeidos
    Load LibrosMega
    Load LibrosQueNoTeGustan
    Load LibrosQueQuieresLeer
    Load LibrosRecomendados
    
    Dim frm As Form
    For Each frm In Forms
        frm.BackColor = newBackColor
        
        Dim ctrl As Control
        For Each ctrl In frm.Controls
            If TypeOf ctrl Is Label Or TypeOf ctrl Is TextBox Or TypeOf ctrl Is CommandButton Then
                ctrl.BackColor = newBackColor
                If TypeOf ctrl Is Label Then
                    ctrl.ForeColor = newForeColor
                End If
            End If
        Next ctrl
    Next frm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseDatabaseConnection
End Sub

