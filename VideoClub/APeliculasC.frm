VERSION 5.00
Begin VB.Form APeliculasC 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Consulta de Peliculas"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   9630
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   0
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   3375
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5640
      MaskColor       =   &H00C0E0FF&
      Picture         =   "APeliculasC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      MaskColor       =   &H00C0E0FF&
      Picture         =   "APeliculasC.frx":0700
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4920
      MaskColor       =   &H00C0E0FF&
      Picture         =   "APeliculasC.frx":0E00
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      MaskColor       =   &H00C0E0FF&
      Picture         =   "APeliculasC.frx":1F98
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Codigo Pelicula:"
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Titulo:"
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Genero:"
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Actores:"
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Numero de Copia:"
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "FechaAlta:"
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   4680
      Width           =   1455
   End
End
Attribute VB_Name = "APeliculasC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset
Dim db As Database

Private Sub Command2_Click()
Unload Me
Mystring = (80000000)
End Sub

Private Sub Command3_Click()
rs.MoveLast
If rs.BOF Then
    rs.MoveLast
End If
muestropeliculas
End Sub

Private Sub Command4_Click()
rs.MoveNext
If rs.EOF Then
    rs.MoveFirst
End If
muestropeliculas
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("\\Direccion\Alumnos\Tocino\Biblioteca\biblioteca.mdb")
Set rs = db.OpenRecordset("select * from Peliculas")
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text7.Text = rs.Fields(6)
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
End Sub
Public Sub muestropeliculas()
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text6.Text = rs.Fields(5)
Text7.Text = rs.Fields(6)
End Sub

