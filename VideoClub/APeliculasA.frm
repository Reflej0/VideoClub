VERSION 5.00
Begin VB.Form APeliculasA 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Alta de Peliculas"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   DrawStyle       =   1  'Dash
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6150
   ScaleWidth      =   9630
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   1920
      TabIndex        =   13
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dar el Alta"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   2040
      TabIndex        =   11
      Top             =   4800
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "FechaAlta:"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Numero de Copia:"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Actores:"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Genero:"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Titulo:"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Codigo Pelicula:"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "APeliculasA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset

Private Sub Command1_Click()
respu = MsgBox("Quiere dar de alta el registro actual?", vbYesNo + vbInformation, "Alta del registro")
If respu = vbYes Then
Dim CodPel As Integer
Dim Titulo As String
Dim Genero As String
Dim Actores As String
Dim NumCopia As Integer
Dim FechaAlta As Date
CodPel = Val(Text1.Text)
Titulo = Text2.Text
Genero = Text3.Text
Actores = Val(Text4.Text)
NumCopia = Val(Text6.Text)
FechaAlta = Text7.Text
db.Execute "insert into Peliculas(CodPel,Titulo,Genero,Actores,NumCopia,FechaAlta) values (" & CodPel & ",'" & Titulo & "','" & Genero & "','" & Actores & "'," & NumCopia & ",#" & FechaAlta & "#)"
MsgBox " Se logro dar de alta el registro recien creado", vbOKOnly + vbInformation, "Alta del registro"
Else
Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
Mystring = (80000000)
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("\\Direccion\Alumnos\Tocino\Biblioteca\biblioteca.mdb")
End Sub
