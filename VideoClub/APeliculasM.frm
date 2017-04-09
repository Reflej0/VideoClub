VERSION 5.00
Begin VB.Form APeliculasM 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Modificaciones de Peliculas"
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4560
      MaskColor       =   &H00C0E0FF&
      Picture         =   "APeliculasM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      MaskColor       =   &H00C0E0FF&
      Picture         =   "APeliculasM.frx":1190
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3840
      MaskColor       =   &H00C0E0FF&
      Picture         =   "APeliculasM.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5880
      MaskColor       =   &H00C0E0FF&
      Picture         =   "APeliculasM.frx":2A28
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar la modificacion"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Actores:"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Genero:"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Titulo:"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Codigo Pelicula:"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "APeliculasM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Private Sub Command1_Click()
rs.Edit
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text4.Text
rs.Update
MsgBox "Registro modificado con exito", vbOKOnly
End Sub

Private Sub Command2_Click()
Unload Me
Mystring = (80000000)
End Sub

Private Sub Command3_Click()
rs.MoveNext
If rs.EOF Then
    rs.MoveFirst
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
Dim CodPel As Integer
Dim Titulo As String
Dim Genero As String
Dim Actores As String
CodPel = Val(Text1.Text)
Titulo = Text2.Text
Genero = Text3.Text
Actores = Val(Text4.Text)
Set db = OpenDatabase("\\Direccion\Alumnos\Tocino\Biblioteca\biblioteca.mdb")
Set rs = db.OpenRecordset("select * from Peliculas")
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text1.Enabled = False
End Sub
Public Sub muestropeliculas()
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
End Sub
