VERSION 5.00
Begin VB.Form AsociosM 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Modificaciones de Socios"
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
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00C0E0FF&
      Picture         =   "AsociosM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      MaskColor       =   &H00C0E0FF&
      Picture         =   "AsociosM.frx":0700
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1560
      MaskColor       =   &H00C0E0FF&
      Picture         =   "AsociosM.frx":0E00
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      MaskColor       =   &H00C0E0FF&
      Picture         =   "AsociosM.frx":1F98
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   495
      Left            =   3120
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   15
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar la Modificacion"
      Height          =   495
      Left            =   3120
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   14
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   13
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   12
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   11
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Celular"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Localidad:"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Direccion:"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "F Nac:"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "DNI:"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Nombre:"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Apellido: "
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "AsociosM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset
Dim db As Database
Private Sub Command1_Click()
rs.Edit
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Text5.Text
rs.Fields(5) = Text6.Text
rs.Fields(6) = Text7.Text
rs.Update
MsgBox "Registro modificado con exito", vbOKOnly
End Sub

Private Sub Command2_Click()
Unload Me
Mystring = (80000000)
End Sub

Private Sub Command3_Click()
rs.MovePrevious
If rs.BOF Then
    rs.MoveLast
End If
muestrosocios
End Sub

Private Sub Command4_Click()
rs.MoveNext
If rs.EOF Then
    rs.MoveFirst
End If
muestrosocios
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("\\Direccion\Alumnos\Tocino\Biblioteca\biblioteca.mdb")
Set rs = db.OpenRecordset("select * from Socios")
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(5)
Text7.Text = rs.Fields(6)
End Sub
Public Sub muestrosocios()
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(5)
Text7.Text = rs.Fields(6)
Text8.Text = rs.Fields(7)
End Sub
