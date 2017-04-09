VERSION 5.00
Begin VB.Form ASociosA 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Alta de Socios"
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
   ScaleMode       =   0  'User
   ScaleWidth      =   9630
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   495
      Left            =   3120
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   17
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dar La Alta"
      Height          =   495
      Left            =   3120
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   16
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   6360
      TabIndex        =   15
      Top             =   5400
      Width           =   3135
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   14
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   13
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   12
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   11
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "FecAlta:"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Celular"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Localidad:"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Direccion:"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "F Nac:"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "DNI:"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Nombre:"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Apellido: "
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "ASociosA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Private Sub Command1_Click()
respu = MsgBox("Quiere dar de alta el registro actual?", vbYesNo + vbInformation, "Alta del registro")
If respu = vbYes Then
Dim Apellido As String
Dim Nombre As String
Dim DNI As String
Dim FNac As Date
Dim Direccion As String
Dim Localidad As String
Dim Celular As String
Dim FecAlta As Date
Apellido = Text1.Text
Nombre = Text2.Text
DNI = Text3.Text
FNac = Text4.Text
Direccion = Text5.Text
Localidad = Text6.Text
Celular = Text7.Text
FecAlta = Text8.Text
db.Execute "insert into Socios(Apellido,Nombre,DNI,FNac,Direccion,Localidad,Celular,FecAlta) values ('" & Apellido & "','" & Nombre & "','" & DNI & "',#" & FNac & "#,'" & Direccion & "','" & Localidad & "'," & Celular & ",#" & FecAlta & "#)"
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
