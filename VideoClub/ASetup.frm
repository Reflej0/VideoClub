VERSION 5.00
Begin VB.Form ASetup 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Setup"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   5880
      TabIndex        =   15
      Top             =   5520
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar Cambios"
      Height          =   495
      Left            =   5880
      TabIndex        =   14
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Dias:"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Precio Recargo:"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Precio Alquiler:"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "                     NUEVOS VALORES"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3720
      Width           =   8775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "                    VALORES ACTUALES"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Dias:"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Precio Recargo:"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Precio Alquiler:"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "Asetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label6_Click()

End Sub
