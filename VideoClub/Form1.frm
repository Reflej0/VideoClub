VERSION 5.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   5115
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MnuArchivos 
      Caption         =   "&Archivos"
      Begin VB.Menu MnuArSocios 
         Caption         =   "&Socios"
         Begin VB.Menu MnuArSoAltas 
            Caption         =   "&Altas"
         End
         Begin VB.Menu MnuArSoBajas 
            Caption         =   "&Bajas"
         End
         Begin VB.Menu MnuArSoModificaciones 
            Caption         =   "&Modificaciones"
         End
         Begin VB.Menu MnuArSoConsultas 
            Caption         =   "&Consultas"
         End
      End
      Begin VB.Menu MnuArPeliculas 
         Caption         =   "&Peliculas"
         Begin VB.Menu MnuArPeAltas 
            Caption         =   "&Altas"
         End
         Begin VB.Menu MnuArPeBajas 
            Caption         =   "&Bajas"
         End
         Begin VB.Menu MnuArPeModificaciones 
            Caption         =   "&Modificaciones"
         End
         Begin VB.Menu MnuArPeConsultas 
            Caption         =   "&Consultas"
         End
      End
      Begin VB.Menu MnuArSetup 
         Caption         =   "&Setup"
      End
      Begin VB.Menu MnuArSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu MnuOperaciones 
      Caption         =   "&Operaciones"
      Begin VB.Menu MnuOpAlquiler 
         Caption         =   "&Alquiler"
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuOpDevoluciones 
         Caption         =   "&Devoluciones"
         Shortcut        =   ^D
      End
      Begin VB.Menu MnuOpCaja 
         Caption         =   "&Caja"
         Begin VB.Menu MnuOpCobros 
            Caption         =   "&Cobros"
            Shortcut        =   {F2}
         End
         Begin VB.Menu MnuOpPagos 
            Caption         =   "&Pagos"
            Shortcut        =   {F3}
         End
      End
   End
   Begin VB.Menu MnuConsultas 
      Caption         =   "&Consultas"
      Begin VB.Menu MnuCoCaja 
         Caption         =   "&Caja"
         Begin VB.Menu MnuCoDia 
            Caption         =   "&Dia"
         End
         Begin VB.Menu MnuCoMes 
            Caption         =   "&Mes"
         End
      End
      Begin VB.Menu MnuCoAlquileres 
         Caption         =   "&Alquilares"
         Begin VB.Menu MnuCoAlPeli 
            Caption         =   "&Peliculas Alquiladas"
         End
      End
      Begin VB.Menu MnuCoSocios 
         Caption         =   "&Socios"
         Begin VB.Menu MnuCoSoAlq 
            Caption         =   "&Peliculas Alquiladas por Socio"
         End
         Begin VB.Menu MnuCoSoSocios 
            Caption         =   "&Socios Actualmente con Peliculas"
         End
      End
   End
   Begin VB.Menu MnuListados 
      Caption         =   "&Listados"
      Begin VB.Menu MnuSocios 
         Caption         =   "&Socios"
      End
      Begin VB.Menu MnuLiPeliculas 
         Caption         =   "&Peliculas"
      End
      Begin VB.Menu MnuLiCaja 
         Caption         =   "&Caja"
         Begin VB.Menu MnuLiCaMovFec 
            Caption         =   "&Movimientos por Fecha"
         End
         Begin VB.Menu MnuLiCaMovMes 
            Caption         =   "&Movimientos por Mes"
         End
      End
   End
   Begin VB.Menu MnuAcerca 
      Caption         =   "&Acerca de"
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

End Sub

Private Sub MnuAcerca_Click()
Acercade.Show
End Sub

Private Sub MnuArSalir_Click()
Unload Me
End Sub

Private Sub MnuArSoAltas_Click()
ASociosA.Show
End Sub
