VERSION 5.00
Begin VB.MDIForm Menu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Menu Principal"
   ClientHeight    =   6315
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "MDIForm1"
   Picture         =   "Menu.frx":0000
   StartUpPosition =   1  'CenterOwner
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
         Caption         =   "&Alquileres"
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
      Begin VB.Menu MnuLiSocios 
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
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MnuAcerca_Click()
Acercade.Show
End Sub

Private Sub MnuArPeAltas_Click()
APeliculasA.Show
End Sub

Private Sub MnuArPeBajas_Click()
APeliculasB.Show
End Sub

Private Sub MnuArPeConsultas_Click()
APeliculasC.Show
End Sub

Private Sub MnuArPeModificaciones_Click()
APeliculasM.Show
End Sub

Private Sub MnuArSalir_Click()
Unload Me
End Sub

Private Sub MnuArSetup_Click()

End Sub

Private Sub MnuArSoAltas_Click()
ASociosA.Show
End Sub

Private Sub MnuArSoBajas_Click()
ASociosB.Show
End Sub

Private Sub MnuArSoConsultas_Click()
AsociosC.Show
End Sub

Private Sub MnuArSoModificaciones_Click()
AsociosM.Show
End Sub

Private Sub MnuCoAlPeli_Click()
CAlquileresP.Show
End Sub

Private Sub MnuCoDia_Click()
CCajaD.Show
End Sub

Private Sub MnuCoMes_Click()
CCajaM.Show
End Sub

Private Sub MnuCoSoAlq_Click()
CSociosP.Show
End Sub

Private Sub MnuCoSoSocios_Click()
CSociosS.Show
End Sub

Private Sub MnuLiCaMovFec_Click()
LCajaD.Show
End Sub

Private Sub MnuLiCaMovMes_Click()
LCajaM.Show
End Sub

Private Sub MnuLiPeliculas_Click()
LPeliculas.Show
End Sub

Private Sub MnuLiSocios_Click()
Lsocios.Show
End Sub

Private Sub MnuOpAlquiler_Click()
OAlquiler.Show
End Sub

Private Sub MnuOpCobros_Click()
OCajaC.Show
End Sub

Private Sub MnuOpDevoluciones_Click()
ODevoluciones.Show
End Sub

Private Sub MnuOpPagos_Click()
OCajaP.Show
End Sub
