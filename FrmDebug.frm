VERSION 5.00
Begin VB.Form FrmDebug 
   Caption         =   "Debug"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   2955
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChQuitarRetardos 
      Caption         =   "Quitar Retardos"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CheckBox ChSaltarPergamino 
      Caption         =   "Saltar Pergamino"
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox ChDesconectarDimensionesAmpliadas 
      Caption         =   "Desconec.Dimens.Apliadas"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Frame FrPersonajes 
      Caption         =   "Personajes"
      Height          =   2295
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1335
      Begin VB.CommandButton BtPersonajesNinguno 
         Caption         =   "Ninguno"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton BtPersonajesTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox ChPersonajesSeverino 
         Caption         =   "Severino"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox ChPersonajesBerengario 
         Caption         =   "Berengario"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox ChPersonajesAbad 
         Caption         =   "Abad"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox ChPersonajesMalaquias 
         Caption         =   "Malaquías"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox ChPersonajesAdso 
         Caption         =   "Adso"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
   End
   Begin VB.Frame FrLuz 
      Caption         =   "Luz"
      Height          =   1695
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.CheckBox ChLampara 
         Caption         =   "Lámpara"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox ChLuzGuillermo 
         Caption         =   "En Guillermo"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton OpLuzOff 
         Caption         =   "Todo OFF"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OpLuzON 
         Caption         =   "Todo ON"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton OpLuzNormal 
         Caption         =   "Normal"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CargandoDatos As Boolean
Private Sub BtPersonajesNinguno_Click()
    ChPersonajesAdso.Value = 0
    ChPersonajesMalaquias.Value = 0
    ChPersonajesAbad.Value = 0
    ChPersonajesBerengario.Value = 0
    ChPersonajesSeverino.Value = 0
End Sub

Private Sub BtPersonajesTodos_Click()
    ChPersonajesAdso.Value = 1
    ChPersonajesMalaquias.Value = 1
    ChPersonajesAbad.Value = 1
    ChPersonajesBerengario.Value = 1
    ChPersonajesSeverino.Value = 1
End Sub

Private Sub ChDesconectarDimensionesAmpliadas_Click()
    If CargandoDatos Then Exit Sub
    If ChDesconectarDimensionesAmpliadas.Value = 1 Then
        Depuracion.DeshabilitarCalculoDimensionesAmpliadas = True
    Else
        Depuracion.DeshabilitarCalculoDimensionesAmpliadas = False
    End If
End Sub

Private Sub ChLampara_Click()
    If CargandoDatos Then Exit Sub
    If ChLampara.Value = 1 Then
        Depuracion.Lampara = True
    Else
        Depuracion.Lampara = False
    End If
End Sub

Private Sub ChLuzGuillermo_Click()
    If CargandoDatos Then Exit Sub
    If ChLuzGuillermo.Value = 1 Then
        Depuracion.LuzEnGuillermo = True
    Else
        Depuracion.LuzEnGuillermo = False
    End If
End Sub

Private Sub ChQuitarRetardos_Click()
    If CargandoDatos Then Exit Sub
    If ChQuitarRetardos.Value = 1 Then
        Depuracion.QuitarRetardos = True
    Else
        Depuracion.QuitarRetardos = False
    End If

End Sub

Private Sub ChSaltarPergamino_Click()
    If CargandoDatos Then Exit Sub
    If ChSaltarPergamino.Value = 1 Then
        Depuracion.SaltarPergamino = True
    Else
        Depuracion.SaltarPergamino = False
    End If
End Sub

Private Sub OpLuzNormal_Click()
    If CargandoDatos Then Exit Sub
    modAbadia.Depuracion.Luz = EnumTipoLuz_Normal
End Sub

Private Sub OpLuzOff_Click()
    If CargandoDatos Then Exit Sub
    modAbadia.Depuracion.Luz = EnumTipoLuz_Off
End Sub

Private Sub OpLuzON_Click()
    If CargandoDatos Then Exit Sub
    modAbadia.Depuracion.Luz = EnumTipoLuz_ON
End Sub

Private Sub ChPersonajesAbad_Click()
    If CargandoDatos Then Exit Sub
    If ChPersonajesAbad.Value = 1 Then
        Depuracion.PersonajesAbad = True
    Else
        Depuracion.PersonajesAbad = False
    End If
End Sub

Private Sub ChPersonajesAdso_Click()
    If CargandoDatos Then Exit Sub
    If ChPersonajesAdso.Value = 1 Then
        Depuracion.PersonajesAdso = True
    Else
        Depuracion.PersonajesAdso = False
    End If
End Sub

Private Sub ChPersonajesBerengario_Click()
    If CargandoDatos Then Exit Sub
    If ChPersonajesBerengario.Value = 1 Then
        Depuracion.PersonajesBerengario = True
    Else
        Depuracion.PersonajesBerengario = False
    End If
End Sub

Private Sub ChPersonajesMalaquias_Click()
    If CargandoDatos Then Exit Sub
    If ChPersonajesMalaquias.Value = 1 Then
        Depuracion.PersonajesMalaquias = True
    Else
        Depuracion.PersonajesMalaquias = False
    End If
End Sub

Private Sub ChPersonajesSeverino_Click()
    If CargandoDatos Then Exit Sub
    If ChPersonajesSeverino.Value = 1 Then
        Depuracion.PersonajesSeverino = True
    Else
        Depuracion.PersonajesSeverino = False
    End If
End Sub

Private Sub Form_Load()
    CargandoDatos = True
    With Depuracion
            ChPersonajesAdso.Value = -CInt(Depuracion.PersonajesAdso)
            ChPersonajesMalaquias.Value = -CInt(Depuracion.PersonajesMalaquias)
            ChPersonajesAbad.Value = -CInt(Depuracion.PersonajesAbad)
            ChPersonajesBerengario.Value = -CInt(Depuracion.PersonajesBerengario)
            ChPersonajesSeverino.Value = -CInt(Depuracion.PersonajesSeverino)
            Select Case .Luz
                Case Is = EnumTipoLuz_Normal
                    OpLuzNormal.Value = True
                Case Is = EnumTipoLuz_ON
                    OpLuzON = True
                Case Is = EnumTipoLuz_Off
                    OpLuzOff = True
            End Select
            ChLampara.Value = -CInt(.Lampara)
            ChLuzGuillermo.Value = -CInt(.LuzEnGuillermo)
            ChDesconectarDimensionesAmpliadas.Value = -CInt(.DeshabilitarCalculoDimensionesAmpliadas)
            ChSaltarPergamino.Value = -CInt(.SaltarPergamino)
            ChQuitarRetardos.Value = -CInt(.QuitarRetardos)
    End With
    
    CargandoDatos = False
End Sub
