VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Form1"
   ClientHeight    =   6915
   ClientLeft      =   5400
   ClientTop       =   150
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   461
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   Begin VB.TextBox TxEscaleras 
      Height          =   285
      Left            =   4920
      TabIndex        =   26
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton BtCheck 
      Caption         =   "Check"
      Height          =   375
      Left            =   1320
      TabIndex        =   25
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton BtPruebas2 
      Caption         =   "Pruebas2"
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   2160
      TabIndex        =   23
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton BtParar 
      Caption         =   "Stop"
      Height          =   255
      Left            =   2160
      TabIndex        =   22
      Top             =   6360
      Width           =   495
   End
   Begin VB.CommandButton BtGuardarPosicion 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   5520
      TabIndex        =   21
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton BtDebug 
      Appearance      =   0  'Flat
      Caption         =   "Debug"
      Height          =   375
      Left            =   3600
      TabIndex        =   20
      Top             =   6360
      Width           =   855
   End
   Begin VB.CheckBox ChPararPaso 
      Caption         =   "Parar al avanzar"
      Height          =   255
      Left            =   6360
      TabIndex        =   19
      Top             =   6600
      Width           =   1815
   End
   Begin VB.TextBox TxOrientacion 
      Height          =   285
      Left            =   2520
      TabIndex        =   18
      Top             =   6000
      Width           =   375
   End
   Begin VB.TextBox TxZ 
      Height          =   285
      Left            =   4320
      TabIndex        =   17
      Top             =   6000
      Width           =   495
   End
   Begin VB.TextBox TxY 
      Height          =   285
      Left            =   3720
      TabIndex        =   16
      Top             =   6000
      Width           =   495
   End
   Begin VB.TextBox TxX 
      Height          =   285
      Left            =   3120
      TabIndex        =   15
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton BtDerecha 
      Caption         =   "R"
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   6120
      Width           =   375
   End
   Begin VB.CommandButton BtIzquierda 
      Caption         =   "L"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   6120
      Width           =   375
   End
   Begin VB.CommandButton BtArriba 
      Caption         =   "Up"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   6120
      Width           =   375
   End
   Begin VB.TextBox TxProfundidad 
      Height          =   285
      Left            =   8280
      TabIndex        =   11
      Text            =   "&Hff"
      Top             =   6240
      Width           =   495
   End
   Begin VB.TextBox TxNbloque 
      Height          =   285
      Left            =   7680
      TabIndex        =   9
      Text            =   "10"
      Top             =   6240
      Width           =   495
   End
   Begin VB.TextBox TxDeltaY 
      Height          =   285
      Left            =   7080
      TabIndex        =   7
      Text            =   "0"
      Top             =   6240
      Width           =   495
   End
   Begin VB.TextBox TxDeltaX 
      Height          =   285
      Left            =   6360
      TabIndex        =   4
      Text            =   "0"
      Top             =   6240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eva"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   6480
      Width           =   495
   End
   Begin VB.TextBox TxNumeroHabitacion 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Text            =   "&H24"
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton BtPruebas 
      Caption         =   "Pruebas"
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   6480
      Width           =   975
   End
   Begin VB.PictureBox pbPantalla 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6000
      Left            =   0
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   0
      Width           =   9600
      Begin VB.Timer TmTemporizador 
         Enabled         =   0   'False
         Left            =   9000
         Top             =   0
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Prof."
      Height          =   255
      Left            =   8280
      TabIndex        =   10
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Nbloque"
      Height          =   255
      Left            =   7680
      TabIndex        =   8
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "DeltaY"
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "DeltaX"
      Height          =   255
      Left            =   6360
      TabIndex        =   5
      Top             =   6000
      Width           =   615
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FrmDebugVisible As Boolean





Private Sub BtArriba_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    modTeclado.KeyDown TeclaArriba
End Sub

Private Sub BtArriba_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    modTeclado.KeyUp TeclaArriba
End Sub

Private Sub BtCheck_Click()
    frmCheck.Show
End Sub

Private Sub BtDebug_Click()
    FrmDebug.Show
End Sub

Private Sub BtDerecha_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    modTeclado.KeyDown TeclaDerecha
End Sub

Private Sub BtDerecha_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    modTeclado.KeyUp TeclaDerecha
End Sub

Private Sub BtGuardarPosicion_Click()
    Open "Posiciones.txt" For Append As #1
    Print #1, TxNumeroHabitacion.Text + "," + TxOrientacion.Text + "," + TxX.Text + "," + TxY.Text + "," + TxZ.Text + "," + TxEscaleras.Text
    Close #1
End Sub

Private Sub BtIzquierda_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    modTeclado.KeyDown TeclaIzquierda
End Sub

Private Sub BtIzquierda_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    modTeclado.KeyUp TeclaIzquierda
End Sub

Private Sub BtParar_Click()
    Dim Contador As Long
    modAbadia.PararAbadia
    'While modAbadia.Parado = False
    '    DoEvents
    '    Contador = Contador + 1
    'Wend
    'MsgBox CStr(Contador)
End Sub

Private Sub BtPruebas_Click()
Dim NumeroHabitacion As Long
'On Error Resume Next
Pintar = True
NumeroHabitacion = CInt(TxNumeroHabitacion.Text)
modPantalla.DibujarRectangulo 0, 0, 319, 160, &H7D7D00
PunteroPantallaActual_156A = BuscarHabitacionProvisional(NumeroHabitacion)
HabitacionOscura_156C = False
DibujarPantalla_19D8
'NumeroHabitacion = NumeroHabitacion + 2
'TxNumeroHabitacion.Text = "&H" + Hex$(NumeroHabitacion)
End Sub

Private Sub BtPruebas2_Click()
    Dim archivo1(2) As Byte
    Dim archivo2(2) As Byte
    archivo1(1) = 1
    archivo2(1) = 2
    modFunciones.CompararArchivos archivo1, archivo2
End Sub

Private Sub Command1_Click()
modPantalla.DibujarRectangulo 0, 0, 319, 199, &H7D7D00
Pintar = True
LimpiarRejilla_1A70 0 'limpia la rejilla y rellena un rectángulo de 256x160 a partir de (32, 0) con el color de fondo


'GenerarBloqueSuelto CInt(frmPrincipal.TxNbloque.Text), 15, 17, CInt(frmPrincipal.TxDeltaX.Text), CInt(frmPrincipal.TxDeltaY.Text), &HFF
HabitacionOscura_156C = False
DibujarPantalla_4EB2
modPantalla.Refrescar
End Sub

Private Sub Command2_Click()
    InicializarJuego_249A
End Sub

Private Sub Form_Activate()
Static Activado As Boolean
If Not Activado Then
    Activado = True
    InicializarJuego_249A
End If
modPantalla.SeleccionarPaleta 2

End Sub





Private Sub Form_Load()
Dim nose As Long
Dim Direccion As Long
'BtParar_Click
InicializarPantalla 2, pbPantalla

'Exit Sub

'DibujarPunto 160, 100, vbBlue

'PantallaCGA2PC &H0&, &H35

'PantallaCGA2PC &H800&, &H35


'BtPruebas_Click
'PunteroPantallaActual_156A = BuscarHabitacionProvisional(&H0)
'HabitacionIluminada_156C = True
'DibujarPantalla_19D8
End Sub




Public Sub Retardo(Tiempo As Long)
'hace una pausa de la duración indicada en "tiempo" (ms)
Dim Contador As Integer
TmTemporizador.Interval = Tiempo
TmTemporizador.Enabled = True
Do While TmTemporizador.Enabled = True
    Contador = Contador + 1
    If Contador = 10 Then
        DoEvents
        Contador = 0
    End If
Loop
End Sub



Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub pbPantalla_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 32
        modTeclado.KeyDown TeclaEspacio
    Case Is = 37
        modTeclado.KeyDown TeclaIzquierda
    Case Is = 38
        modTeclado.KeyDown TeclaArriba
    Case Is = 39
        modTeclado.KeyDown TeclaDerecha
    Case Is = 40
        modTeclado.KeyDown TeclaAbajo
End Select
End Sub




Private Sub TmTemporizador_Timer()
TmTemporizador.Enabled = False
End Sub

Private Sub TxNumeroHabitacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then BtPruebas_Click
End Sub
