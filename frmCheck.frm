VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCheck 
   Caption         =   "Check"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   10170
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChMostrarSoloErrores 
      Caption         =   "Solo Errores"
      Height          =   255
      Left            =   7200
      TabIndex        =   16
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.TextBox TxRutaModelos 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Text            =   "D:\datos\proyectos\16_Abadía\ModelosCheck"
      Top             =   1680
      Width           =   4575
   End
   Begin VB.TextBox TxArchivoModelo 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Text            =   "D:\datos\proyectos\16_Abadía\ModelosCheck\Modificado4.dsk"
      Top             =   1080
      Width           =   4575
   End
   Begin VB.CommandButton BtGenerarModelos 
      Caption         =   "Generar Modelos .dsk"
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox TxRutaArchivoPosiciones 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "D:\datos\proyectos\16_Abadía\Vbasic\Abadia3\Posiciones.txt"
      Top             =   360
      Width           =   4575
   End
   Begin VB.CommandButton BtComparar 
      Caption         =   "Comparar"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton BtGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      ToolTipText     =   "Crear los archivos de prueba"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton BtDespiezarVolcados 
      Caption         =   "Despiezar Volcados .bin"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      ToolTipText     =   "Separa los volcados en archivos independientes"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox TxRutaCheck 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "D:\datos\proyectos\16_Abadía\Vbasic\Abadia3\Check"
      Top             =   3120
      Width           =   4575
   End
   Begin VB.TextBox TxRutaVolcados 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "D:\datos\proyectos\16_Abadía\Volcados"
      Top             =   2400
      Width           =   4575
   End
   Begin MSForms.TextBox TxInforme 
      Height          =   6375
      Left            =   240
      TabIndex        =   15
      Top             =   3600
      Width           =   8295
      VariousPropertyBits=   -1400879077
      ScrollBars      =   2
      Size            =   "14631;11245"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label6 
      Caption         =   "Ruta de volcados"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Archivo Modelo"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Archivo de posiciones"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   $"frmCheck.frx":0000
      Height          =   1335
      Left            =   5160
      TabIndex        =   4
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Ruta de archivos de prueba"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Ruta de archivos modelo"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Check As New cCheckEnvironment

Private Sub BtComparar_Click()
    Check.CompararArchivosCheck TxRutaVolcados.Text, TxRutaCheck.Text, ChMostrarSoloErrores.Value
End Sub

Private Sub BtDespiezarVolcados_Click()
    Dim Result As Long
    Dim Log As String
    Result = Check.Init("D:\datos\proyectos\16_Abadía\Volcados", "D:\datos\proyectos\16_Abadía\Vbasic\Abadia3\Check", Log)
    TxInforme.Text = Log
End Sub

Private Sub BtGenerar_Click()
    Dim Log As String
    Check.GenerarTablasCheck TxRutaArchivoPosiciones.Text, TxRutaCheck.Text, Log
    TxInforme.Text = Log
End Sub


Private Sub BtGenerarModelos_Click()
    Dim Log As String
    Check.GenerarModelos TxRutaArchivoPosiciones.Text, TxRutaModelos.Text, TxArchivoModelo.Text, Log
    TxInforme.Text = Log
End Sub


Private Sub TxInforme_DblClick(Cancel As MSForms.ReturnBoolean)
    TxInforme.Text = ""
End Sub
