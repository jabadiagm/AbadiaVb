VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCheckEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'funciones para generar los archivos de comprobación entre la abadía real y la simulada

Dim ListaTipos As New cListaTipoTabla

Public Function Init(DumpFolder As String, CheckFolder As String, Log As String) As Long
    'comprueba que la rutas son correctas y que están todas las tablas correspondientes cada volcado
    'si falta alguna tabla, la genera
    Dim ArchivosDump As cStringList
    Dim Contador As Long
    Dim Contador2 As Long
    Dim NombreVolcado As String
    Dim PrefijoTabla As String
    Dim NombreTabla As String
    Dim TipoTabla As cTipoTabla
    CargarTiposTabla
    Set ArchivosDump = modFunciones.DirFolder(DumpFolder, "*.bin")
    Log = "Encontrados " + CStr(ArchivosDump.Lenght) + " archivos .bin" + vbCrLf
    For Contador = 0 To ArchivosDump.Lenght - 1
        NombreVolcado = ArchivosDump.ElementAt(Contador)
        PrefijoTabla = Left$(NombreVolcado, Len(NombreVolcado) - 3)
        For Contador2 = 0 To ListaTipos.Lenght - 1
            Set TipoTabla = ListaTipos.ElementAt(Contador2)
            NombreTabla = PrefijoTabla + TipoTabla.Extension
            If Dir(NombreTabla) = "" Then 'hay que generar la tabla
                Log = Log + Dir(NombreVolcado) + ": Generando " + TipoTabla.Descripcion + vbCrLf
                GenerarTabla NombreVolcado, NombreTabla, TipoTabla
            End If
        Next
    Next
    Log = Log + "Archivos de volcado procesados" + vbCrLf + "-------------------------------------------" + vbCrLf
    
End Function

Private Sub CargarTiposTabla()
    Dim Tipo As New cTipoTabla
    Tipo.Init "ALT", &H1C0, &H240, "Buffer de Alturas"
    ListaTipos.Append Tipo
    Set Tipo = New cTipoTabla
    Tipo.Init "TIL", &H8D80&, &H780, "Buffer de Tiles"
    ListaTipos.Append Tipo
    Set Tipo = New cTipoTabla
    Tipo.Init "TSP", &H2E17&, &H1CD, "Tabla de Sprites"
    ListaTipos.Append Tipo
    Set Tipo = New cTipoTabla
    Tipo.Init "PUE", &H2FE4&, &H24, "Tabla de Puertas"
    ListaTipos.Append Tipo
    Set Tipo = New cTipoTabla
    Tipo.Init "OBJ", &H3008&, &H2E, "Posición de  Objetos"
    ListaTipos.Append Tipo
    Set Tipo = New cTipoTabla
    Tipo.Init "PER", &H3036&, &H5A, "Características de Personajes"
    ListaTipos.Append Tipo
    Set Tipo = New cTipoTabla
    Tipo.Init "ANI", &H319F&, &H60, "Animación de Personajes"
    ListaTipos.Append Tipo
    Set Tipo = New cTipoTabla
    Tipo.Init "BSP", &H9500&, &H780, "Buffer de Sprites"
    ListaTipos.Append Tipo
    Set Tipo = New cTipoTabla
    Tipo.Init "GRA", &HA300&, &H859, "Gráficos de Guillermo, Adso y Puertas"
    ListaTipos.Append Tipo
    Set Tipo = New cTipoTabla
    Tipo.Init "MON", &HAB59&, &H8A7, "Gráficos de Monjes"
    ListaTipos.Append Tipo
    Set Tipo = New cTipoTabla
    Tipo.Init "CGA", &HC000&, &H4000, "CGA"
    ListaTipos.Append Tipo
End Sub

Private Function GenerarTabla(RutaVolcado As String, RutaTabla As String, TipoTabla As cTipoTabla) As Long
    Dim Volcado() As Byte
    Dim Tabla() As Byte
    ReDim Tabla(TipoTabla.Tamaño - 1)
    modFunciones.CargarArchivo RutaVolcado, Volcado
    CargarTablaArchivo Volcado, Tabla, TipoTabla.DireccionInicio
    GuardarArchivo RutaTabla, Tabla
End Function

Public Function GenerarTablasCheckPantalla(Posicion As cPosicion, RutaCheck As String) As Long
    With Posicion
        modAbadia.CheckDefinir .NumeroPantalla, .Orientacion, .X, .Y, .Z, .Escaleras, RutaCheck
    End With
    modAbadia.InicializarJuego_249A
    SavePicture frmPrincipal.pbPantalla.Image, modFunciones.FixPath(RutaCheck) + Posicion.NumeroPantallaHex + ".bmp"
End Function

Public Sub CompararArchivosCheck(RutaVolcados As String, RutaCheck As String, MostrarSoloErrores As Boolean)
    Dim ArchivosVolcado As New cStringList
    Dim ArchivosPrueba As New cStringList
    Dim Contador As Byte
    Dim Contador2 As Long
    Dim Index As Long
    Dim Retorno As Long
    Dim RutaArchivo1 As String
    Dim RutaArchivo2 As String
    Dim Log As String
    Set ArchivosPrueba = modFunciones.DirFolder(RutaCheck, "", True)
    For Contador = 0 To 254
        Set ArchivosVolcado = modFunciones.DirFolder(RutaVolcados, Byte2AsciiHex(Contador) + ".*", True)
        For Contador2 = 0 To ArchivosVolcado.Lenght - 1
            Index = ArchivosPrueba.Index(ArchivosVolcado.ElementAt(Contador2), True)
            If Index <> -1 Then
                RutaArchivo1 = modFunciones.FixPath(RutaVolcados) + ArchivosVolcado.ElementAt(Contador2)
                RutaArchivo2 = modFunciones.FixPath(RutaCheck) + ArchivosPrueba.ElementAt(Index)
                Retorno = modFunciones.CompararArchivosRuta(RutaArchivo1, RutaArchivo2, Log)
                If Not MostrarSoloErrores Or MostrarSoloErrores And Retorno <> 0 Then AñadirTextoInforme Log
            End If
        Next
    Next
    AñadirTextoInforme "Fin" + vbCrLf
    
End Sub

Public Sub GenerarTablasCheck(RutaArchivoPosiciones As String, RutaCheck As String, Log As String)
    Dim Result As Long
    Dim ArchivoPosiciones As String
    Dim Contador As Long
    Dim Posicion As cPosicion
    Dim ListaPosiciones As New cListaPosiciones
    ArchivoPosiciones = RutaArchivoPosiciones
    Set ListaPosiciones = LeerArchivoPosiciones(RutaArchivoPosiciones)
    For Contador = 0 To ListaPosiciones.Lenght - 1
        Set Posicion = ListaPosiciones.ElementAt(Contador)
        Log = Log + "Generando tablas para la pantalla &H" + Posicion.NumeroPantallaHex + " ->"
        Result = GenerarTablasCheckPantalla(Posicion, RutaCheck)
        If Result = 0 Then
            Log = Log + "Tablas Generadas" + vbCrLf
        Else
            Log = Log + "Error" + vbCrLf
        End If
        DoEvents
    Next
    AñadirTextoInforme "Terminado" + vbCrLf
End Sub

Private Sub AñadirTextoInforme(Texto As String)
    With frmCheck.TxInforme
        .Text = .Text + Texto
        .SelStart = Len(.Text)
        .SelLength = 0
    End With
End Sub

Public Sub GenerarModelos(RutaArchivoPosiciones As String, RutaModelos As String, RutaModelo As String, Log As String)
    'genera los archivos DSK con las posiciones de inicio indicadas
    Dim ArchivoMaestro() As Byte
    Dim ListaPosiciones As New cListaPosiciones
    Dim Posicion As cPosicion
    Dim Contador As Long
    Set ListaPosiciones = LeerArchivoPosiciones(RutaArchivoPosiciones)
    modFunciones.CargarArchivo RutaModelo, ArchivoMaestro
    For Contador = 0 To ListaPosiciones.Lenght - 1
        Set Posicion = ListaPosiciones.ElementAt(Contador)
        Log = Log + "Generando modelo para la pantalla &H" + Posicion.NumeroPantallaHex + vbCrLf
        GenerarModelo ArchivoMaestro, Posicion, modFunciones.FixPath(RutaModelos) + Posicion.NumeroPantallaHex + ".dsk"
    Next
    
End Sub


Private Sub GenerarModelo(ByRef ArchivoMaestro() As Byte, Posicion As cPosicion, ByVal ArchivoSalida As String)
    'parchea el archivo maestro colocando la ubicación indicada y guardando el contenido
    With Posicion
        ArchivoMaestro(&HF1C4&) = .Escaleras
        ArchivoMaestro(&HF1C5&) = .Z
        ArchivoMaestro(&HF1C6&) = .Y
        ArchivoMaestro(&HF1C7&) = .X
        ArchivoMaestro(&HF1C8&) = .Orientacion
    End With
    modFunciones.GuardarArchivo ArchivoSalida, ArchivoMaestro
End Sub

Private Function LeerArchivoPosiciones(RutaArchivoPosiciones As String) As cListaPosiciones
    'devuelve la lista de posiciones a evaluar
    Dim NumeroPantalla As Byte
    Dim Orientacion As Byte
    Dim X As Byte
    Dim Y As Byte
    Dim Z As Byte
    Dim Escaleras As Byte
    Dim Posicion As cPosicion
    Dim Result As New cListaPosiciones
    Set LeerArchivoPosiciones = Result
    If Dir(RutaArchivoPosiciones) = "" Then Exit Function
    Open RutaArchivoPosiciones For Input As #2
    While Not EOF(2)
        Line Input #2, Linea
        Campos = Split(Linea, ",")
        If UBound(Campos) = 5 Then
            If IsNumeric(Campos(0)) And IsNumeric(Campos(1)) And IsNumeric(Campos(3)) And IsNumeric(Campos(4)) And IsNumeric(Campos(5)) Then
                NumeroPantalla = CByte(Campos(0))
                Orientacion = CByte(Campos(1))
                X = CByte(Campos(2))
                Y = CByte(Campos(3))
                Z = CByte(Campos(4))
                Escaleras = CByte(Campos(5))
                Set Posicion = New cPosicion
                Posicion.Init NumeroPantalla, Orientacion, X, Y, Z, Escaleras
                Result.Append Posicion
                DoEvents
            End If
        End If
        DoEvents
    Wend
    Close #2
    
End Function
