Attribute VB_Name = "modAbadia"
Option Explicit

'Dim bug As Boolean

Public Depuracion As New cDepuracion
'tablas del juego
Private Parar As Boolean
Public Parado As Boolean
Private Check As Boolean 'true para hacer una pasada por el bucle principal, ajustando la posici�n y orientaci�n de guillermo, y guardando las tablas en disco
Private CheckPantalla As String
Private CheckOrientacion As Byte
Private CheckX As Byte
Private CheckY As Byte
Private CheckZ As Byte
Private CheckEscaleras As Byte
Private CheckRuta As String

Dim TablaBufferAlturas_01C0(&H23F) As Byte '576 bytes (24x24) = (4 + 16 + 4)x2  RAM
Dim TablaBloquesPantallas_156D(&HBF) As Byte
Dim DatosAlturaEspejoCerrado_34DB(4) As Byte  'datos de altura si el espejo est� cerrado
Dim TablaRutinasConstruccionBloques_1FE0(&H37) As Byte 'no se usa
Dim VariablesBloques_1FCD(&H12) As Byte
'Dim DatosTilesBloques_1693(&H92) As Byte
Dim TablaCaracteristicasMaterial_1693(&H924) As Byte

Dim TablaHabitaciones_2255(&HFF) As Byte '(realmente empieza en 0x2265 porque en Y = 0 no hay nada)
Dim TablaAvancePersonaje4Tiles_284D(31) As Byte
Dim TablaAvancePersonaje1Tile_286D(31) As Byte

Dim TablaDatosPersonajes_2BAE(&H3D) As Byte 'tabla con datos para mover los personajes
Dim TablaVariablesAuxiliares_2D8D(&H4B) As Byte 'variables auxiliares de algunas rutinas
Dim TablaPermisosPuertas_2DD9(18) As Byte 'copiado en 0x122-0x131. puertas a las que pueden entrar los personajes
Dim CopiaTablaPermisosPuertas_2DD9(18) As Byte
Dim TablaObjetosPersonajes_2DEC(&H2A) As Byte 'copiado en 0x132-0x154. objetos de los personajes
Dim CopiaTablaObjetosPersonajes_2DEC(&H2A) As Byte
Dim TablaSprites_2E17(&H1CC) As Byte 'sprites de los personajes, puertas y objetos
Dim TablaDatosPuertas_2FE4(&H23) As Byte 'datos de las puertas del juego. 5 bytes por entrada
Dim CopiaTablaDatosPuertas_2FE4(&H23) As Byte
Dim TablaPosicionObjetos_3008(&H2D) As Byte 'posici�n de los objetos del juego 5 bytes por entrada
Dim CopiaTablaPosicionObjetos_3008(&H2D) As Byte
Dim TablaCaracteristicasPersonajes_3036(&H59) As Byte
Dim TablaPunterosCarasMonjes_3097(&H7) As Byte
Dim TablaDesplazamientoAnimacion_309F(&HFF) As Byte 'tabla para el c�lculo del desplazamiento seg�n la animaci�n de una entidad del juego
Dim TablaAnimacionPersonajes_319F(&H5F) As Byte
Dim TablaAccesoHabitaciones_3C67(&H1D) As Byte
Dim TablaVariablesLogica_3C85(&H20) As Byte
Dim TablaPosicionesPredefinidasMalaquias_3CA8(&H1D) As Byte
Dim TablaPosicionesPredefinidasAbad_3CC6(&H20) As Byte
Dim TablaPosicionesPredefinidasBerengario_3CE7(&H17) As Byte 'berengario/bernardo gui/encapuchado/jorge
Dim TablaPosicionesPredefinidasSeverino_3CFF(&H11) As Byte 'severino/jorge
Dim TablaPosicionesPredefinidasAdso_3D11(&HB) As Byte
Dim TablaPunterosVariablesScript_3D1D(&H81) As Byte 'tabla de asociaci�n de constantes a direcciones de memoria importantes para el programa (usado por el sistema de script)
Dim DatosHabitaciones_4000(&H2329) As Byte
Dim TablaPunterosTrajesMonjes_48C8(&H1F) As Byte
Dim TablaPatronRellenoLuz_48E8(&H1F) As Byte
Dim TablaAlturasPantallas_4A00(&HA1F) As Byte
Dim TablaEtapasDia_4F7A(&H72) As Byte '4F7A:tabla de duraci�n de las etapas del d�a para cada d�a y periodo del d�a 4FA7:tabla usada para rellenar el n�mero del d�a en el marcador 4FBC:tabla de los nombres de los momentos del d�a
Dim DatosMarcador_6328(&H7FF) As Byte 'datos del marcador (de 0x6328 a 0x6b27)
Dim DatosCaracteresPergamino_6947(&H9B8) As Byte
Dim PunterosCaracteresPergamino_680C(&HB9) As Byte
Dim TilesAbadia_6D00(&H1FFF) As Byte
Dim TablaRellenoBugTiles_8D00(&H7F) As Byte
Dim TextoPergaminoPresentacion_7300(&H589) As Byte
Dim DatosGraficosPergamino_788A(&H5FF) As Byte
Dim BufferTiles_8D80(&H77F) As Byte
Dim BufferSprites_9500(&H77F) As Byte
Dim TablasAndOr_9D00(&H3FF) As Byte
Dim TablaFlipX_A100(&HFF) As Byte
Dim TablaGraficosObjetos_A300(&H858) As Byte 'gr�ficos de guillermo, adso y las puertas
Dim DatosMonjes_AB59(&H8A6) As Byte 'gr�ficos de los movimientos de los monjes ab59-ae2d normal, ae2e-b102 flipx, 0xb103-0xb3ff caras y trajes
Dim BufferComandosMonjes_A200(&HFF) As Byte 'buffer de comandos de los movimientos de los monjes y adso
'Dim TablaCarasTrajesMonjes_B103(&H2FC) As Byte 'caras y trajes de los monjes
Dim PantallaCGA(&H3FFF) As Byte


Dim ContadorAnimacionGuillermo_0990 As Byte 'contador de la animaci�n de guillermo
Dim PintarPantalla_0DFD As Boolean 'usada en las rutinas de las puertas indicando que pinta la pantalla
Dim RedibujarPuerta_0DFF As Boolean 'indica que se redibuje el sprite
Dim TempoMusica_1086 As Byte
Public HabitacionOscura_156C As Boolean 'lee si es una habitaci�n iluminada
Public PunteroPantallaActual_156A As Long 'direcci�n de los datos de inicio de la pantalla actual
Dim PunteroPlantaActual_23EF As Long 'direcci�n del mapa de la planta
Dim OrientacionPantalla_2481 As Byte
Dim VelocidadPasosGuillermo_2618 As Byte
Dim MinimaPosicionYVisible_279D As Byte 'm�nima posici�n y visible en pantalla
Dim MinimaPosicionXVisible_27A9 As Byte 'm�nima posici�n x visible en pantalla
Dim MinimaAlturaVisible_27BA As Byte 'm�nima altura visible en pantalla
Dim EstadoGuillermo_288F As Byte
Dim AjustePosicionYSpriteGuillermo_28B1 As Integer
Dim PunteroRutinaFlipPersonaje_2A59 As Long 'rutina a la que llamar si hay que flipear los gr�ficos
Dim PunteroTablaAnimacionesPersonaje_2A84 As Long 'direcci�n de la tabla de animaciones para el personaje
Dim LimiteInferiorVisibleX_2AE1 As Byte 'limite inferior visible de X
Dim LimiteInferiorVisibleY_2AEB As Byte 'limite inferior visible de y
Dim AlturaBasePlantaActual_2AF9 As Byte 'altura base de la planta
Dim RutinaCambioCoordenadas_2B01 As Long 'rutina que cambia el sistema de coordenadas dependiendo de la orientaci�n de la pantalla
Dim ModificarPosicionSpritePantalla_2B2F As Boolean 'true para modificar la posici�n del sprite en pantalla dentro de &H2ADD
Dim ContadorInterrupcion_2D4B As Byte 'contador que se incrementa en la interrupci�n
Dim PosicionXPersonajeActual_2D75 As Byte 'posici�n en x del personaje que se muestra en pantalla
Dim PosicionYPersonajeActual_2D76 As Byte 'posici�n en y del personaje que se muestra en pantalla
Dim PosicionZPersonajeActual_2D77 As Byte 'posici�n en z del personaje que se muestra en pantalla
Dim NumeroDia_2D80 As Byte 'n�mero de d�a (del 1 al 7)
Dim MomentoDia_2D81 As Byte 'momento del d�a 0=noche, 1=prime,2=tercia,4=nona,5=v�speras,6=completas
Dim HabitacionEspejoCerrada_2D8C As Boolean 'si vale true indica que no se ha abierto el espejo
Dim ScrollCambioMomentoDia_2DA5 As Byte 'posiciones para realizar el scroll del cambio del momento del d�a
Dim PuertaRequiereFlip_2DAF As Boolean 'si la puerta necesita gr�ficos flipeados o no
Dim CambioPantalla_2DB8 As Boolean 'indica que ha habido un cambio de pantalla
Dim AlturaBasePlantaActual_2DBA As Byte 'altura base de la planta en la que est� el personaje de la rejilla ###en 2af9 hay otra
Dim NumeroRomanoHabitacionEspejo_2DBC As Byte 'si es != 0, contiene el n�mero romano generado para el enigma de la habitaci�n del espejo
Dim NumeroPantallaActual_2DBD As Byte 'pantalla del personaje al que sigue la c�mara
Dim MovimientoRealizado_2DC1 As Boolean 'indica que ha habido movimiento
Dim GuillermoMuerto_3C97 As Boolean
Dim PunteroDatosAlturaHabitacionEspejo_34D9 As Long
Dim PunteroHabitacionEspejo_34E0 As Long
Dim PersonajeSeguidoPorCamara_3C8F As Byte 'personaje al que sigue la c�mara

Dim MalaquiasAscendiendo_4384 As Boolean 'indica que malaqu�as est� ascendiendo mientras se est� muriendo
Dim SpriteLuzAdsoX_4B89 As Byte 'posici�n x del sprite de adso dentro del tile
Dim SpriteLuzAdsoX_4BB5 As Byte '4 - (posici�n x del sprite de adso & 0x03)
Dim SpriteLuzTipoRelleno_4B6B As Byte 'bytes a rellenar (tile/ tile y medio)
Dim SpriteLuzTipoRelleno_4BD1 As Byte 'bytes a rellenar (tile y medio / tile)
Dim SpriteLuzFlip_4BA0 As Boolean 'true si los gr�ficos de adso est�n flipeados

Dim SpritesPilaProcesados_4D85 As Boolean 'false si no ha terminado de procesar los sprites de la pila. true: limpia el bit 7 de (ix+0) del buffer de tiles (si es una posici�n v�lida del buffer)
Dim PunteroPantalla As Long 'posici�n actual dentro de la pantalla mientras se procesa
Dim InvertirDireccionesGeneracionBloques As Boolean
Dim Pila(100) As Long
Dim PunteroPila As Long
Enum EnumIncremento
    IncSumarX
    IncRestarX
    IncRestarY
End Enum
Public Pintar As Boolean
'
'Variables que necesitan un valor inicial
Dim Obsequium_2D7F As Byte
Dim PunteroProximaHoraDia_2D82 As Long  'puntero a la pr�xima hora del d�a
Dim PunteroTablaDesplazamientoAnimacion_2D84 As Long 'direcci�n de la tabla para el c�lculo del desplazamiento seg�n la animaci�n de una entidad del juego para la orientaci�n de la pantalla actual
Dim TiempoRestanteMomentoDia_2D86 As Long 'cantidad de tiempo a esperar para que avance el momento del d�a (siempre y cuando sea distinto de cero)
Dim PunteroDatosPersonajeActual_2D88 As Long 'puntero a los datos del personaje actual que se sigue la c�mara
Dim PunteroBufferAlturas_2D8A As Long 'puntero al buffer de alturas de la pantalla actual (buffer de 576 (24*24) bytes)

Dim PuertasAbribles_3CA6 As Byte
Dim InvestigacionNoTerminada_3CA7 As Boolean

Public Sub PararAbadia()
    Parar = True
End Sub

Public Sub CheckDefinir(ByVal NumeroPantalla As Byte, ByVal Orientacion As Byte, ByVal X As Byte, ByVal Y As Byte, ByVal Z As Byte, ByVal Escaleras As Byte, ByVal RutaCheck As String)
    'guarda las variables del modo check para hacer un bucle y guardar las tablas
    Dim Pantalla As String
    Check = True
    Pantalla = Hex$(NumeroPantalla)
    If Len(Pantalla) < 2 Then Pantalla = "0" + Pantalla
    CheckPantalla = Pantalla
    CheckOrientacion = Orientacion
    CheckX = X
    CheckY = Y
    CheckZ = Z
    CheckEscaleras = Escaleras
    CheckRuta = modFunciones.FixPath(RutaCheck)
End Sub



Public Sub CargarDatos()

Dim Archivo() As Byte
'abadia1.bin
CargarArchivo "D:\datos\vbasic\Abadia\Abadia2\abadia1.bin", Archivo
CargarTablaArchivo Archivo, TablaBloquesPantallas_156D, &H156D
CargarTablaArchivo Archivo, TablaRutinasConstruccionBloques_1FE0, &H1FE0
'CargarTablaArchivo Archivo, DatosTilesBloques_1693, &H1693
CargarTablaArchivo Archivo, TablaCaracteristicasMaterial_1693, &H1693
CargarTablaArchivo Archivo, TablaHabitaciones_2255, &H2255

CargarTablaArchivo Archivo, TablaAvancePersonaje4Tiles_284D, &H284D
CargarTablaArchivo Archivo, TablaAvancePersonaje1Tile_286D, &H286D
CargarTablaArchivo Archivo, TablaDatosPersonajes_2BAE, &H2BAE
CargarTablaArchivo Archivo, TablaPermisosPuertas_2DD9, &H2DD9
CargarTablaArchivo Archivo, TablaObjetosPersonajes_2DEC, &H2DEC
CargarTablaArchivo Archivo, TablaSprites_2E17, &H2E17
CargarTablaArchivo Archivo, TablaDatosPuertas_2FE4, &H2FE4
CargarTablaArchivo Archivo, TablaPosicionObjetos_3008, &H3008
CargarTablaArchivo Archivo, TablaCaracteristicasPersonajes_3036, &H3036
CargarTablaArchivo Archivo, TablaPunterosCarasMonjes_3097, &H3097
CargarTablaArchivo Archivo, TablaDesplazamientoAnimacion_309F, &H309F
CargarTablaArchivo Archivo, TablaAnimacionPersonajes_319F, &H319F
CargarTablaArchivo Archivo, DatosAlturaEspejoCerrado_34DB, &H34DB
CargarTablaArchivo Archivo, TablaAccesoHabitaciones_3C67, &H3C67
CargarTablaArchivo Archivo, TablaVariablesLogica_3C85, &H3C85
CargarTablaArchivo Archivo, TablaPosicionesPredefinidasMalaquias_3CA8, &H3CA8
CargarTablaArchivo Archivo, TablaPosicionesPredefinidasAbad_3CC6, &H3CC6
CargarTablaArchivo Archivo, TablaPosicionesPredefinidasBerengario_3CE7, &H3CE7
CargarTablaArchivo Archivo, TablaPosicionesPredefinidasSeverino_3CFF, &H3CFF
CargarTablaArchivo Archivo, TablaPosicionesPredefinidasAdso_3D11, &H3D11
CargarTablaArchivo Archivo, TablaPunterosVariablesScript_3D1D, &H3D1D

'abadia2.bin
CargarArchivo "D:\datos\vbasic\Abadia\Abadia2\abadia2.bin", Archivo
CargarTablaArchivo Archivo, TablaPunterosTrajesMonjes_48C8, &H8C8&
CargarTablaArchivo Archivo, TablaPatronRellenoLuz_48E8, &H8E8&
CargarTablaArchivo Archivo, TablaEtapasDia_4F7A, &HF7A&
CargarTablaArchivo Archivo, PunterosCaracteresPergamino_680C, &H280C
CargarTablaArchivo Archivo, DatosCaracteresPergamino_6947, &H2947
CargarTablaArchivo Archivo, TextoPergaminoPresentacion_7300, &H3300
CargarTablaArchivo Archivo, DatosGraficosPergamino_788A, &H388A

'abadia3.bin
CargarArchivo "D:\datos\vbasic\Abadia\Abadia2\abadia3.bin", Archivo
CargarTablaArchivo Archivo, TilesAbadia_6D00, &H300
CargarTablaArchivo Archivo, TablaRellenoBugTiles_8D00, &HD00&
CargarTablaArchivo Archivo, BufferSprites_9500, &H1500
CargarTablaArchivo Archivo, TablaGraficosObjetos_A300, &H2300
CargarTablaArchivo Archivo, DatosMonjes_AB59, &H2B59
'CargarTablaArchivo Archivo, TablaCarasTrajesMonjes_B103, &H3103

'abadia7.bin -> alturas de las pantallas
CargarArchivo "D:\datos\vbasic\Abadia\Abadia2\abadia7.bin", Archivo
CargarTablaArchivo Archivo, TablaAlturasPantallas_4A00, &HA00

'abadia8.bin -> datos de las pantallas
CargarArchivo "D:\datos\vbasic\Abadia\Abadia2\abadia8.bin", Archivo
CargarTablaArchivo Archivo, DatosHabitaciones_4000, 0 '0x0000-0x2237 datos sobre los bloques que forman las pantallas
CargarTablaArchivo Archivo, DatosMarcador_6328, &H2328 'datos del marcador (de 0x6328 a 0x6b27)



End Sub

Private Sub InicializarVariablesROM()
    PunteroPlantaActual_23EF = &H2255
    PosicionZPersonajeActual_2D77 = 0
End Sub

Public Sub CargarTablaArchivo(ByRef Archivo() As Byte, ByRef Tabla() As Byte, ByVal Puntero As Long)
'rellena la tabla con los datos del archivo desde la posici�n indicada
Dim Contador As Long
For Contador = 0 To UBound(Tabla)
    Tabla(Contador) = Archivo(Puntero + Contador)
Next
End Sub

Sub CopiarTabla(ByRef TablaOrigen() As Byte, ByRef TablaDestino() As Byte)
Dim Contador As Long
For Contador = 0 To UBound(TablaOrigen)
    TablaDestino(Contador) = TablaOrigen(Contador)
Next
End Sub



Public Sub DibujarPantalla_19D8()
'dibuja la pantalla que hay en el buffer de tiles
Dim ColorFondo As Byte
If Not HabitacionOscura_156C Then
    ColorFondo = 0  'color de fondo = azul
Else
    ColorFondo = &HFF 'color de fondo = negro
End If
LimpiarRejilla_1A70 ColorFondo 'limpia la rejilla y rellena un rect�ngulo de 256x160 a partir de (32, 0) con el color de fondo
PunteroPantalla = PunteroPantallaActual_156A + 1 'avanza el byte de longitud
GenerarEscenario_1A0A 'genera el escenerio y lo proyecta a la rejilla
'si es una habitaci�n iluminada, dibuja en pantalla el contenido de la rejilla desde el centro hacia afuera

If Not HabitacionOscura_156C Then DibujarPantalla_4EB2 'dibuja en pantalla el contenido de la rejilla desde el centro hacia afuera
End Sub

Public Sub LimpiarRejilla_1A70(ByVal ColorFondo As Byte)
'limpia la rejilla y rellena en pantalla un rect�ngulo de 256x160 a partir de (32, 0) con el color indicado
Dim Contador As Long
Dim Linea As Long
Dim Columna As Long
Dim PunteroPantalla As Long
For Contador = 0 To UBound(BufferTiles_8D80)
    BufferTiles_8D80(Contador) = 0 'limpia 0x8d80-0x94ff
Next
'rellena un rect�ngulo de 160 de alto por 256 de ancho a partir de la posici�n (32, 0) con ColorFondo
PunteroPantalla = &H8&    'posici�n (32, 0)
For Linea = 1 To 160
    For Columna = 0 To 63 'rellena 64 bytes (256 pixels)
        PantallaCGA(PunteroPantalla + Columna) = ColorFondo
        PantallaCGA2PC PunteroPantalla + Columna, ColorFondo
    Next
    PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla)
Next
modPantalla.Refrescar
End Sub

Function DireccionSiguienteLinea_3A4D_68F2(ByVal PunteroPantalla As Long) As Long
'devuelve la direcci�n de la siguiente l�nea de pantalla
Dim Puntero As Long
Puntero = PunteroPantalla + &H800 'pasa al siguiente banco
If Puntero > &H3FFF& Then
    Puntero = PunteroPantalla And &H7FF&
    Puntero = Puntero + &H50
End If
'pasa a la siguiente l�nea y ajusta para que est� en el rango 0xc000-0xffff
DireccionSiguienteLinea_3A4D_68F2 = Puntero
End Function



Public Sub GenerarEscenario_1A0A()
'genera el escenerio con los datos de abadia8 y lo proyecta a la rejilla
'lee la entrada de abadia8 con un bloque de construcci�n de la pantalla y llama a 0x1bbc
Dim Bloque As Long
Dim Byte1 As Byte
Dim Byte2 As Byte
Dim Byte3 As Byte
Dim Byte4 As Byte
Dim X As Byte 'pos en x del elemento (sistema de coordenadas del buffer de tiles)
Dim nX As Byte 'longitud del elemento en x
Dim Y As Byte 'pos en y del elemento (sistema de coordenadas del buffer de tiles)
Dim nY As Byte 'longitud del elemento en y
Dim PunteroCaracteristicasBloque As Long 'puntero a las caracter�sitcas del bloque
Dim PunteroTilesBloque As Long 'puntero del material a los tiles que forman el bloque
Dim PunteroRutinasBloque As Long 'puntero al resto de caracter�sticas del material
Dim salir As Boolean
Dim BloqueHex As String
Dim Eva As Long
'PunteroPantalla = 2445

Do 'provisional
    Pintar = True
    Bloque = DatosHabitaciones_4000(PunteroPantalla)
    BloqueHex = Hex$(Bloque)
'1A0D
    If Bloque = 255 Then Exit Sub '0xff indica el fin de pantalla
    'Bloque = Bloque And &HFE& 'desprecia el bit inferior para indexar
'1A10
    PunteroCaracteristicasBloque = Leer16(TablaBloquesPantallas_156D, Bloque And &HFE&) 'desprecia el bit inferior para indexar
'1A21
    Byte1 = DatosHabitaciones_4000(PunteroPantalla + 1)
'1A24
    X = Byte1 And &H1F 'pos en x del elemento (sistema de coordenadas del buffer de tiles)
'1A28
    nX = modFunciones.shr(Byte1, 5) And &H7 'longitud del elemento en x
'1A2F
    Byte2 = DatosHabitaciones_4000(PunteroPantalla + 2)
'1A32
    Y = Byte2 And &H1F 'pos en y del elemento (sistema de coordenadas del buffer de tiles)
'1A36
    nY = modFunciones.shr(Byte2, 5) And &H7 'longitud del elemento en y
'1A3D
    VariablesBloques_1FCD(&H1FDE - &H1FCD) = 0 'inicia a (0, 0) la posici�n del bloque en la rejilla (sistema de coordenadas local de la rejilla)
    VariablesBloques_1FCD(&H1FDF - &H1FCD) = 0 'inicia a (0, 0) la posici�n del bloque en la rejilla (sistema de coordenadas local de la rejilla)
'1A47
    PunteroPantalla = PunteroPantalla + 3
    If Bloque Mod 2 = 0 Then
        Byte3 = &HFF 'la entrada es de 3 bytes
    Else
'1A53
        Byte3 = DatosHabitaciones_4000(PunteroPantalla)
        PunteroPantalla = PunteroPantalla + 1
    End If
'1A58
    VariablesBloques_1FCD(&H1FDD - &H1FCD) = Byte3
    PunteroTilesBloque = Leer16(TablaCaracteristicasMaterial_1693, PunteroCaracteristicasBloque - &H1693)
    PunteroRutinasBloque = PunteroCaracteristicasBloque + 2
'1A69
    ConstruirBloque_1BBC X, nX, Y, nY, Byte3, PunteroTilesBloque, PunteroRutinasBloque, True
    If salir Then Exit Sub
    Pintar = False
Loop

Exit Sub
modFunciones.GuardarArchivo "BufferTiles0", BufferTiles_8D80 '&H77f

End Sub

Public Sub DibujarPantalla_4EB2()
'dibuja en pantalla el contenido de la rejilla desde el centro hacia afuera
Dim PunteroPantalla As Long
Dim PunteroRejilla As Long
Dim NAbajo As Long 'n� de posiciones a dibujar hacia abajo
Dim NArriba As Long 'n� de posiciones a dibujar hacia arriba
Dim NDerecha As Long 'n� de posiciones a dibujar hacia la derecha
Dim NIzquierda As Long  'n� de posiciones a dibujar hacia la izquierda
Dim NTiles As Long 'n� de posiciones a dibujar
Dim DistanciaRejilla As Long 'distancia entre elementos consecutivos en la rejilla. cambia si se dibuja en vertical o en horizontal
Dim DistanciaPantalla As Long 'distancia entre elementos consecutivos en la pantalla. cambia si se dibuja en vertical o en horizontal
PunteroPantalla = &H2A4&  '(144, 64) coordenadas de pantalla
PunteroRejilla = &H90AA& '(7, 8) coordenadas de rejilla
NAbajo = 4
NArriba = 5
NDerecha = 1
NIzquierda = 2
Do
    If NAbajo >= 20 Then Exit Sub 'si dibuja m�s de 20 posiciones verticales, sale
    NTiles = NAbajo
    NAbajo = NAbajo + 2 'en la pr�xima iteraci�n dibujar� 2 posiciones verticales m�s hacia abajo
    DistanciaRejilla = &H60 'tama�o entre l�neas de la rejilla
    DistanciaPantalla = &H50 'tama�o entre l�neas en la memoria de v�deo
    DibujarTiles_4F18 NTiles, DistanciaRejilla, DistanciaPantalla, PunteroRejilla, PunteroPantalla 'dibuja posiciones verticales de la rejilla en la memoria de video
    modPantalla.Refrescar
    NTiles = NDerecha
    NDerecha = NDerecha + 2 'en la pr�xima iteraci�n dibujar� 2 posiciones horizontales m�s hacia la derecha
    DistanciaRejilla = 6 'tama�o entre posiciones x de la rejilla
    DistanciaPantalla = 4 'tama�o entre cada 16 pixels en la memoria de video
    DibujarTiles_4F18 NTiles, DistanciaRejilla, DistanciaPantalla, PunteroRejilla, PunteroPantalla 'dibuja posiciones horizontales de la rejilla en la memoria de video
    modPantalla.Refrescar
    NTiles = NArriba
    NArriba = NArriba + 2 'en la pr�xima iteraci�n dibujar� 2 posiciones verticales m�s hacia arriba
    DistanciaRejilla = -&H60 'valor para volver a la l�nea anterior de la rejilla
    DistanciaPantalla = -&H50 'valor para volver a la l�nea anterior de la pantalla
    DibujarTiles_4F18 NTiles, DistanciaRejilla, DistanciaPantalla, PunteroRejilla, PunteroPantalla 'dibuja  posiciones verticales de la rejilla en la memoria de video
    modPantalla.Refrescar
    NTiles = NIzquierda
    NIzquierda = NIzquierda + 2 'en la pr�xima iteraci�n dibujar� 2 posiciones horizontales m�s hacia la izquierda
    DistanciaRejilla = -6 'valor para volver a la anterior posicion x de la rejilla
    DistanciaPantalla = -4 'valor para volver a la anterior posicion x de la pantalla
    DibujarTiles_4F18 NTiles, DistanciaRejilla, DistanciaPantalla, PunteroRejilla, PunteroPantalla ' dibuja posiciones horizontales de la rejilla en la memoria de video
    modPantalla.Refrescar
Loop 'repite hasta que se termine

End Sub

Sub DibujarTiles_4F18(ByVal NTiles As Long, ByVal DistanciaRejilla As Long, ByVal DistanciaPantalla As Long, ByRef PunteroRejilla As Long, ByRef PunteroPantalla As Long)
'dibuja NTiles posiciones horizontales o verticales de la rejilla en la memoria de video
'NTiles = n�mero de posiciones a dibujar
'DistanciaRejilla = tama�o entre posiciones de la rejilla
'DistanciaPantalla = tama�o entre posiciones en la memoria de v�deo
'PunteroRejilla = posici�n en el buffer
'PunteroPantalla = posici�n en la memoria de v�deo
Dim Contador As Long
Dim NumeroTile As Byte
For Contador = 1 To NTiles 'n�mero de posiciones a dibujar
    NumeroTile = BufferTiles_8D80(PunteroRejilla + 2 - &H8D80&) 'lee el n�mero de gr�fico a dibujar (fondo)
    If NumeroTile <> 0 Then
        DibujarTile_4F3D NumeroTile, PunteroPantalla 'copia un gr�fico 16x8 a la memoria de video, combinandolo con lo que hab�a
    End If
    NumeroTile = BufferTiles_8D80(PunteroRejilla + 5 - &H8D80&) 'lee el n�mero de gr�fico a dibujar (fondo)
    If NumeroTile <> 0 Then
        DibujarTile_4F3D NumeroTile, PunteroPantalla 'copia un gr�fico 16x8 a la memoria de video, combinandolo con lo que hab�a
    End If
    PunteroRejilla = PunteroRejilla + DistanciaRejilla
    PunteroPantalla = PunteroPantalla + DistanciaPantalla
Next
End Sub

Public Sub DibujarTile_4F3D(ByVal NumeroTile As Byte, ByVal PunteroPantalla As Long)
'copia el gr�fico NumeroTile (16x8) en la memoria de video (PunteroPantalla), combinandolo con lo que hab�a
'NumeroTile = bits 7-0: n�mero de gr�fico. El bit 7 = indica qu� color sirve de m�scara (el 2 o el 1)
'PunteroPantalla = posici�n en la memoria de video
Dim PunteroTile As Long 'apunta al gr�fico correspondiente
Dim PunteroAndOr As Long 'valor de la tabla AND/OR
Dim ValorAND As Byte
Dim ValorOR As Byte
Dim ValorGrafico As Byte
Dim ValorPantalla As Byte
Dim Linea As Long
Dim Columna As Long
PunteroTile = 32 * NumeroTile 'direcci�n del gr�fico
If (NumeroTile And &H80) <> 0 Then 'dependiendo del bit 7 escoge una tabla AND y OR
    PunteroAndOr = &H200
End If
For Linea = 0 To 7 '8 pixels de alto
    For Columna = 0 To 3 '4 bytes de ancho (16 pixels)
        ValorGrafico = TilesAbadia_6D00(PunteroTile + 4 * Linea + Columna) 'lee un byte del gr�fico
        ValorOR = TablasAndOr_9D00(PunteroAndOr + ValorGrafico) 'valor de la tabla OR
        ValorAND = TablasAndOr_9D00(PunteroAndOr + &H100 + ValorGrafico) 'valor de la tabla AND
        ValorPantalla = PantallaCGA(PunteroPantalla + Columna + Linea * &H800)
        ValorPantalla = ValorPantalla And ValorAND
        ValorPantalla = ValorPantalla Or ValorOR
        PantallaCGA(PunteroPantalla + Columna + Linea * &H800) = ValorPantalla
        PantallaCGA2PC PunteroPantalla + Columna + Linea * &H800, ValorPantalla
    Next
Next
End Sub

Public Function BuscarHabitacionProvisional(ByVal NumeroPantalla As Long) As Long
'devuelve el puntero al primer byte de la habitaci�n indicada
Dim Contador As Long
Dim Puntero As Long
Puntero = 0
Do
    If Contador >= NumeroPantalla Then
        BuscarHabitacionProvisional = Puntero
        Exit Function
    End If
    Contador = Contador + 1
    Puntero = Puntero + DatosHabitaciones_4000(Puntero)
Loop
End Function


Public Sub ConstruirBloque_1BBC(ByVal X As Byte, ByVal nX As Byte, ByVal Y As Byte, ByVal nY As Byte, ByVal Altura As Byte, ByVal PunteroTilesBloque As Long, ByVal PunteroRutinasBloque As Long, ActualizarVariablesTiles As Boolean)
'inicia el buffer para la construcci�n del bloque actual y evalua los par�metros de construcci�n del bloque
Dim Contador As Long
If ActualizarVariablesTiles Then
    For Contador = 0 To 11
        VariablesBloques_1FCD(Contador + 2) = TablaCaracteristicasMaterial_1693(PunteroTilesBloque - &H1693 + Contador) '1FCF = buffer de destino
    Next
End If
TransformarPosicionBloqueCoordenadasRejilla_1FB8 X, Y, Altura
GenerarBloque_2018 X, nX, Y, nY, PunteroRutinasBloque
End Sub


Public Sub TransformarPosicionBloqueCoordenadasRejilla_1FB8(ByVal X As Byte, ByVal Y As Byte, ByVal Altura As Byte)
Dim Xr As Long
Dim Yr As Long
'si la entrada es de 4 bytes, transforma la posici�n del bloque a coordenadas de la rejilla
' las ecuaciones de cambio de sistema de coordenadas son:
' mapa de tiles -> rejilla
' Xrejilla = Ymapa + Xmapa - 15
' Yrejilla = Ymapa - Xmapa + 16
' rejilla -> mapa de tiles
' Xmapa = Xrejilla - Ymapa + 15
' Ymapa = Yrejilla + Xmapa - 16
' de esta forma los datos de la rejilla se almacenan en el mapa de tiles de forma que la conversi�n a la pantalla es directa
If Altura = &HFF Then Exit Sub
Xr = CLng(Y) + CLng(X) + CLng(Altura / 2) - 15
Yr = CLng(Y) - CLng(X) + CLng(Altura / 2) + 16
If Xr < 0 Then
    Xr = Xr + 256
End If
If Yr < 0 Then
    Yr = Yr + 256
End If
VariablesBloques_1FCD(&H1FDE - &H1FCD) = Xr
VariablesBloques_1FCD(&H1FDF - &H1FCD) = Yr
 'comprobar


End Sub

Public Sub GenerarBloque_2018(ByVal X As Byte, ByVal nX As Byte, ByVal Y As Byte, ByVal nY As Byte, ByVal PunteroRutinasBloque As Long)
'inicia el proceso de interpretaci�n los bytes de construcci�n de bloques
VariablesBloques_1FCD(&H1FDB - &H1FCD) = nX
VariablesBloques_1FCD(&H1FDC - &H1FCD) = nY
EvaluarDatosBloque_201E X, nX, Y, nY, PunteroRutinasBloque
End Sub

Sub EvaluarDatosBloque_201E(ByVal X As Byte, ByVal nX As Byte, ByVal Y As Byte, ByVal nY As Byte, ByVal PunteroRutinasBloque As Long)
'eval�a los datos de construcci�n del bloque
'x = pos inicial del bloque en y (sistema de coordenadas del buffer de tiles)
'y = pos inicial del bloque en x (sistema de coordenadas del buffer de tiles)
'ny = lgtud del elemento en y
'nx = lgtud del elemento en x
'PunteroRutinasBloque = puntero a los datos de construcci�n del bloque
Dim Rutina As String
Dim DatosBloque As String
Static TerminarEvaluacion As Boolean
TerminarEvaluacion = False
Do
    DatosBloque = Bytes2AsciiHex(VariablesBloques_1FCD)
    Rutina = Hex$(TablaCaracteristicasMaterial_1693(PunteroRutinasBloque - &H1693))
    PunteroRutinasBloque = PunteroRutinasBloque + 1
    If Len(Rutina) < 2 Then Rutina = "0" + Rutina
    Select Case Rutina
        Case Is = "E4" 'interpreta otro bloque sin modificar los valores de los tiles a usar, y cambiando el sentido de las x
            Rutina_E4_21AA X, nX, Y, nY, PunteroRutinasBloque
        Case Is = "E5" 'cambia las instrucciones que actualizan la coordenada x de los tiles (incx -> decx)
            Rutina_E9_218D
        Case Is = "E6" 'cambia las instrucciones que actualizan la coordenada x de los tiles (incx -> decx)
            Rutina_E9_218D
        Case Is = "E7" 'cambia las instrucciones que actualizan la coordenada x de los tiles (incx -> decx)
            Rutina_E9_218D
        Case Is = "E8" 'cambia las instrucciones que actualizan la coordenada x de los tiles (incx -> decx)
            Rutina_E9_218D
        Case Is = "E9" 'cambia las instrucciones que actualizan la coordenada x de los tiles (incx -> decx)
            Rutina_E9_218D
        Case "EA" 'cambia el puntero a los datos de construcci�n del bloque con la primera direcci�n leida en los datos
            Rutina_EA_21A1 X, nX, Y, nY, PunteroRutinasBloque
        Case Is = "EB"
            Stop
        Case Is = "EC" 'interpreta otro bloque modificando los valores de los tiles a usar
            Rutina_EC_21B4 X, nX, Y, nY, PunteroRutinasBloque, True
        Case Is = "ED"
            Stop
        Case Is = "EE"
            Stop
        Case Is = "EF" 'incrementa la longitud del bloque en x en el buffer de construcci�n del bloque
            Rutina_EF_2071 PunteroRutinasBloque
        Case Is = "F0" 'incrementa la longitud del bloque en y en el buffer de construcci�n del bloque
            Rutina_F0_2077 PunteroRutinasBloque
        Case Is = "F1" 'modifica la posici�n en x con la expresi�n leida
            Rutina_F1_2066 X, PunteroRutinasBloque
        Case Is = "F2" 'modifica la posici�n en y con la expresi�n leida
            Rutina_F2_205B Y, PunteroRutinasBloque
        Case Is = "F3" 'cambia la posici�n de x (x--)
            Rutina_F3_2058 X
        Case Is = "F4" 'cambia la posici�n de Y (y--)
            Rutina_F4_2055 Y
        Case Is = "F5" 'cambia la posici�n de x (x++)
            Rutina_F5_2052 X
        Case Is = "F6" 'cambia la posici�n de Y (y++)
            Rutina_F6_204F Y
        Case Is = "F7" ' modifica una posici�n del buffer de construcci�n del bloque con una expresi�n calculada
            Rutina_F7_2141 nX, PunteroRutinasBloque
        Case Is = "F8" 'pinta el tile indicado por X,Y con el siguiente byte leido y cambia la posici�n de X,Y (x++) � x-- si hay inversi�n
            Rutina_F8_20F5 X, Y, PunteroRutinasBloque
        Case Is = "F9" 'pinta el tile indicado por X,Y con el siguiente byte leido y cambia la posici�n de X,Y (y--)
            Rutina_F9_20E7 X, Y, PunteroRutinasBloque
        Case Is = "FA" 'recupera la longitud y si no es 0, vuelve a saltar a procesar las instrucciones desde la direcci�n que se guard�. En otro caso, limpia la pila y contin�a
            Rutina_FA_20D7 PunteroRutinasBloque
        Case Is = "FB" 'recupera de la pila la posici�n almacenada en el buffer de tiles
            Rutina_FB_20D3 X, Y
        Case Is = "FC" 'guarda en la pila la posici�n actual en el buffer de tiles
            Rutina_FC_20CF X, Y
        Case Is = "FD" 'guarda en la pila la longitud del bloque en y? y la posici�n actual de los datos de construcci�n del bloque
            Rutina_FD_209E PunteroRutinasBloque
        Case Is = "FE" 'guarda en la pila la longitud del bloque en x? y la posici�n actual de los datos de construcci�n del bloque
            Rutina_FE_2091 PunteroRutinasBloque
        Case "FF" 'si se cambiaron las coordenadas x (x = -x), deshace el cambio
            Rutina_FF_2032
            TerminarEvaluacion = True
            Exit Sub 'recupera la direcci�n del siguiente bloque a procesar
    End Select
    If TerminarEvaluacion Then
        If Rutina = "EC" Or Rutina = "E4" Then
            TerminarEvaluacion = False
        Else
            Exit Sub
        End If
    End If
Loop
End Sub



Sub Rutina_FF_2032()
'si se cambiaron las coordenadas x (x = -x), marca para deshacer el cambio la siguiente vez que pase por aqu�
If VariablesBloques_1FCD(&H1FCE - &H1FCD) <> 0 Then
    VariablesBloques_1FCD(&H1FCE - &H1FCD) = 0 'borra el indicador, pero mantiene InvertirDireccionesGeneracionBloques a true hasta la siguiente vez
Else
    InvertirDireccionesGeneracionBloques = False
End If
End Sub

Sub Rutina_F7_2141(ByVal nX As Byte, ByRef PunteroRutinasBloque As Long)
'modifica la posici�n del buffer de construcci�n del bloque (indicada en el primer byte)
'con una expresi�n calculada (indicada por los siguientes de bytes)
'0x61 = 0x1fcf datos de materiales 1
'0x62 = 0x1fd0 datos de materiales 2
'0x63 = 0x1fd1 datos de materiales 3
'0x64 = 0x1fd2 datos de materiales 4
'0x65 = 0x1fd3 datos de materiales 5
'0x66 = 0x1fd4 datos de materiales 6
'0x67 = 0x1fd5 datos de materiales 7
'0x68 = 0x1fd6 datos de materiales 8
'0x69 = 0x1fd7 datos de materiales 9
'0x6a = 0x1fd8 datos de materiales 10
'0x6b = 0x1fd9 datos de materiales 11
'0x6c = 0x1fda datos de materiales 12
'0x6d = 0x1fdb longitud de elemento en x
'0x6e = 0x1fdc longitud de elemento en y
'0x6f = 0x1fdd 0xff � altura
'0x70 = 0x1fde posici�n x del bloque en la rejilla
'0x71 = 0x1fdf posici�n y del bloque en la rejilla
Dim Registro As Byte
Dim Valor As Byte
Dim PunteroRegistro As Long
Dim Resultado As Integer
Dim ValorAnterior
Dim PunteroRegistroGuardado As Long
Registro = TablaCaracteristicasMaterial_1693(PunteroRutinasBloque - &H1693)
LeerPosicionBufferConstruccionBloque_2214 PunteroRutinasBloque, PunteroRegistro  'lee una posici�n del buffer de construcci�n del bloque y guarda en PunteroRegistro la direcci�n accedida
PunteroRegistroGuardado = PunteroRegistro 'guarda la direcci�n del buffer obtenida en la rutina anterior
Valor = LeerPosicionBufferConstruccionBloque_2214(PunteroRutinasBloque, PunteroRegistro) ' valor inicial
Resultado = EvaluarExpresionContruccionBloque_2166(CInt(Valor), PunteroRutinasBloque, PunteroRegistro)
If Resultado < 0 Then Resultado = Resultado + 256
PunteroRegistro = PunteroRegistroGuardado 'recupera la direcci�n obtenida con el primer byte
If Registro >= &H70 Then 'si accede a la posici�n Y del bloque en la rejilla. por qu� con 0x70 no hace lo mismo?
    ValorAnterior = VariablesBloques_1FCD(PunteroRegistro - &H1FCD)
    If ValorAnterior = 0 Then Exit Sub
    If Resultado < 0 Or Resultado > 100 Then Resultado = 0 'ajusta el valor a grabar entre 0x00 y 0x64 (0 y 100). en otro caso lo pone a 0
End If
VariablesBloques_1FCD(PunteroRegistro - &H1FCD) = CByte(Resultado) 'actualiza el valor calculado
'nX = CByte(Resultado)
End Sub

Function LeerPosicionBufferConstruccionBloque_2214(ByRef PunteroRutinasBloque As Long, ByRef PunteroRegistro As Long) As Byte
'lee un byte de los datos de construcci�n del bloque, avanzando el puntero.
'Si ley� un dato del buffer de construcci�n del bloque,
'a la salida, PunteroRegistro apuntar� a dicho registro
'si el byte leido es < 0x60, es un valor y lo devuelve
'si el byte leido es 0x82, sale devolviendo el siguiente byte
'en otro caso, es una operaci�n de lectura de registro de las caracter�sticas del bloque
Dim Valor As Byte
Valor = TablaCaracteristicasMaterial_1693(PunteroRutinasBloque - &H1693) 'lee el byte actual e incrementa el puntero
PunteroRutinasBloque = PunteroRutinasBloque + 1
LeerPosicionBufferConstruccionBloque_2214 = LeerRegistroBufferConstruccionBloque_2219(Valor, PunteroRegistro, PunteroRutinasBloque)

End Function

Function LeerRegistroBufferConstruccionBloque_2219(ByVal Registro As Byte, ByRef PunteroRegistro As Long, ByRef PunteroRutinasBloque As Long) As Byte
'lee un byte de los datos de construcci�n del bloque, avanzando el puntero.
'Si ley� un dato del buffer de construcci�n del bloque,
'a la salida, PunteroRegistro apuntar� a dicho registro
'si el byte leido es < 0x60, es un valor y lo devuelve
'si el byte leido es 0x82, sale devolviendo el siguiente byte
'en otro caso, es una operaci�n de lectura de registro de las caracter�sticas del bloque
PunteroRegistro = &H1FCF 'apunta al buffer de datos sobre la textura
If Registro < &H60 Then
    LeerRegistroBufferConstruccionBloque_2219 = Registro
    Exit Function
Else
    If Registro = &H82 Then
        LeerRegistroBufferConstruccionBloque_2219 = TablaCaracteristicasMaterial_1693(PunteroRutinasBloque - &H1693) 'lee el byte actual e incrementa el puntero
        PunteroRutinasBloque = PunteroRutinasBloque + 1
        Exit Function
    Else
        If Registro >= &H70 And InvertirDireccionesGeneracionBloques Then Registro = Registro Xor &H1
        LeerRegistroBufferConstruccionBloque_2219 = VariablesBloques_1FCD(Registro - &H61 + 2) '0x61=�ndice en el buffer de construcci�n del bloque
        PunteroRegistro = PunteroRegistro + Registro - &H61

    End If
End If
End Function

Function EvaluarExpresionContruccionBloque_2166(ByVal Operando1 As Integer, ByRef PunteroRutinasBloque As Long, ByVal PunteroRegistro As Long) As Integer
'modifica c con sumas de valores o registros y cambios de signo
'leidos de los datos de la construcci�n del bloque
Dim Valor As Byte
Dim Operando2 As Integer
Valor = TablaCaracteristicasMaterial_1693(PunteroRutinasBloque - &H1693)
If Valor >= &HC8 Then
    EvaluarExpresionContruccionBloque_2166 = Operando1
    Exit Function
End If
If Valor = &H84 Then 'si es 0x84, avanza el puntero y niega el byte leido
    PunteroRutinasBloque = PunteroRutinasBloque + 1
    EvaluarExpresionContruccionBloque_2166 = EvaluarExpresionContruccionBloque_2166(-Operando1, PunteroRutinasBloque, PunteroRegistro)
    Exit Function
End If
'si llega aqu� es porque accede a un registro o es un valor inmediato
Operando2 = LeerPosicionBufferConstruccionBloque_2214(PunteroRutinasBloque, PunteroRegistro) 'obtiene el siguiente byte
If Operando2 >= 128 Then Operando2 = Operando2 - 256
EvaluarExpresionContruccionBloque_2166 = EvaluarExpresionContruccionBloque_2166(Operando1 + Operando2, PunteroRutinasBloque, PunteroRegistro)
End Function

Sub Rutina_E4_21AA(ByVal X As Byte, ByVal nX As Byte, ByVal Y As Byte, ByVal nY As Byte, ByRef PunteroRutinasBloque As Long)
'interpreta otro bloque sin modificar los valores de los tiles a usar, y cambiando el sentido de las x
VariablesBloques_1FCD(&H1FCE - &H1FCD) = 1
'InvertirDireccionesGeneracionBloques = True 'marca que se realiz� un cambio en las operaciones que trabajan con coordenadas x en los tiles
Rutina_EC_21B4 X, nX, Y, nY, PunteroRutinasBloque, False
'InvertirDireccionesGeneracionBloques = False
End Sub

Sub Rutina_E9_218D()
'cambia las instrucciones que actualizan la coordenada x de los tiles (incx -> decx)
InvertirDireccionesGeneracionBloques = True
End Sub

Sub Rutina_EA_21A1(ByVal X As Byte, ByVal nX As Byte, ByVal Y As Byte, ByVal nY As Byte, ByVal PunteroRutinasBloque As Long)
Dim AnteriorPunteroRutinasBloque As Long
AnteriorPunteroRutinasBloque = PunteroRutinasBloque
PunteroRutinasBloque = Leer16(TablaCaracteristicasMaterial_1693, PunteroRutinasBloque - &H1693)
EvaluarDatosBloque_201E X, nX, Y, nY, PunteroRutinasBloque
PunteroRutinasBloque = AnteriorPunteroRutinasBloque
End Sub

Sub Rutina_EC_21B4(ByVal X As Byte, ByVal nX As Byte, ByVal Y As Byte, ByVal nY As Byte, ByRef PunteroRutinasBloque As Long, ByVal ActualizarVariablesTiles As Boolean)
'interpreta otro bloque modificando (o n�) los valores de los tiles a usar
Dim InvertirDireccionesGeneracionBloquesAntiguo As Boolean
Dim PunteroCaracteristicasBloque As Long
Dim PunteroTilesBloque As Long
Dim PunteroRutinasBloqueAnterior As Long
Dim Valor As Long
Dim Altura As Byte

'On Error Resume Next

PunteroRutinasBloqueAnterior = PunteroRutinasBloque + 2 'direcci�n para continuar con el proceso
PunteroCaracteristicasBloque = Leer16(TablaCaracteristicasMaterial_1693, PunteroRutinasBloque - &H1693)

PunteroTilesBloque = Leer16(TablaCaracteristicasMaterial_1693, PunteroCaracteristicasBloque - &H1693)
PunteroRutinasBloque = PunteroCaracteristicasBloque + 2

InvertirDireccionesGeneracionBloquesAntiguo = InvertirDireccionesGeneracionBloques 'obtiene las instrucciones que se usan para tratar las x
Push CLng(X)
Push CLng(Y)
Push CLng(VariablesBloques_1FCD(&H1FDE - &H1FCD)) 'obtiene las posiciones en el sistema de coordenadas de la rejilla y los guarda en pila
Push CLng(VariablesBloques_1FCD(&H1FDF - &H1FCD)) 'obtiene las posiciones en el sistema de coordenadas de la rejilla y los guarda en pila
Push CLng(VariablesBloques_1FCD(&H1FDB - &H1FCD)) 'obtiene los par�metros para la construcci�n del bloque y los guarda en pila
nX = VariablesBloques_1FCD(&H1FDB - &H1FCD)
nY = VariablesBloques_1FCD(&H1FDC - &H1FCD)
Push CLng(VariablesBloques_1FCD(&H1FDC - &H1FCD))  'obtiene los par�metros para la construcci�n del bloque y los guarda en pila
Altura = VariablesBloques_1FCD(&H1FDD - &H1FCD)
Push CLng(Altura) 'obtiene el par�metro dependiente del byte 4 y lo guarda en pila
ConstruirBloque_1BBC X, nX, Y, nY, Altura, PunteroTilesBloque, PunteroRutinasBloque, ActualizarVariablesTiles
PunteroRutinasBloque = PunteroRutinasBloqueAnterior 'restaura la direcci�n de los datos de la rutina actual
VariablesBloques_1FCD(&H1FDD - &H1FCD) = CByte(Pop)
VariablesBloques_1FCD(&H1FDC - &H1FCD) = CByte(Pop)
VariablesBloques_1FCD(&H1FDB - &H1FCD) = CByte(Pop)
VariablesBloques_1FCD(&H1FDF - &H1FCD) = CByte(Pop)
VariablesBloques_1FCD(&H1FDE - &H1FCD) = CByte(Pop)
Y = CByte(Pop)
X = CByte(Pop)
InvertirDireccionesGeneracionBloques = InvertirDireccionesGeneracionBloquesAntiguo
End Sub

Sub Rutina_EF_2071(ByVal PunteroRutinasBloque As Long)
'incrementa la longitud del bloque en x
IncrementarRegistroConstruccionBloque_2087 &H6E, 1, PunteroRutinasBloque
End Sub

Sub Rutina_F0_2077(ByVal PunteroRutinasBloque As Long)
'incrementa la longitud del bloque en y
IncrementarRegistroConstruccionBloque_2087 &H6D, 1, PunteroRutinasBloque
End Sub

Sub Rutina_F1_2066(ByRef X As Byte, ByVal PunteroRutinasBloque As Long)
'modifica la posici�n en x con la expresi�n leida
Dim Valor As Byte
Dim Resultado As Integer
Dim PunteroRegistro As Long
Valor = LeerPosicionBufferConstruccionBloque_2214(PunteroRutinasBloque, PunteroRegistro) ' lee un valor inmediato o un registro
Resultado = EvaluarExpresionContruccionBloque_2166(CInt(Valor), PunteroRutinasBloque, PunteroRegistro)
X = CByte(X + Resultado)
End Sub

Sub Rutina_F2_205B(ByRef Y As Byte, ByVal PunteroRutinasBloque As Long)
'modifica la posici�n en y con la expresi�n leida
Dim Valor As Byte
Dim Resultado As Integer
Dim PunteroRegistro As Long
Valor = LeerPosicionBufferConstruccionBloque_2214(PunteroRutinasBloque, PunteroRegistro) ' lee un valor inmediato o un registro
Resultado = EvaluarExpresionContruccionBloque_2166(CInt(Valor), PunteroRutinasBloque, PunteroRegistro)
If CLng(Y + Resultado) >= 256 Then
    Y = CByte(Y + Resultado - 256)
Else
    Y = CByte(Y + Resultado)
End If
End Sub

Sub Rutina_F3_2058(ByRef X As Byte)
'cambia la posici�n de x (x--)
If Not InvertirDireccionesGeneracionBloques Then
    If X = 0 Then
        X = 255
    Else
        X = X - 1
    End If
Else
    X = X + 1
End If
End Sub

Sub Rutina_F4_2055(ByRef Y As Byte)
'cambia la posici�n de Y (y--)
If Y = 0 Then
    Y = 255
Else
    Y = Y - 1
End If
End Sub

Sub Rutina_F5_2052(ByRef X As Byte)
'cambia la posici�n de x (x++)
If Not InvertirDireccionesGeneracionBloques Then
    X = X + 1
Else
    X = X - 1
End If
End Sub

Sub Rutina_F6_204F(ByRef Y As Byte)
'cambia la posici�n de Y (y++)
If Y = 255 Then
    Y = 0
Else
    Y = Y + 1
End If
End Sub

Sub Rutina_F8_20F5(ByRef X As Byte, ByRef Y As Byte, ByVal PunteroRutinasBloque As Long)
'pinta el tile indicado por X,Y con el siguiente byte leido y cambia la posici�n de X,Y (x++) � x-- si hay inversi�n
If Not InvertirDireccionesGeneracionBloques Then
    PintarTileBuffer_20FC X, Y, IncSumarX, PunteroRutinasBloque
Else
    PintarTileBuffer_20FC X, Y, IncRestarX, PunteroRutinasBloque
End If
End Sub

Sub Rutina_F9_20E7(ByRef X As Byte, ByRef Y As Byte, ByRef PunteroRutinasBloque As Long)
'pinta el tile indicado por X,Y con el siguiente byte leido y cambia la posici�n de X,Y (y--)
PintarTileBuffer_20FC X, Y, IncRestarY, PunteroRutinasBloque
End Sub

Sub Rutina_FA_20D7(ByRef PunteroRutinasBloque As Long)
'recupera la longitud y si no es 0, vuelve a saltar a procesar las instrucciones desde la direcci�n que se guard�.
'En otro caso, limpia la pila y contin�a
Dim Longitud As Long
Dim PunteroRutinasBloquePila As Long
Longitud = Pop 'recupera de la pila la longitud del bloque (bien sea en x o en y)
Longitud = Longitud - 1 'decrementa la longitud
If Longitud = 0 Then 'si se ha terminado la longitud, saca el otro valor de la pila y salta
    Pop 'recupera la posici�n actual de los datos de construcci�n del bloque
Else 'en otro caso, recupera los datos de la secuencia, guarda la posici�n decrementada y vuelve a procesar el bloque
    PunteroRutinasBloque = Pop
    Push PunteroRutinasBloque
    Push Longitud
End If
End Sub

Sub Rutina_FB_20D3(ByRef X As Byte, ByRef Y As Byte)
'recupera de la pila la posici�n almacenada en el buffer de tiles
Y = Pop
X = Pop
End Sub

Sub Rutina_FC_20CF(ByVal X As Byte, ByVal Y As Byte)
'guarda en la pila la posici�n actual en el buffer de tiles
Push CLng(X)
Push CLng(Y)
End Sub

Sub Rutina_FE_2091(ByRef PunteroRutinasBloque As Long)
'guarda en la pila la longitud del bloque en x? y la posici�n actual de los datos de construcci�n del bloque
Dim Registro As Byte
Dim PunteroRegistro As Long
Registro = LeerRegistroBufferConstruccionBloque_2219(&H6D, PunteroRegistro, PunteroRutinasBloque)
If Registro <> 0 Then 'si es != 0, sigue procesando el material, en otro caso salta s�mbolos hasta que se acaben los datos de construcci�n
    Push PunteroRutinasBloque
    Push CLng(Registro)
    Exit Sub
End If
'si el bucle no se ejecuta, se salta los comandos intermedios
SaltarComandosIntermedios_20A5 PunteroRutinasBloque
End Sub

Sub Rutina_FD_209E(ByRef PunteroRutinasBloque As Long)
'guarda en la pila la longitud del bloque en y? y la posici�n actual de los datos de construcci�n del bloque
Dim Registro As Byte
Dim PunteroRegistro As Long
Registro = LeerRegistroBufferConstruccionBloque_2219(&H6E, PunteroRegistro, PunteroRutinasBloque)
If Registro <> 0 Then 'si es != 0, sigue procesando el material, en otro caso salta s�mbolos hasta que se acaben los datos de construcci�n
    Push PunteroRutinasBloque
    Push CLng(Registro)
    Exit Sub
End If
'si el bucle no se ejecuta, se salta los comandos intermedios
SaltarComandosIntermedios_20A5 PunteroRutinasBloque
End Sub

Sub IncrementarRegistroConstruccionBloque_2087(ByVal Registro As Byte, ByVal Incremento As Integer, ByVal PunteroRutinasBloque As Long)
'modifica el registro del buffer de construcci�n del bloque, sum�ndole el valor indicado
Dim PunteroRegistro As Long
LeerRegistroBufferConstruccionBloque_2219 Registro, PunteroRegistro, PunteroRutinasBloque
VariablesBloques_1FCD(Registro - &H61 + 2) = VariablesBloques_1FCD(Registro - &H61 + 2) + Incremento '0x61=�ndice en el buffer de construcci�n del bloque
End Sub

Sub SaltarComandosIntermedios_20A5(ByRef PunteroRutinasBloque As Long)
'si el bucle while no se ejecuta, se salta los comandos intermedios
Dim NBucles As Long 'contador de bucles
Dim Valor As Byte
NBucles = 1 'inicialmente estamos dentro de un while
Do
    Valor = TablaCaracteristicasMaterial_1693(PunteroRutinasBloque - &H1693)
    If Valor = &H82 Then 'si es 0x82 (marcador), avanza de 2 en 2
        PunteroRutinasBloque = PunteroRutinasBloque + 2
    Else 'en otro caso, de 1 en 1
        PunteroRutinasBloque = PunteroRutinasBloque + 1
    End If
    If Valor = &HFE Or Valor = &HFD Or Valor = &HE8 Or Valor = &HE7 Then 'si encuentra 0xfe y 0xfd (nuevo while) o 0xe8 y 0xe7 (parcheadas???), sigue avanzando
        NBucles = NBucles + 1
    Else 'sigue pasando hasta encontrar un fin while
        If Valor = &HFA Then NBucles = NBucles - 1
        If NBucles = 0 Then Exit Sub 'repite hasta que se llegue al fin del primer bucle
    End If
Loop
End Sub

Sub Push(ByVal Valor As Long)
Pila(PunteroPila) = Valor
PunteroPila = PunteroPila + 1
If PunteroPila > UBound(Pila) Then Stop
End Sub

Function Pop() As Long
If PunteroPila = 0 Then Stop
PunteroPila = PunteroPila - 1
Pop = Pila(PunteroPila)
Pila(PunteroPila) = 0
End Function





Sub PintarTileBuffer_20FC(ByRef X As Byte, ByRef Y As Byte, ByVal Incremento As EnumIncremento, ByRef PunteroRutinasBloqueIX As Long)
'lee un byte del buffer de construcci�n del bloque que indica el n�mero de tile, lee el siguiente byte y lo pinta en X,Y, modificando X,Y
'si el siguiente byte >= 0xc8, pinta y sale
'si el siguiente byte leido es 0x80 dibuja el tile en X,Y, actualiza las coordenadas y sigue procesando
'si el siguiente byte leido es 0x81, dibuja el tile en X,Y y sigue procesando
'si es otra cosa != 0x00, dibuja el tile en X,Y, actualiza las coordenadas las veces que haya leido, mira a ver si salta un byte y sale
'si es otra cosa = 0x00, mira a ver si salta un byte y sale
Dim Valor1 As Byte
Dim Valor2 As Byte
Dim PunteroRegistro As Long
Dim Nveces As Long
Valor1 = LeerPosicionBufferConstruccionBloque_2214(PunteroRutinasBloqueIX, PunteroRegistro) 'lee una posici�n del buffer de construcci�n del bloque o un operando
Valor2 = TablaCaracteristicasMaterial_1693(PunteroRutinasBloqueIX - &H1693) 'lee el siguiente byte de los datos de construcci�n
If Valor2 >= &HC8 Then 'si es >= 0xc8, pinta, cambia X,Y seg�n la operaci�n y saleX,Y es visible, y si es as�, actualiza el buffer de tiles
    PintarTileBuffer_1633 X, Y, Valor1, PunteroRutinasBloqueIX
    'If Not InvertirDireccionesGeneracionBloques Then
        AplicarIncrementoXY X, Y, Incremento
    'Else
    '    AplicarIncrementoXY X, Y, InvertirIncremento(Incremento)
    'End If
    Exit Sub
End If
PunteroRutinasBloqueIX = PunteroRutinasBloqueIX + 1
If Valor2 = &H80 Then 'dibuja el tile en X, Y, actualiza las coordenadas y sigue procesando
    PintarTileBuffer_1633 X, Y, Valor1, PunteroRutinasBloqueIX
    AplicarIncrementoXY X, Y, Incremento
    PintarTileBuffer_20FC X, Y, Incremento, PunteroRutinasBloqueIX
    Exit Sub
End If
If Valor2 = &H81 Then 'dibuja el tile en X, Y  y sigue procesando
    PintarTileBuffer_1633 X, Y, Valor1, PunteroRutinasBloqueIX
    PintarTileBuffer_20FC X, Y, Incremento, PunteroRutinasBloqueIX
    Exit Sub
End If
'aqu� llega si el byte leido no es 0x80 ni 0x81
Nveces = LeerPosicionBufferConstruccionBloque_2214(PunteroRutinasBloqueIX, PunteroRegistro) 'n�mero de veces que realizar la operaci�n
If Nveces > 0 Then
    Do 'si lo leido es != 0, pinta  y realiza la operaci�n nveces
        PintarTileBuffer_1633 X, Y, Valor1, PunteroRutinasBloqueIX
        AplicarIncrementoXY X, Y, Incremento
        Nveces = Nveces - 1
    Loop While Nveces > 0
End If
Valor2 = TablaCaracteristicasMaterial_1693(PunteroRutinasBloqueIX - &H1693) 'lee el siguiente byte de los datos de construcci�n
If Valor2 >= &HC8 Then Exit Sub
PunteroRutinasBloqueIX = PunteroRutinasBloqueIX + 1
PintarTileBuffer_20FC X, Y, Incremento, PunteroRutinasBloqueIX 'salta y sigue procesando
End Sub

Sub AplicarIncrementoXY(ByRef X As Byte, ByRef Y As Byte, ByVal Incremento As EnumIncremento)
'cambia X,Y seg�n la operaci�n
Select Case Incremento
    Case Is = IncSumarX
        If X = 255 Then
            X = 0
        Else
            X = X + 1
        End If
    Case Is = IncRestarX
        If X = 0 Then
            X = 255
        Else
            X = X - 1
        End If
    Case Is = IncRestarY
        If Y = 0 Then
            Y = 255
        Else
            Y = Y - 1
        End If
End Select
End Sub

Sub AplicarIncrementoReversibleXY(ByRef X As Byte, ByRef Y As Byte, ByVal Incremento As EnumIncremento)
'cambia X,Y seg�n la operaci�n, pero invierte la direcci�n si InvertirDireccionesGeneracionBloques=true
If (Incremento = IncSumarX And Not InvertirDireccionesGeneracionBloques) Or (Incremento = IncRestarX And InvertirDireccionesGeneracionBloques) Then
    If X = 255 Then
        X = 0
    Else
        X = X + 1
    End If
    Exit Sub
End If
If (Incremento = IncRestarX And Not InvertirDireccionesGeneracionBloques) Or (Incremento = IncSumarX And InvertirDireccionesGeneracionBloques) Then
    If X = 0 Then
        X = 255
    Else
        X = X - 1
    End If
End If
If Incremento = IncRestarY Then
    If Y = 0 Then
        Y = 255
    Else
        Y = Y - 1
    End If
End If
End Sub

Function InvertirIncremento(ByRef Incremento As EnumIncremento) As EnumIncremento
'devuelve la operaci�n contraria
If Incremento = IncRestarX Then
    InvertirIncremento = IncSumarX
ElseIf Incremento = IncSumarX Then
    InvertirIncremento = IncRestarX
Else
    InvertirIncremento = Incremento 'Y no se ve afectada por la inversi�n
End If
End Function

Sub PintarTileBuffer_1633(ByVal X As Byte, ByVal Y As Byte, ByVal Tile As Byte, ByVal PunteroRutinasBloqueIX As Long)
'comprueba si el tile indicado por X,Y es visible, y si es as�, actualiza el tile mostrado en esta posici�n y los datos de profundidad asociados
'Y = pos en y usando el sistema de coordenadas del buffer de tiles
'X = pos en x usando el sistema de coordenadas del buffer de tiles
'Tile = n�mero de tile a poner
'PunteroRutinasBloque = puntero a los datos de construcci�n del bloque
Dim Xr As Long 'coordenadas de la rejilla
Dim Yr As Long
Dim PunteroBufferTiles As Long
Dim ProfundidadAnteriorX As Byte
Dim ProfundidadAnteriorY As Byte
Dim TileAnterior As Byte
Dim ProfundidadNuevaX As Byte
Dim ProfundidadNuevaY As Byte
'el buffer de tiles es de 16x20, aunque la rejilla es de 32x36. La parte de la rejilla que se mapea en el buffer de tiles es la central
'(quitandole 8 unidades a la izquierda, derecha arriba y abajo)
Yr = Y - 8 'traslada la posici�n y 8 unidades hacia arriba para tener la coordenada en el origen
If Yr >= 20 Or Yr < 0 Then Exit Sub
Xr = X - 8
If Xr >= 16 Or Xr < 0 Then Exit Sub
'1641
PunteroBufferTiles = 96 * Yr + 6 * Xr
'graba los datos del tile que hay en PunteroBufferTiles, seg�n lo que valgan las coordenadas de profundidad actual y Tile (tile a escribir)
'si ya se hab�a proyectado un tile antes, el nuevo tiene mayor prioridad sobre el viejo
'PunteroBufferTiles = puntero a los datos del tile actual en el buffer de tiles
'Tile = n�mero de tile a poner
'166e
ProfundidadAnteriorX = BufferTiles_8D80(PunteroBufferTiles + 3)
ProfundidadAnteriorY = BufferTiles_8D80(PunteroBufferTiles + 4)
TileAnterior = BufferTiles_8D80(PunteroBufferTiles + 5) 'tile anterior con mayor prioridad
If Pintar Then BufferTiles_8D80(PunteroBufferTiles + 2) = TileAnterior '(el tile anterior pasa a tener ahora menor prioridad)
ProfundidadNuevaX = VariablesBloques_1FCD(&H1FDE - &H1FCD)
If Pintar Then BufferTiles_8D80(PunteroBufferTiles + 3) = ProfundidadNuevaX 'nueva profundidad del tile en la rejilla (sistema de coordenadas local de la rejilla)
ProfundidadNuevaY = VariablesBloques_1FCD(&H1FDF - &H1FCD)
If ProfundidadNuevaX < ProfundidadAnteriorX Then
    If ProfundidadNuevaY < ProfundidadAnteriorY Then
        ProfundidadAnteriorX = ProfundidadNuevaX
        ProfundidadAnteriorY = ProfundidadNuevaY
    End If
End If
'1689
If Pintar Then BufferTiles_8D80(PunteroBufferTiles + 4) = ProfundidadNuevaY 'nueva profundidad del tile en la rejilla (sistema de coordenadas local de la rejilla)
If Pintar Then BufferTiles_8D80(PunteroBufferTiles + 0) = ProfundidadAnteriorX 'vieja profundidad en X, modificado por anterior
If Pintar Then BufferTiles_8D80(PunteroBufferTiles + 1) = ProfundidadAnteriorY 'vieja profundidad en y, modificado por anterior
If Pintar Then BufferTiles_8D80(PunteroBufferTiles + 5) = Tile
End Sub

Public Sub GenerarTablasAndOr_3AD1()
'genera 4 tablas de 0x100 bytes para el manejo de pixels mediante operaciones AND y OR
'TablasAndOr
Dim Puntero As Long
Dim Contador As Long
Dim a As Long
Dim b As Long
Dim c As Long
Dim d As Long
Dim e As Long
For Contador = 0 To 255
    a = Contador            'ld   a,c      ; a = b7 b6 b5 b4 b3 b2 b1 b0
    a = a And &HF0&         'and  $F0      ; a = b7 b6 b5 b4 0 0 0 0
    d = a                   'ld   d,a      ; d = b7 b6 b5 b4 0 0 0 0
    a = Contador            'ld   a,c      ; a = b7 b6 b5 b4 b3 b2 b1 b0
    a = ror8(a, 1)          'rrca          ; a = b0 b7 b6 b5 b4 b3 b2 b1
    a = ror8(a, 1)          'rrca          ; a = b1 b0 b7 b6 b5 b4 b3 b2
    a = ror8(a, 1)          'rrca          ; a = b2 b1 b0 b7 b6 b5 b4 b3
    a = ror8(a, 1)          'rrca          ; a = b3 b2 b1 b0 b7 b6 b5 b4
    e = a                   'ld   e,a      ; e = b3 b2 b1 b0 b7 b6 b5 b4
    a = a And Contador      'and  c        ; a = b3&b7 b2&b6 b1&b5 b0&b4 b3&b7 b2&b6 b1&b5 b0&b4
    a = a And &HF&          'and  $0F      ; a = 0 0 0 0 b3&b7 b2&b6 b1&b5 b0&b4
    a = a Or d              'or   d        ; a = b7 b6 b5 b4 b3&b7 b2&b6 b1&b5 b0&b4
    TablasAndOr_9D00(Puntero) = a 'ld   (bc),a   ; graba pixel i = (Pi1&Pi0 Pi0) (0->0, 1->1, 2->0, 3->3)

    Puntero = Puntero + 256 'inc  b        ; apunta a la siguiente tabla
    a = e                   'ld   a,e      ; a = b3 b2 b1 b0 b7 b6 b5 b4
    a = a Xor Contador      'xor  c        ; a = b3^b7 b2^b6 b1^b5 b0^b4 b3^b7 b2^b6 b1^b5 b0^b4
    a = a And Contador      'and  c        ; a = (b3^b7)&b7 (b2^b6)&b6 (b1^b5)&b5 (b0^b4)&b4 (b3^b7)&b3 (b2^b6)&b2 (b1^b5)&b1 (b0^b4)&b0
    a = a And &HF&          'and  $0F      ; a = 0 0 0 0 (b3^b7)&b3 (b2^b6)&b2 (b1^b5)&b1 (b0^b4)&b0
    d = a                   'ld   d,a      ; d = 0 0 0 0 (b3^b7)&b3 (b2^b6)&b2 (b1^b5)&b1 (b0^b4)&b0
    a = shl(a, 1)           'add  a,a      ; a = 0 0 0 (b3^b7)&b3 (b2^b6)&b2 (b1^b5)&b1 (b0^b4)&b0 0
    a = shl(a, 1)           'add  a,a      ; a = 0 0 (b3^b7)&b3 (b2^b6)&b2 (b1^b5)&b1 (b0^b4)&b0 0 0
    a = shl(a, 1)           'add  a,a      ; a = 0 (b3^b7)&b3 (b2^b6)&b2 (b1^b5)&b1 (b0^b4)&b0 0 0 0
    a = shl(a, 1)           'add  a,a      ; a = (b3^b7)&b3 (b2^b6)&b2 (b1^b5)&b1 (b0^b4)&b0 0 0 0 0
    a = a Or d              'or   d        ; a = (b3^b7)&b3 (b2^b6)&b2 (b1^b5)&b1 (b0^b4)&b0 (b7^b3)&b3 (b6^b2)&b2 (b5^b1)&b1 (b0^b4)&b0
    TablasAndOr_9D00(Puntero) = a 'ld   (bc),a   ; graba pixel i = ((Pi1^Pi0)&Pi1 (Pi1^Pi0)&Pi1) (0->0, 1->0, 2->3, 3->0)

    Puntero = Puntero + 256 'inc  b        ; apunta a la siguiente tabla
    a = Contador            'ld   a,c      ; a = b7 b6 b5 b4 b3 b2 b1 b0
    a = a And &HF&          'and  $0F      ; a = 0 0 0 0 b3 b2 b1 b0
    d = a                   'ld   d,a      ; d = 0 0 0 0 b3 b2 b1 b0
    a = e                   'ld   a,e      ; a = b3 b2 b1 b0 b7 b6 b5 b4
    a = a And Contador      'and  c        ; a = b3&b7 b2&b6 b1&b5 b0&b4 b3&b7 b2&b6 b1&b5 b0&b4
    a = a And &HF0&         'and  $F0      ; a = b3&b7 b2&b6 b1&b5 b0&b4 0 0 0 0
    a = a Or d              'or   d        ; a = b3&b7 b2&b6 b1&b5 b0&b4 b3 b2 b1 b0
    TablasAndOr_9D00(Puntero) = a 'ld   (bc),a   ; graba pixel i = (Pi1 Pi1&Pi0) (0->0, 1->0, 2->2, 3->3)

    Puntero = Puntero + 256 'inc  b        ; apunta a la siguiente tabla
    a = e                   'ld   a,e      ; a = b3 b2 b1 b0 b7 b6 b5 b4
    a = a Xor Contador      'xor  c        ; a = b3^b7 b2^b6 b1^b5 b0^b4 b7^b3 b6^b2 b5^b1 b4^b0
    a = a And Contador      'and  c        ; a = (b3^b7)&b7 (b2^b6)&b6 (b1^b5)&b5 (b0^b4)&b4 (b7^b3)&b3 (b6^b2)&b2 (b5^b1)&b1 (b4^b0)&b0
    a = a And &HF0&         'and  $F0      ; a = (b3^b7)&b7 (b2^b6)&b6 (b1^b5)&b5 (b0^b4)&b4 0 0 0 0
    d = a                   'ld   d,a      ; d = (b3^b7)&b7 (b2^b6)&b6 (b1^b5)&b5 (b0^b4)&b4 0 0 0 0
    a = shr(a, 1)           'srl  a        ; a = 0 (b3^b7)&b7 (b2^b6)&b6 (b1^b5)&b5 (b0^b4)&b4 0 0 0
    a = shr(a, 1)           'srl  a        ; a = 0 0 (b3^b7)&b7 (b2^b6)&b6 (b1^b5)&b5 (b0^b4)&b4 0 0
    a = shr(a, 1)           'srl  a        ; a = 0 0 0 (b3^b7)&b7 (b2^b6)&b6 (b1^b5)&b5 (b0^b4)&b4 0
    a = shr(a, 1)           'srl  a        ; a = 0 0 0 0 (b3^b7)&b7 (b2^b6)&b6 (b1^b5)&b5 (b0^b4)&b4
    a = a Or d 'or   d        ; a = (b3^b7)&b7 (b2^b6)&b6 (b1^b5)&b5 (b0^b4)&b4 (b3^b7)&b7 (b2^b6)&b6 (b1^b5)&b5 (b0^b4)&b4
    TablasAndOr_9D00(Puntero) = a 'ld   (bc),a   ; graba pixel i = ((Pi1^Pi0)&Pi0 (Pi1^Pi0)&Pi0) (0->0, 1->3, 2->0, 3->0)
    Puntero = Puntero - 767 '; apunta a la tabla inicial
Next
End Sub

Sub DeshabilitarInterrupcion()

End Sub

Sub HabilitarInterrupcion()

End Sub

Sub ColocarVectorInterrupcion()

End Sub

Public Sub InicializarJuego_249A()
    Depuracion.Init
    'inicio real del programa
    Static Inicializado_00FE As Boolean
    DeshabilitarInterrupcion
    CargarDatos
    InicializarVariablesROM
    If Not Inicializado_00FE Or Check = True Then 'comprueba si es la primera vez que llega aqu�
        Inicializado_00FE = True
        modPantalla.DefinirModo 1 'fija el modo 0 (320x200 4 colores)
        modPantalla.SeleccionarPaleta 0 'pone una paleta de colores negra
        'InicializarInterrupcion 'coloca el c�digo a ejecutar al producirse una interrupci�n ###pendiente
        'InicializarTablaSonido_103F() ' inicializa la tabla del sonido y habilita las interrupciones ###pendiente
        DeshabilitarInterrupcion
        If Not Depuracion.SaltarPergamino Then DibujarPergaminoIntroduccion_659D 'dibuja el Pergamino y cuenta la introducci�n. De aqu� vuelve al pulsar espacio
        DeshabilitarInterrupcion
        'ApagarSonido_1376 'apaga el sonido '###pendiente
        modPantalla.SeleccionarPaleta 2  'pone los colores de la paleta a negro
        Limpiar40LineasInferioresPantalla_2712
        CopiarVariables_37B6 'copia cosas de muchos sitios en 0x0103-0x01a9 (pq??z)
        RellenarTablaFlipX_3A61
        CerrarEspejo_3A7E
        GenerarTablasAndOr_3AD1
        InicializarPartida_2509
        
    End If
End Sub

Sub Limpiar40LineasInferioresPantalla_2712()
Dim Banco As Long
Dim PunteroPantalla As Long
Dim Contador As Long
PunteroPantalla = &H640 'apunta a memoria de video
For Banco = 1 To 8 'repite el proceso para 8 bancos
    For Contador = 0 To &H18F '5 l�neas
        PantallaCGA(PunteroPantalla + Contador) = &HFF
        PantallaCGA2PC PunteroPantalla + Contador, &HFF
    Next
    PunteroPantalla = PunteroPantalla + &H800 'apunta al siguiente banco
Next


End Sub
Sub DibujarPergaminoIntroduccion_659D()
'dibuja el pergamino
modPantalla.SeleccionarPaleta 1 'coloca la paleta negra
DibujarPergamino_65AF 'dibuja el pergamino
modPantalla.SeleccionarPaleta 1 'coloca la paleta del pergamino
DibujarTextosPergamino_6725 'dibuja los textos en el Pergamino mientras no se pulse el espacio







End Sub


Sub DibujarPergamino_65AF()
Dim Contador As Long
Dim Linea As Long
Dim Relleno As Long
Dim Puntero As Long
For Contador = 0 To &H3FFF& 'limpia la memoria de video
    PantallaCGA(Contador) = 0
Next
modPantalla.DibujarRectanguloCGA 0, 0, 319, 199, 0

'deja un rect�ngulo de 192 pixels de ancho en el medio de la pantalla, el resto limpio
Contador = 0
For Linea = 1 To 200 'n�mero de l�neas a rellenar
    For Relleno = 0 To 15 '16, ancho de los rellenos
        '&HF0=240, valor con el que rellenar
        PantallaCGA(Contador + Relleno) = &HF0 'apunta al relleno por la izquierda
        PantallaCGA2PC Contador + Relleno, &HF0
        PantallaCGA(Contador + &H40 + Relleno) = &HF0 'apunta al relleno por la derecha. &H40=64, salto entre rellenos
        PantallaCGA2PC Contador + Relleno + &H40, &HF0
    Next 'completa 16 bytes (64 pixels)
    Contador = DireccionSiguienteLinea_3A4D_68F2(Contador) 'pasa a la siguiente l�nea de pantalla
Next 'repite para 200 lineas
modPantalla.Refrescar
'limpia las 8 l�neas de debajo de la pantalla
Contador = &H780  'apunta a una l�nea (la octava empezando por abajo)
For Linea = 0 To 7 'repetir para 8 l�neas
    For Relleno = 1 To &H4F
        PantallaCGA(Contador + Relleno) = PantallaCGA(Contador) 'copia lo que hay en la primera posici�n de la l�nea para el resto de pixels de la l�nea
        PantallaCGA2PC Contador + Relleno, PantallaCGA(Contador)
    Next
    Contador = DireccionSiguienteLinea_3A4D_68F2(Contador) 'avanza hl 0x0800 bytes y si llega al final, pasa a la siguiente l�nea (+0x50)
Next
PunteroPantalla = CalcularDesplazamientoPantalla_68C7(32, 0) 'calcula el desplazamiento en pantalla
DibujarParteSuperiorInferiorPergamino_661B PunteroPantalla, &H788A - &H788A 'dibuja la parte superior del pergamino
PunteroPantalla = CalcularDesplazamientoPantalla_68C7(218, 0) 'calcula el desplazamiento en pantalla
DibujarParteDerechaIzquierdaPergamino_662E PunteroPantalla, &H7A0A - &H788A 'dibuja la parte derecha del pergamino
PunteroPantalla = CalcularDesplazamientoPantalla_68C7(32, 0) 'calcula el desplazamiento en pantalla
DibujarParteDerechaIzquierdaPergamino_662E PunteroPantalla, &H7B8A - &H788A 'dibuja la parte derecha del pergamino
PunteroPantalla = CalcularDesplazamientoPantalla_68C7(32, 184) 'calcula el desplazamiento en pantalla
DibujarParteSuperiorInferiorPergamino_661B PunteroPantalla, &H7D0A - &H788A 'dibuja la parte superior del pergamino

End Sub

Function CalcularDesplazamientoPantalla_68C7(ByVal X As Long, ByVal Y As Long) As Long
'dados X,Y (coordenadas en pixels), calcula el desplazamiento correspondiente en pantalla
'el valor calculado se hace partiendo de la coordenada x multiplo de 4 m�s cercana y sumandole 32 pixels a la derecha
Dim Valor1 As Long
Dim Valor2 As Long
Dim Valor3 As Long
Valor1 = modFunciones.shr(X, 2) 'l / 4 (cada 4 pixels = 1 byte)
Valor2 = Y And &HF8& 'obtiene el valor para calcular el desplazamiento dentro del banco de VRAM
Valor2 = Valor2 * 10 'dentro de cada banco, la l�nea a la que se quiera ir puede calcularse como (y & 0xf8)*10
Valor3 = Y And 7 '3 bits menos significativos en y (para calcular al banco de VRAM al que va)
Valor3 = shl(Valor3, 3)
Valor3 = shl(Valor3, 8) Or Valor2 'completa el c�lculo del banco
Valor3 = Valor3 + Valor1 'suma el desplazamiento en x
CalcularDesplazamientoPantalla_68C7 = Valor3 + 8 'ajusta para que salga 32 pixels m�s a la derecha
End Function

Sub DibujarParteSuperiorInferiorPergamino_661B(ByVal PunteroPantalla As Long, ByVal PunteroDatos As Long)
'rellena la parte superior (o inferior del pergamino)
Dim Linea As Long
Dim Contador As Long
Dim PunteroPantallaAnterior As Long
PunteroPantallaAnterior = PunteroPantalla
For Contador = 1 To 48 '48 bytes (= 192 pixels a rellenar)
    For Linea = 0 To 7 '8 l�neas de alto
        PantallaCGA(PunteroPantalla) = DatosGraficosPergamino_788A(PunteroDatos + Linea)
        PantallaCGA2PC PunteroPantalla, DatosGraficosPergamino_788A(PunteroDatos + Linea)
        PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla)
    Next
    PunteroDatos = PunteroDatos + Linea
    PunteroPantalla = PunteroPantallaAnterior + Contador
Next
End Sub

Sub DibujarParteDerechaIzquierdaPergamino_662E(ByVal PunteroPantalla As Long, ByVal PunteroDatos As Long)
'rellena la parte superior (o inferior del pergamino)
Dim Linea As Long
Dim Contador As Long
For Contador = 1 To 192 '192 l�neas
    PantallaCGA(PunteroPantalla) = DatosGraficosPergamino_788A(PunteroDatos)
    PantallaCGA2PC PunteroPantalla, DatosGraficosPergamino_788A(PunteroDatos)
    PantallaCGA(PunteroPantalla + 1) = DatosGraficosPergamino_788A(PunteroDatos + 1)
    PantallaCGA2PC PunteroPantalla + 1, DatosGraficosPergamino_788A(PunteroDatos + 1)
    PunteroDatos = PunteroDatos + 2
    PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla)
Next
End Sub


Sub DibujarTextosPergamino_6725()
'dibuja los textos en el Pergamino mientras no se pulse el espacio
Dim PunteroDatosPergamino As Long
Dim Caracter As Byte 'caracter a imprimir
Dim PosicionPergaminoY_680A As Long
Dim PosicionPergaminoX_680B As Long
Dim ColorLetra_67C0 As Byte
Dim PunteroCaracter As Long
PosicionPergaminoY_680A = 16
PosicionPergaminoX_680B = 44
Do
    'LeerEstadoTeclas_32BC ###pendiente 'lee el estado de las teclas
    If TeclaPulsadaNivel_3482(&H2F) Then Exit Sub '###pendiente 'comprueba si se puls� el espacio
    Caracter = TextoPergaminoPresentacion_7300(PunteroDatosPergamino) 'lee el caracter a imprimir
    'si ha encontrado el car�cter de fin de pergamino (&H1A), espera a que se pulse espacio para terminar
    If Caracter <> &H1A Then
        PunteroDatosPergamino = PunteroDatosPergamino + 1 'apunta al siguiente car�cter
        Select Case Caracter
            Case Is = &HD 'salto de l�nea
                ImprimirRetornoCarroPergamino_67DE PosicionPergaminoX_680B, PosicionPergaminoY_680A
            Case Is = &H20 'espacio
                ImprimirEspacioPergamino_67CD &HA, PosicionPergaminoX_680B 'espera un poco y avanza la posici�n en 10 pixels
            Case Is = &HA  'avanzar una p�gina. dibuja el tri�ngulo
                PasarPaginaPergamino_67F0 PosicionPergaminoX_680B, PosicionPergaminoY_680A
            Case Else
                If (Caracter And &H60) = &H40 Then
                    ColorLetra_67C0 = &HFF 'may�sculas en rojo
                Else
                    ColorLetra_67C0 = &HF 'min�sculas en negro
                End If
                PunteroCaracter = Caracter - &H20 'solo tiene caracteres a partir de 0x20
                PunteroCaracter = 2 * PunteroCaracter 'cada entrada ocupa 2 bytes
                PunteroCaracter = PunterosCaracteresPergamino_680C(PunteroCaracter) + 256 * PunterosCaracteresPergamino_680C(PunteroCaracter + 1)
                DibujarCaracterPergamino_6781 Caracter, PosicionPergaminoX_680B, PosicionPergaminoY_680A, PunteroCaracter, ColorLetra_67C0
        End Select
    End If
    DoEvents
Loop
End Sub

Sub Retardo_67C6(ByVal Ciclos As Long)
'retardo hasta que Ciclos = 0x0000. Cada iteraci�n son 32 ciclos (aprox 10 microsegundos, puesto
'que aunque el Z80 funciona a 4 MHz, la arquitectura del CPC tiene una sincronizaci�n para los
'el video que hace que funcione de forma efectiva entorno a los 3.2 MHz)
Dim Milisegundos As Long
'Do
'    Ciclos = Ciclos - 1
'    DoEvents
'Loop While Ciclos > 0
Milisegundos = 0.01 * Ciclos
frmPrincipal.Retardo Milisegundos
End Sub

Sub ImprimirRetornoCarroPergamino_67DE(ByRef X As Long, ByRef Y As Long)
Retardo_67C6 &HEA60& 'espera un rato (aprox. 600 ms)
'calcula la posici�n de la siguiente l�nea
X = &H2C
Y = Y + &H10
If Y > &HA4 Then PasarPaginaPergamino_67F0 X, Y 'se ha llegado a fin de hoja?
End Sub

Sub DibujarTrianguloRectanguloPergamino_6906_nousar(ByVal PixelX As Long, ByVal PixelY As Long, ByVal Lado As Long)
'dibuja un tri�ngulo rect�ngulo con los catetos paralelos a los ejes de coordenadas y de longitud Lado
Dim PunteroPantalla As Long
Dim RellenoTriangular_6943(3) As Byte
Dim Relleno As Long
Dim d As Byte
Dim e As Byte
Dim a As Byte
Dim b As Byte
Dim LadoAnterior As Long
Dim c As Byte
Dim c2 As Byte
Dim b2 As Byte
Dim PunteroRelleno As Long
Dim X2 As Byte
Dim Y2 As Byte
Dim Valor_6932 As Byte
Dim PunteroPantallaAnterior As Long
RellenoTriangular_6943(0) = &HF0
RellenoTriangular_6943(1) = &HE0
RellenoTriangular_6943(2) = &HC0
RellenoTriangular_6943(3) = &H80
PunteroPantalla = CalcularDesplazamientoPantalla_68C7(PixelX, PixelY)
b = Lado
d = 0
H690B:
c = 4
H690D:
c2 = c
b2 = b
X2 = PixelX
Y2 = PixelY
PunteroPantallaAnterior = PunteroPantalla
a = c
a = a - 1
PunteroRelleno = a
a = RellenoTriangular_6943(PunteroRelleno)
e = 0
Valor_6932 = a
H691F:
a = d
If a = e Then GoTo H6931
PantallaCGA(PunteroPantalla) = &HF0
PantallaCGA2PC PunteroPantalla, &HF0
b = b - 1
If b <> 0 Then GoTo H692D
e = e + 1
PunteroPantalla = PunteroPantalla + 1
b = b + 1
GoTo H691F

H692D:
b = b + 1
PunteroPantalla = PunteroPantalla + a

H6931:
PantallaCGA(PunteroPantalla) = Valor_6932
PantallaCGA2PC PunteroPantalla, Valor_6932
PunteroPantalla = PunteroPantalla + 1
PantallaCGA(PunteroPantalla) = 0
PantallaCGA2PC PunteroPantalla, 0
PunteroPantalla = PunteroPantallaAnterior
PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla)
c = c2
b = b2
c = c - 1
If c <> 0 Then GoTo H690D
e = e + 1
d = d + 1
b = b - 1
If b <> 0 Then GoTo H690B
End Sub

Sub DibujarTrianguloRectanguloPergamino_6906(ByVal PixelX As Long, ByVal PixelY As Long, ByVal Lado As Long)
'dibuja un tri�ngulo rect�ngulo con los catetos paralelos a los ejes de coordenadas y de longitud Lado
Dim PunteroPantalla As Long
Dim RellenoTriangular_6943(3) As Byte
Dim Relleno As Long
'Dim d As Long
Dim Aux As Long
Dim Distancia As Long 'separaci�n en bytes entre la parte derecha y la izquierda del tri�ngulo
Dim ContadorLado As Long
Dim Linea As Long
Dim PunteroRelleno As Long
Dim Valor_6932 As Byte
Dim PunteroPantallaAnterior As Long
RellenoTriangular_6943(0) = &HF0
RellenoTriangular_6943(1) = &HE0
RellenoTriangular_6943(2) = &HC0
RellenoTriangular_6943(3) = &H80
PunteroPantalla = CalcularDesplazamientoPantalla_68C7(PixelX, PixelY)
Distancia = 0
For ContadorLado = Lado To 1 Step -1
    For Linea = 4 To 1 Step -1
        Aux = 0
        PunteroPantallaAnterior = PunteroPantalla
        PunteroRelleno = Linea - 1
        Valor_6932 = RellenoTriangular_6943(PunteroRelleno)
        Do
            If Distancia = Aux Then
                DibujarTrianguloRectanguloPergamino_Parte2 Valor_6932, PunteroPantalla, PunteroPantallaAnterior, 0
                Exit Do
            Else
                PantallaCGA(PunteroPantalla) = &HF0
                PantallaCGA2PC PunteroPantalla, &HF0
                If ContadorLado > 1 Then
                    DibujarTrianguloRectanguloPergamino_Parte2 Valor_6932, PunteroPantalla, PunteroPantallaAnterior, Distancia
                    Exit Do
                End If
                Aux = Aux + 1
                PunteroPantalla = PunteroPantalla + 1
            End If
        Loop
    Next
    Aux = Aux + 1
    Distancia = Distancia + 1
Next
End Sub

Sub DibujarTrianguloRectanguloPergamino_Parte2(ByVal Valor As Byte, ByRef PunteroPantalla As Long, ByVal PunteroPantallaAnterior As Long, ByVal Incremento As Long)
PunteroPantalla = PunteroPantalla + Incremento
PantallaCGA(PunteroPantalla) = Valor
PantallaCGA2PC PunteroPantalla, Valor
PunteroPantalla = PunteroPantalla + 1
PantallaCGA(PunteroPantalla) = 0
PantallaCGA2PC PunteroPantalla, 0
PunteroPantalla = PunteroPantallaAnterior
PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla)
End Sub

Sub PasarPaginaPergamino_67F0(ByRef X As Long, ByRef Y As Long)
X = &H2C 'reinicia la posici�n al principio de la l�nea
Y = &H10 'reinicia la posici�n al principio de la l�nea
Retardo_67C6 3 * 65536 '(aprox. 655 ms), repite 3 veces los retardos
PasarPaginaPergamino_6697 'pasa de hoja
End Sub

Sub PasarPaginaPergamino_6697()
Dim Linea As Long
Dim X As Long
Dim Y As Long
Dim Tama�oTriangulo As Long
Dim PunteroPantalla As Long
Dim PunteroDatos As Long
Dim Contador As Long
Dim PunteroPantallaAnterior As Long
X = 211 - 4 * Linea '(00, 211) -> posici�n de inicio
Y = 0
Tama�oTriangulo = 3
For Linea = 0 To &H2C 'repite para 45 l�neas
    DibujarTrianguloRectanguloPergamino_6906 X, Y, Tama�oTriangulo 'dibuja un tri�ngulo rect�ngulo de lado Tama�oTriangulo
    If Linea Mod 2 = 0 Then modPantalla.Refrescar
    Retardo_67C6 &H7D0 'peque�o retardo (20 ms)
    LimpiarParteSuperiorDerechaPergamino_663E X, Y, Tama�oTriangulo
    X = X - 4
    Tama�oTriangulo = Tama�oTriangulo + 1
Next
LimpiarParteSuperiorDerechaPergamino_663E X, Y, Tama�oTriangulo
X = 32 '(32, 4) -> posici�n de inicio
Y = 4
Tama�oTriangulo = &H2F
For Contador = 0 To &H2D 'repite 46 veces
    DibujarTrianguloRectanguloPergamino_6906 X, Y, Tama�oTriangulo 'dibuja un tri�ngulo rect�ngulo de lado Tama�oTriangulo
    If Contador Mod 2 = 0 Then modPantalla.Refrescar
    Retardo_67C6 &H7D0 'peque�o retardo (20 ms)
    PunteroPantalla = CalcularDesplazamientoPantalla_68C7(X, Y) ' - 4)
    PunteroDatos = 2 * Y + &H7B8A - &H788A 'desplazamiento de los datos borrados de la parte izquierda del pergamino
    For Linea = 0 To 3 '4 l�neas de alto
        PantallaCGA(PunteroPantalla) = DatosGraficosPergamino_788A(PunteroDatos + 2 * Linea)
        PantallaCGA2PC PunteroPantalla, DatosGraficosPergamino_788A(PunteroDatos + 2 * Linea)
        PantallaCGA(PunteroPantalla + 1) = DatosGraficosPergamino_788A(PunteroDatos + 2 * Linea + 1)
        PantallaCGA2PC PunteroPantalla + 1, DatosGraficosPergamino_788A(PunteroDatos + 2 * Linea + 1)
        PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla)
    Next
    LimpiarParteInferiorPergamino_6705 Tama�oTriangulo
    Y = Y + 4
    Tama�oTriangulo = Tama�oTriangulo - 1
Next
LimpiarParteInferiorPergamino_6705 Tama�oTriangulo
LimpiarParteInferiorPergamino_6705 0
modPantalla.Refrescar
End Sub

Sub LimpiarParteSuperiorDerechaPergamino_663E(ByVal PixelX As Long, ByVal PixelY As Long, ByVal LadoTriangulo As Long)
Dim PunteroDatos As Long
Dim PunteroPantalla As Long
Dim PunteroPantallaAnterior As Long
Dim NumeroPixel As Byte 'n�mero de pixel despu�s del tri�ngulo en la parte superior del pergamino
Dim Linea As Long
Dim XBorde As Long 'coordenadas del borde derecho a restaurar
Dim YBorde As Long
NumeroPixel = &H30 - LadoTriangulo 'halla la parte del pergamino que falta por procesar
NumeroPixel = NumeroPixel * 4 'pasa a pixels
PunteroDatos = NumeroPixel * 2
PunteroPantalla = CalcularDesplazamientoPantalla_68C7(PixelX + 4, PixelY) 'pasa la posici�n actual a direcci�n de VRAM
PunteroPantallaAnterior = PunteroPantalla
For Linea = 0 To 7 '8 l�neas de alto
    PantallaCGA(PunteroPantalla) = DatosGraficosPergamino_788A(PunteroDatos + Linea)
    PantallaCGA2PC PunteroPantalla, DatosGraficosPergamino_788A(PunteroDatos + Linea)
    PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla)
Next 'completa las 8 l�neas
PunteroPantalla = PunteroPantallaAnterior 'recupera la posici�n actual
PunteroPantalla = PunteroPantalla + 1 'avanza 4 pixels en x
For Linea = 8 To 15 'copia los siguientes 4 pixels de otras 8 l�neas
    PantallaCGA(PunteroPantalla) = DatosGraficosPergamino_788A(PunteroDatos + Linea)
    PantallaCGA2PC PunteroPantalla, DatosGraficosPergamino_788A(PunteroDatos + Linea)
    PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla)
Next 'completa las 8 l�neas
YBorde = (LadoTriangulo - 3) * 4
XBorde = &HDA 'x = pixel 218
PunteroDatos = 2 * YBorde + &H7A0A - &H788A
PunteroPantalla = CalcularDesplazamientoPantalla_68C7(XBorde, YBorde) 'pasa la posici�n actual a direcci�n de VRAM
For Linea = 0 To 7 '8 l�neas de alto
    PantallaCGA(PunteroPantalla) = DatosGraficosPergamino_788A(PunteroDatos + 2 * Linea)
    PantallaCGA2PC PunteroPantalla, DatosGraficosPergamino_788A(PunteroDatos + 2 * Linea)
    PantallaCGA(PunteroPantalla + 1) = DatosGraficosPergamino_788A(PunteroDatos + 2 * Linea + 1)
    PantallaCGA2PC PunteroPantalla + 1, DatosGraficosPergamino_788A(PunteroDatos + 2 * Linea + 1)
    PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla)
Next 'completa las 8 l�neas
End Sub

Sub LimpiarParteInferiorPergamino_6705(ByVal Tama�oTriangulo As Long)
'restaura la parte inferior del pergamino modificada por lado Tama�oTriangulo
Dim PunteroDatos As Long
Dim PunteroPantalla As Long
Dim X As Long
Dim Y As Long
Dim Contador As Long
X = &H20 + 4 * Tama�oTriangulo
Y = &HB8 'y = 184
PunteroPantalla = CalcularDesplazamientoPantalla_68C7(X, Y) 'calcula el desplazamiento de las coordenadas en pantalla
PunteroDatos = 4 * Tama�oTriangulo * 2 + &H7D0A - &H788A 'desplazamiento de los datos borrados de la parte inferior del pergamino
For Contador = 0 To 7 '8 l�neas
    PantallaCGA(PunteroPantalla) = DatosGraficosPergamino_788A(PunteroDatos + Contador)
    PantallaCGA2PC PunteroPantalla, DatosGraficosPergamino_788A(PunteroDatos + Contador)
    PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla)
Next

End Sub

Sub ImprimirEspacioPergamino_67CD(ByVal Hueco As Byte, ByRef X As Long)
'a�ade un hueco del tama�o indicado, en p�xeles
Retardo_67C6 &HBB8 'espera un poco (aprox. 30 ms)
X = X + Hueco
End Sub

Sub DibujarCaracterPergamino_6781(ByVal Caracter As Byte, X As Long, Y As Long, ByVal PunteroCaracter As Long, ByVal Color As Byte)
'dibuja un car�cter en el pergamino
Dim Valor As Byte 'dato del car�cter
Dim AvanceX As Long
Dim AvanceY As Long
Dim PunteroPantalla As Long
Dim Pixel As Long
Dim InversaMascaraAND As Byte
Dim MascaraOr As Byte
Dim MascaraAnd As Byte
Dim PunteroHex As String
Dim ValorHex As String
Dim Contador As Long
Do
    If Not Depuracion.QuitarRetardos Then Retardo_67C6 (&H320) 'peque�o retardo (aprox. 8 ms)
    Valor = DatosCaracteresPergamino_6947(PunteroCaracter - &H6947)
    PunteroCaracter = PunteroCaracter + 1
    If (Valor And &HF0) = &HF0 Then 'si es el �ltimo byte del car�cter
        ImprimirEspacioPergamino_67CD Valor And &HF, X 'imprime un espacio y sale al bucle para imprimir m�s caracteres
        Exit Sub
    End If
    AvanceX = Valor And &HF 'avanza la posici�n x seg�n los 4 bits menos significativos del byte leido de dibujo del caracter
    AvanceY = modFunciones.shr(Valor, 4) And &HF& 'avanza la posici�n y seg�n los 4 bits m�s significativos del byte leido de dibujo del caracter
    PunteroPantalla = CalcularDesplazamientoPantalla_68C7(X + AvanceX, Y + AvanceY) 'calcula el desplazamiento de las coordenadas en pantalla
    Pixel = (X + AvanceX) And &H3&        'se queda con los 2 bits menos significativos de la posici�n para saber que pixel pintar
    MascaraAnd = modFunciones.ror8(&H88, Pixel)
    InversaMascaraAND = MascaraAnd Xor &HFF&
    MascaraOr = InversaMascaraAND And PantallaCGA(PunteroPantalla) 'obtiene el valor del resto de pixels de la pantalla
    PunteroHex = Hex$(PunteroPantalla)
    ValorHex = Hex$((Color And MascaraAnd) Or MascaraOr)
    PantallaCGA(PunteroPantalla) = (Color And MascaraAnd) Or MascaraOr 'combina con los pixels de pantalla. actualiza la memoria de video con el nuevo pixel
    PantallaCGA2PC PunteroPantalla, (Color And MascaraAnd) Or MascaraOr
    Contador = Contador + 1
    If Contador > 2 Then
        modPantalla.Refrescar
        Contador = 0
    End If
    
Loop
End Sub

Sub RellenarTablaFlipX_3A61()
'crea una tabla para hacer flip en x a 4 pixels
Dim Contador As Long
Dim Contador2 As Long
Dim NibbleSuperior As Byte
Dim NibbleInferior As Byte
Dim AcarreoI As Byte 'acarreo por la izquierda
Dim AcarreoD As Byte 'acarreo por la derecha

For Contador = 0 To &HFF
    NibbleSuperior = Contador And &HF0
    NibbleInferior = Contador And &HF
    If (NibbleSuperior And &H80) <> 0 Then
        AcarreoI = &H80
    Else
        AcarreoI = 0
    End If
    NibbleSuperior = modFunciones.rol8(NibbleSuperior And &H7F, 1)
    For Contador2 = 1 To 4
        If (NibbleInferior And &H1) <> 0 Then
            AcarreoD = 1
        Else
            AcarreoD = 0
        End If
        NibbleInferior = modFunciones.ror8(NibbleInferior And &HFE, 1) Or AcarreoI
        If (NibbleSuperior And &H80) <> 0 Then
            AcarreoI = &H80
        Else
            AcarreoI = 0
        End If
        NibbleSuperior = modFunciones.rol8(NibbleSuperior And &H7F, 1) Or AcarreoD
    Next
    TablaFlipX_A100(Contador) = NibbleSuperior Or NibbleInferior
Next
End Sub

Sub CerrarEspejo_3A7E()
'obtiene la direcci�n en donde est� la altura del espejo, obtiene la direcci�n del bloque
'que forma el espejo y si estaba abierto, lo cierra
Dim PunteroEspejo As Long
Dim Puntero
Dim Valor As Byte
Dim Contador As Long
PunteroEspejo = &H5086 'apunta a datos de altura de la planta 2
Do
    Valor = TablaAlturasPantallas_4A00(PunteroEspejo - &H4A00)
    If Valor = &HFF Then Exit Do '0xff indica el final
    If (Valor And &H8) <> 0 Then PunteroEspejo = PunteroEspejo + 1 'incrementa la direcci�n 4 o 5 bytes dependiendo del bit 3
    PunteroEspejo = PunteroEspejo + 4
    DoEvents
Loop
PunteroDatosAlturaHabitacionEspejo_34D9 = PunteroEspejo 'guarda el puntero de fin de tabla (que apunta a los datos de la habitaci�n del espejo)
PunteroEspejo = &H4000 'apunta a los datos de bloques de la pantallas
For Contador = 1 To &H72 '114 pantallas
    Valor = DatosHabitaciones_4000(PunteroEspejo - &H4000) 'lee el n�mero de bytes de la pantalla
    PunteroEspejo = PunteroEspejo + Valor
Next
'PunteroEspejo apunta a la habitaci�n del espejo
For Contador = 0 To 255 'hasta 256 bloques
    Valor = DatosHabitaciones_4000(PunteroEspejo - &H4000) 'lee un byte de la habitaci�n del espejo
    PunteroEspejo = PunteroEspejo + 1
    If Valor = &H1F Then 'si es 0x1f, lee los 2 bytes siguientes
        If DatosHabitaciones_4000(PunteroEspejo - &H4000) = &HAA And DatosHabitaciones_4000(PunteroEspejo + 1 - &H4000) = &H51 Then
            'si llega aqu�, los datos de la habitaci�n indican que el espejo est� abierto
            PunteroEspejo = PunteroEspejo + 1
            DatosHabitaciones_4000(PunteroEspejo - &H4000) = &H11 'por lo que modifica la habitaci�n para que el espejo se cierre
            PunteroHabitacionEspejo_34E0 = PunteroEspejo 'guarda el desplazamiento de la pantalla del espejo
        End If
    End If
Next
End Sub

Sub InicializarPartida_2509()
Dim Contador As Long
modTeclado.Inicializar
'aqu� ya se ha completado la inicializaci�n de datos para el juego
'ahora realiza la inicializaci�n para poder empezar a jugar una partida
DeshabilitarInterrupcion
'ApagarSonido_1376 'apaga el sonido '###pendiente
'LeerEstadoTeclas_32BC ###pendiente 'lee el estado de las teclas y lo guarda en los buffers de teclado
Do
    DoEvents
Loop While TeclaPulsadaNivel_3482(&H2F)  'mientras no se suelte el espacio, espera
InicializarVariables_381E
DibujarAreaJuego_275C 'dibuja un rect�ngulo de 256 de ancho en las 160 l�neas superiores de pantalla
DibujarMarcador_272C
'2520
TempoMusica_1086 = 6 'coloca el nuevo tempo de la m�sica
ColocarVectorInterrupcion
VelocidadPasosGuillermo_2618 = 36
'254e
TablaCaracteristicasPersonajes_3036(&H3038 - &H3036) = &H88 'coloca la posici�n inicial de guillermo
TablaCaracteristicasPersonajes_3036(&H3039 - &H3036) = &HA8 'coloca la posici�n inicial de guillermo
TablaCaracteristicasPersonajes_3036(&H3047 - &H3036) = &H88 - 2 'coloca la posici�n inicial de adso
TablaCaracteristicasPersonajes_3036(&H3048 - &H3036) = &HA8 + 2 'coloca la posici�n inicial de adso
TablaCaracteristicasPersonajes_3036(&H303A - &H3036) = 0 'coloca la altura inicial de guillermo
TablaCaracteristicasPersonajes_3036(&H3049 - &H3036) = 0 'coloca la altura inicial de adso
For Contador = 0 To &H2D4 'apunta a los gr�ficos de los movimientos de los monjes
    DatosMonjes_AB59(&H2D5 + Contador) = DatosMonjes_AB59(Contador) 'copia 0xab59-0xae2d a 0xae2e-0xb102
Next
'obtiene en 0xae2e-0xb102 los gr�ficos de los monjes flipeados con respecto a x
GirarGraficosRespectoX_3552 DatosMonjes_AB59, &HAE2E - &HAB59, 5, &H91 'gr�ficos de 5 bytes de ancho, 0x91 bloques de 5 bytes (= 0x2d5)
InicializarEspejo_34B0 'inicia la habitaci�n del espejo y las variables relacionadas con el espejo
InicializarDiaMomento_54D2 'inicia el d�a y el momento del d�a en el que se est�
'habilita los comandos cuando procese el comportamiento
BufferComandosMonjes_A200(&HA2C0 - &HA200) = &H11 'inicia el comando de adso
BufferComandosMonjes_A200(&HA200 - &HA200) = &H11 'inicia el comando de malaqu�as
BufferComandosMonjes_A200(&HA230 - &HA200) = &H11 'inicia el comando del abad
BufferComandosMonjes_A200(&HA260 - &HA200) = &H11 'inicia el comando de berengario
BufferComandosMonjes_A200(&HA290 - &HA200) = &H11 'inicia el comando de severino
ContadorInterrupcion_2D4B = 0 'resetea el contador de la interrupci�n
PintarPantalla_0DFD = True 'a�adido para que corresponda con lo que hace realmente
'For Contador = 0 To UBound(BufferSprites_9500)
'    BufferSprites_9500(Contador) = &HFF 'rellena el buffer de sprites con un relleno para depuraci�n
'Next
InicializarPartida_258F
End Sub

Function TeclaPulsadaNivel_3482(ByVal CodigoTecla As Byte)
'comprueba si se est� pulsando la tecla con el c�digo indicado. si no est� siedo pulsada, devuelve true
TeclaPulsadaNivel_3482 = modTeclado.TeclaPulsadaNivel(TraducirCodigoTecla(CodigoTecla))
End Function

Function TeclaPulsadaNivel_3472(ByVal CodigoTecla As Byte)
'comprueba si ha sido pulsanda la tecla con el c�digo indicado. si no ha sido pulsada o ya se ha preguntado antes, devuelve true
TeclaPulsadaNivel_3472 = modTeclado.TeclaPulsadaFlanco(TraducirCodigoTecla(CodigoTecla))
End Function

Function TraducirCodigoTecla(ByVal CodigoTecla As Byte) As EnumTecla
    Select Case CodigoTecla
        Case Is = &H0
            TraducirCodigoTecla = TeclaArriba
        Case Is = &H2
            TraducirCodigoTecla = TeclaAbajo
        Case Is = &H8
            TraducirCodigoTecla = TeclaIzquierda
        Case Is = &H1
            TraducirCodigoTecla = TeclaDerecha
        Case Is = &H2F
            TraducirCodigoTecla = TeclaEspacio
        Case Is = &H44
            TraducirCodigoTecla = TeclaTabulador
        Case Is = &H17
            TraducirCodigoTecla = TeclaControl
        Case Is = &H15
            TraducirCodigoTecla = TeclaMayusculas
        Case Is = &H6
            TraducirCodigoTecla = TeclaEnter
        Case Is = &H4F
            TraducirCodigoTecla = TeclaSuprimir
        Case Is = &H42
            TraducirCodigoTecla = TeclaEscape
        Case Is = &H7
            TraducirCodigoTecla = TeclaPunto
        Case Is = &H3C
            TraducirCodigoTecla = TeclaS
        Case Is = &H2E
            TraducirCodigoTecla = TeclaN
        Case Is = &H43
            TraducirCodigoTecla = TeclaQ
        Case Is = &H32
            TraducirCodigoTecla = TeclaR
    End Select
End Function

Sub InicializarVariables_381E()
'inicia la memoria
Dim Contador As Long
Dim Puntero As Long
Dim Valor As Byte
For Contador = 0 To UBound(TablaVariablesLogica_3C85)
    TablaVariablesLogica_3C85(Contador) = 0
Next
For Contador = 0 To UBound(TablaVariablesAuxiliares_2D8D)
    TablaVariablesAuxiliares_2D8D(Contador) = 0
Next
RestaurarVariables_37B9 'copia cosas de 0x0103-0x01a9 a muchos sitios (nota: al inicializar se hizo la operaci�n inversa)
Puntero = &H2E17 'apunta a la tabla con datos de los sprites
Contador = &H14 'cada sprite ocupa 20 bytes
Do
    Valor = TablaSprites_2E17(Puntero - &H2E17)
    If Valor = &HFF Then Exit Do
    TablaSprites_2E17(Puntero - &H2E17) = &HFE 'pone todos los sprites como no visibles
    Puntero = Puntero + Contador
    DoEvents
Loop
Puntero = &H3036 'apunta a la tabla de caracter�sticas de los personajes
For Contador = 0 To 5 '6 entradas
    TablaCaracteristicasPersonajes_3036(Puntero - &H3036) = 0 'pone a 0 el contador de la animaci�n del personaje
    TablaCaracteristicasPersonajes_3036(Puntero + 1 - &H3036) = 0 'fija la orientaci�n del personaje mirando a +x
    TablaCaracteristicasPersonajes_3036(Puntero + 5 - &H3036) = 0 'inicialmente el personaje ocupa 4 posiciones
    TablaCaracteristicasPersonajes_3036(Puntero + 9 - &H3036) = 0 'indica que no hay movimientos del personaje que procesar
    TablaCaracteristicasPersonajes_3036(Puntero + &HA - &H3036) = &HFD 'acci�n que se est� ejecutando actualmente
    TablaCaracteristicasPersonajes_3036(Puntero + &HB - &H3036) = 0 'inicia el �ndice en la tabla de comandos de movimiento
    Puntero = Puntero + &HF 'cada entrada ocupa 15 bytes
Next
End Sub

Sub DibujarAreaJuego_275C()
'dibuja un rect�ngulo de 256 de ancho en las 160 l�neas superiores de pantalla
Dim PunteroPantalla As Long
Dim Contador As Long
Dim Contador2 As Long
PunteroPantalla = 0
For Contador = 1 To &HA0 '160 l�neas
    For Contador2 = 0 To 7 'rellena 8 bytes con 0xff (32 pixels)
        PantallaCGA(PunteroPantalla + Contador2) = &HFF
        PantallaCGA2PC PunteroPantalla + Contador2, &HFF
    Next
    For Contador2 = 0 To &H40 'rellena 64 bytes con 0x00 (256 pixels)
        PantallaCGA(PunteroPantalla + 8 + Contador2) = 0
        PantallaCGA2PC PunteroPantalla + 8 + Contador2, 0
    Next
    For Contador2 = 0 To 7 'rellena 8 bytes con 0xff (32 pixels)
        PantallaCGA(PunteroPantalla + 72 + Contador2) = &HFF
        PantallaCGA2PC PunteroPantalla + 72 + Contador2, &HFF
    Next
    PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla)
    DoEvents
Next
modPantalla.Refrescar
End Sub

Sub DibujarMarcador_272C()
Dim PunteroDatos As Long
Dim PunteroPantalla As Long
Dim PunteroPantallaAnterior As Long
Dim Contador As Long
Dim Contador2 As Long
Dim Contador3 As Long
PunteroDatos = &H6328 'apunta a datos del marcador (de 0x6328 a 0x6b27)
PunteroPantalla = &H648 'apunta a la direcci�n en memoria donde se coloca el marcador (32, 160)
For Contador = 0 To 3
    PunteroPantallaAnterior = PunteroPantalla
    For Contador2 = 0 To 7 '8 l�neas
        For Contador3 = 0 To &H3F 'copia 64 bytes a pantalla (256 pixels)
            PantallaCGA(PunteroPantalla + Contador3) = DatosMarcador_6328(PunteroDatos - &H6328)
            PantallaCGA2PC PunteroPantalla + Contador3, DatosMarcador_6328(PunteroDatos - &H6328)
            PunteroDatos = PunteroDatos + 1
        Next
        PunteroPantalla = PunteroPantalla + &H800
    Next
    PunteroPantalla = PunteroPantallaAnterior
    PunteroPantalla = PunteroPantalla + &H50
    modPantalla.Refrescar
Next



End Sub


Sub GirarGraficosRespectoX_3552(ByRef Tabla() As Byte, ByVal PunteroTablaHL As Long, ByVal AnchoC As Byte, ByVal NGraficosB As Byte)
'gira con respecto a x una serie de datos gr�ficos que se le pasan en Tabla
'el ancho de los gr�ficos se pasa en Ancho y en NGraficos un n�mero para calcular cuantos gr�ficos girar
Dim Bloque As Long 'contador de l�neas
Dim Contador As Long 'contador dentro de la l�nea
Dim NumeroCambios As Long
Dim Valor1 As Byte
Dim Valor2 As Byte
Dim PunteroValor1 As Long
Dim PunteroValor2 As Long
Dim HL As String
HL = Hex$(PunteroTablaHL)
NumeroCambios = Int(AnchoC + 1) / 2
For Bloque = 0 To NGraficosB - 1
    For Contador = 0 To NumeroCambios - 1
        PunteroValor1 = PunteroTablaHL + AnchoC * Bloque + Contador 'valor por la izquierda
        PunteroValor2 = PunteroTablaHL + AnchoC * Bloque + AnchoC - 1 - Contador 'valor por la derecha
        Valor1 = Tabla(PunteroValor1)
        Valor2 = Tabla(PunteroValor2)
        'se usa la tabla auxiliar para flipx
        Valor1 = TablaFlipX_A100(Valor1)
        Valor2 = TablaFlipX_A100(Valor2)
        'intercambia los registros
        Tabla(PunteroValor1) = Valor2
        Tabla(PunteroValor2) = Valor1
    Next
Next
'3584
End Sub

Sub InicializarEspejo_34B0()
'inicializa la habitaci�n del espejo y sus variables
HabitacionEspejoCerrada_2D8C = True 'inicialmente la habitaci�n secreta detr�s del espejo no est� abierta
NumeroRomanoHabitacionEspejo_2DBC = 0 'indica que el n�mero romano de la habitaci�n del espejo no se ha generado todav�a
InicializarEspejo_34B9
End Sub

Sub InicializarEspejo_34B9()
Dim Contador As Long
DeshabilitarInterrupcion
For Contador = 0 To 4
     TablaAlturasPantallas_4A00(PunteroDatosAlturaHabitacionEspejo_34D9 + Contador - &H4A00) = DatosAlturaEspejoCerrado_34DB(Contador)
Next
EscribirValorBloqueHabitacionEspejo_336F &H11 'modifica la habitaci�n del espejo para que el espejo aparezca cerrado
EscribirValorBloqueHabitacionEspejo_3372 &H1F 'modifica la habitaci�n del espejo para que la trampa est� cerrada
HabilitarInterrupcion
End Sub

Sub EscribirValorBloqueHabitacionEspejo_336F(ByVal Valor As Byte)
'graba el valor en el bloque que forma el espejo en la habitaci�n el espejo
DatosHabitaciones_4000(PunteroHabitacionEspejo_34E0 - &H4000) = Valor
End Sub

Sub EscribirValorBloqueHabitacionEspejo_3372(ByVal Valor As Byte)
'graba el valor dos posiciones antes del bloque que forma el espejo en la habitaci�n el espejo
DatosHabitaciones_4000(PunteroHabitacionEspejo_34E0 - 2 - &H4000) = Valor
End Sub

Sub InicializarDiaMomento_54D2()
'inicia el d�a y el momento del d�a en el que se est�
NumeroDia_2D80 = 1 'primer d�a
MomentoDia_2D81 = 4 '4=nona
End Sub

Sub InicializarPartida_258F()
'segunda parte de la inicializaci�n. cuando carga una partida tambi�n se llega aqu�
DeshabilitarInterrupcion
PosicionXPersonajeActual_2D75 = 0 'inicia la pantalla en la que est� el personaje
EstadoGuillermo_288F = 0 'inicia el estado de guillermo
AjustePosicionYSpriteGuillermo_28B1 = 2
'DibujarAreaJuego_275C 'dibuja un rect�ngulo de 256 de ancho en las 160 l�neas superiores de pantalla
'ApagarSonido_1376 'apaga el sonido '###pendiente
InicializarEspejo_34B9 'inicia la habitaci�n del espejo
DibujarObjetosMarcador_51D4 'dibuja los objetos que tenemos en el marcador
FijarPaletaMomentoDia_54DF 'fija la paleta seg�n el momento del d�a, muestra el n�mero de d�a y avanza el momento del d�a
DecrementarObsequium_55D3 0 'decrementa el obsequium 0 unidades
LimpiarZonaFrasesMarcador_5001 'limpia la parte del marcador donde se muestran las frases
If Not Check Then
    BuclePrincipal_25B7 'el bucle principal del juego empieza aqu�
Else
    BuclePrincipal_Check
End If
End Sub


Public Sub GenerarBloqueSuelto(ByVal Bloque As Byte, ByVal X As Byte, ByVal Y As Byte, ByVal nX As Byte, ByVal nY As Byte, ByVal Byte3 As Byte)
'genera el escenerio con los datos de abadia8 y lo proyecta a la rejilla
'lee la entrada de abadia8 con un bloque de construcci�n de la pantalla y llama a 0x1bbc

Dim PunteroCaracteristicasBloque As Long 'puntero a las caracter�sitcas del bloque
Dim PunteroTilesBloque As Long 'puntero del material a los tiles que forman el bloque
Dim PunteroRutinasBloque As Long 'puntero al resto de caracter�sticas del material
Dim BloqueHex As String
'PunteroPantalla = 2445

BloqueHex = Hex$(Bloque)

If Bloque = 255 Then Exit Sub '0xff indica el fin de pantalla
PunteroCaracteristicasBloque = Leer16(TablaBloquesPantallas_156D, Bloque And &HFE&) 'desprecia el bit inferior para indexar
VariablesBloques_1FCD(&H1FDE - &H1FCD) = 0 'inicia a (0, 0) la posici�n del bloque en la rejilla (sistema de coordenadas local de la rejilla)
VariablesBloques_1FCD(&H1FDF - &H1FCD) = 0 'inicia a (0, 0) la posici�n del bloque en la rejilla (sistema de coordenadas local de la rejilla)
VariablesBloques_1FCD(&H1FDD - &H1FCD) = Byte3
PunteroTilesBloque = Leer16(TablaCaracteristicasMaterial_1693, PunteroCaracteristicasBloque - &H1693)
PunteroRutinasBloque = PunteroCaracteristicasBloque + 2
ConstruirBloque_1BBC X, nX, Y, nY, Byte3, PunteroTilesBloque, PunteroRutinasBloque, True





End Sub


Sub CopiarVariables_37B6()
CopiarTabla TablaPermisosPuertas_2DD9, CopiaTablaPermisosPuertas_2DD9 'puertas a las que pueden entrar los personajes
CopiarTabla TablaObjetosPersonajes_2DEC, CopiaTablaObjetosPersonajes_2DEC 'objetos de los personajes
CopiarTabla TablaDatosPuertas_2FE4, CopiaTablaDatosPuertas_2FE4 'datos de las puertas del juego
CopiarTabla TablaPosicionObjetos_3008, CopiaTablaPosicionObjetos_3008 'posici�n de los objetos

End Sub

Sub RestaurarVariables_37B9()
PuertasAbribles_3CA6 = &HEF ' m�scara para las puertas donde cada bit indica que puerta se comprueba si se abre
InvestigacionNoTerminada_3CA7 = True
TablaPosicionesPredefinidasMalaquias_3CA8(0) = &HFA
TablaPosicionesPredefinidasMalaquias_3CA8(1) = 0
TablaPosicionesPredefinidasMalaquias_3CA8(2) = 0
TablaPosicionesPredefinidasAbad_3CC6(0) = &HFA
TablaPosicionesPredefinidasAbad_3CC6(1) = 0
TablaPosicionesPredefinidasAbad_3CC6(2) = 0
TablaPosicionesPredefinidasBerengario_3CE7(0) = &HFA
TablaPosicionesPredefinidasBerengario_3CE7(1) = 0
TablaPosicionesPredefinidasBerengario_3CE7(2) = 0
TablaPosicionesPredefinidasSeverino_3CFF(0) = &HFA
TablaPosicionesPredefinidasSeverino_3CFF(1) = 0
TablaPosicionesPredefinidasSeverino_3CFF(2) = 0
TablaPosicionesPredefinidasAdso_3D11(0) = &HFF
TablaPosicionesPredefinidasAdso_3D11(1) = 0
TablaPosicionesPredefinidasAdso_3D11(2) = 0
Obsequium_2D7F = &H1F
NumeroDia_2D80 = 1
MomentoDia_2D81 = 4
PunteroProximaHoraDia_2D82 = &H4FBC
PunteroTablaDesplazamientoAnimacion_2D84 = &H309F
TiempoRestanteMomentoDia_2D86 = &HDAC
PunteroDatosPersonajeActual_2D88 = &H3036
PunteroBufferAlturas_2D8A = &H1C0
HabitacionEspejoCerrada_2D8C = True
CopiarTabla CopiaTablaPermisosPuertas_2DD9, TablaPermisosPuertas_2DD9 'puertas a las que pueden entrar los personajes
CopiarTabla CopiaTablaObjetosPersonajes_2DEC, TablaObjetosPersonajes_2DEC 'objetos de los personajes
CopiarTabla CopiaTablaDatosPuertas_2FE4, TablaDatosPuertas_2FE4 'datos de las puertas del juego
CopiarTabla CopiaTablaPosicionObjetos_3008, TablaPosicionObjetos_3008 'posici�n de los objetos
TablaCaracteristicasPersonajes_3036(&H3038 - &H3036) = &H22 'posici�n de guillermo
TablaCaracteristicasPersonajes_3036(&H3039 - &H3036) = &H22 'posici�n de guillermo
TablaCaracteristicasPersonajes_3036(&H303A - &H3036) = 0 'posici�n de guillermo
TablaCaracteristicasPersonajes_3036(&H3047 - &H3036) = &H24 'posici�n de adso
TablaCaracteristicasPersonajes_3036(&H3048 - &H3036) = &H24 'posici�n de adso
TablaCaracteristicasPersonajes_3036(&H3049 - &H3036) = 0 'posici�n de adso
TablaCaracteristicasPersonajes_3036(&H3056 - &H3036) = &H26 'posici�n de malaqu�as
TablaCaracteristicasPersonajes_3036(&H3057 - &H3036) = &H26 'posici�n de malaqu�as
TablaCaracteristicasPersonajes_3036(&H3058 - &H3036) = &HF  'posici�n de malaqu�as
TablaCaracteristicasPersonajes_3036(&H3065 - &H3036) = &H88 'posici�n del abad
TablaCaracteristicasPersonajes_3036(&H3066 - &H3036) = &H84 'posici�n del abad
TablaCaracteristicasPersonajes_3036(&H3067 - &H3036) = &H2  'posici�n del abad
TablaCaracteristicasPersonajes_3036(&H3074 - &H3036) = &H28 'posici�n de berengario
TablaCaracteristicasPersonajes_3036(&H3075 - &H3036) = &H48 'posici�n de berengario
TablaCaracteristicasPersonajes_3036(&H3076 - &H3036) = &HF  'posici�n de berengario
TablaCaracteristicasPersonajes_3036(&H3083 - &H3036) = &HC8 'posici�n de severino
TablaCaracteristicasPersonajes_3036(&H3084 - &H3036) = &H28 'posici�n de severino
TablaCaracteristicasPersonajes_3036(&H3085 - &H3036) = 0  'posici�n de severino
End Sub

Sub DibujarObjetosMarcador_51D4()
'dibuja los objetos que tiene guillermo en el marcador
Dim Objetos As Byte
Objetos = TablaObjetosPersonajes_2DEC(&H2DEF - &H2DEC) 'lee los objetos que tenemos
ActualizarMarcador_51DA Objetos, &HFF 'comprobar todos los objetos posibles. y si est�n, se dibujan
End Sub

Sub ActualizarMarcador_51DA(ByVal Objetos As Byte, ByVal Mascara As Byte)
'comprueba si se tienen los objetos que se le pasan (se comprueban los indicados por la m�scara), y si se tienen se dibujan
Dim PunteroPosicionesObjetos As Long
Dim PunteroSpritesObjetos As Long
Dim PunteroPantalla As Long
Dim PunteroPantallaAnterior As Long
Dim PunteroGraficosObjeto As Long
Dim Contador As Long
Dim Alto As Long
Dim Ancho As Long
Dim ContadorAncho As Long
Dim ContadorAlto As Long
PunteroPosicionesObjetos = &H3008 'apunta a las posiciones sobre los objetos del juego
PunteroSpritesObjetos = &H2F1B
PunteroPantalla = &H6F9& 'apunta a la memoria de video del primer hueco (100, 176)
For Contador = 1 To 6 'hay 6 huecos donde colocar los objetos
    If Mascara = 0 Then Exit Sub 'ya no hay objetos por comprobar
    If (Mascara And &H80) <> 0 Then 'comprobar objeto
        PunteroPantallaAnterior = PunteroPantalla
        If (Objetos And &H80) <> 0 Then 'objeto presente. lo dibuja
            Alto = TablaSprites_2E17(PunteroSpritesObjetos + 6 - &H2E17) 'lee el alto del objeto
            Ancho = TablaSprites_2E17(PunteroSpritesObjetos + 5 - &H2E17) 'lee el ancho del objeto
            Ancho = Ancho And &H7F& 'pone a 0 el bit 7
            PunteroGraficosObjeto = Leer16(TablaSprites_2E17, PunteroSpritesObjetos + 7 - &H2E17)
            For ContadorAlto = 0 To Alto - 1
                For ContadorAncho = 0 To Ancho - 1
                    PantallaCGA(PunteroPantalla + ContadorAncho) = TilesAbadia_6D00(PunteroGraficosObjeto + ContadorAlto * Ancho + ContadorAncho - &H6D00)
                    PantallaCGA2PC PunteroPantalla + ContadorAncho, TilesAbadia_6D00(PunteroGraficosObjeto + ContadorAlto * Ancho + ContadorAncho - &H6D00)
                Next
                PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla)
            Next
        Else 'objeto ausente. limpia el hueco
            For Alto = 0 To 11
                For Ancho = 0 To 3
                    PantallaCGA(PunteroPantalla + Ancho) = 0  'limpia el pixel actual
                    PantallaCGA2PC PunteroPantalla + Ancho, 0
                Next
                PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla)
            Next
        End If
        modPantalla.Refrescar
        PunteroPantalla = PunteroPantallaAnterior
    End If
    PunteroPantalla = PunteroPantalla + 5 'pasa al siguiente hueco
    PunteroPosicionesObjetos = PunteroPosicionesObjetos + 5 'avanza las posiciones sobre los objetos del juego
    PunteroSpritesObjetos = PunteroSpritesObjetos + &H14 'avanza a la siguiente entrada de las caracter�sticas del objeto
    If Contador = 3 Then PunteroPantalla = PunteroPantalla + 1 'al pasar del hueco 3 al 4 hay 4 pixels extra
    Mascara = Mascara And &H7F
    Mascara = Mascara * 2 'desplaza la m�scara un bit hacia la izquierda
    Objetos = Objetos And &H7F
    Objetos = Objetos * 2 'desplaza los objetos un bit hacia la izquierda
Next
End Sub



Sub ActualizarDiaMarcador_5559(ByVal Dia As Byte)
'actualiza el d�a, reflej�ndolo en el marcador
NumeroDia_2D80 = Dia 'actualiza el d�a
Dim PunteroDia As Long
Dim PunteroPantalla As Long
PunteroDia = &H4FA7 + (Dia - 1) * 3 'indexa en la tabla de los d�as. ajusta el �ndice a 0. cada entrada en la tabla ocupa 3 bytes
PunteroPantalla = &HEE51 - &HC000 'apunta a pantalla (68, 165)
DibujarNumeroDia_5583 PunteroDia, PunteroPantalla 'coloca el primer n�mero de d�a en el marcador
DibujarNumeroDia_5583 PunteroDia, PunteroPantalla 'coloca el segundo n�mero de d�a en el marcador
DibujarNumeroDia_5583 PunteroDia, PunteroPantalla 'coloca el tercer n�mero de d�a en el marcador
InicializarScrollMomentoDia_5575 0 'pone la primera hora del d�a
End Sub

Sub InicializarScrollMomentoDia_5575(ByVal MomentoDia As Byte)
MomentoDia_2D81 = MomentoDia
ScrollCambioMomentoDia_2DA5 = 9 '9 posiciones para realizar el scroll del cambio del momento del d�a
ColocarDiaHora_550A 'pone en 0x2d86 un valor dependiente del d�a y la hora
End Sub
Sub DibujarNumeroDia_5583(ByRef PunteroDia As Long, ByRef PunteroPantalla As Long)
'pone un n�mero de d�a
Dim Sumar As Boolean
Dim Valor As Byte
Dim PunteroGraficos As Long
Dim Contador As Long
Dim PunteroPantallaAnterior As Long
PunteroPantallaAnterior = PunteroPantalla
Sumar = True
Valor = TablaEtapasDia_4F7A(PunteroDia - &H4F7A) 'lee un byte de los datos que forman el n�mero del d�a
Select Case Valor
    Case Is = 2
        PunteroGraficos = &HAB49&
    Case Is = 1
        PunteroGraficos = &HAB39&
    Case Else
        PunteroGraficos = &HA3BB&   'cambiado para que apunte a una zona con FF FF en TablaGraficosObjetos_A300
        Sumar = False
End Select
For Contador = 0 To 7 'rellena las 8 l�neas que ocupa la letra (8x8)
    PantallaCGA(PunteroPantalla) = TablaGraficosObjetos_A300(PunteroGraficos - &HA300&)
    PantallaCGA2PC PunteroPantalla, TablaGraficosObjetos_A300(PunteroGraficos - &HA300&)
    PunteroPantalla = PunteroPantalla + 1
    PunteroGraficos = PunteroGraficos + 1
    PantallaCGA(PunteroPantalla) = TablaGraficosObjetos_A300(PunteroGraficos - &HA300&)
    PantallaCGA2PC PunteroPantalla, TablaGraficosObjetos_A300(PunteroGraficos - &HA300&)
    PunteroPantalla = PunteroPantalla - 1
    If Sumar Then
        PunteroGraficos = PunteroGraficos + 1
    Else
        PunteroGraficos = PunteroGraficos - 1
    End If
    PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla)
Next
modPantalla.Refrescar
PunteroPantalla = PunteroPantallaAnterior + 2
PunteroDia = PunteroDia + 1
End Sub

Sub ColocarDiaHora_550A()
'pone en 0x2d86 un valor dependiente del d�a y la hora
Dim PunteroDuracionEtapaDia As Long
PunteroDuracionEtapaDia = &H4F7A + 7 * NumeroDia_2D80 + MomentoDia_2D81
TiempoRestanteMomentoDia_2D86 = Leer16(TablaEtapasDia_4F7A, PunteroDuracionEtapaDia - &H4F7A)
End Sub

Sub FijarPaletaMomentoDia_54DF()
'fija la paleta seg�n el momento del d�a y muestra el n�mero de d�a
Dim MomentoDia_2D81Anterior As Byte
MomentoDia_2D81Anterior = MomentoDia_2D81
If MomentoDia_2D81 < 6 Then
    modPantalla.SeleccionarPaleta 2 'paleta de d�a
Else
    modPantalla.SeleccionarPaleta 3 'paleta de noche
End If
ActualizarDiaMarcador_5559 NumeroDia_2D80 'dibuja el n�mero de d�a en el marcador
MomentoDia_2D81 = MomentoDia_2D81Anterior - 1 'recupera el momento del d�a en el que estaba
PunteroProximaHoraDia_2D82 = &H4FBC + 7 * (MomentoDia_2D81 + 1) 'apunta al nombre del momento del d�a
ActualizarMomentoDia_553E 'avanza el momento del d�a
End Sub

Sub ActualizarMomentoDia_553E()
'actualiza el momento del d�a
MomentoDia_2D81 = MomentoDia_2D81 + 1 'avanza la hora del d�a
If MomentoDia_2D81 = 7 Then 'si se sali� de la tabla vuelve al primer momento del d�a
    PunteroProximaHoraDia_2D82 = &H4FBC
    NumeroDia_2D80 = NumeroDia_2D80 + 1 'avanza un d�a
    If NumeroDia_2D80 = 8 Then
        ActualizarDiaMarcador_5559 1 'en el caso de que se haya pasado del s�ptimo d�a, vuelve al primer d�a
    End If
Else
    InicializarScrollMomentoDia_5575 MomentoDia_2D81
End If
End Sub

Sub DecrementarObsequium_55D3(ByVal Decremento As Byte)
'decrementa y actualiza en pantalla la barra de energ�a (obsequium)
Dim TablaRellenoObsequium(3) As Byte 'tabla con pixels para rellenar los 4 �ltimos pixels de la barra de obsequium
Dim PunteroRelleno As Byte 'apunta a una tabla de pixels para los 4 �ltimos pixels de la vida
Dim Valor As Byte
Dim PunteroPantalla As Long
TablaRellenoObsequium(0) = &HFF
TablaRellenoObsequium(1) = &H7F
TablaRellenoObsequium(2) = &H3F
TablaRellenoObsequium(3) = &H1F
Obsequium_2D7F = Obsequium_2D7F - Decremento 'lee la energ�a y le resta las unidades leidas
If Obsequium_2D7F < 0 Then 'aqu� llega si ya no queda energ�a
    If Not GuillermoMuerto_3C97 Then
        TablaPosicionesPredefinidasAbad_3CC6(&H3CC7 - &H3CC6) = &HB 'cambia el estado del abad para que le eche de la abad�a
    End If
    Obsequium_2D7F = 0 ' actualiza el contador de energ�a
End If
PunteroRelleno = Obsequium_2D7F And &H3
Valor = TablaRellenoObsequium(PunteroRelleno) 'indexa en la tabla seg�n los 2 bits menos significativos
PunteroPantalla = &HF1C&  'apunta a pantalla (252, 177)
DibujarBarraObsequium_560E Int(Obsequium_2D7F / 4), &HF, PunteroPantalla 'dibuja la primera parte de la barra de vida.  calcula el ancho de la barra de vida readondeada al m�ltiplo de 4 m�s cercano
DibujarBarraObsequium_560E 1, Valor, PunteroPantalla '4 pixels de ancho+valor a escribir dependiendo de la vida que quede. dibuja la segunda parte de la barra de vida
DibujarBarraObsequium_560E 7 - Int(Obsequium_2D7F / 4), &HFF, PunteroPantalla 'obtiene la vida que ha perdido y rellena de negro
End Sub

Sub DibujarBarraObsequium_560E(ByVal Ancho As Byte, ByVal Relleno As Byte, ByRef PunteroPantalla As Long)
'dibuja un rect�ngulo de Ancho bytes de ancho y 6 l�neas de alto (graba valor de relleno)
If Ancho = 0 Then Exit Sub
Dim Contador As Long
Dim Contador2 As Long
Dim PunteroPantallaAnterior As Long
For Contador = 1 To Ancho
    PunteroPantallaAnterior = PunteroPantalla
    For Contador2 = 1 To 6 '6 l�neas de alto
        PantallaCGA(PunteroPantalla) = Relleno
        PantallaCGA2PC PunteroPantalla, Relleno
        PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla)
    Next
    PunteroPantalla = PunteroPantallaAnterior + 1
Next
modPantalla.Refrescar
End Sub

Sub LimpiarZonaFrasesMarcador_5001()
'limpia la parte del marcador donde se muestran las frases
Dim Contador As Long
Dim Contador2 As Long
Dim PunteroPantalla As Long
PunteroPantalla = &H2658 'apunta a pantalla (96, 164)
For Contador = 1 To 8 '8 l�neas de alto
    For Contador2 = 0 To &H1F 'repite hasta rellenar 128 pixels de esta l�nea
        PantallaCGA(PunteroPantalla + Contador2) = &HFF
        PantallaCGA2PC PunteroPantalla + Contador2, &HFF
    Next
    PunteroPantalla = DireccionSiguienteLinea_3A4D_68F2(PunteroPantalla) 'pasa a la siguiente l�nea de pantalla
Next
modPantalla.Refrescar
End Sub

Sub BuclePrincipal_25B7()
    Dim Contador As Long
    Dim PunteroPersonajeHL As Long
    
    
    
    
    'el bucle principal del juego empieza aqu�
    
    'puerta en primera pantalla:
    'TablaDatosPuertas_2FE4(17) = &H88
    'TablaDatosPuertas_2FE4(18) = &HAD
    'TablaDatosPuertas_2FE4(19) = 0
    'libro en la primera pantalla:4
    'TablaPosicionObjetos_3008(2) = &H81
    'TablaPosicionObjetos_3008(3) = &HA6
    'TablaPosicionObjetos_3008(4) = 0
    
    
    'TablaCaracteristicasPersonajes_3036(&H3047 - &H3036) = &H88 - 0 'coloca la posici�n inicial de adso
    'TablaCaracteristicasPersonajes_3036(&H3048 - &H3036) = &HA8 + 0 'coloca la posici�n inicial de adso
    
    'guillermo enla biblioteca
    'TablaCaracteristicasPersonajes_3036(&H3036 + 2 - &H3036) = &H21
    'TablaCaracteristicasPersonajes_3036(&H3036 + 3 - &H3036) = &H69
    'TablaCaracteristicasPersonajes_3036(&H3036 + 4 - &H3036) = &H1A
    
    'bug de tiles
    'TablaCaracteristicasPersonajes_3036(&H3036 + 1 - &H3036) = &H2
    'TablaCaracteristicasPersonajes_3036(&H3036 + 2 - &H3036) = &H41
    'TablaCaracteristicasPersonajes_3036(&H3036 + 3 - &H3036) = &H7D
    'TablaCaracteristicasPersonajes_3036(&H3036 + 4 - &H3036) = &H4
    
    'bug de biblioteca
    'TablaCaracteristicasPersonajes_3036(&H3036 + 1 - &H3036) = &H3
    'TablaCaracteristicasPersonajes_3036(&H3036 + 2 - &H3036) = &H45
    'TablaCaracteristicasPersonajes_3036(&H3036 + 3 - &H3036) = &H28
    'TablaCaracteristicasPersonajes_3036(&H3036 + 4 - &H3036) = &H20
    
    'bug de puertas
    'TablaCaracteristicasPersonajes_3036(&H3036 + 1 - &H3036) = &H2
    'TablaCaracteristicasPersonajes_3036(&H3036 + 2 - &H3036) = &H72
    'TablaCaracteristicasPersonajes_3036(&H3036 + 3 - &H3036) = &H8C
    'TablaCaracteristicasPersonajes_3036(&H3036 + 4 - &H3036) = &H0
    
    '44
    'TablaCaracteristicasPersonajes_3036(&H3036 + 1 - &H3036) = &H0
    'TablaCaracteristicasPersonajes_3036(&H3036 + 2 - &H3036) = &H28
    'TablaCaracteristicasPersonajes_3036(&H3036 + 3 - &H3036) = &H27
    'TablaCaracteristicasPersonajes_3036(&H3036 + 4 - &H3036) = &HF
    'TablaCaracteristicasPersonajes_3036(&H3036 + 5 - &H3036) = &H0
    
    
    'cambia de sitio al abad
    'TablaCaracteristicasPersonajes_3036(&H3063 + 2 - &H3036) = &H88
    'TablaCaracteristicasPersonajes_3036(&H3063 + 3 - &H3036) = &HA6
    'TablaCaracteristicasPersonajes_3036(&H3063 + 4 - &H3036) = &H0
    'TablaSprites_2E17(&H2E53 - &H2E17) = TablaSprites_2E17(&H2E2B - &H2E17)
    
    'bug &H2d
    'TablaCaracteristicasPersonajes_3036(&H3036 + 1 - &H3036) = &H1
    'TablaCaracteristicasPersonajes_3036(&H3036 + 2 - &H3036) = &HC2
    'TablaCaracteristicasPersonajes_3036(&H3036 + 3 - &H3036) = &H3F
    'TablaCaracteristicasPersonajes_3036(&H3036 + 4 - &H3036) = &H2
    'TablaCaracteristicasPersonajes_3036(&H3036 + 5 - &H3036) = &H0
    
    'bug &H14
    'TablaCaracteristicasPersonajes_3036(&H3036 + 1 - &H3036) = &H2
    'TablaCaracteristicasPersonajes_3036(&H3036 + 2 - &H3036) = &H7A
    'TablaCaracteristicasPersonajes_3036(&H3036 + 3 - &H3036) = &H8A
    'TablaCaracteristicasPersonajes_3036(&H3036 + 4 - &H3036) = &H0
    'TablaCaracteristicasPersonajes_3036(&H3036 + 5 - &H3036) = &H0
    
    'bug &H40
    'TablaCaracteristicasPersonajes_3036(&H3036 + 1 - &H3036) = &H3
    'TablaCaracteristicasPersonajes_3036(&H3036 + 2 - &H3036) = &H34
    'TablaCaracteristicasPersonajes_3036(&H3036 + 3 - &H3036) = &H5C
    'TablaCaracteristicasPersonajes_3036(&H3036 + 4 - &H3036) = &HF
    'TablaCaracteristicasPersonajes_3036(&H3036 + 5 - &H3036) = &H0
    
    'bug &H33
    'TablaCaracteristicasPersonajes_3036(&H3036 + 1 - &H3036) = &H1
    'TablaCaracteristicasPersonajes_3036(&H3036 + 2 - &H3036) = &HC1
    'TablaCaracteristicasPersonajes_3036(&H3036 + 3 - &H3036) = &H5E
    'TablaCaracteristicasPersonajes_3036(&H3036 + 4 - &H3036) = &H0
    'TablaCaracteristicasPersonajes_3036(&H3036 + 5 - &H3036) = &H0
    
    Parado = False
    Do
        ContadorAnimacionGuillermo_0990 = TablaCaracteristicasPersonajes_3036(&H3036 - &H3036) 'obtiene el contador de la animaci�n de guillermo
        
    
        
        
        '25e4
        ComprobarCambioPantalla_2355 'comprueba si el personaje que se muestra ha cambiado de pantalla y si es as� hace muchas cosas
        
        'modFunciones.GuardarArchivo "Buffer0", BufferAlturas
        'modFunciones.GuardarArchivo "Sprites0", TablaSprites_2E17
        'modFunciones.GuardarArchivo "Puertas0", TablaDatosPuertas_2FE4
        'modFunciones.GuardarArchivo "Objetos0", TablaPosicionObjetos_3008
        'modFunciones.GuardarArchivo "Perso0", TablaCaracteristicasPersonajes_3036
        'modFunciones.GuardarArchivo "PersoAnim0",TablaAnimacionPersonajes_319F
        'nose = modFunciones.Bytes2AsciiHex(TablaCaracteristicasPersonajes_3036)
        
        '25E7
        If CambioPantalla_2DB8 Then
            DibujarPantalla_19D8 'si hay que redibujar la pantalla
            PintarPantalla_0DFD = True 'modifica una instrucci�n de las rutinas de las puertas indicando que pinta la pantalla

        Else
            PintarPantalla_0DFD = False
        End If
        'GuardarArchivo "D:\datos\vbasic\Abadia\Abadia2\Buffertiles", BufferTiles_8D80
        
        PunteroPersonajeHL = &H2BAE 'hl apunta a la tabla de guillermo
        
        'AvanzarPersonaje_27CB 0, &H3036, nose1, nose2
        '25fe
        ActualizarDatosPersonaje_291D PunteroPersonajeHL 'comprueba si guillermo puede moverse a donde quiere y actualiza su sprite y el buffer de alturas
        
        
    '2604
        CambioPantalla_2DB8 = False 'indica que no hay que redibujar la pantalla
        
    '260b
        ModificarCaracteristicasSpriteLuz_26A3 'modifica las caracter�sticas del sprite de la luz si puede ser usada por adso
        
    '2627
        DibujarSprites_2674 'dibuja los sprites
    If Not Depuracion.QuitarRetardos Then
        For Contador = 1 To 10
            DoEvents
            DoEvents
            Sleep (10)
        Next
    End If
    frmPrincipal.TxOrientacion = Hex$(TablaCaracteristicasPersonajes_3036(1))
    frmPrincipal.TxX = "&H" + Hex$(TablaCaracteristicasPersonajes_3036(2))
    frmPrincipal.TxY = "&H" + Hex$(TablaCaracteristicasPersonajes_3036(3))
    frmPrincipal.TxZ = "&H" + Hex$(TablaCaracteristicasPersonajes_3036(4))
    frmPrincipal.TxEscaleras = "&H" + Hex$(TablaCaracteristicasPersonajes_3036(5))
    If Parar Then
        Parar = False
        MsgBox "Parado"
        Exit Do
    End If
    '2632
    Loop
    Parado = True
    Exit Sub

    modFunciones.GuardarArchivo "Buffer0", TablaBufferAlturas_01C0 '&H23F
    modFunciones.GuardarArchivo "BufferTiles0", BufferTiles_8D80 '&H77f
    modFunciones.GuardarArchivo "Sprites0", TablaSprites_2E17 '&H1CC
    modFunciones.GuardarArchivo "Puertas0", TablaDatosPuertas_2FE4 '&H23
    modFunciones.GuardarArchivo "Objetos0", TablaPosicionObjetos_3008 '&H2D
    modFunciones.GuardarArchivo "Perso0", TablaCaracteristicasPersonajes_3036 '&H59
    modFunciones.GuardarArchivo "PersoAnim0", TablaAnimacionPersonajes_319F '&H5F
    modFunciones.GuardarArchivo "BufferSprites", BufferSprites_9500 '&H77F
    modFunciones.GuardarArchivo "Graficos0", TablaGraficosObjetos_A300 '&H858
    modFunciones.GuardarArchivo "CGA0", PantallaCGA '&H2000
End Sub

Sub BuclePrincipal_Check()
    Dim PunteroPersonajeHL As Long
    'el bucle principal del juego empieza aqu�
    
    'coloca a Guillermo en posici�n
    TablaCaracteristicasPersonajes_3036(&H3036 + 1 - &H3036) = CheckOrientacion
    TablaCaracteristicasPersonajes_3036(&H3036 + 2 - &H3036) = CheckX
    TablaCaracteristicasPersonajes_3036(&H3036 + 3 - &H3036) = CheckY
    TablaCaracteristicasPersonajes_3036(&H3036 + 4 - &H3036) = CheckZ
    TablaCaracteristicasPersonajes_3036(&H3036 + 5 - &H3036) = CheckEscaleras
    
    Parado = False
    
    'contenido del bucle principal
    ContadorAnimacionGuillermo_0990 = TablaCaracteristicasPersonajes_3036(&H3036 - &H3036) 'obtiene el contador de la animaci�n de guillermo
    '25e4
    ComprobarCambioPantalla_2355 'comprueba si el personaje que se muestra ha cambiado de pantalla y si es as� hace muchas cosas
    '25E7
    If CambioPantalla_2DB8 Then
        DibujarPantalla_19D8 'si hay que redibujar la pantalla
        PintarPantalla_0DFD = True 'modifica una instrucci�n de las rutinas de las puertas indicando que pinta la pantalla

    Else
        PintarPantalla_0DFD = False
    End If
    PunteroPersonajeHL = &H2BAE 'hl apunta a la tabla de guillermo
    '25fe
    ActualizarDatosPersonaje_291D PunteroPersonajeHL 'comprueba si guillermo puede moverse a donde quiere y actualiza su sprite y el buffer de alturas
'2604
    CambioPantalla_2DB8 = False 'indica que no hay que redibujar la pantalla
'260b
    ModificarCaracteristicasSpriteLuz_26A3 'modifica las caracter�sticas del sprite de la luz si puede ser usada por adso
'2627
    DibujarSprites_2674 'dibuja los sprites
    
    frmPrincipal.TxOrientacion = Hex$(TablaCaracteristicasPersonajes_3036(1))
    frmPrincipal.TxX = "&H" + Hex$(TablaCaracteristicasPersonajes_3036(2))
    frmPrincipal.TxY = "&H" + Hex$(TablaCaracteristicasPersonajes_3036(3))
    frmPrincipal.TxZ = "&H" + Hex$(TablaCaracteristicasPersonajes_3036(4))
    frmPrincipal.TxEscaleras = "&H" + Hex$(TablaCaracteristicasPersonajes_3036(5))
    '2632
    Parado = True
    Check = False
    modFunciones.GuardarArchivo CheckRuta + CheckPantalla + ".alt", TablaBufferAlturas_01C0 '&H23F
    modFunciones.GuardarArchivo CheckRuta + CheckPantalla + ".til", BufferTiles_8D80 '&H77f
    modFunciones.GuardarArchivo CheckRuta + CheckPantalla + ".tsp", TablaSprites_2E17 '&H1CC
    modFunciones.GuardarArchivo CheckRuta + CheckPantalla + ".pue", TablaDatosPuertas_2FE4 '&H23
    modFunciones.GuardarArchivo CheckRuta + CheckPantalla + ".obj", TablaPosicionObjetos_3008 '&H2D
    modFunciones.GuardarArchivo CheckRuta + CheckPantalla + ".per", TablaCaracteristicasPersonajes_3036 '&H59
    modFunciones.GuardarArchivo CheckRuta + CheckPantalla + ".ani", TablaAnimacionPersonajes_319F '&H5F
    modFunciones.GuardarArchivo CheckRuta + CheckPantalla + ".bsp", BufferSprites_9500 '&H77F
    modFunciones.GuardarArchivo CheckRuta + CheckPantalla + ".gra", TablaGraficosObjetos_A300 '&H858
    modFunciones.GuardarArchivo CheckRuta + CheckPantalla + ".mon", DatosMonjes_AB59 '&H8A6
    modFunciones.GuardarArchivo CheckRuta + CheckPantalla + ".cga", PantallaCGA '&H2000
End Sub

Sub ComprobarCambioPantalla_2355()

'comprueba si el personaje que se muestra ha cambiado de pantalla y si es as�, obtiene los datos de alturas de la nueva pantalla,
'modifica los valores de las posiciones del motor ajustados para la nueva pantalla, inicia los sprites de las puertas y de los objetos
'del juego con la orientaci�n de la pantalla actual y modifica los sprites de los personajes seg�n la orientaci�n de pantalla
Dim Cambios As Byte 'inicialmente no ha habido cambios
Dim PosicionX As Byte
Dim PosicionY As Byte
Dim PosicionZ As Byte
Dim AlturaBase As Byte
Dim PosX As Byte 'parte alta de la posici�n en X del personaje actual (en los 4 bits inferiores)
Dim PosY As Byte 'parte alta de la posici�n en Y del personaje actual (en los 4 bits inferiores)
Dim PunteroHabitacion As Long
Dim PantallaActual As Byte
Dim PunteroDatosPersonajesHL As Long
Dim PunteroSpritePersonajeIX As Long 'direcci�n del sprite asociado al personaje
Dim PunteroDatosPersonajeIY As Long 'direcci�n a los datos de posici�n del personaje asociado al sprite
Dim PunteroRutinaScriptPersonaje As Long 'direcci�n de la rutina en la que el personaje piensa
Dim ValorBufferAlturas As Byte 'valor a poner en las posiciones que ocupa el personaje en el buffer de alturas
PosicionX = TablaCaracteristicasPersonajes_3036(PunteroDatosPersonajeActual_2D88 + 2 - &H3036) 'lee la posici�n en X del personaje actual
'2361
PosicionX = PosicionX And &HF0
If PosicionX <> PosicionXPersonajeActual_2D75 Then 'posici�n X ha cambiado
'2366
    Cambios = Cambios + 1 'indica el cambio
    PosicionXPersonajeActual_2D75 = PosicionX 'actualiza la posici�n de la pantalla actual
    LimiteInferiorVisibleX_2AE1 = PosicionX - 12 'limite inferior visible de X
End If
PosicionY = TablaCaracteristicasPersonajes_3036(PunteroDatosPersonajeActual_2D88 + 3 - &H3036) 'lee la posici�n en Y del personaje actual
PosicionY = PosicionY And &HF0
If PosicionY <> PosicionYPersonajeActual_2D76 Then 'posici�n Y ha cambiado
'2376
    Cambios = Cambios + 1 'indica el cambio
    PosicionYPersonajeActual_2D76 = PosicionY 'actualiza la posici�n de la pantalla actual
    LimiteInferiorVisibleY_2AEB = PosicionY - 12 'limite inferior visible de y
End If
'237D
PosicionZ = TablaCaracteristicasPersonajes_3036(PunteroDatosPersonajeActual_2D88 + 4 - &H3036) 'lee la posici�n en Z del personaje actual
'2381
AlturaBase = LeerAlturaBasePlanta_2473(PosicionZ) 'dependiendo de la altura, devuelve la altura base de la planta
'2384
If AlturaBase <> PosicionZPersonajeActual_2D77 Then 'altura Z ha cambiado
'2388
    AlturaBasePlantaActual_2AF9 = AlturaBase 'altura base de la planta
'238B
    PosicionZPersonajeActual_2D77 = AlturaBase
'238C
    Cambios = Cambios + 1 'indica el cambio
    Select Case AlturaBase
        Case Is = 0
            PunteroPlantaActual_23EF = &H2255 'apunta a los datos de la planta baja
        Case Is = &HB
            PunteroPlantaActual_23EF = &H22E5 'apunta a los datos de la primera planta
        Case Else
            PunteroPlantaActual_23EF = &H22ED 'apunta a los datos de la segunda planta
    End Select
End If
'23A0
If Cambios = 0 Then Exit Sub ' si no ha habido ning�n cambio de pantalla, sale
'23A3
CambioPantalla_2DB8 = True 'indica que ha habido un cambio de pantalla
'23A6
'averigua si es una habitaci�n iluminada o no
HabitacionOscura_156C = False
If PosicionZPersonajeActual_2D77 = &H16 Then 'si est� en la segunda planta
    'en la segunda planta las habitaciones iluminadas son la 67, 73 y 72
    If PosicionXPersonajeActual_2D75 >= &H20 Then '67 y 73 tienen x<&H20
        If PosicionYPersonajeActual_2D76 <> &H60 Then '60 tiene Y=&H60
            HabitacionOscura_156C = True 'la pantalla no est� iluminada
        End If
    End If
End If
If Depuracion.Luz = EnumTipoLuz_ON Then
    HabitacionOscura_156C = False '###depuraci�n
ElseIf Depuracion.Luz = EnumTipoLuz_Off Then
    HabitacionOscura_156C = True '###depuraci�n
End If
'23C9
'aqu� se llega con HabitacionIluminada_156C a true o false
TablaSprites_2E17(&H2FCF - &H2E17) = &HFE 'marca el sprite de la luz como no visible
'23CE
PosX = (PosicionXPersonajeActual_2D75 And &HF0) / 16 'pone en los 4 bits menos significativos de Valor los 4 bits m�s significativos de PosicionX
PosY = (PosicionYPersonajeActual_2D76 And &HF0) / 16 'pone en los 4 bits menos significativos de Valor los 4 bits m�s significativos de PosicionY
OrientacionPantalla_2481 = (((PosY And &H1) Xor (PosX And &H1)) Or ((PosX And &H1) * 2))
PunteroHabitacion = ((PosicionYPersonajeActual_2D76 And &HF0) Or PosX) + PunteroPlantaActual_23EF '(Y, X) (desplazamiento dentro del mapa de la planta)
'23F2
PantallaActual = TablaHabitaciones_2255(PunteroHabitacion - &H2255) 'lee la pantalla actual
frmPrincipal.TxNumeroHabitacion.Text = "&H" + Hex$(PantallaActual)
'23F3
GuardarDireccionPantalla_2D00 PantallaActual 'guarda en 0x156a-0x156b la direcci�n de los datos de la pantalla actual
'23F6
RellenarBufferAlturas_2D22 PunteroDatosPersonajeActual_2D88 'rellena el buffer de alturas con los datos leidos de la abadia y recortados para la pantalla actual
'23F9
PunteroTablaDesplazamientoAnimacion_2D84 = shl(OrientacionPantalla_2481, 6) 'coloca la orientaci�n en los 2 bits superiores para indexar en la tabla (cada entrada son 64 bytes)
PunteroTablaDesplazamientoAnimacion_2D84 = PunteroTablaDesplazamientoAnimacion_2D84 + &H309F 'apunta a la tabla para el c�lculo del desplazamiento seg�n la animaci�n de una entidad del juego
'tabla de rutinas a llamar en 0x2add seg�n la orientaci�n de la c�mara
'225D:
'    248A 2485 248B 2494
Select Case OrientacionPantalla_2481
    Case Is = 0
        RutinaCambioCoordenadas_2B01 = &H248A
    Case Is = 1
        RutinaCambioCoordenadas_2B01 = &H2485
    Case Is = 2
        RutinaCambioCoordenadas_2B01 = &H248B
    Case Is = 3
        RutinaCambioCoordenadas_2B01 = &H2494
End Select
'241A
InicializarSpritesPuertas_0D30 'inicia los sprites de las puertas del juego para la habitaci�n actual
'241D
InicializarObjetos_0D23 'inicia los sprites de los objetos del juego para la habitaci�n actual
'2420
PunteroDatosPersonajesHL = &H2BAE 'apunta a la tabla con datos para los sprites de los personajes
Dim DE As String
Dim HL As String
Do
'2423
    PunteroSpritePersonajeIX = Leer16(TablaDatosPersonajes_2BAE, PunteroDatosPersonajesHL - &H2BAE) 'direcci�n del sprite asociado al personaje
    DE = Hex$(PunteroSpritePersonajeIX)
    If PunteroSpritePersonajeIX = &HFFFF& Then Exit Sub
    'mientras no lea 0xff, contin�a
'242a
    PunteroDatosPersonajeIY = Leer16(TablaDatosPersonajes_2BAE, PunteroDatosPersonajesHL + 2 - &H2BAE) 'direcci�n a los datos de posici�n del personaje asociado al sprite
    HL = Hex$(PunteroDatosPersonajesHL + 2)
    DE = Hex$(PunteroDatosPersonajeIY)
'242f
    'la rutina de script no se usa
    'PunteroRutinaScriptPersonaje = Leer16(TablaDatosPersonajes_2BAE, PunteroDatosPersonajesHL + 4 - &H2BAE) 'direcci�n de la rutina en la que el personaje piensa
    'HL = Hex$(PunteroDatosPersonajesHL + 4)
    'DE = Hex$(PunteroRutinaScriptPersonaje)
'2436
    PunteroRutinaFlipPersonaje_2A59 = Leer16(TablaDatosPersonajes_2BAE, PunteroDatosPersonajesHL + 6 - &H2BAE) 'rutina a la que llamar si hay que flipear los gr�ficos
    HL = Hex$(PunteroDatosPersonajesHL + 6)
    DE = Hex$(PunteroRutinaFlipPersonaje_2A59)
'2441
    PunteroTablaAnimacionesPersonaje_2A84 = Leer16(TablaDatosPersonajes_2BAE, PunteroDatosPersonajesHL + 8 - &H2BAE) 'direcci�n de la tabla de animaciones para el personaje
    HL = Hex$(PunteroDatosPersonajesHL + 8)
    DE = Hex$(PunteroTablaAnimacionesPersonaje_2A84)
'2449
    ProcesarPersonaje_2468 PunteroSpritePersonajeIX, PunteroDatosPersonajeIY, PunteroDatosPersonajesHL + &HA 'procesa los datos del personaje para cambiar la animaci�n y posici�n del sprite e indicar si es visible o no
'2455
    ValorBufferAlturas = TablaCaracteristicasPersonajes_3036(PunteroDatosPersonajeIY + &HE - &H3036) 'valor a poner en las posiciones que ocupa el personaje en el buffer de alturas
'2458
    RellenarBufferAlturasPersonaje_28EF PunteroDatosPersonajeIY, ValorBufferAlturas
'245B
    PunteroDatosPersonajesHL = PunteroDatosPersonajesHL + 10 'pasa al siguiente personaje
    DoEvents
Loop
End Sub

Function LeerAlturaBasePlanta_2473(ByVal PosicionZ As Byte) As Byte
'dependiendo de la altura indicada, devuelve la altura base de la planta
If PosicionZ < 13 Then
    LeerAlturaBasePlanta_2473 = 0 'si la altura es < 13 sale con 0 (00-12 -> planta baja)
ElseIf PosicionZ >= 24 Then
    LeerAlturaBasePlanta_2473 = 22 'si la altura es >= 24 sale con b = 22 (24- -> segunda planta)
Else
    LeerAlturaBasePlanta_2473 = 11 'si la altura es >= 13 y < 24 sale con b = 11 (13-23 -> primera planta)
End If
End Function

Sub GuardarDireccionPantalla_2D00(ByVal NumeroPantalla As Byte)
'guarda en 0x156a-0x156b la direcci�n de los datos de la pantalla a
Dim PunteroDatosPantalla As Long
Dim Tama�oPantalla As Byte
Dim Contador As Long
NumeroPantallaActual_2DBD = NumeroPantalla 'guarda la pantalla actual
PunteroDatosPantalla = 0
If NumeroPantalla <> 0 Then 'si la pantalla actual  est� definida (o no es la n�mero 0)
    For Contador = 1 To NumeroPantalla
        Tama�oPantalla = DatosHabitaciones_4000(PunteroDatosPantalla)
        PunteroDatosPantalla = PunteroDatosPantalla + Tama�oPantalla
    Next
End If
PunteroPantallaActual_156A = PunteroDatosPantalla
End Sub

Sub RellenarBufferAlturas_2D22(ByVal PunteroDatosPersonaje As Long)
'rellena el buffer de alturas indicado por 0x2d8a con los datos leidos de abadia7 y recortados para la pantalla del personaje que se le pasa en iy
Dim Contador As Long
Dim AlturaBase As Byte 'altura base de la planta
Dim PunteroAlturasPantalla As Long
For Contador = 0 To &H23F
    TablaBufferAlturas_01C0(Contador) = 0 'limpia 576 bytes (24x24) = (4 + 16 + 4)x2
Next
AlturaBase = CalcularMinimosVisibles_0B8F(PunteroDatosPersonaje) 'calcula los m�nimos valores visibles de pantalla para la posici�n del personaje
Select Case AlturaBase
    Case Is = 0
        PunteroAlturasPantalla = &H4A00 'valores de altura de la planta baja
    Case Is = &HB
        PunteroAlturasPantalla = &H4F00 'valores de altura de la primera planta
    Case Else
        PunteroAlturasPantalla = &H5080 'valores de altura de la segunda planta
End Select
RellenarBufferAlturas_3945_3973 PunteroAlturasPantalla 'rellena el buffer de pantalla con los datos leidos de la abadia recortados para la pantalla actual
'GuardarArchivo "BufferAlturas", BufferAlturas
End Sub

Function CalcularMinimosVisibles_0B8F(ByVal PunteroDatosPersonaje As Long) As Byte
'dada la posici�n de un personaje, calcula los m�nimos valores visibles de pantalla y devuelve la altura base de la planta
Dim PosicionX As Byte
Dim PosicionY As Byte
Dim Altura As Byte
PosicionX = TablaCaracteristicasPersonajes_3036(PunteroDatosPersonaje + 2 - &H3036) 'lee la posici�n en x del personaje
PosicionX = (PosicionX And &HF0) - 4 'se queda con la m�nima posici�n visible en X de la parte m�s significativa
MinimaPosicionXVisible_27A9 = PosicionX
PosicionY = TablaCaracteristicasPersonajes_3036(PunteroDatosPersonaje + 3 - &H3036) 'lee la posici�n en y del personaje
PosicionY = (PosicionY And &HF0) - 4 'se queda con la m�nima posici�n visible en y de la parte m�s significativa
MinimaPosicionYVisible_279D = PosicionY
Altura = TablaCaracteristicasPersonajes_3036(PunteroDatosPersonaje + 4 - &H3036) 'lee la altura del personaje
MinimaAlturaVisible_27BA = LeerAlturaBasePlanta_2473(Altura) 'dependiendo de la altura, devuelve la altura base de la planta
AlturaBasePlantaActual_2DBA = MinimaAlturaVisible_27BA
CalcularMinimosVisibles_0B8F = MinimaAlturaVisible_27BA
End Function

Sub RellenarBufferAlturas_3945_3973(ByVal PunteroAlturasPantalla As Long)
'rellena el buffer de pantalla con los datos leidos de la abadia recortados para la pantalla actual
'entradas:
'    byte 0
'        bits 7-4: valor inicial de altura
'        bit 3: si es 0, entrada de 4 bytes. si es 1, entrada de 5 bytes
'        bit 2-0: tipo de elemento de la pantalla
'            si es 0, 6 o 7, sale
'            si es del 1 al 4 recorta (altura cambiante)
'            si es 5, recorta (altura constante)
'    byte 1: coordenada X de inicio
'    byte 2: coordenada Y de inicio
'    byte 3: si longitud == 4 bytes
'        bits 7-4: n�mero de unidades en X
'        bits 3-0: n�mero de unidades en Y
'            si longitud == 5 bytes
'        bits 7-0: n�mero de unidades en X
'    byte 4 n�mero de unidades en Y
Dim Byte0 As Byte
Dim Byte1 As Byte
Dim Byte2 As Byte
Dim Byte3 As Byte
Dim Byte4 As Byte
Dim X As Byte 'coordenada X de inicio
Dim Y As Byte 'coordenada Y de inicio
Dim Z As Byte 'valor inicial de altura
Dim nX As Byte 'n�mero de unidades en X
Dim nY As Byte 'n�mero de unidades en Y
Dim Xrecortada As Byte
Dim Yrecortada As Byte
Dim PunteroBufferAlturas As Long
Dim Ancho As Long
Dim Alto As Long
Do
    Byte0 = TablaAlturasPantallas_4A00(PunteroAlturasPantalla - &H4A00) 'lee un byte
    If Byte0 = &HFF Then Exit Sub 'si ha llegado al final de los datos, sale
    If (Byte0 And &H7) = 0 Then Exit Sub 'si los 3 bits menos significativos del byte leido son 0, sale
    If (Byte0 And &H7) >= 6 Then Exit Sub 'si el (dato & 0x07) >= 0x06, sale
    Byte3 = TablaAlturasPantallas_4A00(PunteroAlturasPantalla + 3 - &H4A00) 'lee un byte
    If (Byte0 And &H8) = 0 Then 'si la entrada es de 4 bytes
        nY = Byte3 And &HF
        nX = CByte(shr(Byte3, 4)) And &HF 'a = 4 bits m�s significativos del byte 3
    Else ' si la entrada es de 5 bytes, lee el �ltimo byte
        Byte4 = TablaAlturasPantallas_4A00(PunteroAlturasPantalla + 4 - &H4A00) 'lee el �ltimo byte
        nX = Byte3
        nY = Byte4
    End If
    Z = CByte(shr(Byte0, 4)) And &HF 'obtiene los 4 bits superiores del byte 0
    Byte1 = TablaAlturasPantallas_4A00(PunteroAlturasPantalla + 1 - &H4A00) 'lee un byte
    Byte2 = TablaAlturasPantallas_4A00(PunteroAlturasPantalla + 2 - &H4A00) 'lee un byte
    X = Byte1
    Y = Byte2
    If (Byte0 And &H8) <> 0 Then 'si la entrada es de 5 bytes
        PunteroAlturasPantalla = PunteroAlturasPantalla + 1
    End If
    PunteroAlturasPantalla = PunteroAlturasPantalla + 4
    nX = nX + 1
    nY = nY + 1
    'If X >= MinimaPosicionXVisible_27A9 Then
    '    If X >= &H18 Then
    '        salta
    '    End If
    'Else
    '    If (X + nX) >= MinimaPosicionXVisible_27A9 Then sigue
    'End If
    'comprueba si se ve en x
'39b5
    If (X >= MinimaPosicionXVisible_27A9 And X < (&H18 + MinimaPosicionXVisible_27A9)) Or ((X < MinimaPosicionXVisible_27A9) And (X + nX) >= MinimaPosicionXVisible_27A9) Then
        'comprueba si se ve en y
'39c8
        If (Y >= MinimaPosicionYVisible_279D And Y < (&H18 + MinimaPosicionYVisible_279D)) Or ((Y < MinimaPosicionYVisible_279D) And (Y + nY) >= MinimaPosicionYVisible_279D) Then
            'si entra aqu�, es porque algo de la entrada es visible
'39d8
            If (Byte0 And &H7) = 5 Then 'si es 5, recorta (altura constante)
                'a partir de aqu�, X e Y son incrementos respecto del borde de la pantalla
'39ee
                If X >= MinimaPosicionXVisible_27A9 Then
'39ff
                    X = X - MinimaPosicionXVisible_27A9
                    If (X + nX) >= &H18 Then nX = &H18 - X
                Else
'39f3
                    If (X + nX - MinimaPosicionXVisible_27A9) > &H18 Then 'si la �ltima coordenada X > limite superior en X
                        nX = &H18
                    Else 'si la �ltima coordenada X <= limite superior en X, salta
                        nX = X + nX - MinimaPosicionXVisible_27A9
                    End If
                    X = 0
                End If
                'pasa a recortar en Y
'3a09
                If Y >= MinimaPosicionYVisible_279D Then 'si la coordenada Y > limite inferior en Y, salta
'3a1a
                    Y = Y - MinimaPosicionYVisible_279D
                    If (Y + nY) >= &H18 Then nY = &H18 - Y
                Else
'3a0e
                    If (Y + nY - MinimaPosicionYVisible_279D) > &H18 Then 'si la �ltima coordenada y > limite superior en y
                        nY = &H18
                    Else 'si la �ltima coordenada y <= limite superior en y, salta
                        nY = Y + nY - MinimaPosicionYVisible_279D
                    End If
                    Y = 0
                End If
 '3a24
                'aqu� llega la entrada una vez que ha sido recortada en X y en Y
                'X = posici�n inicial en X
                'Y = posici�n inicial en Y
                'nX = n�mero de elementos a dibujar en X
                'nY = n�mero de elementos a dibujar en Y
                For Alto = 0 To nY - 1
                    For Ancho = 0 To nX - 1
                        PunteroBufferAlturas = 24 * (Y + Alto) + X + Ancho 'cada l�nea ocupa 24 bytes
                        TablaBufferAlturas_01C0(PunteroBufferAlturas) = Z
                    Next
                Next
            Else 'si es del 1 al 4 recorta (altura cambiante)
'39DF
                RellenarAlturas_38FD X, Y, Z, nX, nY, Byte0 And &H7
            End If
        End If
    End If
    DoEvents
Loop
End Sub

Sub RellenarAlturas_38FD(ByVal X As Byte, ByVal Y As Byte, ByVal Z As Byte, ByVal nX As Byte, ByVal nY As Byte, ByVal TipoIncremento As Byte)
'rutina para rellenar alturas
'X(L)=posicion X inicial
'Y(H)=posicion Y inicial
'Z(a)=valor de la altura inicial de bloque
'nX(c)=n�mero de unidades en X
'nY(b)=n�mero de unidades en Y
Dim Incremento1 As Long
Dim Incremento2 As Long
Dim Alto As Long
Dim Ancho As Long
Dim Altura As Long
Dim AlturaAnterior As Byte
'On Error Resume Next
'tabla de instrucciones para modificar un bucle del c�lculo de alturas
'38EF:   00 00 -> 0 nop, nop (caso imposible)
'        3C 00 -> 1 inc a, nop
'        00 3D -> 2 nop, dec a
'        3D 00 -> 3 dec a, nop
'        00 3C -> 4 nop, inc a
'        00 00 -> 5 nop, nop (caso imposible)
Select Case TipoIncremento
    Case Is = 0 'caso imposible
        Incremento1 = 0
        Incremento2 = 0
    Case Is = 1
        Incremento1 = 1
        Incremento2 = 0
    Case Is = 2
        Incremento1 = 0
        Incremento2 = -1
    Case Is = 3
        Incremento1 = -1
        Incremento2 = 0
    Case Is = 4
        Incremento1 = 0
        Incremento2 = 1
    Case Is = 5 'caso imposible
        Incremento1 = 0
        Incremento2 = 0
End Select
Altura = Z
For Alto = 0 To nY - 1
    AlturaAnterior = Altura
    For Ancho = 0 To nX - 1
        If Altura >= 0 Then
            EscribirAlturaBufferAlturas_391D X + Ancho, Y + Alto, CByte(Altura)
        Else
            EscribirAlturaBufferAlturas_391D X + Ancho, Y + Alto, CByte(256 + Altura)
        End If
        Altura = Altura + Incremento1
    Next
'3915
    Altura = AlturaAnterior + Incremento2
Next
End Sub

Sub EscribirAlturaBufferAlturas_391D(ByVal X As Byte, ByVal Y As Byte, ByVal Z As Byte)
'si la posici�n X (L) ,Y (H) est� dentro del buffer, lo modifica con la altura Z (C)
Dim PunteroBufferAlturas As Long
Dim XAjustada As Long
Dim YAjustada As Long
YAjustada = CLng(Y) - MinimaPosicionYVisible_279D 'ajusta la coordenada al principio de lo visible en Y
'3920
If YAjustada < 0 Then Exit Sub 'si no es visible, sale
'3921
If (YAjustada - &H18) >= 0 Then Exit Sub 'si no es visible, sale
'3924
PunteroBufferAlturas = 24 * YAjustada
'392f
PunteroBufferAlturas = PunteroBufferAlturas + PunteroBufferAlturas_2D8A
'3936
XAjustada = CLng(X) - MinimaPosicionXVisible_27A9
'3939
If XAjustada < 0 Then Exit Sub 'si no es visible, sale
'393a
If (XAjustada - &H18) >= 0 Then Exit Sub 'si no es visible, sale
'393d
PunteroBufferAlturas = PunteroBufferAlturas + XAjustada
TablaBufferAlturas_01C0(PunteroBufferAlturas - &H1C0) = Z
'If Y < MinimaPosicionYVisible_279D Or Y > (&H18 + MinimaPosicionYVisible_279D) Then Exit Sub 'si no es visible, sale
'If X < MinimaPosicionXVisible_27A9 Or X > (&H18 + MinimaPosicionXVisible_27A9) Then Exit Sub 'si no es visible, sale
'PunteroBufferAlturas = 24 * Y + X + PunteroBufferAlturas_2D8A
'TablaBufferAlturas_01C0(PunteroBufferAlturas) = Z
End Sub

Sub InicializarObjetos_0D23()
Dim PunteroRutinaProcesarObjetos As Long
Dim PunteroSpritesObjetos As Long
Dim PunteroDatosObjetos As Long
PunteroRutinaProcesarObjetos = &HDBB 'rutina a la que saltar para procesar los objetos del juego
PunteroSpritesObjetos = &H2F1B 'apunta a los sprites de los objetos del juego
PunteroDatosObjetos = &H3008 'apunta a los datos de posici�n de los objetos del juego
ProcesarObjetos_0D3B PunteroRutinaProcesarObjetos, PunteroSpritesObjetos, PunteroDatosObjetos, False
End Sub

Sub InicializarSpritesPuertas_0D30()
Dim PunteroRutinaProcesarPuertas As Long
Dim PunteroSpritesPuertas As Long
Dim PunteroDatosPuertas As Long
PunteroRutinaProcesarPuertas = &HDD2 'rutina a la que saltar para procesar los sprites de las puertas
PunteroSpritesPuertas = &H2E8F 'apunta a los sprites de las puertas
PunteroDatosPuertas = &H2FE4 'apunta a los datos de las puertas
ProcesarObjetos_0D3B PunteroRutinaProcesarPuertas, PunteroSpritesPuertas, PunteroDatosPuertas, False
End Sub

Sub ProcesarObjetos_0D3B(ByVal PunteroRutinaProcesarObjetos As Long, ByVal PunteroSpritesObjetosIX As Long, ByVal PunteroDatosObjetosIY As Long, ByVal ProcesarSoloUno As Boolean)
Dim Valor As Byte
Dim Visible As Boolean
Dim X As Byte
Dim Y As Byte
Dim Z As Byte
Dim Yp As Byte
Dim PunteroSpritesObjetosIXAnterior As Long
Do
    If PunteroDatosObjetosIY < &H3008 Then 'el puntero apunta a la tabla de puertas
        Valor = TablaDatosPuertas_2FE4(PunteroDatosObjetosIY - &H2FE4) 'lee un byte y si encuentra 0xff termina
    Else 'el puntero apunta a la tabla de objetos
        Valor = TablaPosicionObjetos_3008(PunteroDatosObjetosIY - &H3008) 'lee un byte y si encuentra 0xff termina
    End If
    If Valor = &HFF Then Exit Sub
    Visible = ObtenerCoordenadasObjeto_0E4C(PunteroSpritesObjetosIX, PunteroDatosObjetosIY, X, Y, Z, Yp) 'obtiene en X,Y,Z la posici�n en pantalla del objeto. Si no es visible devuelve False
    If Visible Then 'si el objeto es visible, salta a la rutina siguiente
        PunteroSpritesObjetosIXAnterior = PunteroSpritesObjetosIX
        Select Case PunteroRutinaProcesarObjetos
            Case Is = &HDD2 'rutina a la que saltar para procesar los sprites de las puertas
                ProcesarPuertaVisible_0DD2 PunteroSpritesObjetosIX, PunteroDatosObjetosIY, X, Y, Yp
            Case Is = &HDBB 'rutina a la que saltar para procesar los objetos del juego
                ProcesarObjetoVisible_0DBB PunteroSpritesObjetosIX, PunteroDatosObjetosIY, X, Y, Yp
        End Select
        PunteroSpritesObjetosIX = PunteroSpritesObjetosIXAnterior
    End If
    'pone la posici�n actual del sprite como la posici�n antigua
    TablaSprites_2E17(PunteroSpritesObjetosIX + 3 - &H2E17) = TablaSprites_2E17(PunteroSpritesObjetosIX + 1 - &H2E17)
    TablaSprites_2E17(PunteroSpritesObjetosIX + 4 - &H2E17) = TablaSprites_2E17(PunteroSpritesObjetosIX + 2 - &H2E17)
    PunteroDatosObjetosIY = PunteroDatosObjetosIY + 5 'avanza a la siguiente entrada
    PunteroSpritesObjetosIX = PunteroSpritesObjetosIX + &H14 'apunta al siguiente sprite
    If ProcesarSoloUno Then Exit Sub
    DoEvents
Loop
End Sub

Function ObtenerCoordenadasObjeto_0E4C(ByVal PunteroSpritesObjetosIX As Long, ByVal PunteroDatosObjetosIY As Long, ByRef X As Byte, ByRef Y As Byte, ByRef Z As Byte, ByRef Yp As Byte) As Boolean
'devuelve la posici�n la entidad en coordenadas de pantalla. Si no es visible sale con False
'si es visible devuelve en Z la profundidad del sprite y en X,Y devuelve la posici�n en pantalla del sprite
Dim Visible As Boolean
ModificarPosicionSpritePantalla_2B2F = False
Visible = ProcesarObjeto_2ADD(PunteroSpritesObjetosIX, PunteroDatosObjetosIY, X, Y, Z, Yp)
ModificarPosicionSpritePantalla_2B2F = True
If Not Visible Then
    TablaSprites_2E17(PunteroSpritesObjetosIX + 0 - &H2E17) = &HFE 'marca el sprite como no visible
Else
    ObtenerCoordenadasObjeto_0E4C = Visible
End If
End Function

Function LeerDatoObjeto(ByVal PunteroDatosObjeto As Long) As Byte
'devuelve un valor de la tabla TablaDatosPuertas_2FE4 � TablaPosicionObjetos_3008
If PunteroDatosObjeto < &H3008 Then 'el objeto es una puerta
    LeerDatoObjeto = TablaDatosPuertas_2FE4(PunteroDatosObjeto - &H2FE4)
ElseIf PunteroDatosObjeto < &H3036 Then 'objetos del juego
    LeerDatoObjeto = TablaPosicionObjetos_3008(PunteroDatosObjeto - &H3008)
Else 'personajes
    LeerDatoObjeto = TablaCaracteristicasPersonajes_3036(PunteroDatosObjeto - &H3036)
End If
End Function

Function LeerDatoGrafico(ByVal PunteroDatosGrafico As Long) As Byte
'devuelve un valor de la tabla TilesAbadia_6D00, TablaGraficosObjetos_A300 � DatosMonjes_AB59
If PunteroDatosGrafico < &HA300& Then 'TilesAbadia_6D00
    LeerDatoGrafico = TilesAbadia_6D00(PunteroDatosGrafico - &H6D00)
ElseIf PunteroDatosGrafico < &HAB59& Then 'TablaGraficosObjetos_A300
    LeerDatoGrafico = TablaGraficosObjetos_A300(PunteroDatosGrafico - &HA300&)
Else 'DatosMonjes_AB59
    LeerDatoGrafico = DatosMonjes_AB59(PunteroDatosGrafico - &HAB59&)
End If
End Function

Function LeerByteTablaCualquiera(ByVal Puntero As Long) As Byte
    'lee un byte que puede pertenecer a cualquier tabla. usado en los errores de overflow del programa original
    If PunteroPerteneceTabla(Puntero, TablaBufferAlturas_01C0, &H1C0&) Then
        LeerByteTablaCualquiera = TablaBufferAlturas_01C0(Puntero - &H1C0&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaBloquesPantallas_156D, &H156D&) Then
        LeerByteTablaCualquiera = TablaBloquesPantallas_156D(Puntero - &H156D&)
    End If
    If PunteroPerteneceTabla(Puntero, DatosAlturaEspejoCerrado_34DB, &H34DB&) Then
        LeerByteTablaCualquiera = DatosAlturaEspejoCerrado_34DB(Puntero - &H34DB&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaRutinasConstruccionBloques_1FE0, &H1FE0&) Then
        LeerByteTablaCualquiera = TablaRutinasConstruccionBloques_1FE0(Puntero - &H1FE0&)
    End If
    If PunteroPerteneceTabla(Puntero, VariablesBloques_1FCD, &H1FCD&) Then
        LeerByteTablaCualquiera = VariablesBloques_1FCD(Puntero - &H1FCD&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaCaracteristicasMaterial_1693, &H1693&) Then
        LeerByteTablaCualquiera = TablaCaracteristicasMaterial_1693(Puntero - &H1693&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaHabitaciones_2255, &H2255&) Then
        LeerByteTablaCualquiera = TablaHabitaciones_2255(Puntero - &H2255&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaAvancePersonaje4Tiles_284D, &H284D&) Then
        LeerByteTablaCualquiera = TablaAvancePersonaje4Tiles_284D(Puntero - &H284D&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaAvancePersonaje1Tile_286D, &H286D&) Then
        LeerByteTablaCualquiera = TablaAvancePersonaje1Tile_286D(Puntero - &H286D&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaDatosPersonajes_2BAE, &H2BAE&) Then
        LeerByteTablaCualquiera = TablaDatosPersonajes_2BAE(Puntero - &H2BAE&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaVariablesAuxiliares_2D8D, &H2D8D&) Then
        LeerByteTablaCualquiera = TablaVariablesAuxiliares_2D8D(Puntero - &H2D8D&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaPermisosPuertas_2DD9, &H2DD9&) Then
        LeerByteTablaCualquiera = TablaPermisosPuertas_2DD9(Puntero - &H2DD9&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaObjetosPersonajes_2DEC, &H2DEC&) Then
        LeerByteTablaCualquiera = TablaObjetosPersonajes_2DEC(Puntero - &H2DEC&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaSprites_2E17, &H2E17&) Then
        LeerByteTablaCualquiera = TablaSprites_2E17(Puntero - &H2E17&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaDatosPuertas_2FE4, &H2FE4&) Then
        LeerByteTablaCualquiera = TablaDatosPuertas_2FE4(Puntero - &H2FE4&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaDatosPuertas_2FE4, &H2FE4&) Then
        LeerByteTablaCualquiera = TablaDatosPuertas_2FE4(Puntero - &H2FE4&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaPosicionObjetos_3008, &H3008&) Then
        LeerByteTablaCualquiera = TablaPosicionObjetos_3008(Puntero - &H3008&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaCaracteristicasPersonajes_3036, &H3036&) Then
        LeerByteTablaCualquiera = TablaCaracteristicasPersonajes_3036(Puntero - &H3036&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaPunterosCarasMonjes_3097, &H3097&) Then
        LeerByteTablaCualquiera = TablaPunterosCarasMonjes_3097(Puntero - &H3097&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaDesplazamientoAnimacion_309F, &H309F&) Then
        LeerByteTablaCualquiera = TablaDesplazamientoAnimacion_309F(Puntero - &H309F&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaAnimacionPersonajes_319F, &H319F&) Then
        LeerByteTablaCualquiera = TablaAnimacionPersonajes_319F(Puntero - &H319F&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaAccesoHabitaciones_3C67, &H3C67&) Then
        LeerByteTablaCualquiera = TablaAccesoHabitaciones_3C67(Puntero - &H3C67&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaVariablesLogica_3C85, &H3C85&) Then
        LeerByteTablaCualquiera = TablaVariablesLogica_3C85(Puntero - &H3C85&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaPosicionesPredefinidasMalaquias_3CA8, &H3CA8&) Then
        LeerByteTablaCualquiera = TablaPosicionesPredefinidasMalaquias_3CA8(Puntero - &H3CA8&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaPosicionesPredefinidasAbad_3CC6, &H3CC6&) Then
        LeerByteTablaCualquiera = TablaPosicionesPredefinidasAbad_3CC6(Puntero - &H3CC6&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaPosicionesPredefinidasBerengario_3CE7, &H3CE7&) Then
        LeerByteTablaCualquiera = TablaPosicionesPredefinidasBerengario_3CE7(Puntero - &H3CE7&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaPosicionesPredefinidasSeverino_3CFF, &H3CFF&) Then
        LeerByteTablaCualquiera = TablaPosicionesPredefinidasSeverino_3CFF(Puntero - &H3CFF&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaPosicionesPredefinidasAdso_3D11, &H3D11&) Then
        LeerByteTablaCualquiera = TablaPosicionesPredefinidasAdso_3D11(Puntero - &H3D11&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaPunterosVariablesScript_3D1D, &H3D1D&) Then
        LeerByteTablaCualquiera = TablaPunterosVariablesScript_3D1D(Puntero - &H3D1D&)
    End If
    If PunteroPerteneceTabla(Puntero, DatosHabitaciones_4000, &H4000&) Then
        LeerByteTablaCualquiera = DatosHabitaciones_4000(Puntero - &H4000&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaPunterosTrajesMonjes_48C8, &H48C8&) Then
        LeerByteTablaCualquiera = TablaPunterosTrajesMonjes_48C8(Puntero - &H48C8&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaPatronRellenoLuz_48E8, &H48E8&) Then
        LeerByteTablaCualquiera = TablaPatronRellenoLuz_48E8(Puntero - &H48E8&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaAlturasPantallas_4A00, &H4A00&) Then
        LeerByteTablaCualquiera = TablaAlturasPantallas_4A00(Puntero - &H4A00&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaEtapasDia_4F7A, &H4F7A&) Then
        LeerByteTablaCualquiera = TablaEtapasDia_4F7A(Puntero - &H4F7A&)
    End If
    If PunteroPerteneceTabla(Puntero, DatosMarcador_6328, &H6328&) Then
        LeerByteTablaCualquiera = DatosMarcador_6328(Puntero - &H6328&)
    End If
    If PunteroPerteneceTabla(Puntero, DatosCaracteresPergamino_6947, &H6947&) Then
        LeerByteTablaCualquiera = DatosCaracteresPergamino_6947(Puntero - &H6947&)
    End If
    If PunteroPerteneceTabla(Puntero, PunterosCaracteresPergamino_680C, &H680C&) Then
        LeerByteTablaCualquiera = PunterosCaracteresPergamino_680C(Puntero - &H680C&)
    End If
    If PunteroPerteneceTabla(Puntero, TilesAbadia_6D00, &H6D00&) Then
        LeerByteTablaCualquiera = TilesAbadia_6D00(Puntero - &H6D00&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaRellenoBugTiles_8D00, &H8D00&) Then
        LeerByteTablaCualquiera = TablaRellenoBugTiles_8D00(Puntero - &H8D00&)
    End If
    If PunteroPerteneceTabla(Puntero, TextoPergaminoPresentacion_7300, &H7300&) Then
        LeerByteTablaCualquiera = TextoPergaminoPresentacion_7300(Puntero - &H7300&)
    End If
    If PunteroPerteneceTabla(Puntero, DatosGraficosPergamino_788A, &H788A&) Then
        LeerByteTablaCualquiera = DatosGraficosPergamino_788A(Puntero - &H788A&)
    End If
    If PunteroPerteneceTabla(Puntero, BufferTiles_8D80, &H8D80&) Then
        LeerByteTablaCualquiera = BufferTiles_8D80(Puntero - &H8D80&)
    End If
    If PunteroPerteneceTabla(Puntero, BufferSprites_9500, &H9500&) Then
        LeerByteTablaCualquiera = BufferSprites_9500(Puntero - &H9500&)
    End If
    If PunteroPerteneceTabla(Puntero, TablasAndOr_9D00, &H9D00&) Then
        LeerByteTablaCualquiera = TablasAndOr_9D00(Puntero - &H9D00&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaFlipX_A100, &HA100&) Then
        LeerByteTablaCualquiera = TablaFlipX_A100(Puntero - &HA100&)
    End If
    If PunteroPerteneceTabla(Puntero, TablaGraficosObjetos_A300, &HA300&) Then
        LeerByteTablaCualquiera = TablaGraficosObjetos_A300(Puntero - &HA300&)
    End If
    If PunteroPerteneceTabla(Puntero, DatosMonjes_AB59, &HAB59&) Then
        LeerByteTablaCualquiera = DatosMonjes_AB59(Puntero - &HAB59&)
    End If
    If PunteroPerteneceTabla(Puntero, BufferComandosMonjes_A200, &HA200&) Then
        LeerByteTablaCualquiera = BufferComandosMonjes_A200(Puntero - &HA200&)
    End If
End Function



Function ProcesarObjeto_2ADD(ByVal PunteroSpritesObjetosIX As Long, ByVal PunteroDatosObjetosIY As Long, ByRef X As Byte, ByRef Y As Byte, ByRef Z As Byte, ByRef Yp As Byte) As Boolean
'comprueba si el sprite est� dentro de la zona visible de pantalla.
'Si no es as�, sale. Si est� dentro de la zona visible lo transforma
'a otro sistema de coordenadas. Dependiendo de un par�metro sigue o no.
'Si sigue actualiza la posici�n seg�n la orientaci�n
'si no es visible, sale. Si es visible, sale 2 veces (2 pop de pila)
Dim ValorX As Long
Dim ValorY As Long
Dim ValorZ As Byte
Dim AlturaBase As Byte
ValorX = CLng(LeerDatoObjeto(PunteroDatosObjetosIY + 2)) - LimiteInferiorVisibleX_2AE1
ValorY = CLng(LeerDatoObjeto(PunteroDatosObjetosIY + 3)) - LimiteInferiorVisibleY_2AEB
ValorZ = LeerDatoObjeto(PunteroDatosObjetosIY + 4)
If ValorX < 0 Or ValorX > &H28 Then 'si el objeto en X es < limite inferior visible de X o el objeto en X es >= limite superior visible de X, sale
    ProcesarObjeto_2ADD = False
    Exit Function
End If
If ValorY < 0 Or ValorY > &H28 Then 'si el objeto en Y es < limite inferior visible de Y o el objeto en Y es >= limite superior visible de Y, sale
    ProcesarObjeto_2ADD = False
    Exit Function
End If
'2af4
AlturaBase = LeerAlturaBasePlanta_2473(ValorZ) 'dependiendo de la altura, devuelve la altura base de la planta
If AlturaBase <> AlturaBasePlantaActual_2AF9 Then 'si el objeto no est� en la misma planta, sale
    ProcesarObjeto_2ADD = False
    Exit Function
End If
X = CByte(ValorX) 'coordenada X del objeto en la pantalla
Y = CByte(ValorY) 'coordenada Y del objeto en la pantalla
Z = ValorZ - AlturaBase 'altura del objeto ajustada para esta pantalla
'2b00
'al llegar aqu� los par�metros son:
'X = coordenada X del objeto en la rejilla
'Y = coordenada Y del objeto en la rejilla
'Z = altura del objeto en la rejilla ajustada para esta planta
Select Case RutinaCambioCoordenadas_2B01 'rutina que cambia el sistema de coordenadas dependiendo de la orientaci�n de la pantalla
    Case Is = &H248A
        CambiarCoordenadasOrientacion0_248A X, Y
    Case Is = &H2485
        CambiarCoordenadasOrientacion1_2485 X, Y
    Case Is = &H248B
        CambiarCoordenadasOrientacion2_248B X, Y
    Case Is = &H2494
        CambiarCoordenadasOrientacion3_2494 X, Y
End Select
TablaSprites_2E17(PunteroSpritesObjetosIX + &H12 - &H2E17) = X 'graba las nuevas coordenadas x e y en el sprite
TablaSprites_2E17(PunteroSpritesObjetosIX + &H13 - &H2E17) = Y 'graba las nuevas coordenadas x e y en el sprite
'2b09
'convierte las coordenadas en la rejilla a coordenadas de pantalla
Dim Xcalc As Long
Dim Ycalc As Long
Dim Ypantalla As Long
'2b09
Ycalc = X + Y 'pos x + pos y = coordenada y en pantalla
'2B0B
Ypantalla = Ycalc
'2B0C
Ycalc = Ycalc - Z 'le resta la altura (cuanto m�s alto es el objeto, menor y tiene en pantalla)
'2B0D
If Ycalc < 0 Then Exit Function
'2B0E
Ycalc = Ycalc - 6 'y calc = y calc - 6 (traslada 6 unidades arriba)
'2b10
If Ycalc < 0 Then Exit Function 'si y calc < 0, sale
'2b11
If Ycalc < 8 Then Exit Function 'si y calc < 8, sale
'2b14
If Ycalc >= &H3A Then Exit Function 'si y calc  >= 58, sale
'llega aqu� si la y calc est� entre 8 y 57
'2b17
Ycalc = 4 * (Ycalc + 1)
Xcalc = 2 * (CLng(X) - CLng(Y)) + &H50 - &H28
If Xcalc < 0 Then Exit Function
If Xcalc >= &H50 Then Exit Function
'2b26
X = CByte(Xcalc) 'pos x con nuevo sistema de coordenadas
Y = CByte(Ycalc) 'pos y con nuevo sistema de coordenadas
ProcesarObjeto_2ADD = True 'el objeto es visible
Ypantalla = Ypantalla - &H10
If Ypantalla < 0 Then Ypantalla = 0 'si la posici�n en y < 16, pos y = 0
Yp = Long2Byte(Ypantalla)
If Not ModificarPosicionSpritePantalla_2B2F Then Exit Function
'si llega aqu� modifica la posici�n del sprite en pantalla
'2B30
Dim Entrada As Byte
Entrada = 0 'primera entrada
If (LeerDatoObjeto(PunteroDatosObjetosIY + 5) And &H80) <> 0 Then
    GoTo H2B78 'si el personaje ocupa una posici�n
Else
    GoTo H2B3A
End If

H2B3A:
If (LeerDatoObjeto(PunteroDatosObjetosIY + 0) And 1) = 0 Then 'lee el bit 0 del contador de animaci�n '5?
    GoTo H2B41 'si es 1, avanza a la siguiente entrada
Else
    Entrada = Entrada + 1
End If

H2B41:
Dim Puntero As Long
Dim Orientacion As Byte
Dim Desplazamiento As Long
Orientacion = ModificarOrientacion_2480(LeerDatoObjeto(PunteroDatosObjetosIY + 1)) 'obtiene la orientaci�n del personaje. modifica la orientaci�n que se le pasa en a con la orientaci�n de la pantalla actual
'2b4b
Puntero = (shl(Orientacion, 4) And &H30) + 2 * Entrada + PunteroTablaDesplazamientoAnimacion_2D84
'2b58
'Desplazamiento = TablaDesplazamientoAnimacion_309F(Puntero - &H309F) 'lee un byte de la tabla
Desplazamiento = Leer8Signo(TablaDesplazamientoAnimacion_309F, Puntero - &H309F) 'lee un byte de la tabla
'2b59
Desplazamiento = Desplazamiento + X 'le suma la x del nuevo sistema de coordenadas
'2b5a
'Desplazamiento = Desplazamiento - (256 - LeerDatoObjeto(PunteroDatosObjetosIY + 7)) 'le suma un desplazamieno
Desplazamiento = Desplazamiento + Leer8Signo(TablaCaracteristicasPersonajes_3036, PunteroDatosObjetosIY + 7 - &H3036) 'le suma un desplazamieno
If Desplazamiento >= 0 Then
    X = Desplazamiento 'actualiza la x
Else
    X = 256 + Desplazamiento 'no deber�an aparecer coordenadas negativas. bug del original?
End If
Puntero = Puntero + 1
'Desplazamiento = TablaDesplazamientoAnimacion_309F(Puntero - &H309F) 'lee un byte de la tabla
Desplazamiento = Leer8Signo(TablaDesplazamientoAnimacion_309F, Puntero - &H309F) 'lee un byte de la tabla
Desplazamiento = Desplazamiento + Y 'le suma la Y del nuevo sistema de coordenadas
'Desplazamiento = Desplazamiento - (256 - LeerDatoObjeto(PunteroDatosObjetosIY + 8)) 'le suma un desplazamieno
Desplazamiento = Desplazamiento + Leer8Signo(TablaCaracteristicasPersonajes_3036, PunteroDatosObjetosIY + 8 - &H3036) 'le suma un desplazamieno
Y = Desplazamiento 'actualiza la Y
TablaSprites_2E17(PunteroSpritesObjetosIX + 1 - &H2E17) = X 'graba la posici�n x del sprite (en bytes)
TablaSprites_2E17(PunteroSpritesObjetosIX + 2 - &H2E17) = Y 'graba la posici�n y del sprite (en pixels)
If TablaSprites_2E17(PunteroSpritesObjetosIX + 0 - &H2E17) <> &HFE Then Exit Function
'si el sprite no es visible, continua
TablaSprites_2E17(PunteroSpritesObjetosIX + 3 - &H2E17) = X 'graba la posici�n anterior x del sprite (en bytes)
TablaSprites_2E17(PunteroSpritesObjetosIX + 4 - &H2E17) = Y 'graba la posici�n anterior y del sprite (en pixels)
Exit Function

H2B78:
'aqu� llega si el personaje ocupa una posici�n (porque est� en los escalones)
Entrada = Entrada + 2 'avanza a la tercera entrada
If (LeerDatoObjeto(PunteroDatosObjetosIY + 5) And 32) <> 0 Then GoTo H2B3A
Entrada = Entrada + 2 'avanza a la quinta entrada

H2B82:
'aqu� llega si el personaje ocupa una posici�n y est� orientado para subir o bajar las escaleras (ya est� apuntando a la 5� entrada)
If (LeerDatoObjeto(PunteroDatosObjetosIY + 5) And &H3) <> 0 Then GoTo H2B99  'esto nunca pasa???
If (LeerDatoObjeto(PunteroDatosObjetosIY + 0) And &H1) = 0 Then GoTo H2B41
Entrada = Entrada + 1 'avanza una entrada

H2B90:
If (LeerDatoObjeto(PunteroDatosObjetosIY + 5) And &H10) = 0 Then GoTo H2B41
Entrada = Entrada + 1 'avanza una entrada
GoTo H2B41

H2B99:
'??? cuando llega aqu�???
Entrada = Entrada + 3 'avanza a la octava entrada
If (LeerDatoObjeto(PunteroDatosObjetosIY + 5) And &H40) <> 0 Then GoTo H2BA6
Entrada = Entrada + 4 'avanza a la 12� entrada
H2BA6:
If (LeerDatoObjeto(PunteroDatosObjetosIY + 5) And &H3) <> 1 Then GoTo H2B90  'si los bits 0 y 1 de (iy+05) != 1, salta (entrada 12 o 13)
Entrada = Entrada + 2 'avanza a la 14� entrada
GoTo H2B90
End Function

Sub CambiarCoordenadasOrientacion0_248A(ByRef X As Byte, ByRef Y As Byte)
'realiza el cambio de coordenadas si la orientaci�n la c�mara es del tipo 0
'no hace ning�n cambio
End Sub

Sub CambiarCoordenadasOrientacion1_2485(ByRef X As Byte, ByRef Y As Byte)
'realiza el cambio de coordenadas si la orientaci�n la c�mara es del tipo 1
Dim Valor As Byte
Valor = Y 'guarda Y
Y = X
X = &H28 - Valor
End Sub

Sub CambiarCoordenadasOrientacion2_248B(ByRef X As Byte, ByRef Y As Byte)
'realiza el cambio de coordenadas si la orientaci�n la c�mara es del tipo 2
Y = &H28 - Y
X = &H28 - X
End Sub

Sub CambiarCoordenadasOrientacion3_2494(ByRef X As Byte, ByRef Y As Byte)
'realiza el cambio de coordenadas si la orientaci�n la c�mara es del tipo 1
Dim Valor As Byte
Valor = X 'guarda x
X = Y
Y = &H28 - Valor
End Sub

Function ModificarOrientacion_2480(ByVal Orientacion As Byte) As Byte
'modifica la orientaci�n que se le pasa en a con la orientaci�n de la pantalla actual
Dim Resultado As Long
Resultado = (CLng(Orientacion) - OrientacionPantalla_2481) And &H3
ModificarOrientacion_2480 = Long2Byte(Resultado)
'If Orientacion < OrientacionPantalla_2481 Then
'    ModificarOrientacion_2480 = 3
'    Exit Function
'End If
'ModificarOrientacion_2480 = (Orientacion - OrientacionPantalla_2481) And &H3
End Function

Sub ProcesarPuertaVisible_0DD2(ByVal PunteroSpriteIX As Long, ByVal PunteroDatosIY As Long, ByVal X As Byte, ByVal Y As Byte, ByVal Z As Byte)
'rutina llamada cuando las puertas son visibles en la pantalla actual
'se encarga de modificar la posici�n del sprite seg�n la orientaci�n, modificar el buffer de alturas para indicar si se puede pasar
'por la zona de la puerta o no, colocar el gr�fico de las puertas y modificar 0x2daf
'PunteroSprite apunta al sprite de una puerta
'PunteroDatos apunta a los datos de la puerta
'X,Y contienen la posici�n en pantalla del objeto
'Z tiene la profundidad de la puerta en pantalla
Dim DeltaX As Long
Dim DeltaY As Long
Dim DeltaBuffer As Long
Dim Orientacion As Byte
Dim TablaDesplazamientoOrientacionPuertas_05AD(31) As Long
Dim Valor As Long
Dim PunteroBufferAlturasIX As Long
'tabla de desplazamientos relacionada con las orientaciones de las puertas
'cada entrada ocupa 8 bytes
'byte 0: relacionado con la posici�n x de pantalla
'byte 1: relacionado con la posici�n y de pantalla
'byte 2: relacionado con la profundidad de los sprites
'byte 3: indica el estado de flipx de los gr�ficos que necesita la puerta
'byte 4: relacionado con la posici�n x de la rejilla
'byte 5: relacionado con la posici�n y de la rejilla
'byte 6-7: no usado, pero es el desplazamiento en el buffer de alturas
'05AD:   FF DE 01 00 00 00 0001 -> -01 -34  +01  00    00  00   +01
'        FF D6 00 01 00 00 FFE8 -> -01 -42   00 +01    00  00   -24
'        FB D6 00 00 00 00 FFFF -> -05 -42   00  00    00  00   -01
'        FB DE 01 01 01 01 0018 -> -05 -34  +01 +01   +01 +01   +24
TablaDesplazamientoOrientacionPuertas_05AD(0) = -1
TablaDesplazamientoOrientacionPuertas_05AD(1) = -34
TablaDesplazamientoOrientacionPuertas_05AD(2) = 1
TablaDesplazamientoOrientacionPuertas_05AD(7) = 1

TablaDesplazamientoOrientacionPuertas_05AD(8) = -1
TablaDesplazamientoOrientacionPuertas_05AD(9) = -42
TablaDesplazamientoOrientacionPuertas_05AD(11) = 1
TablaDesplazamientoOrientacionPuertas_05AD(14) = -1
TablaDesplazamientoOrientacionPuertas_05AD(15) = -24

TablaDesplazamientoOrientacionPuertas_05AD(16) = -5
TablaDesplazamientoOrientacionPuertas_05AD(17) = -42
TablaDesplazamientoOrientacionPuertas_05AD(22) = -1
TablaDesplazamientoOrientacionPuertas_05AD(23) = -1

TablaDesplazamientoOrientacionPuertas_05AD(24) = -5
TablaDesplazamientoOrientacionPuertas_05AD(25) = -34
TablaDesplazamientoOrientacionPuertas_05AD(26) = 1
TablaDesplazamientoOrientacionPuertas_05AD(27) = 1
TablaDesplazamientoOrientacionPuertas_05AD(28) = 1
TablaDesplazamientoOrientacionPuertas_05AD(29) = 1
TablaDesplazamientoOrientacionPuertas_05AD(31) = 24

DefinirDatosSpriteComoAntiguos_2AB0 PunteroSpriteIX
LeerOrientacionPuerta_0E7C PunteroSpriteIX, DeltaX, DeltaY  'lee 2 valores relacionados con la orientaci�n y modifica la posici�n del sprite (en coordenadas locales) seg�n la orientaci�n
Orientacion = TablaDatosPuertas_2FE4(PunteroDatosIY + 0 - &H2FE4) 'lee la orientaci�n de la puerta
Orientacion = ModificarOrientacion_2480(Orientacion And &H3)  'modifica la orientaci�n que se le pasa con la orientaci�n de la pantalla actual
'0deb
Valor = TablaDesplazamientoOrientacionPuertas_05AD(Orientacion * 8) 'indexa en la tabla
TablaSprites_2E17(PunteroSpriteIX + 1 - &H2E17) = Long2Byte(Valor + DeltaX + Byte2Long(X)) 'modifica la posici�n x del sprite
'0df1
Valor = TablaDesplazamientoOrientacionPuertas_05AD(Orientacion * 8 + 1) 'indexa en la tabla
TablaSprites_2E17(PunteroSpriteIX + 2 - &H2E17) = Long2Byte(Valor + DeltaY + Byte2Long(Y)) 'modifica la posici�n y del sprite
'0df8
Valor = TablaDesplazamientoOrientacionPuertas_05AD(Orientacion * 8 + 2) 'indexa en la tabla
Valor = Valor + Byte2Long(Z)
If PintarPantalla_0DFD Then Valor = Valor Or &H80 'Si se pinta la pantalla, 0x80, en otro caso 0
If RedibujarPuerta_0DFF Then Valor = Valor Or &H80 'Si se pinta la puerta, 0x80, en otro caso 0
'0e00
TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17) = Long2Byte(Valor)
If TablaDesplazamientoOrientacionPuertas_05AD(Orientacion * 8 + 3) <> 0 Then PuertaRequiereFlip_2DAF = True
'modifica la posici�n x e y del sprite en la rejilla seg�n los 2 siguientes valores de la tabla
Valor = TablaDesplazamientoOrientacionPuertas_05AD(Orientacion * 8 + 4) 'indexa en la tabla
TablaSprites_2E17(PunteroSpriteIX + &H12 - &H2E17) = TablaSprites_2E17(PunteroSpriteIX + &H12 - &H2E17) + Valor
Valor = TablaDesplazamientoOrientacionPuertas_05AD(Orientacion * 8 + 5) 'indexa en la tabla
TablaSprites_2E17(PunteroSpriteIX + &H13 - &H2E17) = TablaSprites_2E17(PunteroSpriteIX + &H13 - &H2E17) + Valor
'coloca la direcci�n del gr�fico de la puerta en el sprite (&Haa49)
'0e0e
TablaSprites_2E17(PunteroSpriteIX + 7 - &H2E17) = &H49
TablaSprites_2E17(PunteroSpriteIX + 8 - &H2E17) = &HAA
'si el objeto no es visible, sale. En otro caso, devuelve en ix un puntero a la entrada de la tabla de alturas de la posici�n correspondiente
If Not LeerDesplazamientoPuerta_0E2C(PunteroBufferAlturasIX, PunteroDatosIY, DeltaBuffer) Then Exit Sub
TablaBufferAlturas_01C0(PunteroBufferAlturasIX) = &HF 'marca la altura de esta posici�n del buffer de alturas
TablaBufferAlturas_01C0(PunteroBufferAlturasIX + DeltaBuffer) = &HF 'marca la altura de la siguiente posici�n del buffer de alturas
TablaBufferAlturas_01C0(PunteroBufferAlturasIX + 2 * DeltaBuffer) = &HF 'marca la altura de la siguiente posici�n del buffer de alturas
End Sub

Sub DefinirDatosSpriteComoAntiguos_2AB0(ByVal PunteroSpriteIX As Long)
'pone la posici�n y dimensiones actuales como posici�n y dimensiones antiguas
'copia la posici�n actual en x y en y como la posici�n antigua
TablaSprites_2E17(PunteroSpriteIX + 3 - &H2E17) = TablaSprites_2E17(PunteroSpriteIX + 1 - &H2E17)
TablaSprites_2E17(PunteroSpriteIX + 4 - &H2E17) = TablaSprites_2E17(PunteroSpriteIX + 2 - &H2E17)
'copia el ancho y alto del sprite actual como el ancho y alto antiguos
TablaSprites_2E17(PunteroSpriteIX + 9 - &H2E17) = TablaSprites_2E17(PunteroSpriteIX + 5 - &H2E17)
TablaSprites_2E17(PunteroSpriteIX + 10 - &H2E17) = TablaSprites_2E17(PunteroSpriteIX + 6 - &H2E17)
End Sub

Sub LeerOrientacionPuerta_0E7C(ByVal PunteroSpriteIX As Long, ByRef DeltaX As Long, ByRef DeltaY As Long)
'lee en DeltaX, DeltaY 2 valores relacionados con la orientaci�n y modifica la posici�n del sprite (en coordenadas locales) seg�n la orientaci�n
'PunteroSprite apunta al sprite de una puerta
Dim TablaDesplazamientoOrientacionPuertas_0E9D(15) As Long
Dim Orientacion As Byte
'tabla relacionada con el desplazamiento de las puertas y la orientaci�n
'cada entrada ocupa 4 bytes
'byte 0: valor a sumar a la posici�n x en coordenadas de pantalla del sprite de la puerta
'byte 1: valor a sumar a la posici�n y en coordenadas de pantalla del sprite de la puerta
'byte 2: valor a sumar a la posici�n x en coordenadas locales del sprite de la puerta
'byte 3: valor a sumar a la posici�n y en coordenadas locales del sprite de la puerta
'0E9D:   02 00 00 FF -> +2 00 00 -1
'        00 FC FF FF -> 00 -4 -1 -1
'        FE 00 FF 00 -> -2 00 -1 00
'        00 04 00 00 -> 00 +4 00 00
TablaDesplazamientoOrientacionPuertas_0E9D(0) = 2
TablaDesplazamientoOrientacionPuertas_0E9D(3) = -1

TablaDesplazamientoOrientacionPuertas_0E9D(5) = -4
TablaDesplazamientoOrientacionPuertas_0E9D(6) = -1
TablaDesplazamientoOrientacionPuertas_0E9D(7) = -1

TablaDesplazamientoOrientacionPuertas_0E9D(8) = -2
TablaDesplazamientoOrientacionPuertas_0E9D(10) = -1

TablaDesplazamientoOrientacionPuertas_0E9D(13) = 4

Orientacion = ModificarOrientacion_2480(3) 'modifica la orientaci�n que se le pasa con la orientaci�n de la pantalla actual
'indexa en la tabla. cada entrada ocupa 4 bytes
'lee los valores a sumar a la posici�n en coordenadas de pantalla del sprite de la puerta
DeltaX = TablaDesplazamientoOrientacionPuertas_0E9D(Orientacion * 4)
DeltaY = TablaDesplazamientoOrientacionPuertas_0E9D(Orientacion * 4 + 1)
' modifica la posici�n x de la rejilla seg�n la orientaci�n de la c�mara con el valor leido
TablaSprites_2E17(PunteroSpriteIX + &H12 - &H2E17) = CByte(CLng(TablaSprites_2E17(PunteroSpriteIX + &H12 - &H2E17)) + TablaDesplazamientoOrientacionPuertas_0E9D(Orientacion * 4 + 2))
TablaSprites_2E17(PunteroSpriteIX + &H13 - &H2E17) = CByte(CLng(TablaSprites_2E17(PunteroSpriteIX + &H13 - &H2E17)) + TablaDesplazamientoOrientacionPuertas_0E9D(Orientacion * 4 + 3))
End Sub

Function LeerDesplazamientoPuerta_0E2C(ByRef PunteroBufferAlturasIX As Long, ByVal PunteroDatosIY As Long, ByRef DeltaBuffer As Long) As Boolean
'lee en DeltaBuffer el desplazamiento para el buffer de alturas, y si la puerta es visible devuelve en PunteroBufferAlturasIX un puntero a la entrada de la tabla de alturas de la posici�n correspondiente
'DeltaBuffer=incremento entre posiciones marcadas en el buffer de alturas
'devuelve true si el elemento ocupa una posici�n central
Dim Orientacion As Byte
Dim TablaDesplazamientosBufferPuertas(3) As Long
'tabla de desplazamientos en el buffer de alturas relacionada con la orientaci�n de las puertas
'0E44:   0001 -> +01
'        FFE8 -> -24
'        FFFF -> -01
'        0018 -> +24
TablaDesplazamientosBufferPuertas(0) = 1
TablaDesplazamientosBufferPuertas(1) = -24
TablaDesplazamientosBufferPuertas(2) = -1
TablaDesplazamientosBufferPuertas(3) = 24
Orientacion = LeerDatoObjeto(PunteroDatosIY + 0)  'obtiene la orientaci�n de la puerta
Orientacion = Orientacion And &H3
'Orientacion = Orientacion * 2 'cada entrada ocupa 2 bytes
'DeltaX = TablaDesplazamientosBufferPuertas(Orientacion)
'DeltaY = TablaDesplazamientosBufferPuertas(Orientacion + 1)
DeltaBuffer = TablaDesplazamientosBufferPuertas(Orientacion)
LeerDesplazamientoPuerta_0E2C = DeterminarPosicionCentral_0CBE(PunteroDatosIY, PunteroBufferAlturasIX)
End Function

Function DeterminarPosicionCentral_0CBE(ByVal PunteroDatosIY As Long, ByRef PunteroBufferAlturasIX As Long) As Boolean
'si la posici�n no es una de las del centro de la pantalla o la altura del personaje no coincide con la altura base de la planta, sale con false
'en otro caso, devuelve en PunteroBufferAlturasIX un puntero a la entrada de la tabla de alturas de la posici�n correspondiente
'llamado con PunteroDatosIY = direcci�n de los datos de posici�n asociados al personaje/objeto
Dim Altura As Byte
Dim AlturaBase As Byte
Dim X As Byte
Dim Y As Byte
Altura = LeerDatoObjeto(PunteroDatosIY + 4) 'obtiene la altura del personaje
AlturaBase = LeerAlturaBasePlanta_2473(Altura) 'dependiendo de la altura, devuelve la altura base de la planta
If AlturaBasePlantaActual_2DBA <> AlturaBase Then Exit Function 'si las alturas son distintas, sale con false
X = LeerDatoObjeto(PunteroDatosIY + 2) 'posici�n x del personaje
Y = LeerDatoObjeto(PunteroDatosIY + 3) 'posici�n y del personaje
If Not DeterminarPosicionCentral_279B(X, Y) Then Exit Function 'ajusta la posici�n pasada en X,Y a las 20x20 posiciones centrales que se muestran. Si la posici�n est� fuera, sale
DeterminarPosicionCentral_0CBE = True 'visible
PunteroBufferAlturasIX = 24 * Y + X
End Function

Function DeterminarPosicionCentral_279B(ByRef X As Byte, ByRef Y As Byte) As Boolean
'ajusta la posici�n pasada en X,Y a las 20x20 posiciones centrales que se muestran. Si la posici�n est� fuera, devuelve false
If Y < MinimaPosicionYVisible_279D Then Exit Function 'si la posici�n en y es < el l�mite inferior en y en esta pantalla, sale
Y = Y - MinimaPosicionYVisible_279D 'l�mite inferior en y
If Y < 2 Then Exit Function
If Y >= &H16 Then Exit Function 'si la posici�n en y es > el l�mite superior en y en esta pantalla, sale
If X < MinimaPosicionXVisible_27A9 Then Exit Function ' si la posici�n en x es < el l�mite inferior en x en esta pantalla, sale
X = X - MinimaPosicionXVisible_27A9 'l�mite inferior en x
If X < 2 Then Exit Function
If X >= &H16 Then Exit Function 'si la posici�n en x es > el l�mite superior en x en esta pantalla, sale
DeterminarPosicionCentral_279B = True
End Function

Sub ProcesarObjetoVisible_0DBB(ByVal PunteroSpriteIX As Long, ByVal PunteroDatosIY As Long, ByVal X As Byte, ByVal Y As Byte, ByVal Z As Byte)
'rutina llamada cuando los objetos del juego son visibles en la pantalla actual
'si no se dibujaba el objeto, ajusta la posici�n y lo marca para que se dibuje
'PunteroSpriteIX apunta al sprite del objeto
'PunteroDatosIY apunta a los datos del objeto
'X,Y continene la posici�n en pantalla del objeto
'X = la coordenada y del sprite en pantalla (-16)
If (TablaPosicionObjetos_3008(PunteroDatosIY - &H3008) And &H80) <> 0 Then Exit Sub 'si el objeto ya se ha cogido, sale
TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17) = Z Or &H80  'indica que hay que pintar el objeto y actualiza la profundidad del objeto dentro del buffer de tiles
TablaSprites_2E17(PunteroSpriteIX + 2 - &H2E17) = Y - 8  'modifica la posici�n y del objeto (-8 pixels)
If X >= 2 Then
    TablaSprites_2E17(PunteroSpriteIX + 1 - &H2E17) = X - 2 'modifica la posici�n x del objeto (-2 pixels)
Else
    TablaSprites_2E17(PunteroSpriteIX + 1 - &H2E17) = 256 + X - 2 'evita el bug del pergamino
End If
End Sub

Sub ProcesarPersonaje_2468(ByVal PunteroSpritePersonajeIX As Long, ByVal PunteroDatosPersonajeIY As Long, ByVal PunteroDatosPersonajeHL As Long)
'procesa los datos del personaje para cambiar la animaci�n y posici�n del sprite
'PunteroSpritePersonajeIX = direcci�n del sprite correspondiente
'PunteroDatosPersonajeIY = datos de posici�n del personaje correspondiente
Dim PunteroTablaAnimaciones As Long
Dim Y As Byte
Dim HL As String
Dim IX As String
Dim IY As String
IX = Hex$(PunteroSpritePersonajeIX)
IY = Hex$(PunteroDatosPersonajeIY)
HL = Hex$(PunteroDatosPersonajeHL)
PunteroTablaAnimaciones = CambiarAnimacionTrajesMonjes_2A61(PunteroSpritePersonajeIX, PunteroDatosPersonajeIY) 'cambia la animaci�n de los trajes de los monjes seg�n la posici�n y en contador de animaciones
HL = Hex$(PunteroTablaAnimaciones)
If ComprobarVisibilidadSprite_245E(PunteroSpritePersonajeIX, PunteroDatosPersonajeIY, Y) Then
    ActualizarDatosGraficosPersonaje_2A34 PunteroSpritePersonajeIX, PunteroDatosPersonajeIY, PunteroTablaAnimaciones, Y
End If
End Sub

Function CambiarAnimacionTrajesMonjes_2A61(ByVal PunteroSpritePersonajeIX As Long, ByVal PunteroDatosPersonajeIY As Long) As Long
'cambia la animaci�n de los trajes de los monjes seg�n la posici�n y en contador de animaciones y obtiene la direcci�n de los
'datos de la animaci�n que hay que poner en hl
'PunteroSpritePersonajeIX = direcci�n del sprite correspondiente
'PunteroDatosPersonajeIY = datos de posici�n del personaje correspondiente
'al salir devuelve el �ndice en la tabla de animaciones
Dim AnimacionPersonaje As Byte
Dim AnimacionTraje As Byte
Dim AnimacionSprite As Byte
Dim Orientacion As Byte
Dim PunteroAnimacion As Long
Dim IX As String
Dim IY As String
Dim DE As String
IX = Hex$(PunteroSpritePersonajeIX)
IY = Hex$(PunteroDatosPersonajeIY)
AnimacionPersonaje = TablaCaracteristicasPersonajes_3036(PunteroDatosPersonajeIY - &H3036) 'obtiene la animaci�n del personaje
'2A64
Orientacion = TablaCaracteristicasPersonajes_3036(PunteroDatosPersonajeIY + 1 - &H3036)  'obtiene la orientaci�n del personaje
'2A67
Orientacion = ModificarOrientacion_2480(Orientacion) 'modifica la orientaci�n que se le pasa con la orientaci�n de la pantalla actual
'2A6b
AnimacionTraje = (Orientacion * 4) Or AnimacionPersonaje 'desplaza la orientaci�n 2 a la izquierda y la combina con la animaci�n para obtener la animaci�n del traje de los monjes
'2A6F
AnimacionSprite = TablaSprites_2E17(PunteroSpritePersonajeIX + &HB - &H2E17) 'lee el antiguo valor...
'2A72
AnimacionSprite = AnimacionSprite And &HF0 '...y se queda con los bits que no son de la animaci�n
AnimacionSprite = AnimacionSprite Or AnimacionTraje
'2A75
TablaSprites_2E17(PunteroSpritePersonajeIX + &HB - &H2E17) = AnimacionSprite 'combina el valor anterior con la animaci�n del traje
'2A78
PunteroAnimacion = Orientacion 'recupera la orientaci�n del personaje en la pantalla actual
PunteroAnimacion = PunteroAnimacion + 1
PunteroAnimacion = PunteroAnimacion And 2 'indica si el personaje mira hacia la derecha o hacia la izquierda
PunteroAnimacion = PunteroAnimacion * 2 'desplaza 1 bit a la izquierda
PunteroAnimacion = PunteroAnimacion Or AnimacionPersonaje 'combina con el n�mero de animaci�n actual
PunteroAnimacion = PunteroAnimacion * 4 'desplaza 2 bits a la izquierda (las animaciones de las x y de las y est�n separadas por 8 entradas)
'2A80
'a = 0 0 0 (si se mueve en x, 0, si se mueve en y, 1) (n�mero de la secuencia de animaci�n (2 bits)) 0 0
DE = Hex$(PunteroTablaAnimacionesPersonaje_2A84)
If (PunteroTablaAnimacionesPersonaje_2A84 And &HC000&) <> &HC000& Then
'2A8D
    PunteroAnimacion = PunteroAnimacion + PunteroTablaAnimacionesPersonaje_2A84 'indexa en la tabla
    CambiarAnimacionTrajesMonjes_2A61 = PunteroAnimacion
    Exit Function 'si la direcci�n que se ha puesto en 2A84 empieza por 0xc0, vuelve
End If
'aqu� llega si la direcci�n que se ha puesto en la instrucci�n modificada empieza por 0xc0
'PunteroAnimacion = �ndice en la tabla de animaciones
Dim NumeroMonje As Byte
Dim PunteroCaraMonje As Long
'2A8F
NumeroMonje = CByte(PunteroTablaAnimacionesPersonaje_2A84 And &HFF&) 'n�mero de monje (0, 2, 4 � 6)
'2A96
PunteroCaraMonje = Leer16(TablaPunterosCarasMonjes_3097, NumeroMonje + &H3097 - &H3097)
'2aa0
If (PunteroAnimacion And &H10&) <> 0 Then 'seg�n se mueva en x o en y, pone una cabeza
'2AA5
    PunteroCaraMonje = PunteroCaraMonje + &H32 'si el bit 4 es 1 (se mueve en y), coge la segunda cara
End If
'2AA9
PunteroAnimacion = PunteroAnimacion + &H31DF&
Escribir16 TablaAnimacionPersonajes_319F, PunteroAnimacion - &H319F, PunteroCaraMonje
CambiarAnimacionTrajesMonjes_2A61 = PunteroAnimacion
End Function

Function ComprobarVisibilidadSprite_245E(ByVal PunteroSpritePersonajeIX As Long, ByVal PunteroDatosPersonajeIY As Long, ByRef Ypantalla As Byte) As Boolean
Dim Visible As Boolean
Dim X As Byte
Dim Z As Byte
Dim Y As Byte
Visible = ProcesarObjeto_2ADD(PunteroSpritePersonajeIX, PunteroDatosPersonajeIY, X, Y, Z, Ypantalla) 'comprueba si es visible y si lo es, actualiza su posici�n si fuese necesario
If Not Visible Then
    TablaSprites_2E17(PunteroSpritePersonajeIX + 0 - &H2E17) = &HFE 'marca el sprite como no usado
    Exit Function 'sale con visibilidad=false
End If
ComprobarVisibilidadSprite_245E = Visible
End Function

Sub ActualizarDatosGraficosPersonaje_2A34(ByVal PunteroSpritePersonajeIX As Long, ByVal PunteroDatosPersonajeIY As Long, ByVal PunteroDatosPersonajeHL As Long, Y As Byte)
'aqu� se llega desde fuera si un sprite es visible, despu�s de haber actualizado su posici�n.
'en PunteroDatosPersonajeHL se apunta a la animaci�n correspondiente para el sprite
'PunteroSpritePersonajeIX = direcci�n del sprite correspondiente
'PunteroDatosPersonajeIY = datos de posici�n del personaje correspondiente
'Y = posici�n y en pantalla del sprite
Dim Orientacion As Byte
TablaSprites_2E17(PunteroSpritePersonajeIX + 7 - &H2E17) = TablaAnimacionPersonajes_319F(PunteroDatosPersonajeHL - &H319F) 'actualiza la direcci�n de los gr�ficos del sprite con la animaci�n que toca
'2a38
TablaSprites_2E17(PunteroSpritePersonajeIX + 8 - &H2E17) = TablaAnimacionPersonajes_319F(PunteroDatosPersonajeHL + 1 - &H319F) 'actualiza la direcci�n de los gr�ficos del sprite con la animaci�n que toca
'2a3d
TablaSprites_2E17(PunteroSpritePersonajeIX + 5 - &H2E17) = TablaAnimacionPersonajes_319F(PunteroDatosPersonajeHL + 2 - &H319F) 'actualiza el ancho y alto del sprite seg�n la animaci�n que toca
'2a42
TablaSprites_2E17(PunteroSpritePersonajeIX + 6 - &H2E17) = TablaAnimacionPersonajes_319F(PunteroDatosPersonajeHL + 3 - &H319F) 'actualiza el ancho y alto del sprite seg�n la animaci�n que toca
'2a47
TablaSprites_2E17(PunteroSpritePersonajeIX + 0 - &H2E17) = Y Or &H80 'indica que hay que redibujar el sprite. combina el valor con la posici�n y de pantalla del sprite
'2a4d
Orientacion = ModificarOrientacion_2480(LeerDatoObjeto(PunteroDatosPersonajeIY + 1)) 'obtiene la orientaci�n del personaje. modifica la orientaci�n que se le pasa en a con la orientaci�n de la pantalla actual
'2a53
Orientacion = modFunciones.shr(Orientacion, 1)
'2a55
If Orientacion <> LeerDatoObjeto(PunteroDatosPersonajeIY + 6) Then 'comprueba si ha cambiado la orientaci�n del personaje
    'si es as�, salta al m�todo correspondiente por si hay que flipear los gr�ficos
'2A58
    Select Case PunteroRutinaFlipPersonaje_2A59
        Case Is = &H353B
            FlipearSpritesGuillermo_353B
        Case Is = &H34E2
            FlipearSpritesAdso_34E2
        Case Is = &H34FB
            FlipearSpritesMalaquias_34FB
        Case Is = &H350B
            FlipearSpritesAbad_350B
        Case Is = &H351B
            FlipearSpritesBerengario_351B
        Case Is = &H352B
            FlipearSpritesSeverino_352B
    End Select
End If
'2A5D
MovimientoRealizado_2DC1 = True 'indica que ha habido movimiento
End Sub

Sub FlipearSpritesGuillermo_353B()
'este m�todo se llama cuando cambia la orientaci�n del sprite de guillermo y se encarga de flipear los sprites de guillermo
TablaCaracteristicasPersonajes_3036(&H303C - &H3036) = TablaCaracteristicasPersonajes_3036(&H303C - &H3036) Xor 1 'invierte el estado del flag
'A300 apunta a los gr�ficos de guillermo de 5 bytes de ancho
'5 bytes de ancho y 0x366 bytes (0xae*5)
GirarGraficosRespectoX_3552 TablaGraficosObjetos_A300, &HA300 - &HA300, 5, &HAE
'A666 apunta a los gr�ficos de guillermo de 4 bytes de ancho
'4 bytes de ancho y 0x84 bytes (0x21*4)
GirarGraficosRespectoX_3552 TablaGraficosObjetos_A300, &HA666 - &HA300, 4, &H21
End Sub

Sub FlipearSpritesAdso_34E2()
'este m�todo se llama cuando cambia la orientaci�n del sprite de adso y se encarga de flipear los sprites de adso
TablaCaracteristicasPersonajes_3036(&H304B - &H3036) = TablaCaracteristicasPersonajes_3036(&H304B - &H3036) Xor 1 'flip de adso
'A6EA apunta a los sprites de adso de 5 bytes de ancho
GirarGraficosRespectoX_3552 TablaGraficosObjetos_A300, &HA6EA& - &HA300&, 5, &H5F
'A8C5 apunta a los sprite de adso de 4 bytes de ancho
GirarGraficosRespectoX_3552 TablaGraficosObjetos_A300, &HA8C5& - &HA300&, 4, &H5A
End Sub

Sub FlipearSpritesMalaquias_34FB()
'este m�todo se llama cuando cambia la orientaci�n del sprite de malaqu�as y se encarga de flipear las caras del sprite
Dim PunteroDatos As Long
TablaCaracteristicasPersonajes_3036(&H305A - &H3036) = TablaCaracteristicasPersonajes_3036(&H305A - &H3036) Xor 1 'flip de malaqu�as
PunteroDatos = Leer16(TablaPunterosCarasMonjes_3097, &H3097 - &H3097) 'apunta a los datos de las caras de malaqu�as
GirarGraficosRespectoX_3552 DatosMonjes_AB59, PunteroDatos - &HAB59&, 5, &H14 'flipea las caras de malaqu�as
End Sub

Sub FlipearSpritesAbad_350B()
'este m�todo se llama cuando cambia la orientaci�n del sprite del abad y se encarga de flipear las caras del sprite
Dim PunteroDatos As Long
TablaCaracteristicasPersonajes_3036(&H3069 - &H3036) = TablaCaracteristicasPersonajes_3036(&H3069 - &H3036) Xor 1 'flip de malaqu�as
PunteroDatos = Leer16(TablaPunterosCarasMonjes_3097, &H3099 - &H3097) 'apunta a los datos de las caras del abad
GirarGraficosRespectoX_3552 DatosMonjes_AB59, PunteroDatos - &HAB59&, 5, &H14 'flipea las caras del abad
End Sub

Sub FlipearSpritesBerengario_351B()
'este m�todo se llama cuando cambia la orientaci�n del sprite de berengario y se encarga de flipear las caras del sprite
Dim PunteroDatos As Long
TablaCaracteristicasPersonajes_3036(&H3078 - &H3036) = TablaCaracteristicasPersonajes_3036(&H3078 - &H3036) Xor 1 'flip de malaqu�as
PunteroDatos = Leer16(TablaPunterosCarasMonjes_3097, &H309B - &H3097) 'apunta a los datos de las caras de berengario
GirarGraficosRespectoX_3552 DatosMonjes_AB59, PunteroDatos - &HAB59&, 5, &H14 'flipea las caras de berengario
End Sub

Sub FlipearSpritesSeverino_352B()
'este m�todo se llama cuando cambia la orientaci�n del sprite de severino y se encarga de flipear las caras del sprite
Dim PunteroDatos As Long
TablaCaracteristicasPersonajes_3036(&H3087 - &H3036) = TablaCaracteristicasPersonajes_3036(&H3087 - &H3036) Xor 1 'flip de malaqu�as
PunteroDatos = Leer16(TablaPunterosCarasMonjes_3097, &H309D - &H3097) 'apunta a los datos de las caras de severino
GirarGraficosRespectoX_3552 DatosMonjes_AB59, PunteroDatos - &HAB59&, 5, &H14 'flipea las caras de severino
End Sub

Sub RellenarBufferAlturasPersonaje_28EF(ByVal PunteroDatosPersonajeIY As Long, ByVal ValorBufferAlturas As Byte)
'si la posici�n del sprite es central y la altura est� bien, pone ValorBufferAlturas en las posiciones que ocupa del buffer de alturas
'PunteroDatosPersonajeIY = direcci�n de los datos de posici�n asociados al personaje
'ValorBufferAlturas = valor a poner en las posiciones que ocupa el personaje del buffer de alturas
Dim PunteroBufferAlturasIX As Long
Dim Altura As Byte
If Not DeterminarPosicionCentral_0CBE(PunteroDatosPersonajeIY, PunteroBufferAlturasIX) Then Exit Sub 'si la posici�n no es una de las del centro de la pantalla o la altura del personaje no coincide con la altura base de la planta, sale
'28F3
'en otro caso PunteroBufferAlturasIX apunta a la altura de la pos actual
Altura = TablaBufferAlturas_01C0(PunteroBufferAlturasIX) 'obtiene la entrada del buffer de alturas
'28f6
TablaBufferAlturas_01C0(PunteroBufferAlturasIX) = (Altura And &HF) Or ValorBufferAlturas 'indica que el personaje est� en la posici�n (x, y)
'28FC
If (TablaCaracteristicasPersonajes_3036(PunteroDatosPersonajeIY + 5 - &H3036) And &H80) <> 0 Then Exit Sub 'si el bit 7 del byte 5 est� puesto, sale
'2901
'indica que el personaje tambi�n ocupa la posici�n (x - 1, y)
Altura = TablaBufferAlturas_01C0(PunteroBufferAlturasIX - 1)
TablaBufferAlturas_01C0(PunteroBufferAlturasIX - 1) = (Altura And &HF) Or ValorBufferAlturas 'indica que el personaje est� en la posici�n (x-1, y)
'290A
'indica que el personaje tambi�n ocupa la posici�n (x, y-1)
Altura = TablaBufferAlturas_01C0(PunteroBufferAlturasIX - &H18)
'290D
TablaBufferAlturas_01C0(PunteroBufferAlturasIX - &H18) = (Altura And &HF) Or ValorBufferAlturas 'indica que el personaje est� en la posici�n (x, y-1)
'2913
'indica que el personaje tambi�n ocupa la posici�n (x-1, y-1)
Altura = TablaBufferAlturas_01C0(PunteroBufferAlturasIX - &H19)
TablaBufferAlturas_01C0(PunteroBufferAlturasIX - &H19) = (Altura And &HF) Or ValorBufferAlturas 'indica que el personaje est� en la posici�n (x, y-1)
End Sub

Sub DibujarSprites_2674()
'dibuja los sprites
Dim PunteroSpritesHL As Long
Dim Valor As Byte
If HabitacionOscura_156C Then
    DibujarSprites_267B
Else
    DibujarSprites_4914
End If
End Sub

Sub DibujarSprites_267B()
    'dibuja los sprites
    Dim PunteroSpritesHL As Long
    Dim Valor As Byte

    'dibujo de los sprites cuando la habitaci�n no est� iluminada
    PunteroSpritesHL = &H2E17 'apunta al primer sprite de los personajes
    Do
'2681
        Valor = TablaSprites_2E17(PunteroSpritesHL - &H2E17)
        If Valor = &HFF Then
            Exit Do 'si ha llegado al final, salta
        ElseIf Valor <> &HFE Then 'si es visible, marca el sprite como que no hay que dibujarlo (porque est� oscuro)
'268A
            TablaSprites_2E17(PunteroSpritesHL - &H2E17) = Valor And &H7F
        End If
        PunteroSpritesHL = PunteroSpritesHL + &H14 'longitud de cada sprite
        DoEvents
    Loop
'268F
    If Not Depuracion.LuzEnGuillermo Then
        If TablaSprites_2E17(&H2E2B - &H2E17) = &HFE Then Exit Sub 'si el sprite de adso no es visible, sale '### depuraci�n
    End If
    If (Not Depuracion.LuzEnGuillermo And Depuracion.Luz = EnumTipoLuz_Off) Or Depuracion.Luz = EnumTipoLuz_Normal Then
        If Not Depuracion.Lampara Then
'2695
            If (TablaObjetosPersonajes_2DEC(&H2DF3 - &H2DEC) And &H80) = 0 Then Exit Sub 'si adso no tiene la l�mpara, sale '### depuraci�n
        End If
    End If
    TablaSprites_2E17(&H2FCF - &H2E17) = &HBC 'activa el sprite de la luz
    DibujarSprites_4914
End Sub

Sub DibujarSprites_4914()
Dim Punteros(22) As Long 'punteros a los sprites
Dim NumeroSprites As Long 'n�mero de sprites en la pila
Dim NumeroSpritesVisibles As Long 'n�mero de elementos visibles
Dim PunteroSpriteIX As Long 'sprite original (bucle exterior)
Dim Valor As Byte
Dim NumeroCambios As Byte
Dim Temporal As Long
Dim Contador As Long
Dim Contador2 As Long
Dim Profundidad1 As Byte
Dim Profundidad2 As Byte
Dim Xactual As Byte
Dim Yactual As Byte
Dim nXactual As Byte
Dim nYactual As Byte
Dim Xanterior As Byte
Dim Yanterior As Byte
Dim nXanterior As Byte
Dim nYanterior As Byte
Dim TileX As Byte
Dim TileY As Byte
Dim nXsprite As Byte
Dim nYsprite As Byte
Dim ValorLongDE As Long
Dim PunteroBufferTiles As Long
Dim AltoXanchoSprite As Long
Dim PunteroBufferSprites As Long
Dim PunteroBufferSpritesAnterior As Long
Dim PunteroBufferSpritesLibre As Long '4908
Dim ProfundidadMaxima_4DD9 As Long 'l�mite superior de profundidad de la iteraci�n anterior
Dim PunteroSpriteIY As Long 'sprite actual (bucle interior)
Dim Distancia1X As Byte 'distancia desde el inicio del sprite actual al inicio del sprite original
Dim Distancia2X As Byte 'distancia desde el inicio del sprite original al inicio del sprite actual
Dim LongitudX As Byte 'longitud a pintar del sprite actual
Dim Distancia1Y As Byte
Dim Distancia2Y As Byte
Dim LongitudY As Byte
Dim ProfundidadMaxima As Long 'profundidad m�xima de la iteraci�n actual
Dim PunteroBufferTilesAnterior_3095 As Long

If Not Depuracion.PersonajesAdso Then
    TablaSprites_2E17(&H2E2B + 0 - &H2E17) = &HFE 'desconecta a adso ###depuraci�n
End If
If Not Depuracion.PersonajesMalaquias Then
    TablaSprites_2E17(&H2E3F + 0 - &H2E17) = &HFE 'desconecta a malaqu�as
End If
If Not Depuracion.PersonajesAbad Then
    TablaSprites_2E17(&H2E53 + 0 - &H2E17) = &HFE 'desconecta al abad ###depuraci�n
End If
If Not Depuracion.PersonajesBerengario Then
    TablaSprites_2E17(&H2E67 + 0 - &H2E17) = &HFE 'desconecta a berengario
End If
If Not Depuracion.PersonajesSeverino Then
    TablaSprites_2E17(&H2E7B + 0 - &H2E17) = &HFE 'desconecta a severino
End If


'TablaSprites_2E17(&H2E2B + 1 - &H2E17) = TablaSprites_2E17(&H2E17 + 1 - &H2E17)
'TablaSprites_2E17(&H2E2B + 2 - &H2E17) = TablaSprites_2E17(&H2E17 + 2 - &H2E17)
'TablaSprites_2E17(&H2E2B + 3 - &H2E17) = TablaSprites_2E17(&H2E17 + 3 - &H2E17)


Do
'4918
    PunteroBufferSprites = &H9500& 'apunta al comienzo del buffer para los sprites
    PunteroBufferSpritesLibre = &H9500&
    PunteroSpriteIX = &H2E17 'apunta al primer sprite
    'limpia los punteros de la iteraci�n anterior
    For Contador = 0 To NumeroSprites - 1
        Punteros(Contador) = 0
    Next
    NumeroSprites = 0
    NumeroSpritesVisibles = 0
    Do
'4929
        Valor = TablaSprites_2E17(PunteroSpriteIX - &H2E17)
        If Valor = &HFF Then
            Exit Do 'si ha llegado al final, salta
        ElseIf Valor <> &HFE Then 'si es visible, guarda la direcci�n
'4932
            Punteros(NumeroSprites) = PunteroSpriteIX 'ojo, cambiado.  antes NumeroSpritesVisibles
            NumeroSprites = NumeroSprites + 1
            If (Valor And &H80) <> 0 Then 'hay que dibujar el sprite
                If (TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17) And &H80) <> 0 Then 'hay que dibujar el sprite
                    NumeroSpritesVisibles = NumeroSpritesVisibles + 1
                End If
            End If
        End If
        PunteroSpriteIX = PunteroSpriteIX + &H14 '20 bytes por entrada
        DoEvents
    Loop
'493b
    'aqu� llega una vez que ha metido en la pila las entradas a tratar
    If NumeroSpritesVisibles = 0 Then Exit Sub ' si no hab�a alguna entrada activa, vuelve
'494a
    'aqu� llega si hab�a alguna entrada que hab�a que pintar
    'primero se ordenan las entradas seg�n la profundidad por el m�todo de la burbuja mejorado
    If NumeroSprites > 1 Then
        Do
            NumeroCambios = 0
            For Contador = NumeroSprites - 2 To 0 Step -1
                Profundidad1 = TablaSprites_2E17(Punteros(Contador + 1) - &H2E17) And &H3F
                Profundidad2 = TablaSprites_2E17(Punteros(Contador) - &H2E17) And &H3F
                If Profundidad2 < Profundidad1 Then 'realiza un intercambio
                    Temporal = Punteros(Contador)
                    Punteros(Contador) = Punteros(Contador + 1)
                    Punteros(Contador + 1) = Temporal
                    NumeroCambios = NumeroCambios + 1
                End If
            Next
            If NumeroCambios = 0 Then Exit Do
            DoEvents
        Loop
    End If
    'aqu� llega una vez que las entradas de la pila est�n ordenadas por la profundidad
'4977
    For Contador = NumeroSprites - 1 To 0 Step -1
'498C
        PunteroSpriteIX = Punteros(Contador)
'498F
        TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17) = (TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17) And &HBF) 'pone el bit 6 a 0. sprite no prcesado
        If (TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17) And &H80) <> 0 Then 'el sprite ha cambiado
'4999
            
            Xactual = TablaSprites_2E17(PunteroSpriteIX + 1 - &H2E17) 'posici�n x en bytes
            Yactual = TablaSprites_2E17(PunteroSpriteIX + 2 - &H2E17) 'posici�n y en pixels
            nYactual = TablaSprites_2E17(PunteroSpriteIX + 6 - &H2E17) 'alto en pixels
            nXactual = TablaSprites_2E17(PunteroSpriteIX + 5 - &H2E17) 'ancho en bytes
            nXactual = nXactual And &H7F 'el bit7 de la posici�n 5 no nos interesa ahora
            CalcularDimensionesAmpliadasSprite_4D35 Xactual, Yactual, nXactual, nYactual, nXsprite, nYsprite, TileX, TileY
            Xanterior = TablaSprites_2E17(PunteroSpriteIX + 3 - &H2E17) 'posici�n x en bytes
            Yanterior = TablaSprites_2E17(PunteroSpriteIX + 4 - &H2E17) 'posici�n y en pixels
            nYanterior = TablaSprites_2E17(PunteroSpriteIX + &HA - &H2E17)  'alto en pixels
            nXanterior = TablaSprites_2E17(PunteroSpriteIX + 9 - &H2E17) 'ancho en bytes
            
            'l=X=anterior posici�n x del sprite (en bytes)
            'h=Y=anterior posici�n y del sprite (en pixels)
            'e=nX=anterior ancho del sprite (en bytes)
            'd=nY=anterior alto del sprite (en pixels)
            '2DD5=TileX=posici�n x del tile en el que empieza el sprite
            '2DD6=TileY=posici�n y del tile en el que empieza el sprite
            '2DD7=nXsprite=tama�o en x del sprite
            '2DD8=nYsprite=tama�o en y del sprite
'49BD
            If Not Depuracion.DeshabilitarCalculoDimensionesAmpliadas Then
                CalcularDimensionesAmpliadasSprite_4CBF Xanterior, Yanterior, nXanterior, nYanterior, nXsprite, nYsprite, TileX, TileY '###depuracion
            End If
            
            TablaSprites_2E17(PunteroSpriteIX + &HC - &H2E17) = TileX 'posici�n en x del tile en el que empieza el sprite (en bytes)
            TablaSprites_2E17(PunteroSpriteIX + &HD - &H2E17) = TileY 'posici�n en y del tile en el que empieza el sprite (en pixels
 'dado PunteroSpriteIX, calcula la coordenada correspondiente del buffer de tiles (buffer de tiles de 16x20, donde cada tile ocupa 16x8)
 '49c9
            ValorLongDE = modFunciones.Byte2Long(TileX) And &HFC& 'posici�n en x del tile inicial en el que empieza el sprite (en bytes)
            ValorLongDE = ValorLongDE + modFunciones.shr(ValorLongDE, 1) 'x + x/2 (ya que en cada byte hay 4 pixels y cada entrada en el buffer de tiles es de 6 bytes)
'49d6
            PunteroBufferTiles = modFunciones.Byte2Long(TileY) 'tile inicial en y en el que empieza el sprite (en pixels)
            PunteroBufferTiles = PunteroBufferTiles * 12 + ValorLongDE 'apunta a la l�nea correspondiente en el buffer de tiles
            'TileY tiene valores m�ltiplos de 8, porque utiliza el pixel como unidad. cada tile son 8 p�xeles,
            'por lo que el cambio de tile supone 12*8=96 bytes

            
            'indexa en el buffer de tiles (0x8b94 se corresponde a la posici�n X = -2, Y = -5 en el buffer de tiles)
            'que en pixels es: (X = -32, Y = -40), luego el primer pixel del buffer de tiles en coordenadas de sprite es el (32,40)
'49e1
            PunteroBufferTiles = PunteroBufferTiles + &H8B94&
            
            
            '3095=PunteroBuffertiles
            PunteroBufferTilesAnterior_3095 = PunteroBufferTiles
            TablaSprites_2E17(PunteroSpriteIX + &HE - &H2E17) = nXsprite 'ancho final del sprite (en bytes)
            TablaSprites_2E17(PunteroSpriteIX + &HF - &H2E17) = nYsprite 'alto final del sprite (en pixels)
            AltoXanchoSprite = Byte2Long(nXsprite) * Byte2Long(nYsprite) 'alto del sprite*ancho del sprite
            PunteroBufferSprites = PunteroBufferSpritesLibre
            TablaSprites_2E17(PunteroSpriteIX + &H10 - &H2E17) = LeerByteLong(PunteroBufferSprites, 0) 'direcci�n del buffer de sprites asignada a este sprite
            TablaSprites_2E17(PunteroSpriteIX + &H11 - &H2E17) = LeerByteLong(PunteroBufferSprites, 1) 'direcci�n del buffer de sprites asignada a este sprite
            PunteroBufferSpritesLibre = PunteroBufferSprites + AltoXanchoSprite 'guarda la direcci�n libre del buffer de sprites
            If PunteroBufferSpritesLibre > &H9CFE& Then Exit For '9CFE= l�mite del buffer de sprites. si no hay sitio para el sprite, salta pasa vaciar la lista de los procesados y procesa el resto
'4a13
            'aqu� llega si hay espacio para procesar el sprite
            TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17) = (TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17) Or &H40) 'pone el bit 6 a 1. marca el sprite como procesado
            For Contador2 = PunteroBufferSprites To PunteroBufferSpritesLibre - 1
                BufferSprites_9500(Contador2 - &H9500&) = 0 'limpia la zona asignada del buffer de sprites
            Next
'4A1F
            ProfundidadMaxima_4DD9 = 0
'4a2e
            For Contador2 = NumeroSprites - 1 To 0 Step -1
'4a56
                PunteroSpriteIY = Punteros(Contador2) 'direcci�n de la entrada del sprite actual
                If (TablaSprites_2E17(PunteroSpriteIY + 5 - &H2E17) And &H80) = 0 Then 'si el sprite no va a desaparecer
'4A5F
                    'entrada:
                    'l=PosicionOriginal
                    'h=PosicionActual
                    'e=LongitudOriginal
                    'd=LongitudActual
                    'en a=Longitud devuelve la longitud a pintar del sprite actual para la coordenada que se pasa
                    'en h=Distancia1 devuelve la distancia desde el inicio del sprite actual al inicio del sprite original
                    'en l=Distancia2 devuelve la distancia desde el inicio del sprite original al inicio del sprite actual
                    'si devuelve true, indica que debe evitarse el proceso de esta combinaci�n de sprites
                    'comprueba si el sprite actual puede verse en la zona del sprite original
                    If Not ObtenerDistanciaSprites_4D54(TileX, TablaSprites_2E17(PunteroSpriteIY + 1 - &H2E17), nXsprite, TablaSprites_2E17(PunteroSpriteIY + 5 - &H2E17), Distancia1X, Distancia2X, LongitudX) Then
'4a70                   comprueba si el sprite actual puede verse en la zona del sprite original
                        If Not ObtenerDistanciaSprites_4D54(TileY, TablaSprites_2E17(PunteroSpriteIY + 2 - &H2E17), nYsprite, TablaSprites_2E17(PunteroSpriteIY + 6 - &H2E17), Distancia1Y, Distancia2Y, LongitudY) Then
'4A9A
                            'obtiene la posici�n del sprite en coordenadas de c�mara
                            ProfundidadMaxima = Bytes2Long(TablaSprites_2E17(PunteroSpriteIY + &H12 - &H2E17), TablaSprites_2E17(PunteroSpriteIY + &H13 - &H2E17)) 'combina los dos bytes en un entero largo
                            'obtiene el l�mite superior de profundidad de la iteraci�n anterior y lo coloca como l�mite inferior
                            PunteroBufferSpritesAnterior = PunteroBufferSprites
                            'GuardarArchivo "D:\datos\vbasic\Abadia\Abadia2\BufferSprites", BufferSprites_9500
'4AA0
                            CopiarTilesBufferSprites_4D9E ProfundidadMaxima, ProfundidadMaxima_4DD9, False, PunteroBufferTiles, PunteroBufferSprites, nXsprite, nYsprite 'copia en el buffer de sprites los tiles que est�n detras del sprite
                            'GuardarArchivo "D:\datos\vbasic\Abadia\Abadia2\BufferSprites", BufferSprites_9500
                            ProfundidadMaxima_4DD9 = ProfundidadMaxima
                            PunteroBufferSprites = PunteroBufferSpritesAnterior
                            DibujarSprite_4AA3 PunteroSpriteIY, Distancia1Y, Distancia2Y, Distancia1X, Distancia2X, nXsprite, PunteroBufferSprites, LongitudY, LongitudX 'al llegar aqu� pinta el sprite actual
                            'GuardarArchivo "D:\datos\vbasic\Abadia\Abadia2\BufferSprites", BufferSprites_9500
                        End If
                    End If
                End If
            Next
'4A43
            'aqu� llega si ya se han procesado todos los sprites de la pila (con respecto al sprite actual)
            'fcfc: se le pasa un valor de profundidad muy alto
            'obtiene el l�mite superior de profundidad de la iteraci�n anterior y lo coloca como l�mite inferior
            PunteroBufferTiles = PunteroBufferTilesAnterior_3095
            'GuardarArchivo "D:\datos\vbasic\Abadia\Abadia2\BufferSprites", BufferSprites_9500
'4A4B
            CopiarTilesBufferSprites_4D9E &HFCFC&, ProfundidadMaxima_4DD9, True, PunteroBufferTiles, PunteroBufferSprites, nXsprite, nYsprite 'dibuja en el buffer de sprites los tiles que est�n delante del sprite
            'GuardarArchivo "D:\datos\vbasic\Abadia\Abadia2\BufferSprites", BufferSprites_9500
        End If
    Next
'4BDF
    'aqu� llega una vez ha procesado todos los sprites que hab�a que redibujar (o si no hab�a m�s espacio en el buffer de sprites)
    PunteroSpriteIX = &H2E17 'apunta al primer sprite
    Do
        Valor = TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17)
        If Valor = &HFF Then Exit Do 'cuando encuentra el �ltimo, sale
        If Valor <> &HFE Then
            If (Valor And &H40) <> 0 Then 'si  tiene puesto el bit 6 (sprite procesado)
'4BF2
                CopiarSpritePantalla_4C1A PunteroSpriteIX
                TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17) = TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17) And &H3F 'limpia el bit 6 y 7 del byte 0
                If (TablaSprites_2E17(PunteroSpriteIX + 5 - &H2E17) And &H80) <> 0 Then 'si el sprite va a desaparecer
                    TablaSprites_2E17(PunteroSpriteIX + 5 - &H2E17) = TablaSprites_2E17(PunteroSpriteIX + 5 - &H2E17) And &H7F 'limpia el bit 7
                    TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17) = &HFE 'marca el sprite como inactivo
                End If
            End If
        End If
        PunteroSpriteIX = PunteroSpriteIX + &H14 'pasa al siguiente sprite
        DoEvents
    Loop
    DoEvents
Loop
End Sub

Sub CalcularDimensionesAmpliadasSprite_4D35(ByVal X As Byte, ByVal Y As Byte, ByVal nX As Byte, ByVal nY As Byte, nXsprite As Byte, nYsprite As Byte, ByRef TileX As Byte, ByRef TileY As Byte)
'devuelve en TileX,TileY la posici�n inicial del tile en el que empieza el sprite (TileY = pos inicial Y en pixels, TileX = posici�n inicial X en bytes)
'devuelve en nXsprite,nYsprite las dimensiones del sprite ampliadas para abarcar todos los tiles en los que se va a dibujar el sprite
'en X,Y se le pasa la posici�n inicial (Y = pos Y en pixels, X = pos X en bytes)
'en nX,nY se le pasa las dimensiones del sprite (nY = alto en pixels, nX = ancho en bytes)
Dim b As Byte
Dim c As Byte
c = Y And 7 'pos Y dentro del tile actual (en pixels)
TileY = Y And &HF8 'posici�n del tile actual en Y (en pixels)
b = X And 3 'pos X dentro del tile actual (en bytes)
TileX = X And &HFC 'posici�n del tile actual en X (en bytes)
nYsprite = (nY + c + 7) And &HF8 'calcula el alto del objeto para que abarque todos los tiles en los que se va a dibujar
nXsprite = (nX + b + 3) And &HFC 'calcula el ancho del objeto para que abarque todos los tiles en los que se va a dibujar
End Sub

Sub CalcularDimensionesAmpliadasSprite_4CBF(ByVal X As Byte, ByVal Y As Byte, ByVal nX As Byte, ByVal nY As Byte, ByRef nXsprite As Byte, ByRef nYsprite As Byte, ByRef TileX As Byte, ByRef TileY As Byte)
'comprueba las dimensiones m�nimas del sprite (para borrar el sprite viejo) y actualiza 0x2dd5 y 0x2dd7
'en X,Y se le pasa la posici�n anterior (Y = pos Y en pixels, X = pos X en bytes)
'en nX,nY se le pasa las dimensiones anteriores del sprite (nY = alto en pixels, nX = ancho en bytes)
'l=X=anterior posici�n x del sprite (en bytes)
'h=Y=anterior posici�n y del sprite (en pixels)
'e=nX=anterior ancho del sprite (en bytes)
'd=nY=anterior alto del sprite (en pixels)
'2DD5=TileX=posici�n x del tile en el que empieza el sprite
'2DD6=TileY=posici�n y del tile en el que empieza el sprite
'2DD7=nXsprite=tama�o en x del sprite
'2DD8=nYsprite=tama�o en y del sprite

Dim Valor As Byte
If TileX >= X Then 'si Xtile >= X2
'4cc5
    Valor = TileX - X + nXsprite
    If Valor > nX Then nX = Valor 'si el ancho ampliado es mayor que el m�nimo, e = ancho ampliado + Xtile - Xspr (coge el mayor ancho del sprite)
'4cce
    Valor = X And 3 'posici�n x dentro del tile actual
    TileX = X And &HFC 'actualiza la posici�n inicial en x del tile en el que empieza el sprite
    nXsprite = ((nX + Valor + 3) And &HFC) 'redondea el ancho al tile superior
Else
'4CE3
'aqu� llega si la posici�n del sprite en x > que el inicio de un tile en x
    Valor = X - TileX 'diferencia de posici�n en x del tile a x2
    Valor = Valor + nX 'a�ade al ancho del sprite la diferencia en x entre el inicio del sprite y el del tile asociado al sprite
    If nXsprite < Valor Then 'si el ancho ampliado del sprite < el ancho m�nimo del sprite
        nXsprite = ((Valor + 3) And &HFC)  'amplia el ancho m�nimo del sprite
    End If
End If
'4cf5
'ahora hace lo mismo para y
If TileY >= Y Then 'si ytile >= Y2
'4cfb
    Valor = TileY - Y + nYsprite
    If Valor > nY Then nY = Valor 'si el alto ampliado es mayor que el m�nimo, d = alto ampliado + Ytile - Yspr (coge el mayor alto del sprite)
'4d04
    Valor = Y And 7 'posici�n y dentro del tile actual
    TileY = Y And &HF8 'actualiza la posici�n inicial en y del tile en el que empieza el sprite
    nYsprite = ((nY + Valor + 7) And &HF8) 'redondea el ancho del sprite
    Exit Sub
Else
'4d18
    Valor = Y - TileY 'Y2 - Ytile - Y2
    Valor = Valor + nY 'suma al alto del sprite lo que sobresale del inicio del tile en y
    If nYsprite >= Valor Then Exit Sub 'si el alto del sprite >= el alto m�nimo, sale
    nYsprite = ((Valor + 7) And &HF8) 'redondea el alto al tile superior y actualiza el alto del sprite
    Exit Sub
End If
End Sub

Function ObtenerDistanciaSprites_4D54(ByVal PosicionOriginal As Byte, ByVal PosicionActual As Byte, ByVal LongitudOriginal As Byte, ByVal LongitudActual As Byte, ByRef Distancia1 As Byte, ByRef Distancia2 As Byte, ByRef Longitud As Byte) As Boolean
'dado l y e, y h y d, que son las posiciones iniciales y longitudes de los sprites original y actual, comprueba si el sprite actual puede
'verse en la zona del sprite original. Si puede verse, lo recorta. En otro caso, salta a por otro sprite actual
'entrada:
'l=PosicionOriginal
'h=PosicionActual
'e=LongitudOriginal
'd=LongitudActual
'salida:
'en a=Longitud devuelve la longitud a pintar del sprite actual para la coordenada que se pasa
'en h=Distancia1 devuelve la distancia desde el inicio del sprite actual al inicio del sprite original
'en l=Distancia2 devuelve la distancia desde el inicio del sprite original al inicio del sprite actual
'si devuelve true, indica que debe evitarse el proceso de esta combinaci�n de sprites
If PosicionOriginal = PosicionActual Then 'el sprite original empieza en el mismo punto que el sprite actual
'4d69
    Distancia1 = 0
    Distancia2 = 0
    If LongitudOriginal < LongitudActual Then
        Longitud = LongitudOriginal
    Else
        Longitud = LongitudActual
    End If
ElseIf PosicionOriginal < PosicionActual Then 'el sprite original empieza antes que el actual
'4d71
    Distancia1 = 0
    Distancia2 = PosicionActual - PosicionOriginal 'distancia entre la posici�n inicial del sprite original y del actual
    If Distancia2 > LongitudOriginal Then 'si la distancia entre el origen de los 2 sprites es >= que el ancho ampliado del sprite original
'4D81
        ObtenerDistanciaSprites_4D54 = True
    Else
'4D79
        Longitud = LongitudOriginal - Distancia2 'guarda la longitud de la parte visible del sprite actual en el sprite original
        If Longitud > LongitudActual Then Longitud = LongitudActual 'si esa longitud es > que la longitud del sprite actual, modifica la longitud a pintar del sprite actual
    End If
Else 'si llega aqu�, el sprite actual empieza antes que el sprite original
'4d5a
    If (PosicionOriginal - PosicionActual) >= LongitudActual Then 'si la distancia entre los sprites es >= que el ancho del sprite actual, el sprite actual no es visible
'4D81
        ObtenerDistanciaSprites_4D54 = True
    Else
'4d5d
        Distancia1 = PosicionOriginal - PosicionActual 'distancia desde el inicio del sprite actual al inicio del sprite original
        Distancia2 = 0
        If (PosicionOriginal - PosicionActual + LongitudOriginal) >= LongitudActual Then 'si la distancia entre los sprites + la longitud del sprite original >=LongitudActual
'4D66
            'como el sprite original no est� completamente dentro del sprite actual, dibuja solo la parte del sprite
            'actual que se superpone con el sprite original
            Longitud = LongitudActual - Distancia1
        Else
'4d64
            Longitud = LongitudOriginal
        End If
    End If
End If
End Function

Sub DibujarSprite_4AA3(ByVal PunteroSpriteIY As Long, ByVal Distancia1Y As Byte, ByVal Distancia2Y As Byte, ByVal Distancia1X As Byte, ByVal Distancia2X As Byte, ByVal nXsprite As Byte, ByVal PunteroBufferSprites As Long, ByVal LongitudY As Byte, ByVal LongitudX As Byte)
'pinta el sprite actual
'Distancia1Y=h
'Distancia2Y=l
Dim nX As Byte 'ancho del sprite actual
Dim PunteroDatosGraficosSpriteHL As Long
Dim PunteroDatosGraficosSpriteAnterior As Long
Dim PunteroBufferSpritesDE As Long
Dim PunteroBufferSpritesAnterior As Long
Dim ValorLong As Long
Dim Valor As Byte
Dim DesplazAdsoX As Byte
Dim Contador As Long
Dim Contador2 As Long
Dim MascaraOr As Long
Dim MascaraAnd As Long
Dim Fila As Long
Dim PunteroPatronLuz As Long
Dim DesplazamientoDE As Byte '= 80 (desplazamiento de medio tile)
Dim PunteroBufferSpritesIX As Long
Dim ValorRelleno As Long 'valor de la tabla 48E8 de rellenos de la luz
Dim HL As String
'4AA3
If Distancia1Y < 10 Or (Distancia1Y >= 10 And (TablaSprites_2E17(PunteroSpriteIY + &HB - &H2E17) And &H80) <> 0) Then 'si la distancia en y desde el inicio del sprite actual al inicio del sprite original < 10 o no se trata de un monje
'4AD5
    'calcula la l�nea en la que empezar a dibujar el sprite actual (saltandose la distancia entre el inicio del sprite actual y el inicio del sprite original)
    nX = TablaSprites_2E17(PunteroSpriteIY + 5 - &H2E17) 'obtiene el ancho del sprite actual
    ValorLong = Byte2Long(Distancia1Y) '(distancia en y desde el inicio del sprite actual al incio del sprite original
    ValorLong = ValorLong * nX
    PunteroDatosGraficosSpriteHL = Bytes2Long(TablaSprites_2E17(PunteroSpriteIY + 7 - &H2E17), TablaSprites_2E17(PunteroSpriteIY + 8 - &H2E17)) 'direcci�n de los datos gr�ficos del sprite
    'direcci�n de los datos gr�ficos del sprite (saltando lo que no se superpone con el �rea del sprite original en y)
    PunteroDatosGraficosSpriteHL = PunteroDatosGraficosSpriteHL + ValorLong
    HL = Hex$(PunteroDatosGraficosSpriteHL)

Else
'4AB5
    'si llega aqu� es porque la distancia en y desde el inicio del sprite actual al inicio del sprite original es >= 10, por lo que del sprite
    'actual (que es un monje), ya se ha pasado la cabeza. Por ello, obtiene un puntero al traje del monje
    ValorLong = Byte2Long(Distancia1Y - 10)
    nX = TablaSprites_2E17(PunteroSpriteIY + 5 - &H2E17) 'obtiene el ancho del sprite actual
    ValorLong = ValorLong * nX
    Valor = TablaSprites_2E17(PunteroSpriteIY + &HB - &H2E17) 'animaci�n del traje del monje
    PunteroDatosGraficosSpriteHL = Leer16(TablaPunterosTrajesMonjes_48C8, 2 * Byte2Long(Valor)) 'cada entrada son 2 bytes
    PunteroDatosGraficosSpriteHL = PunteroDatosGraficosSpriteHL + ValorLong
End If
'4ae5
'direcci�n de los datos gr�ficos del sprite (saltando lo que no est� en el �rea del sprite original en x y en y)
PunteroDatosGraficosSpriteHL = PunteroDatosGraficosSpriteHL + Distancia1X 'suma la distancia en x desde el inicio del sprite actual al incio del sprite original
HL = Hex$(PunteroDatosGraficosSpriteHL)
'4AED
'distancia en y desde el inicio del sprite original al inicio del sprite actual * ancho ampliado del sprite original
ValorLong = CLng(Distancia2Y) * CLng(nXsprite)
'PunteroBufferSpritelibre=posici�n inicial del buffer de sprites para este sprite
'direcci�n del buffer de sprites para el sprite original (saltando lo que no puede sobreescribir el sprite actual en y)
PunteroBufferSpritesDE = PunteroBufferSprites + ValorLong
'direcci�n del buffer de sprites para el sprite original (saltando lo que no puede sobreescribir el sprite actual en x y en y)
PunteroBufferSpritesDE = PunteroBufferSpritesDE + Distancia2X
'4b05
If PunteroDatosGraficosSpriteHL <> 0 Then 'si hl <> 0 (no es el sprite de la luz)
'4B0A
    'c=Distancia1Y
    'b'=LongitudY
    'b=LongitudX
    For Fila = 0 To LongitudY - 1
        PunteroDatosGraficosSpriteAnterior = PunteroDatosGraficosSpriteHL
        PunteroBufferSpritesAnterior = PunteroBufferSpritesDE
        For Contador = 0 To LongitudX - 1
            'Valor = TablaGraficosObjetos_A300(PunteroDatosGraficosSpriteHL - &HA300&) 'lee un byte gr�fico
            Valor = LeerDatoGrafico(PunteroDatosGraficosSpriteHL)
            If Valor <> 0 Then 'si es 0, salta al siguiente pixel
'4B18
                MascaraOr = Byte2Long(Valor)                'b7 b6 b5 b4 b3 b2 b1 b0
                ValorLong = modFunciones.rol8(MascaraOr, 4) 'b3 b2 b1 b0 b7 b6 b5 b4
                ValorLong = ValorLong Or MascaraOr   'b7|b3 b6|b2 b5|b1 b4|b0 b7|b3 b6|b2 b5|b1 b4|b0
                If ValorLong <> 0 Then 'si es 0, salta (???, no ser�a 0 antes tb???)
'4B21
                    MascaraAnd = (-ValorLong - 1) And &HFF& 'invierte el byte inferior (los sprites usan el color 0 como transparente)
                    Valor = BufferSprites_9500(PunteroBufferSpritesDE - &H9500&) 'lee un byte del buffer de sprites
                    Valor = Valor And Long2Byte(MascaraAnd)
                End If
'4b27
                Valor = Valor Or Long2Byte(MascaraOr) 'combina el byte leido
                BufferSprites_9500(PunteroBufferSpritesDE - &H9500&) = Valor 'escribe el byte en buffer de sprites despu�s de haberlo combinado
            End If
'4b2a
            PunteroDatosGraficosSpriteHL = PunteroDatosGraficosSpriteHL + 1 'avanza a la siguiente posici�n en x del gr�fico
            PunteroBufferSpritesDE = PunteroBufferSpritesDE + 1 'avanza a la siguiente posici�n en x dentro del buffer de sprites
        Next 'repite para el ancho
'4B2E
        PunteroDatosGraficosSpriteHL = PunteroDatosGraficosSpriteAnterior
        PunteroDatosGraficosSpriteHL = PunteroDatosGraficosSpriteHL + nX 'pasa a la siguiente l�nea del sprite
        PunteroBufferSpritesDE = PunteroBufferSpritesAnterior 'obtiene el puntero al buffer de sprites
        Distancia1Y = Distancia1Y + 1
        If Distancia1Y = 10 And (TablaSprites_2E17(PunteroSpriteIY + &HB - &H2E17) And &H80) = 0 Then
'4B41
            'si llega a 10, cambia la direcci�n de los datos gr�ficos de origen,
            'puesto que se pasa de dibujar la cabeza de un monje a dibujar su traje
            Valor = TablaSprites_2E17(PunteroSpriteIY + &HB - &H2E17) And &H7F 'animaci�n del traje del monje
            PunteroDatosGraficosSpriteHL = &H48C8 'apunta a la tabla de las posiciones de los trajes de los monjes
            PunteroDatosGraficosSpriteHL = PunteroDatosGraficosSpriteHL + 2 * Byte2Long(Valor)
            PunteroDatosGraficosSpriteHL = Leer16(TablaPunterosTrajesMonjes_48C8, PunteroDatosGraficosSpriteHL - &H48C8)
            'modifica la direcci�n de los datos gr�ficos de origen, para que apunte a la animaci�n del traje del monje
            PunteroDatosGraficosSpriteHL = PunteroDatosGraficosSpriteHL + Distancia1X 'distancia en x desde el inicio del sprite actual al incio del sprite original
        End If
'4B53
        PunteroBufferSpritesDE = PunteroBufferSpritesDE + nXsprite 'pasa a la siguiente l�nea del buffer de sprites
    Next 'repite para las l�neas de alto
Else 'si hl == 0 (es el sprite de la luz)
'4B60
    'aqu� llega si el sprite tiene un puntero a datos gr�ficos = 0 (es el sprite de la luz)
    PunteroPatronLuz = &H48E8 'apunta a la tabla con el patr�n de relleno de la luz
    For Contador = 0 To SpriteLuzTipoRelleno_4B6B  'TipoRellenoLuz_4B6B=0x00ef o 0x009f
        BufferSprites_9500(PunteroBufferSpritesDE + Contador - &H9500&) = &HFF 'rellena un tile o tile y medio de negro (la parte superior del sprite de la luz)
    Next
    PunteroBufferSpritesIX = PunteroBufferSpritesDE + Contador 'apunta a lo que hay despu�s del buffer de tiles
    DesplazamientoDE = &H50 'de= 80 (desplazamiento de medio tile)
'4b79
    For Contador = 1 To 15 '15 veces rellena con bloques de 4x4
'4b7b
        PunteroBufferSpritesAnterior = PunteroBufferSpritesIX
        'ValorRelleno = Leer16(TablaPatronRellenoLuz_48E8, PunteroPatronLuz - &H48E8) 'lee un valor de la tabla
        ValorRelleno = shl(TablaPatronRellenoLuz_48E8(PunteroPatronLuz - &H48E8), 8) Or TablaPatronRellenoLuz_48E8(PunteroPatronLuz + 1 - &H48E8) 'lee un valor de la tabla
        PunteroPatronLuz = PunteroPatronLuz + 2
'4B86
        DesplazAdsoX = SpriteLuzAdsoX_4B89 'posici�n x del sprite de adso dentro del tile
        If DesplazAdsoX <> 0 Then
'4b8e
            For Contador2 = 0 To DesplazAdsoX - 1
                BufferSprites_9500(PunteroBufferSpritesIX + 0 - &H9500&) = &HFF 'relleno negro, primera l�nea
                BufferSprites_9500(PunteroBufferSpritesIX + &H14 - &H9500&) = &HFF 'relleno negro, segunda l�nea
                BufferSprites_9500(PunteroBufferSpritesIX + &H28 - &H9500&) = &HFF 'relleno negro, tercera l�nea
                BufferSprites_9500(PunteroBufferSpritesIX + &H3C - &H9500&) = &HFF 'relleno negro, cuarta l�nea
                PunteroBufferSpritesIX = PunteroBufferSpritesIX + 1
            Next 'completa el relleno de la parte izquierda
        End If
'4b9e
        If SpriteLuzFlip_4BA0 Then
            ValorRelleno = shl(ValorRelleno, 1) '0x00 o 0x29 (si los gr�ficos de adso est�n flipeados o no)
        End If
        For Contador2 = 1 To 16 '16 bits tiene el valor de la tabla 48E8
            If (ValorRelleno And &H8000&) = 0 Then 'si el bit m�s significativo es 0, rellena de negro el bloque de 4x4
'4ba4
                BufferSprites_9500(PunteroBufferSpritesIX + 0 - &H9500&) = &HFF 'relleno negro
                BufferSprites_9500(PunteroBufferSpritesIX + &H14 - &H9500&) = &HFF 'relleno negro
                BufferSprites_9500(PunteroBufferSpritesIX + &H28 - &H9500&) = &HFF 'relleno negro
                BufferSprites_9500(PunteroBufferSpritesIX + &H3C - &H9500&) = &HFF 'relleno negro
            End If
'4bb0
            ValorRelleno = shl(ValorRelleno, 1)
            PunteroBufferSpritesIX = PunteroBufferSpritesIX + 1
        Next 'completa los 16 bits
'4BB4
        DesplazAdsoX = SpriteLuzAdsoX_4BB5  '4 - (posici�n x del sprite de adso & 0x03)
        For Contador2 = 1 To DesplazAdsoX  'completa la parte de los 16 pixels que sobra por la derecha seg�n la ampliaci�n de la posici�n x
'4bb6
                BufferSprites_9500(PunteroBufferSpritesIX + 0 - &H9500&) = &HFF 'relleno negro
                BufferSprites_9500(PunteroBufferSpritesIX + &H14 - &H9500&) = &HFF 'relleno negro
                BufferSprites_9500(PunteroBufferSpritesIX + &H28 - &H9500&) = &HFF 'relleno negro
                BufferSprites_9500(PunteroBufferSpritesIX + &H3C - &H9500&) = &HFF 'relleno negro
                PunteroBufferSpritesIX = PunteroBufferSpritesIX + 1
'4BC4
        Next 'completa la parte derecha
'4bc6
        PunteroBufferSpritesIX = PunteroBufferSpritesAnterior
        PunteroBufferSpritesIX = PunteroBufferSpritesIX + DesplazamientoDE
'4bcb
    Next 'repite hasta completar los 15 bloques de 4 pixels de alto
'4BCD
    For Contador = 0 To SpriteLuzTipoRelleno_4BD1  '0x00ef o 0x009f
        BufferSprites_9500(PunteroBufferSpritesIX + Contador - &H9500&) = &HFF 'rellena un tile o tile y medio de negro (la parte inferior del sprite de la luz)
    Next
End If
Exit Sub
GuardarArchivo "BufferSprites", BufferSprites_9500
End Sub

Function EsValidoPunteroBufferTiles(ByVal Puntero As Long) As Boolean
    'comprueba si un puntero al buffer de tiles est� dentro de sus l�mites
    If (Puntero - &H8D80&) >= 0 And (Puntero - &H8D80&) <= UBound(BufferTiles_8D80) Then EsValidoPunteroBufferTiles = True
End Function

Sub CopiarTilesBufferSprites_4D9E(ByVal ProfundidadMaxima As Long, ByVal ProfundidadMinima As Long, ByVal SpritesPilaProcesados As Boolean, ByVal PunteroBufferTilesIX As Long, ByVal PunteroBufferSpritesDE As Long, ByVal nXsprite As Byte, ByVal nYsprite As Byte)
'4dd9=ProfundidadMinima
'4afa=PunteroBufferSpritesDE
'bc=ProfundidadMaxima
'3095=ix=PunteroBufferTilesIX
'2dd7=nXsprite
'2dd8=nYsprite
'copia en el buffer de sprites los tiles que est�n entre la profundidad m�nima y la m�xima
'Exit Sub
Dim NtilesY As Long 'n�mero de tiles que ocupa el sprite en y
Dim NtilesX As Long 'n�mero de tiles que ocupa el sprite en x
Dim PunteroBufferTilesAnterior As Long
Dim PunteroBufferSpritesAnterior As Long
Dim PunteroBufferSpritesAnterior2 As Long
Dim Contador As Long
Dim Contador2 As Long
Dim ProcesarTileDirectamente_4DE4 As Boolean 'true si salta a 4E11 (procesar directamente), false salta a 4DE6 (comprobaciones previas)
Dim Valor As Byte
Dim ProfundidadX As Byte
Dim ProfundidadY As Byte
Dim ProfundidadMinimaX As Byte
Dim ProfundidadMinimaY As Byte
Dim ProfundidadMaximaX As Byte
Dim ProfundidadMaximaY As Byte
Dim ProcesarTile As Boolean
Dim Contador3 As Long
Dim PunteroBufferTilesAnterior3 As Long
Dim BugOverflow As Boolean 'true si el puntero a la tabla de tiles est� fuera
'On Error Resume Next


Dim H4dd9 As String
Dim DE As String
Dim BC As String
Dim IX As String
H4dd9 = Hex$(ProfundidadMaxima)
DE = Hex$(PunteroBufferSpritesDE)
IX = Hex$(PunteroBufferTilesIX)



PunteroBufferTilesAnterior3 = PunteroBufferTilesIX
ProfundidadMaxima = ProfundidadMaxima + 257
ProfundidadMinimaX = LeerByteLong(ProfundidadMinima, 0)
ProfundidadMinimaY = LeerByteLong(ProfundidadMinima, 1)
ProfundidadMaximaX = LeerByteLong(ProfundidadMaxima, 0)
ProfundidadMaximaY = LeerByteLong(ProfundidadMaxima, 1)
'4DB8
NtilesY = shr(Byte2Long(nYsprite), 3) 'nysprite = nysprite/8 (n�mero de tiles que ocupa el sprite en y)
NtilesX = shr(Byte2Long(nXsprite), 2) 'nxsprite = nxsprite/4 (n�mero de tiles que ocupa el sprite en x)
'4dc2
For Contador3 = 1 To NtilesY
    PunteroBufferTilesAnterior = PunteroBufferTilesIX
    PunteroBufferSpritesAnterior = PunteroBufferSpritesDE
    For Contador = 1 To NtilesX
'4DC9
        ProcesarTileDirectamente_4DE4 = False
        For Contador2 = 1 To 2 'cada tile tiene 2 prioridades
'4DD1
            IX = Hex$(PunteroBufferTilesIX)
            If EsValidoPunteroBufferTiles(PunteroBufferTilesIX) Then
                BugOverflow = False
                Valor = BufferTiles_8D80(PunteroBufferTilesIX + 2 - &H8D80&) 'lee el n�mero de tile de la entrada actual del buffer de tiles
            Else 'correcci�n bug del programa original. en algunas pantallas parte de la cabeza de guillermo queda fuera
                BugOverflow = True
                Valor = LeerByteTablaCualquiera(PunteroBufferTilesIX + 2)
            End If
            If Valor <> 0 Then
'4DD7
                ProcesarTile = False
                If Not BugOverflow Then
                    ProfundidadX = BufferTiles_8D80(PunteroBufferTilesIX + 0 - &H8D80&) 'lee la profundidad en x del tile actual
                Else
                    ProfundidadX = LeerByteTablaCualquiera(PunteroBufferTilesIX + 0)
                End If
                'si en esta llamada no se ha pintado en esta posici�n del buffer de tiles, comprueba si hay que pintar el
                'tile que hay en esta capa de profundidad. Si se ha pintado y el tile de esta capa se hab�a pintado
                'en otra iteraci�n anterior, lo combina sin comprobar la profundidad
                If (ProfundidadX And &H80) = 0 Or (((ProfundidadX And &H80)) <> 0 And Not ProcesarTileDirectamente_4DE4) Then
'4de3
                    'If Not ProcesarTileDirectamente_4DE4 Then
'4de6
                        If Not BugOverflow Then
                            ProfundidadY = BufferTiles_8D80(PunteroBufferTilesIX + 1 - &H8D80&) 'lee la profundidad en y del tile actual
                        Else
                            ProfundidadY = LeerByteTablaCualquiera(PunteroBufferTilesIX + 1)
                        End If
                        If (ProfundidadX >= ProfundidadMinimaX Or ProfundidadY >= ProfundidadMinimaY) And _
                           (ProfundidadX < ProfundidadMaximaX And ProfundidadY < ProfundidadMaximaY) And (ProfundidadX And &H80) = 0 Then
                            ProcesarTile = True
'4e00
                            'aqu� llega si el tile tiene mayor profundidad que el m�nimo y menor profundidad que el sprite
                            ProcesarTileDirectamente_4DE4 = True 'modifica un salto para indicar que en esta llamada ha pintado alg�n tile para esta posici�n del buffer de tiles
'4E07
                            If EsDireccionBufferTiles_37A5(PunteroBufferTilesIX) Then 'si ix est� dentro del buffer de tiles
                                If Not BugOverflow Then
                                    BufferTiles_8D80(PunteroBufferTilesIX + 0 - &H8D80&) = BufferTiles_8D80(PunteroBufferTilesIX + 0 - &H8D80&) Or &H80 'indica que se ha procesado este tile
                                End If
                            End If
                            
                        Else
                            ProcesarTile = False
                        End If
                    'Else
                        'ProcesarTile = True
                    'End If
                Else
                    ProcesarTile = True
                End If
'4e11
                If ProcesarTile Then
                    PunteroBufferSpritesAnterior2 = PunteroBufferSpritesDE

                    DE = Hex$(PunteroBufferSpritesDE)
                    IX = Hex$(PunteroBufferTilesIX)
                    CombinarTileBufferSprites_4E49 PunteroBufferTilesIX, PunteroBufferSpritesDE, nXsprite
                    PunteroBufferSpritesDE = PunteroBufferSpritesAnterior2
                End If
            End If
'4E1B
            'avanza al siguiente tile o a la siguiente prioridad
            If EsValidoPunteroBufferTiles(PunteroBufferTilesIX) Then
                LimpiarBit7BufferTiles_4D85 SpritesPilaProcesados, PunteroBufferTilesIX 'ret (si no ha terminado de procesar los sprites de la pila) o limpia el bit 7 de (ix+0) del buffer de tiles (si es una posici�n v�lida del buffer)
            End If
            PunteroBufferTilesIX = PunteroBufferTilesIX + 3 'pasa al tile de mayor prioridad del buffer de tiles
'4e25
        Next 'repite hasta que se hayan completado las prioridades de la entrada del buffer de tiles
'4e27
        PunteroBufferSpritesDE = PunteroBufferSpritesDE + 4 'pasa a la posici�n del siguiente tile en x del buffer de sprites
'4e2d
    Next 'repite mientras no se termine en x
'4e2f
    PunteroBufferSpritesDE = PunteroBufferSpritesAnterior
    PunteroBufferSpritesDE = PunteroBufferSpritesDE + 8 * nXsprite 'pasa a la posici�n del siguiente tile en y del buffer de sprites (ancho del sprite*8)
    PunteroBufferTilesIX = PunteroBufferTilesAnterior 'recupera la posici�n del buffer de tiles
    PunteroBufferTilesIX = PunteroBufferTilesIX + &H60 'pasa a la siguiente l�nea del buffer de tiles
'4e45
Next 'repite hasta que se acaben los tiles en y
PunteroBufferTilesIX = PunteroBufferTilesAnterior3
End Sub

Sub LimpiarBit7BufferTiles_4D85(ByVal SpritesPilaProcesados As Boolean, ByVal PunteroBufferTilesIX As Long)
'vuelve si no ha terminado de procesar los sprites de la pila o limpia el bit 7 de (ix+0) del buffer de tiles (si es una posici�n v�lida del buffer)
If Not SpritesPilaProcesados Then Exit Sub
If EsDireccionBufferTiles_37A5(PunteroBufferTilesIX) Then
    'If PunteroBufferTilesIX + 0 - &H8D80& >= 0 And PunteroBufferTilesIX + 0 - &H8D80& < UBound(BufferTiles_8D80) Then
        BufferTiles_8D80(PunteroBufferTilesIX + 0 - &H8D80&) = BufferTiles_8D80(PunteroBufferTilesIX + 0 - &H8D80&) And &H7F 'limpia el bit mas significativo del buffer de tiles
    'End If
End If
End Sub

Function EsDireccionBufferTiles_37A5(ByVal PunteroBufferTilesIX As Long) As Boolean
'dada una direcci�n, devuelve true si es una direcci�n v�lida del buffer de tiles
If PunteroBufferTilesIX >= &H8D80 Then EsDireccionBufferTiles_37A5 = True '8d80=inicio del buffer de tiles
End Function

Sub CombinarTileBufferSprites_4E49(ByVal PunteroBufferTilesIX As Long, ByVal PunteroBufferSpritesDE As Long, ByVal nXsprite As Byte)
'aqu� entra con PunteroBufferTilesIX apuntando a alguna entrada del buffer de tiles y PunteroBufferSpritesDE apuntando
'a alguna posici�n del buffer de sprites
'combina el tile de la entrada actual de ix en la posici�n actual del buffer de sprites
Dim NumeroTile As Byte
Dim PunteroDatosTile As Long
Dim Contador As Long
Dim Contador2 As Long
Dim PunteroTablasAndOr As Long
Dim MascaraAnd As Byte
Dim MascaraOr As Byte
Dim Valor As Byte
Dim BugOverflow As Boolean
If PunteroPerteneceTabla(PunteroBufferTilesIX, BufferTiles_8D80, &H8D80&) Then
    NumeroTile = BufferTiles_8D80(PunteroBufferTilesIX + 2 - &H8D80&) 'n�mero de tile de la entrada actual
Else
    NumeroTile = LeerByteTablaCualquiera(PunteroBufferTilesIX + 2)
    BugOverflow = True
End If
PunteroDatosTile = Byte2Long(NumeroTile) * 32 'cada tile ocupa 32 bytes
PunteroDatosTile = PunteroDatosTile + &H6D00 'a partir de 0x6d00 est�n los gr�ficos de los tiles que forman las pantallas
If NumeroTile < &HB Then 'si el gr�fico es menor que el 0x0b (gr�ficos sin transparencia, caso m�s sencillo)
'4e92
    'aqu� llega si el n�mero de tile era < 0x0b (son gr�ficos sin transparencia)
    For Contador = 1 To 8 '8 pixels de alto
        For Contador2 = 1 To 4 '4 bytes de ancho (16 pixels)
            BufferSprites_9500(PunteroBufferSpritesDE - &H9500&) = TilesAbadia_6D00(PunteroDatosTile - &H6D00)
            PunteroBufferSpritesDE = PunteroBufferSpritesDE + 1
            PunteroDatosTile = PunteroDatosTile + 1
        Next
'4ea7
        PunteroBufferSpritesDE = PunteroBufferSpritesDE + nXsprite - 4 'pasa a la siguiente l�nea del sprite
'4eae
    Next
'4eb0
Else
'4e60
    'si el gr�fico es mayor o igual que 0x0b (gr�ficos con transparencia)
    If Not BugOverflow Then
        If (BufferTiles_8D80(PunteroBufferTilesIX + 2 - &H8D80&) And &H80) = 0 Then 'comprueba que tabla usar seg�n el n�mero de tile que haya
            PunteroTablasAndOr = &H9D00& 'tablas 0 y 1
        Else
            PunteroTablasAndOr = &H9F00& 'tablas 2 y 3
        End If
    Else
        If (NumeroTile And &H80) = 0 Then 'comprueba que tabla usar seg�n el n�mero de tile que haya
            PunteroTablasAndOr = &H9D00& 'tablas 0 y 1
        Else
            PunteroTablasAndOr = &H9F00& 'tablas 2 y 3
        End If
    
    End If
    For Contador = 1 To 8 '8 pixels de alto
        For Contador2 = 1 To 4 '4 bytes de ancho (16 pixels)
'4e75
            Valor = TilesAbadia_6D00(PunteroDatosTile - &H6D00) 'obtiene un byte del gr�fico
            MascaraOr = TablasAndOr_9D00(PunteroTablasAndOr + Byte2Long(Valor) - &H9D00&) 'obtiene el or
            MascaraAnd = TablasAndOr_9D00(PunteroTablasAndOr + Byte2Long(Valor) + 256 - &H9D00&) 'obtiene el and
            Valor = BufferSprites_9500(PunteroBufferSpritesDE - &H9500&) 'obtiene un valor del buffer de sprites
            Valor = (Valor And MascaraAnd) Or MascaraOr 'aplica el valor de las m�scaras
            BufferSprites_9500(PunteroBufferSpritesDE - &H9500&) = Valor 'graba el valor obtenido combinando el fondo con el sprite
            PunteroBufferSpritesDE = PunteroBufferSpritesDE + 1 'avanza a la siguiente posici�n del buffer
            PunteroDatosTile = PunteroDatosTile + 1 'avanza al siguiente byte del gr�fico
'4e83
            Next
'4e86
            PunteroBufferSpritesDE = PunteroBufferSpritesDE + nXsprite - 4 'pasa a la siguiente l�nea del sprite
        Next 'repite hasta que se complete el alto del tile
'4e91
End If
'GuardarArchivo "D:\datos\vbasic\Abadia\Abadia2\BufferSprites", BufferSprites_9500
End Sub

Sub CopiarSpritePantalla_4C1A(ByVal PunteroSpriteIX As Long)
'vuelca el buffer del sprite a la pantalla
Dim Xnovisible As Byte 'distancia en x de lo que no es visible
Dim Xsprite As Byte 'posici�n en x del tile en el que empieza el sprite (en bytes)
Dim Ysprite As Byte 'posici�n en y del tile en el que empieza el sprite
Dim nXsprite As Byte 'ancho final del sprite (en bytes)
Dim nYsprite As Byte 'alto final del sprite (en pixels)
Dim PunteroBufferSpritesHL As Long 'direcci�n del buffer de sprites asignada a este sprite
Dim PunteroPantallaDE As Long 'posici�n en pantalla donde copiar los bytes
Dim PunteroPantallaAnterior As Long
Dim Contador As Long
Dim Contador2 As Long
Dim ValorPantalla As Byte
'4C1A
Xnovisible = 0 'distancia en x de lo que no es visible
Ysprite = TablaSprites_2E17(PunteroSpriteIX + &HD - &H2E17) 'posici�n en y del tile en el que empieza el sprite
nYsprite = TablaSprites_2E17(PunteroSpriteIX + &HF - &H2E17) 'alto final del sprite (en pixels)
PunteroPantallaDE = 0
If Ysprite >= 200 Then Exit Sub 'si la coordenada y >= 200 (no es visible en pantalla), sale
'4C2D
If Ysprite <= 40 Then 'si la coordenada y <= 40 (no visible o visible en parte en pantalla)
    If (40 - Ysprite) >= nYsprite Then 'si la distancia desde el punto en que comienza el sprite al primer punto visible >= la altura del sprite, sale (no visible)
        Exit Sub
    End If
'4C36
    nXsprite = TablaSprites_2E17(PunteroSpriteIX + &HE - &H2E17)
    PunteroPantallaDE = (40 - Ysprite) * nXsprite 'avanza las l�neas del sprite no visible
    nYsprite = nYsprite - (40 - Ysprite) 'modifica el alto del sprite por el recorte
    Ysprite = 0 'el sprite empieza en y = 0
Else
    Ysprite = Ysprite - 40 'ajusta la coordenada y
End If
'4C45
'direcci�n del buffer de sprites asignada a este sprit
PunteroBufferSpritesHL = Bytes2Long(TablaSprites_2E17(PunteroSpriteIX + &H10 - &H2E17), TablaSprites_2E17(PunteroSpriteIX + &H11 - &H2E17))
PunteroBufferSpritesHL = PunteroBufferSpritesHL + PunteroPantallaDE 'salta los bytes no visibles en y
'4C4E
Xsprite = TablaSprites_2E17(PunteroSpriteIX + &HC - &H2E17) 'posici�n en x del tile en el que empieza el sprite (en bytes)
nXsprite = TablaSprites_2E17(PunteroSpriteIX + &HE - &H2E17) 'ancho final del sprite (en bytes)
If Xsprite >= 72 Then Exit Sub 'sale si la posici�n en x >= (32 + 256 pixels)
'4C58
If Xsprite < 8 Then 'si la coordenada x <= 32 (no visible o visible en parte en pantalla)
'4C5D
    If (8 - Xsprite) >= nXsprite Then 'si la distancia desde el punto en que comienza el sprite al primer punto visible >= la anchura del sprite, sale (no visible)
        Exit Sub
    End If
'4C63
    PunteroBufferSpritesHL = PunteroBufferSpritesHL + 8 - Xsprite 'avanza los pixels recortados
    Xnovisible = 8 - Xsprite
    nXsprite = nXsprite - (8 - Xsprite) 'modifica el ancho a pintar
    Xsprite = 0 'el sprite empieza en x = 0
Else
    Xsprite = Xsprite - 8
End If
'4c72
If (Xsprite + nXsprite) >= 64 Then  'comprueba si el sprite es m�s ancho que la pantalla (64*4 = 256)
    Xnovisible = nXsprite + Xsprite - 64
    nXsprite = nXsprite - Xnovisible 'pone un nuevo ancho para el sprite
End If
'4C7F
If (Ysprite + nYsprite) >= 160 Then 'comprueba si el sprite es m�s alto que la pantalla (160)
'4C8A
    nYsprite = nYsprite - (Ysprite + nYsprite - 160) 'actualiza el alto a pintar
End If
'4C8E
PunteroPantallaDE = ObtenerDesplazamientoPantalla_3C42(Xsprite, Ysprite) 'dadas coordenadas X,Y, calcula el desplazamiento correspondiente en pantalla
'4C95
For Contador = nYsprite To 1 Step -1
'4C9A
    PunteroPantallaAnterior = PunteroPantallaDE
    For Contador2 = nXsprite To 1 Step -1
        ValorPantalla = BufferSprites_9500(PunteroBufferSpritesHL - &H9500&)
        PantallaCGA(PunteroPantallaDE - &HC000&) = ValorPantalla
        PantallaCGA2PC PunteroPantallaDE - &HC000&, ValorPantalla
        PunteroBufferSpritesHL = PunteroBufferSpritesHL + 1
        PunteroPantallaDE = PunteroPantallaDE + 1
    Next
    PunteroBufferSpritesHL = PunteroBufferSpritesHL + Xnovisible
    PunteroPantallaDE = PunteroPantallaAnterior
'4CA7
    'pasa a la siguiente l�nea de pantalla
    PunteroPantallaDE = PunteroPantallaDE + &H800& 'pasa al siguiente banco
    If (PunteroPantallaDE And &H3800&) = 0 Then 'banco inexistente
        PunteroPantallaDE = PunteroPantallaDE - &H800& 'vuelve al banco anterior
        PunteroPantallaDE = PunteroPantallaDE And &HC7FF&
        PunteroPantallaDE = PunteroPantallaDE + &H50 'cada l�nea ocupa 0x50 bytes
    End If
'4CBC
Next
modPantalla.Refrescar
End Sub

Function ObtenerDesplazamientoPantalla_3C42(ByVal X As Byte, ByVal Y As Byte) As Long
'; dados X,Y, calcula el desplazamiento correspondiente en pantalla
'al valor calculado se le suma 32 pixels a la derecha (puesto que el �rea de juego va desde x = 32 a x = 256 + 32 - 1
'l = coordenada X (en bytes)
Dim PunteroPantalla As Long
Dim ValorLong As Long
PunteroPantalla = Byte2Long(Y And &HF8) 'obtiene el valor para calcular el desplazamiento dentro del banco de VRAM
'dentro de cada banco, la l�nea a la que se quiera ir puede calcularse como (y & 0xf8)*10
'o lo que es lo mismo, (y >> 3)*0x50
PunteroPantalla = 10 * PunteroPantalla 'PunteroPantalla = desplazamiento dentro del banco
ValorLong = Byte2Long(Y And &H7) '3 bits menos significativos en y (para calcular al banco de VRAM al que va)
ValorLong = shl(ValorLong, 11) 'ajusta los 3 bits
PunteroPantalla = PunteroPantalla Or ValorLong 'completa el c�lculo del banco
PunteroPantalla = PunteroPantalla Or &HC000&
PunteroPantalla = PunteroPantalla + Byte2Long(X) 'suma el desplazamiento en x
PunteroPantalla = PunteroPantalla + 8 'ajusta para que salga 32 pixels m�s a la derecha
ObtenerDesplazamientoPantalla_3C42 = PunteroPantalla
End Function

Sub ActualizarDatosPersonaje_291D(ByVal PunteroPersonajeHL As Long)
'comprueba si el personaje puede moverse a donde quiere y actualiza su sprite y el buffer de alturas
'PunteroPersonajeHL apunta a la tabla del personaje a mover
'&H2BAE 'guillermo
'&H2BB8 'adso
'&H2BC2 'malaqu�as
'&H2BCC 'abad
'&H2BD6 'berengario
'&H2BE0 'severino
Dim PunteroSpriteIX As Long
Dim PunteroCaracteristicasPersonajeIY As Long
Dim PunteroRutinaComportamientoHL As Long
Dim PunteroRutinaFlipearGraficos As Long
Dim Valor As Byte
PunteroSpriteIX = modFunciones.Leer16(TablaDatosPersonajes_2BAE, PunteroPersonajeHL + 0 - &H2BAE)
PunteroCaracteristicasPersonajeIY = modFunciones.Leer16(TablaDatosPersonajes_2BAE, PunteroPersonajeHL + 2 - &H2BAE)
PunteroRutinaComportamientoHL = modFunciones.Leer16(TablaDatosPersonajes_2BAE, PunteroPersonajeHL + 4 - &H2BAE)
PunteroRutinaFlipearGraficos = modFunciones.Leer16(TablaDatosPersonajes_2BAE, PunteroPersonajeHL + 6 - &H2BAE)
PunteroRutinaFlipPersonaje_2A59 = PunteroRutinaFlipearGraficos
PunteroTablaAnimacionesPersonaje_2A84 = modFunciones.Leer16(TablaDatosPersonajes_2BAE, PunteroPersonajeHL + 8 - &H2BAE)
DefinirDatosSpriteComoAntiguos_2AB0 PunteroSpriteIX 'pone la posici�n y dimensiones actuales del sprite como posici�n y dimensiones antiguas
'si la posici�n del sprite es central y la altura est� bien, limpia las posiciones que ocupaba el sprite en el buffer de alturas
'292f
RellenarBufferAlturasPersonaje_28EF PunteroCaracteristicasPersonajeIY, 0
If MalaquiasAscendiendo_4384 Then
    MalaquiasAscendiendo_4384 = False
'2945
    AvanzarAnimacionSprite_2A27 PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroRutinaFlipearGraficos
Else
'2948
    'lee el contador de la animaci�n
    Valor = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 0 - &H3036)
    If (Valor And &O1&) <> 0 Then
'294d
        IncrementarContadorAnimacionSprite_2A01 PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroRutinaFlipearGraficos
    Else
'2950
        Select Case PunteroRutinaComportamientoHL
            Case Is = &H288D 'guillermo
                EjecutarComportamientoGuillermo_288D PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroRutinaFlipearGraficos
            Case Is = &H2C3A 'resto
                'EjecutarComportamientoPersonaje_2C3A PunteroSpriteIX, PunteroCaracteristicasPersonajeIY
        End Select
    End If
End If
'2940
'lee el valor a poner en el buffer de alturas para indicar que est� el personaje
Valor = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + &HE - &H3036)
'2943
'si la posici�n del sprite es central y la altura est� bien, pone c en las posiciones que ocupa del buffer de alturas
RellenarBufferAlturasPersonaje_28EF PunteroCaracteristicasPersonajeIY, Valor
End Sub

Sub AvanzarAnimacionSprite_2A27(ByVal PunteroSpriteIX As Long, ByVal PunteroCaracteristicasPersonajeIY As Long, ByVal PunteroRutinaFlipearGraficos As Long)
'avanza la animaci�n del sprite y lo redibuja
Dim PunteroTablaAnimacionesHL As Long
Dim Yp As Byte 'posici�n y en pantalla del sprite
Dim Valor As Byte
'cambia la animaci�n de los trajes de los monjes seg�n la posici�n y el contador de animaciones y
'obtiene la direcci�n de los datos de la animaci�n que hay que poner en hl
PunteroTablaAnimacionesHL = CambiarAnimacionTrajesMonjes_2A61(PunteroSpriteIX, PunteroCaracteristicasPersonajeIY)
MovimientoRealizado_2DC1 = True 'indica que ha habido movimiento
If EsSpriteVisible_2AC9(PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, Yp) = True Then
    'aqu� se llega desde fuera si un sprite es visible, despu�s de haber actualizado su posici�n.
    'en PunteroTablaAnimacionesHL se apunta a la animaci�n correspondiente para el sprite
    'PunteroSpriteIX = direcci�n del sprite correspondiente
    'PunteroCaracteristicasPersonajeIY = datos de posici�n del personaje correspondiente
'2a34
    ActualizarDatosGraficosPersonaje_2A34 PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroTablaAnimacionesHL, Yp
End If
End Sub

Function EsSpriteVisible_2AC9(ByVal PunteroSpriteIX As Long, ByVal PunteroCaracteristicasPersonajeIY As Long, ByVal Yp As Byte) As Boolean
Dim Visible As Boolean
Dim X As Byte
Dim Y As Byte
Dim Z As Byte
'comprueba si es visible y si lo es, actualiza su posici�n si fuese necesario.
'Si es visible no vuelve, sino que sale a la rutina que lo llam�
Visible = ProcesarObjeto_2ADD(PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, X, Y, Z, Yp)
If Visible Then
    EsSpriteVisible_2AC9 = True
    Exit Function
End If
'aqu� llega si el sprite no es visible
If TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17) = &HFE Then 'si el sprite no era visible, sale
    Exit Function
Else
    TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17) = &H80 'en otro caso, indica que hay que redibujar el sprite
    TablaSprites_2E17(PunteroSpriteIX + 5 - &H2E17) = TablaSprites_2E17(PunteroSpriteIX + 5 - &H2E17) Or &H80 'indica que el sprite va a pasar a inactivo, y solo se quiere redibujar la zona que ocupaba
End If
End Function

Sub IncrementarContadorAnimacionSprite_2A01(ByVal PunteroSpriteIX As Long, ByVal PunteroCaracteristicasPersonajeIY As Long, ByVal PunteroRutinaFlipearGraficos As Long)
'incrementa el contador de los bits 0 y 1 del byte 0, avanza la animaci�n del sprite y lo redibuja
Dim Valor As Byte
'lee el contador de la animaci�n
Valor = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 0 - &H3036)
Valor = Valor + 1
Valor = Valor And 3
'2a07
TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 0 - &H3036) = Valor
'2A0A
AvanzarAnimacionSprite_2A27 PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroRutinaFlipearGraficos
End Sub

Sub EjecutarComportamientoGuillermo_288D(ByVal PunteroSpriteIX As Long, ByVal PunteroCaracteristicasPersonajeIY As Long, ByVal PunteroRutinaFlipearGraficos As Long)
'rutina del comportamiento de guillermo
'PunteroSpriteIX que apunta al sprite de guillermo
'PunteroCaracteristicasPersonajeIY apunta a los datos de posici�n de guillermo
Dim Valor As Byte
Dim RetornoA As Long
Dim RetornoC As Long
Dim RetornoHL As Long

If EstadoGuillermo_288F <> 0 Then
'2893
    If EstadoGuillermo_288F = 1 Then Exit Sub 'si EstadoGuillermo_288F era 1, sale
    EstadoGuillermo_288F = EstadoGuillermo_288F - 1
    If EstadoGuillermo_288F = &H13 Then
'289C
        'aqu� llega si el estado de guillermo es 0x13
        If AjustePosicionYSpriteGuillermo_28B1 = 2 Then
'28a3
            'decrementa la posici�n en x de guillermo
            Valor = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 2 - &H3036)
            Valor = Valor - 1
            TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 2 - &H3036) = Valor
            'avanza la animaci�n del sprite y lo redibuja
            AvanzarAnimacionSprite_2A27 PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroRutinaFlipearGraficos
            Exit Sub
        End If
'28a9
        'si se modifica la y del sprite con 1, salta y marca el sprite como inactivo
        If EstadoGuillermo_288F <> 1 Then
'28ad
            'modifica la posici�n y del sprite
            Valor = TablaSprites_2E17(PunteroSpriteIX + 2 - &H2E17)
            Valor = Valor + AjustePosicionYSpriteGuillermo_28B1
            TablaSprites_2E17(PunteroSpriteIX + 2 - &H2E17) = Valor
            Valor = TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17)
            Valor = Valor And &H3F
            Valor = Valor Or &H80
            TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17) = Valor 'marca el sprite para dibujar
            MovimientoRealizado_2DC1 = True 'indica que ha habido movimiento
        Else
'28c5
            'aqu� llega si se modifica la y del sprite con 1 y el estado de guillermo es el 0x13
            TablaSprites_2E17(PunteroSpriteIX + 0 - &H2E17) = &HFE 'marca el sprite como inactivo
        End If
    End If
Else
'28ca
    'aqu� llega si el estado de guillermo es 0, que es el estado normal
    If PersonajeSeguidoPorCamara_3C8F <> 0 Then Exit Sub 'si la c�mara no sigue a guillermo, sale
'28CF
    If modTeclado.TeclaPulsadaFlanco(TeclaIzquierda) Then
'2a0c
        ActualizarDatosPersonajeCursorIzquierdaDerecha_2A0C True, PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroRutinaFlipearGraficos
    End If
'28d9
    If modTeclado.TeclaPulsadaFlanco(TeclaDerecha) Then 'comprueba si ha cambiado el estado de cursor derecha
'2a0c
        ActualizarDatosPersonajeCursorIzquierdaDerecha_2A0C False, PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroRutinaFlipearGraficos
    Else
'28e3
        If modTeclado.TeclaPulsadaNivel(TeclaArriba) = False Then Exit Sub 'si no se ha pulsado el cursor arriba, sale
        
'28E9
        ObtenerAlturaDestinoPersonaje_27B8 0, PunteroCaracteristicasPersonajeIY, RetornoA, RetornoC, RetornoHL
        
'28EC
        AvanzarPersonaje_2954 PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroRutinaFlipearGraficos, RetornoA, RetornoC, RetornoHL
    End If

End If



End Sub

Sub ActualizarDatosPersonajeCursorIzquierdaDerecha_2A0C(ByVal Izquierda As Boolean, ByVal PunteroSpriteIX As Long, ByVal PunteroCaracteristicasPersonajeIY As Long, ByVal PunteroRutinaFlipearGraficos As Long)
'aqu� llega si se ha pulsado cursor derecha o izquierda
Dim Valor As Byte
TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 0 - &H3036) = 0 'resetea el contador de la animaci�n
'2A10
If (TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 5 - &H3036) And &H80) <> 0 Then
'2a16
    Valor = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 5 - &H3036)
    Valor = Valor Xor &H20
    TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 5 - &H3036) = Valor
End If
'2a1e
Valor = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 1 - &H3036) 'lee la orientaci�n
'cambia la orientaci�n del personaje
If Izquierda Then
    Valor = (Valor + 1) And &H3
Else
    Valor = (Valor + 255) And &H3
End If

TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 1 - &H3036) = Valor


AvanzarAnimacionSprite_2A27 PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroRutinaFlipearGraficos





End Sub

Sub ObtenerAlturaDestinoPersonaje_27B8(ByVal DiferenciaAlturaA As Byte, ByVal PunteroCaracteristicasPersonajeIY As Long, ByRef Salida1A As Long, ByRef Salida2C As Long, ByRef Salida3HL As Long)
'comprueba la altura de las posiciones a las que va a moverse el personaje y las devuelve en Salida1A y Salida2C
'en Salida3HL devuelve el puntero en la tabla TablaAvancePersonaje con los incrementos necesarios en x e y para avanzar el personaje
'si el personaje no est� en la pantalla actual, se devuelve lo mismo que se pas� en DiferenciaAlturaA (se supone que ya se ha calculado la diferencia de altura fuera)
'DiferenciaAlturaA se usar� si el personaje no est� en la pantalla actual
'en PunteroCaracteristicasPersonajeIY se pasan las caracter�sticas del personaje que se mueve hacia delante
'llamado al pulsar cursor arriba
Dim AlturaPersonaje As Byte
Dim AlturaBasePlanta As Byte
Dim AlturaRelativa As Byte 'altura relativa dentro de la planta
'27b9
AlturaPersonaje = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 4 - &H3036) 'obtiene la altura del personaje
'27bc
AlturaBasePlanta = LeerAlturaBasePlanta_2473(AlturaPersonaje)
If AlturaBasePlanta <> AlturaBasePlantaActual_2DBA Then 'si no coincide la planta en la que est� el personaje con la que se est� mostrando, sale
    Salida1A = DiferenciaAlturaA
    Exit Sub
End If
'27c6
AlturaRelativa = AlturaPersonaje - AlturaBasePlanta
'27CB
ObtenerAlturaDestinoPersonaje_27CB AlturaRelativa, PunteroCaracteristicasPersonajeIY, Salida1A, Salida2C, Salida3HL
End Sub

Sub ObtenerAlturaDestinoPersonaje_27CB(ByVal DiferenciaAlturaA As Byte, ByVal PunteroCaracteristicasPersonajeIY As Long, ByRef Salida1A As Long, ByRef Salida2C As Long, ByRef Salida3HL As Long)
'aqu� llega con DiferenciaAlturaA = altura relativa dentro de la planta
Dim PosicionX As Byte 'posici�n global del personaje
Dim PosicionY As Byte 'posici�n global del personaje
Dim PunteroBufferAlturas As Long
Dim PunteroBufferAlturasAnterior As Long
Dim PunteroTablaAvancePersonaje As Long 'puntero a la tabla de incrementos
Dim IncrementoBucleInterior As Long
Dim IncrementoBucleExterior As Long
Dim IncrementoInicial As Long
Dim ContadorExterior As Long
Dim ContadorInterior As Long
Dim BufferAuxiliar_2DC5(&HF) As Long
Dim PunteroBufferAuxiliar As Long
Dim ValorBufferAlturas As Byte
'obtiene la posici�n global del personaje
PosicionY = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 3 - &H3036)
PosicionX = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 2 - &H3036)
If Not DeterminarPosicionCentral_279B(PosicionX, PosicionY) Then 'PosicionX,PosicionY = posici�n ajustada a las 20x20 posiciones centrales
'27d8
    Salida1A = DiferenciaAlturaA
    Exit Sub
End If
'aqu� llega si la posici�n es visible
'27db
PunteroBufferAlturas = 24 * PosicionY + PosicionX
'27EC
PunteroBufferAlturas = PunteroBufferAlturas + PunteroBufferAlturas_2D8A 'indexa en el buffer de alturas
'27EE
PunteroTablaAvancePersonaje = ObtenerPunteroPosicionVecinaPersonaje_2783(PunteroCaracteristicasPersonajeIY)
'27F1
IncrementoBucleInterior = LeerDatoTablaAvancePersonaje(PunteroTablaAvancePersonaje, 16)
IncrementoBucleExterior = LeerDatoTablaAvancePersonaje(PunteroTablaAvancePersonaje + 2, 16)
IncrementoInicial = LeerDatoTablaAvancePersonaje(PunteroTablaAvancePersonaje + 4, 16)
Salida3HL = PunteroTablaAvancePersonaje + 6
'280A
PunteroBufferAlturas = PunteroBufferAlturas + IncrementoInicial 'suma a la posici�n actual en el buffer de alturas el desplazamiento leido
'280B
PunteroBufferAuxiliar = &H2DC5 'apunta a un buffer auxiliar
'280E
For ContadorExterior = 1 To 4 'el bucle exterior realiza 4 iteraciones
'2811
    PunteroBufferAlturasAnterior = PunteroBufferAlturas
'2812
    For ContadorInterior = 1 To 4 'el bucle interior realiza 4 iteraciones
'2815
        ValorBufferAlturas = TablaBufferAlturas_01C0(PunteroBufferAlturas - &H1C0)
        If ValorBufferAlturas < &H10 Then 'comprueba si en esa posici�n hay algun personaje
'281E
            BufferAuxiliar_2DC5(PunteroBufferAuxiliar - &H2DC5) = CLng(ValorBufferAlturas) - CLng(DiferenciaAlturaA)
        Else
'281A
            BufferAuxiliar_2DC5(PunteroBufferAuxiliar - &H2DC5) = ValorBufferAlturas And &H30&
        End If
'2821
        PunteroBufferAuxiliar = PunteroBufferAuxiliar + 1
'2822
        PunteroBufferAlturas = PunteroBufferAlturas + IncrementoBucleInterior
    Next
'282C
    PunteroBufferAlturas = PunteroBufferAlturasAnterior + IncrementoBucleExterior
Next
'2833
If (TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 5 - &H3036) And &H80) Then  'si el personaje s�lo ocupa 1 posici�n
'2839
    'guarda en a y en c el contenido de las 2 posiciones hacia las que avanza el personaje
    Salida2C = BufferAuxiliar_2DC5(&H2DC6 - &H2DC5)
    Salida1A = BufferAuxiliar_2DC5(&H2DCA - &H2DC5)
Else 'si el personaje ocupa 4 posiciones en el buffer de alturas
'2841
    'si en las 2 posiciones en las que se avanza no hay lo mismo, sale con valores iguales para a y c
    Salida2C = BufferAuxiliar_2DC5(&H2DC6 - &H2DC5)
    Salida1A = BufferAuxiliar_2DC5(&H2DC7 - &H2DC5)
    If Salida1A <> Salida2C Then
        Salida1A = 2 'indica que hay una diferencia entre las alturas > 1
    End If
End If
End Sub

Function ObtenerPunteroPosicionVecinaPersonaje_2783(ByVal PunteroCaracteristicasPersonajeIY As Long) As Long
'devuelve la direcci�n de la tabla para calcular la altura de las posiciones vecinas
'seg�n el tama�o de la posici�n del personaje y la orientaci�n
'iy=3072,a=0->284d
Dim OrientacionA As Long
'obtiene la orientaci�n del personaje
'278f
OrientacionA = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 1 - &H3036)
If (TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 5 - &H3036) And &H80) Then
'2792
    ObtenerPunteroPosicionVecinaPersonaje_2783 = &H286D + 8 * OrientacionA
Else 'si el bit 7 no est� puesto (si el personaje ocupa 4 tiles)
    'apunta a la tabla si el personaje ocupa 4 tiles
'2792
    ObtenerPunteroPosicionVecinaPersonaje_2783 = &H284D + 8 * OrientacionA
End If
End Function

Private Function LeerDatoTablaAvancePersonaje(ByVal PunteroPosicionVecinaPersonajeHL As Long, ByVal NBits As Long) As Long
    'busca en la tabla 284D � 286D, dependiendo del valor de HL, un valor con signo de 8 � 16 bits
    
    '; tabla para el c�lculo del avance de los personajes seg�n la orientaci�n (para personajes que ocupan 4 tiles)
    '; bytes 0-1: desplazamiento en el bucle interior del buffer de tiles
    '; bytes 2-3: desplazamiento en el bucle exterior del buffer de tiles
    '; bytes 4-5: desplazamiento inicial en el buffer de alturas para el bucle
    ': byte 6: valor a sumar a la posici�n x del personaje si avanza en este sentido
    ': byte 7: valor a sumar a la posici�n y del personaje si avanza en este sentido
    '284D:   0018 FFFF FFD1 01 00 -> +24 -1  -47 [+1 00]
    '        0001 0018 FFCE 00 FF -> +1  +24 -50 [00 -1]
    '        FFE8 0001 0016 FF 00 -> -24 +1  +22 [-1 00]
    '        FFFF FFE8 0019 00 01 -> -1  -24 +25 [00 +1]
    
    '; tabla para el c�lculo del avance de los personajes seg�n la orientaci�n (para personajes que ocupan 1 tile)
    '286D:   0018 FFFF FFEA 01 00 -> +24  -1 -22 [+1 00]
    '        0001 0018 FFCF 00 FF -> +1  +24 -49 [00 -1]
    '        FFE8 0001 0016 FF 00 -> -24  +1 +22 [-1 00]
    '        FFFF FFE8 0031 00 01 -> -1  -24 +49 [00 +1]
    If PunteroPosicionVecinaPersonajeHL < &H286D Then 'personaje ocupa 4 tiles
        If NBits = 8 Then
            LeerDatoTablaAvancePersonaje = Leer8Signo(TablaAvancePersonaje4Tiles_284D, PunteroPosicionVecinaPersonajeHL - &H284D)
        ElseIf NBits = 16 Then
            LeerDatoTablaAvancePersonaje = Leer16Signo(TablaAvancePersonaje4Tiles_284D, PunteroPosicionVecinaPersonajeHL - &H284D)
        Else
            Stop
        End If
    Else 'personaje ocupa 1 tile
        If NBits = 8 Then
            LeerDatoTablaAvancePersonaje = Leer8Signo(TablaAvancePersonaje1Tile_286D, PunteroPosicionVecinaPersonajeHL - &H286D)
        ElseIf NBits = 16 Then
            LeerDatoTablaAvancePersonaje = Leer16Signo(TablaAvancePersonaje1Tile_286D, PunteroPosicionVecinaPersonajeHL - &H286D)
        Else
            Stop
        End If
    End If
End Function

Sub AvanzarPersonaje_2954(ByVal PunteroSpriteIX As Long, ByVal PunteroCaracteristicasPersonajeIY As Long, ByVal PunteroRutinaFlipearGraficos As Long, ByVal Altura1A As Long, ByVal Altura2C As Long, ByVal PunteroTablaAvancePersonajeHL As Long)
'; rutina llamada para ver si el personaje avanza
'; en a y en c se pasa la diferencia de alturas a la posici�n a la que quiere avanzar
' en HL se pasa el puntero a la tabla de avence de personaje para actualizar la posici�n del personaje
Dim AlturaPersonajeE As Byte
Dim Tama�oOcupadoA As Byte 'tama�o ocupado por el personaje en el buffer de alturas
TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 5 - &H3036) = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 5 - &H3036) And &HEF 'pone a 0 el bit que indica si el personaje est� bajando o subiendo
'295C
AlturaPersonajeE = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 4 - &H3036) 'altura del personaje
'295F
If (TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 5 - &H3036) And &HF0) = 0 Then ' si el personaje ocupa 4 posiciones
'29b7
'; aqu� salta si el personaje ocupa 4 posiciones. Llega con:
';  Altura1A = diferencia de altura con la posicion 1 m�s cercana al personaje seg�n la orientaci�n
';  Altura2C = diferencia de altura con la posicion 2 m�s cercana al personaje seg�n la orientaci�n
    If Altura1A = 1 Or Altura1A = -1 Then
        If Altura1A = 1 Then 'si se va hacia arriba
'29c3
            'aqu� llega si se sube
            TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 4 - &H3036) = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 4 - &H3036) + 1 'incrementa la altura del personaje
            Tama�oOcupadoA = &H80& 'cambia el tama�o ocupado en el buffer de alturas de 4 a 1
        ElseIf Altura1A = -1 Then 'si se va hacia abajo
'29ca
            TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 4 - &H3036) = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 4 - &H3036) - 1 'decrementa la altura del personaje
            Tama�oOcupadoA = &H90& 'cambia el tama�o ocupado en el buffer de alturas de 4 a 1 e indica que est� bajando
        End If
'29cf
        TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 5 - &H3036) = Tama�oOcupadoA
'29d3
        'actualiza la posici�n en x y en y del personaje seg�n la orientaci�n hacia la que avanza
        AvanzarPersonaje_29E4 PunteroCaracteristicasPersonajeIY, PunteroTablaAvancePersonajeHL
        If ObtenerOrientacion_29AE(PunteroCaracteristicasPersonajeIY) <> 0 Then 'devuelve 0 si la orientaci�n del personaje es 0 o 3, en otro caso devuelve 1
            'actualiza la posici�n en x y en y del personaje seg�n la orientaci�n hacia la que avanza
            AvanzarPersonaje_29E4 PunteroCaracteristicasPersonajeIY, PunteroTablaAvancePersonajeHL
        End If
'29dd
        MovimientoRealizado_2DC1 = True 'indica que ha habido movimiento
'29E2
        'incrementa el contador de los bits 0 y 1 del byte 0, avanza la animaci�n del sprite y lo redibuja
        IncrementarContadorAnimacionSprite_2A01 PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroRutinaFlipearGraficos
        Exit Sub
'29bf
    ElseIf Altura1A <> 0 Then 'en otro caso, sale si quiere subir o bajar m�s de una unidad
'29c0
        Exit Sub
    Else
'29C1
        'si no cambia de altura, actualiza la posici�n seg�n hacia donde se avance, incrementa el contador de los bits 0 y 1 del byte 0, avanza la animaci�n del sprite y lo redibuja
        AvanzarPersonaje_29F4 PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroRutinaFlipearGraficos, Altura1A, Altura2C, PunteroTablaAvancePersonajeHL
    End If
    Exit Sub
Else
'2961
' aqu� llega si el personaje ocupa una sola posici�n
'  Altura1A = diferencia de altura con la posici�n m�s cercana al personaje seg�n la orientaci�n
'  Altura2C = diferencia de altura con la posici�n del personaje + 2 (seg�n la orientaci�n que tenga)
    If Altura2C = &H10 Then Exit Sub 'si en la posici�n del personaje + 2 (seg�n la orientaci�n que tenga) hay un personaje, sale
    If Altura2C = &H20 Then Exit Sub 'si se quiere avanzar a una posici�n donde hay un personaje, sale
'2969
    If (TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 5 - &H3036) And &H20) = 0 Then 'si el personaje no est� girado en el sentido de subir o bajar en el desnivel
'297D
' aqu� salta si el bit 5 es 0. Llega con:
'  Altura1A = diferencia de altura con la posici�n m�s cercana al personaje seg�n la orientaci�n
'  Altura2C = diferencia de altura con la posici�n del personaje + 2 (seg�n la orientaci�n que tenga)
        TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 4 - &H3036) = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 4 - &H3036) + 1 'incrementa la altura del personaje
        If Altura1A <> 1 Then 'si no se est� subiendo una unidad
'2984
            TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 4 - &H3036) = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 4 - &H3036) - 1 'deshace el incremento
            If Altura1A <> -1 Then Exit Sub 'si no se est� bajando una unidad, sale
'298a
            TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 5 - &H3036) = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 5 - &H3036) Or &H10 'indica que est� bajando
'298e
            TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 4 - &H3036) = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 4 - &H3036) - 1 'decrementa la altura del personaje
        End If
'2991
        If Altura1A <> Altura2C Then 'compara la altura de la posici�n m�s cercana al personaje con la siguiente
'2992
            'si las alturas no son iguales, avanza la posici�n
            AvanzarPersonaje_29F4 PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroRutinaFlipearGraficos, Altura1A, Altura2C, PunteroTablaAvancePersonajeHL
        Else
'2994
            'aqu� llega si avanza y las 2 posiciones siguientes tienen la misma altura
            'tan solo deja activo el bit 4, por lo que el personaje pasa de ocupar una posici�n en el buffer de alturas a ocupar 4
            TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 5 - &H3036) = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 5 - &H3036) And &H10
'299C
        'actualiza la posici�n en x y en y del personaje seg�n la orientaci�n hacia la que avanza
        AvanzarPersonaje_29E4 PunteroCaracteristicasPersonajeIY, PunteroTablaAvancePersonajeHL
        If ObtenerOrientacion_29AE(PunteroCaracteristicasPersonajeIY) = 0 Then 'devuelve 0 si la orientaci�n del personaje es 0 o 3, en otro caso devuelve 1
            'actualiza la posici�n en x y en y del personaje seg�n la orientaci�n hacia la que avanza
            AvanzarPersonaje_29E4 PunteroCaracteristicasPersonajeIY, PunteroTablaAvancePersonajeHL
        End If
        MovimientoRealizado_2DC1 = True 'indica que ha habido movimiento
        'incrementa el contador de los bits 0 y 1 del byte 0, avanza la animaci�n del sprite y lo redibuja
        IncrementarContadorAnimacionSprite_2A01 PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroRutinaFlipearGraficos
        End If
    Else
'2970
        Dim Orientacion As Long
        Dim Valor As Long
        Orientacion = ObtenerOrientacion_29AE(PunteroCaracteristicasPersonajeIY) 'devuelve 0 si la orientaci�n del personaje es 0 o 3, en otro caso devuelve 1
'2974
        'cuando va hacia la derecha o hacia abajo, al convertir la posici�n en 4, solo hay 1 de diferencia
        'en cambio, si se va a los otros sentidos al convertir la posici�n a 4 hay 2 de dif
        Valor = Altura1A
'2975
        If Orientacion <> 0 Then
'2977
            Valor = Altura2C
        End If
'2978
        If Valor <> 0 Then Exit Sub 'si no est� a ras de suelo, sale?
'297a
        'aunque en realidad se llama a 29FE, la primera parte no hace nada, as� que es lo mismo llamar a 29F4
        AvanzarPersonaje_29F4 PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroRutinaFlipearGraficos, Altura1A, Altura2C, PunteroTablaAvancePersonajeHL
    End If
End If
End Sub


Sub AvanzarPersonaje_29F4(ByVal PunteroSpriteIX As Long, ByVal PunteroCaracteristicasPersonajeIY As Long, ByVal PunteroRutinaFlipearGraficos As Long, ByVal Altura1A As Long, ByVal Altura2C As Long, ByVal PunteroTablaAvancePersonajeHL As Long)
'; actualiza la posici�n seg�n hacia donde se avance, incrementa el contador de los bits 0 y 1 del byte 0, avanza la animaci�n del sprite y lo redibuja
'; aqu� salta si las alturas de las 2 posiciones no son iguales. Llega con:
';  Altura1A = diferencia de altura con la posici�n m�s cercana al personaje seg�n la orientaci�n
';  Altura2C = diferencia de altura con la posici�n del personaje + 2 (seg�n la orientaci�n que tenga)
'   PunteroTablaAvancePersonajeHL=puntero a la tabla de avance del personaje
Dim DiferenciaAlturaA As Long
DiferenciaAlturaA = Altura1A - Altura2C + 1
'29F8
AvanzarPersonaje_29E4 PunteroCaracteristicasPersonajeIY, PunteroTablaAvancePersonajeHL
'modFunciones.GuardarArchivo "Perso0", TablaCaracteristicasPersonajes_3036
'2a01
IncrementarContadorAnimacionSprite_2A01 PunteroSpriteIX, PunteroCaracteristicasPersonajeIY, PunteroRutinaFlipearGraficos
End Sub

Sub AvanzarPersonaje_29E4(ByVal PunteroCaracteristicasPersonajeIY As Long, ByVal PunteroTablaAvancePersonajeHL As Long)
'actualiza la posici�n en x y en y del personaje seg�n la orientaci�n hacia la que avanza
Dim AvanceX As Long
Dim AvanceY As Long
AvanceX = LeerDatoTablaAvancePersonaje(PunteroTablaAvancePersonajeHL, 8)
'29e5
TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 2 - &H3036) = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 2 - &H3036) + AvanceX
'29eb
AvanceY = LeerDatoTablaAvancePersonaje(PunteroTablaAvancePersonajeHL + 1, 8)
'29EC
TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 3 - &H3036) = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 3 - &H3036) + AvanceY
End Sub

Function ObtenerOrientacion_29AE(ByVal PunteroCaracteristicasPersonajeIY As Long) As Byte
'devuelve 0 si la orientaci�n del personaje es 0 o 3, en otro caso devuelve 1
Dim Valor As Byte
Valor = TablaCaracteristicasPersonajes_3036(PunteroCaracteristicasPersonajeIY + 1 - &H3036) 'lee la orientaci�n del personaje
'29b1
Valor = Valor And &H3
If Valor = 0 Then
    ObtenerOrientacion_29AE = 0
Else
'29B4
    ObtenerOrientacion_29AE = Valor Xor &H3
End If
End Function

Sub ModificarCaracteristicasSpriteLuz_26A3()
    'modifica las caracter�sticas del sprite de la luz si puede ser usada por adso
    Dim PosicionX As Long 'posici�n x del sprite de la luz
    Dim PosicionY As Long 'posici�n y del sprite de la luz
    TablaSprites_2E17(&H2FCF - &H2E17) = &HFE 'desactiva el sprite de la luz
    If Not HabitacionOscura_156C Then Exit Sub 'si la habitaci�n est� iluminada, sale
'26ad
    'aqui llega si es una habitaci�n oscura
    If Not Depuracion.LuzEnGuillermo Then
        If TablaSprites_2E17(&H2E2B - &H2E17) = &HFE Then DibujarSprites_267B 'si el sprite de adso no es visible, evita que se redibujen los sprites y sale '###depuracion
    Else
        If TablaSprites_2E17(&H2E17 - &H2E17) = &HFE Then DibujarSprites_267B 'si el sprite de guillermo no es visible, evita que se redibujen los sprites y sale '###depuracion
    End If
'26B4
    If Not Depuracion.LuzEnGuillermo Then
        PosicionX = TablaSprites_2E17(&H2E2C - &H2E17) 'posici�n x del sprite de adso '###depuraci�n
    Else
        PosicionX = TablaSprites_2E17(&H2E17 + 1 - &H2E17) 'posici�n x del sprite de guillermo '###depuraci�n
    End If
    SpriteLuzAdsoX_4B89 = PosicionX And &H3 'posici�n x del sprite de adso dentro del tile
'26BD
    SpriteLuzAdsoX_4BB5 = 4 - SpriteLuzAdsoX_4B89 '4 - (posici�n x del sprite de adso & 0x03)
'26C4
    TablaSprites_2E17(&H2FCF + &H12 - &H2E17) = &HFE 'le da la m�xima profundidad al sprite
    TablaSprites_2E17(&H2FCF + &H13 - &H2E17) = &HFE 'le da la m�xima profundidad al sprite
'26d1
    PosicionX = (PosicionX And &HFC) - 8
    If PosicionX < 0 Then PosicionX = 0
    TablaSprites_2E17(&H2FCF + 1 - &H2E17) = Long2Byte(PosicionX) 'fija la posici�n x del sprite
    TablaSprites_2E17(&H2FCF + 3 - &H2E17) = Long2Byte(PosicionX) 'fija la posici�n anterior x del sprite
'26de
    If Not Depuracion.LuzEnGuillermo Then
        PosicionY = TablaSprites_2E17(&H2E2D - &H2E17) 'obtiene la posici�n y del sprite de adso '###depuraci�n
    Else
        PosicionY = TablaSprites_2E17(&H2E17 + 2 - &H2E17) 'obtiene la posici�n y del sprite de guillermo '###depuraci�n
    End If
    If (PosicionY And &H7) >= 4 Then 'si el desplazamiento dentro del tile en y >=4...
        SpriteLuzTipoRelleno_4B6B = &HEF 'bytes a rellenar (tile y medio)
        SpriteLuzTipoRelleno_4BD1 = &H9F 'bytes a rellenar (tile)
    Else 'si es <4, intercambia los rellenos
        SpriteLuzTipoRelleno_4B6B = &H9F 'bytes a rellenar (tile)
        SpriteLuzTipoRelleno_4BD1 = &HEF 'bytes a rellenar (tile y medio)
    End If
'26F6
    PosicionY = (PosicionY And &HF8) - &H18 'ajusta la posici�n y del sprite de adso al tile m�s cercano y la traslada
    If PosicionY < 0 Then PosicionY = 0
'26FE
    TablaSprites_2E17(&H2FCF + 2 - &H2E17) = Long2Byte(PosicionY) 'modifica la posici�n y del sprite
    TablaSprites_2E17(&H2FCF + 4 - &H2E17) = Long2Byte(PosicionY) 'modifica la posici�n anterior y del sprite
'2704
    If TablaCaracteristicasPersonajes_3036(&H304B - &H3036) <> 0 Then 'si los gr�ficos estan flipeados
        SpriteLuzFlip_4BA0 = True
    Else
        SpriteLuzFlip_4BA0 = False
    End If
End Sub
