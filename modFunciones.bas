Attribute VB_Name = "modFunciones"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function shr(ByVal Value As Long, ByVal Shift As Byte) As Long
    Dim i As Byte
    Dim m As Long
    shr = Value
    m = shr And &H80000000
    If m <> 0 Then
        shr = shr / 2
        Shift = Shift - 1
        shr = shr And &H7FFFFFFF
    End If
    If Shift > 0 Then
        shr = Int(shr / (2 ^ Shift))
    End If
End Function

Public Function ror(ByVal Value As Long, ByVal Shift As Byte) As Long
ror = rol(Value, 32 - Shift)
End Function


Public Function shl(ByVal Value As Long, ByVal Shift As Byte) As Long
    shl = Value
    If Shift > 0 Then
        Dim i As Byte
        Dim m As Long
        For i = 1 To Shift
            m = shl And &H40000000
            shl = (shl And &H3FFFFFFF) * 2
            If m <> 0 Then
                shl = shl Or &H80000000
            End If
        Next i
    End If
End Function

Public Function rol(ByVal Value As Long, ByVal Shift As Byte) As Long
    rol = Value
    If Shift > 0 Then
        Dim i As Byte
        Dim m As Long
        Dim n As Long
        For i = 1 To Shift
            m = rol And &H40000000
            n = rol And &H80000000
            rol = (rol And &H3FFFFFFF) * 2
            If m <> 0 Then
                rol = rol Or &H80000000
            End If
            If n <> 0 Then
                rol = rol Or &H1&
            End If
        Next i
    End If
End Function

Public Function rol8(ByVal Value As Long, ByVal Shift As Byte) As Long
    rol8 = Value
    If Shift > 0 Then
        Dim i As Byte
        Dim m As Long
        Dim n As Long
        For i = 1 To Shift
            n = rol8 And &H80&
            rol8 = (rol8 And &H7F&) * 2
            If n <> 0 Then
                rol8 = rol8 Or &H1&
            End If
        Next i
    End If
End Function

Public Function ror8(ByVal Value As Long, ByVal Shift As Byte) As Long
ror8 = rol8(Value, 8 - Shift)
End Function

Public Function Leer16(Bytes() As Byte, Posicion As Long) As Long
'lee un valor de 16 bits de una cadena de bytes
Leer16 = shl(Bytes(Posicion + 1), 8) + Bytes(Posicion)
End Function

Public Function Leer16Signo(Bytes() As Byte, Posicion As Long) As Long
'lee un valor de 16 bits con signo de una cadena de bytes
Dim Valor As Long
Valor = shl(Bytes(Posicion + 1), 8) + Bytes(Posicion)
If Valor >= 32768 Then 'complemento a 2
    Leer16Signo = Valor - 65536
Else
    Leer16Signo = Valor
End If
End Function

Public Function Leer8Signo(Bytes() As Byte, Posicion As Long) As Long
'lee un valor de 16 bits con signo de una cadena de bytes
Dim Valor As Long
Valor = Bytes(Posicion)
If Valor >= 128 Then 'complemento a 2
    Leer8Signo = Valor - 256
Else
    Leer8Signo = Valor
End If
End Function

Public Sub Escribir16(Bytes() As Byte, Posicion As Long, Valor As Long)
'escribe un valor de 16 bits de una cadena de bytes
Bytes(Posicion) = Valor And &HFF&
Bytes(Posicion + 1) = shr(Valor And &HFF00&, 8)
End Sub

Function Bytes2AsciiHex(Entrada() As Byte) As String
'convierte una serie de bytes en una cadena hexadecimal
Dim Contador As Long
Dim Limite As Long
Dim Cadena As String
Dim Caracter_Hex As String
Limite = UBound(Entrada)
For Contador = 0 To Limite
    Caracter_Hex = Byte2AsciiHex(Entrada(Contador))
    Cadena = Cadena + Caracter_Hex
    If Contador <> Limite Then Cadena = Cadena + " "
Next
Bytes2AsciiHex = Cadena
End Function

Function Byte2AsciiHex(Entrada As Byte) As String
    'convierte un byte en una cadena de texto con el valor hexadecimal
    Byte2AsciiHex = Hex$(Entrada)
    If Len(Byte2AsciiHex) <> 2 Then Byte2AsciiHex = "0" + Byte2AsciiHex
End Function

Function Long2AsciiHex(Entrada As Long, NCaracteres As Long) As String
    'convierte un long en una cadena de texto con el valor hexadecimal del n�mero de caracteres indicado
    Long2AsciiHex = Hex$(Entrada)
    While Len(Long2AsciiHex) < NCaracteres
        Long2AsciiHex = "0" + Long2AsciiHex
    Wend
End Function


Function CargarArchivo(NombreArchivo As String, Archivo() As Byte) As Long
    On Error GoTo CatchError
    ReDim Archivo(FileLen(NombreArchivo) - 1)
    Open NombreArchivo For Binary As #1
    Get #1, , Archivo
    Close #1
    Exit Function

CatchError:
    CargarArchivo = 1 'error leyendo el archivo
End Function

Sub GuardarArchivo(NombreArchivo As String, Archivo() As Byte)
Open NombreArchivo For Binary As #1
Put #1, , Archivo
Close #1
End Sub

Function Long2Byte(Valor As Long) As Byte
'pasa un entero largo de 32 bits a un byte. si el valor ewst� fuera de l�mites, da un error
'un byte s�lo puede contener enteros entre 0 y 255
If Valor < -128 Or Valor > 255 Then Stop
If Valor >= 0 Then
    Long2Byte = CByte(Valor)
Else
    Long2Byte = CByte(256 + Valor)
End If
End Function

Function Byte2Long(Valor As Byte) As Long
'pasa un byte a entero largo de 32 bits
Byte2Long = CLng(Valor)
End Function

Function LeerByteLong(Valor As Long, NumeroByte As Byte) As Byte
'devuelve el byte indicado de un entero largo
'el byte menos significativo es el 0
Dim Desplazamiento As Byte
Dim Resultado As Long
If NumeroByte > 3 Then Exit Function
Select Case NumeroByte
    Case Is = 0
        Resultado = Valor And &HFF&
        Desplazamiento = 0
    Case Is = 1
        Resultado = Valor And &HFF00&
        Desplazamiento = 8
    Case Is = 2
        Resultado = Valor And &HFF0000
        Desplazamiento = 16
    Case Is = 3
        Resultado = Valor And &HFF000000
        Desplazamiento = 24
End Select
If Desplazamiento > 0 Then
    Resultado = shr(Resultado, Desplazamiento)
End If
LeerByteLong = Long2Byte(Resultado)
End Function

Function Bytes2Long(Byte0 As Byte, Byte1 As Byte) As Long
'devuelve un entero largo con los dos primeros bytes indicados
Dim Resultado As Long
Resultado = Byte2Long(Byte1)
Resultado = shl(Resultado, 8)
Resultado = Resultado Or Byte0
Bytes2Long = Resultado
End Function

Function FixPath(Path As String) As String
    'append "\" at the end of the path, if not present
    If Path = "" Then Exit Function
    FixPath = Path
    If Right$(Path, 1) <> "\" Then FixPath = FixPath + "\"
End Function

Function DirFolder(Folder As String, Optional Mask As String, Optional OnlyName As Boolean) As cStringList
    'enumerate files in folder and return stringlist
    Dim Path As String
    Dim Result As New cStringList
    Dim FileName As String
    Set DirFolder = Result
    Path = FixPath(Folder)
    If Mask <> "" Then
        FileName = Dir(Path + Mask, vbArchive)
    Else
        FileName = Dir(Path, vbArchive)
    End If
    
    While FileName <> ""
        If OnlyName Then
            Result.Append FileName
        Else
            Result.Append Path + FileName
        End If
        FileName = Dir
        DoEvents
    Wend
End Function

Public Function CompararArchivos(archivo1() As Byte, archivo2() As Byte, Optional ByRef Log As String, Optional ByVal NombreArchivo1 As String, Optional ByVal NombreArchivo2 As String) As Long
    'devuelve 0 si los archivos son iguales, 1 si hay alg�n error y -1 si son diferentes
    Dim Limite As Long
    Dim MensajeFinal As String
    Dim Contador As Long
    Dim Diferente As Boolean
    Dim Linea As String
    On Error GoTo CatchError
    If NombreArchivo1 = "" Then NombreArchivo1 = "Archivo1"
    If NombreArchivo2 = "" Then NombreArchivo2 = "Archivo2"
    Log = "Comparando archivos " + NombreArchivo1 + " y " + NombreArchivo2 + vbCrLf
    If UBound(archivo1) > UBound(archivo2) Then
        Limite = UBound(archivo2)
        MensajeFinal = NombreArchivo1 + " es mayor que " + NombreArchivo2
    ElseIf UBound(archivo2) > UBound(archivo1) Then
        Limite = UBound(archivo1)
        MensajeFinal = NombreArchivo2 + " es mayor que " + NombreArchivo1
    Else 'igual duraci�n
        Limite = UBound(archivo1)
        MensajeFinal = ""
    End If
    For Contador = 0 To Limite
        If archivo1(Contador) <> archivo2(Contador) Then
            Diferente = True
            Linea = Long2AsciiHex(Contador, 8)
            Linea = Linea + ": "
            Linea = Linea + Byte2AsciiHex(archivo1(Contador)) + " "
            Linea = Linea + Byte2AsciiHex(archivo2(Contador)) + vbCrLf
            Log = Log + Linea
        End If
    Next
    If MensajeFinal <> "" Then Log = Log + MensajeFinal
    If Diferente Then
        CompararArchivos = 1
    Else
        Log = Log + "No se han encontrado diferencias" + vbCrLf
    End If
    
    
    Exit Function
    
CatchError:
        CompararArchivos = -1 'error en el acceso a archivos
End Function

Public Function CompararArchivosRuta(RutaArchivo1 As String, RutaArchivo2 As String, Optional ByRef Log As String) As Long
    'devuelve 0 si los archivos son iguales, 1 si hay alg�n error y -1 si son diferentes
    Dim archivo1() As Byte
    Dim archivo2() As Byte
    Dim NombreArchivo1 As String
    Dim NombreArchivo2 As String
    On Error GoTo CatchError
    NombreArchivo1 = Dir(RutaArchivo1)
    NombreArchivo2 = Dir(RutaArchivo2)
    If NombreArchivo1 = "" Or NombreArchivo2 = "" Then
        CompararArchivosRuta = 1 'archivo noencontrado
        Exit Function
    End If
    CargarArchivo RutaArchivo1, archivo1
    CargarArchivo RutaArchivo2, archivo2
    CompararArchivosRuta = CompararArchivos(archivo1, archivo2, Log, NombreArchivo1, NombreArchivo2)
    Exit Function
    
CatchError:
        CompararArchivosRuta = -1 'error en el acceso a archivos
    
End Function

Public Function PunteroPerteneceTabla(ByVal Puntero As Long, ByRef Tabla() As Byte, ByVal Origen As Long) As Boolean
    'devuelve true si el puntero apunta a una posici�n de la tabla indicada
    If (Puntero - Origen) >= 0 And (Puntero - Origen) <= UBound(Tabla) Then PunteroPerteneceTabla = True
End Function

