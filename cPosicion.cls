VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPosicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public NumeroPantalla As Byte
Public NumeroPantallaHex As String
Public Orientacion As Byte
Public X As Byte
Public Y As Byte
Public Z As Byte
Public Escaleras As Byte

Public Sub Init(NumeroPantalla_ As Byte, Orientacion_ As Byte, X_ As Byte, Y_ As Byte, Z_ As Byte, Escaleras_ As Byte)
    NumeroPantalla = NumeroPantalla_
    Orientacion = Orientacion_
    X = X_
    Y = Y_
    Z = Z_
    Escaleras = Escaleras_
    NumeroPantallaHex = modFunciones.Byte2AsciiHex(NumeroPantalla)
End Sub
