Attribute VB_Name = "modTeclado"
Option Explicit
Const NumeroTeclas = 10
Dim TeclasNivel(NumeroTeclas) As Boolean 'interesa su estado
Dim TeclasFlanco(NumeroTeclas) As Boolean 'interesa su pulsaci�n

Public Enum EnumTecla
    TeclaArriba = 0
    TeclaAbajo = 1
    TeclaIzquierda = 2
    TeclaDerecha = 3
    TeclaEspacio = 4
    TeclaTabulador = 5
    TeclaControl = 6
    TeclaMayusculas = 7
    TeclaEnter = 8
    TeclaSuprimir = 9
    TeclaEscape = 10
    TeclaPunto = 11
    TeclaS = 12
    TeclaN = 13
    TeclaQ = 13
    TeclaR = 14
End Enum

Public Sub Inicializar()
'borra el estado de las teclas
Dim Contador As Integer
For Contador = 0 To UBound(TeclasNivel)
    TeclasNivel(Contador) = False
    TeclasFlanco(Contador) = False
Next
End Sub

Public Sub KeyDown(Tecla As EnumTecla)
TeclasNivel(Tecla) = True
TeclasFlanco(Tecla) = True
End Sub

Public Sub KeyUp(Tecla As EnumTecla)
TeclasNivel(Tecla) = False
End Sub

Public Function TeclaPulsadaNivel(Tecla As EnumTecla) As Boolean
'devuelve true si una tecla se mantiene pulsada
TeclaPulsadaNivel = TeclasNivel(Tecla)
TeclasNivel(Tecla) = False '### depuraci�n
End Function

Public Function TeclaPulsadaFlanco(Tecla As EnumTecla) As Boolean
'devuelve true si una tecla ha sido pulsada y no se hab�a llamado todav�a a esta funci�n.
'si se vuelve a llamar, aunque la tecla siga f�sicamente pulsada, se devolver� false
TeclaPulsadaFlanco = TeclasFlanco(Tecla) 'devuelve el estado del flanco
TeclasFlanco(Tecla) = False 'y lo borra si estaba a true
End Function
