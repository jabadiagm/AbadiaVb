VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDepuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Enum EnumTipoLuz
    EnumTipoLuz_Normal
    EnumTipoLuz_ON
    EnumTipoLuz_Off
End Enum


Public PersonajesAdso As Boolean
Public PersonajesMalaquias As Boolean
Public PersonajesAbad As Boolean
Public PersonajesBerengario As Boolean 'berengario/bernardo gui/encapuchado/jorge
Public PersonajesSeverino As Boolean 'severino/jorge
Public LuzEnGuillermo As Boolean
Public Lampara As Boolean 'lámpara siempre disponible
Public DeshabilitarCalculoDimensionesAmpliadas As Boolean 'true para evitar el uso de la función CalcularDimensionesAmpliadasSprite_4CBF
Public Luz As EnumTipoLuz
Public QuitarRetardos As Boolean
Public SaltarPergamino As Boolean


Public Sub Init()
    Luz = EnumTipoLuz_ON
    LuzEnGuillermo = False
    Lampara = False
    PersonajesAdso = True
    PersonajesMalaquias = True
    PersonajesAbad = True
    PersonajesBerengario = True
    PersonajesSeverino = True
    DeshabilitarCalculoDimensionesAmpliadas = True
    QuitarRetardos = False
    SaltarPergamino = True
End Sub
