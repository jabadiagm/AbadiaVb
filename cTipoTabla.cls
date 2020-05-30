VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTipoTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Extension As String
Public DireccionInicio As Long
Public Tamaño As Long
Public Descripcion As String
Private Archivo() As Byte

Public Sub Init(cExtension As String, cDireccionInicio As Long, cTamaño As Long, cDescripcion As String)
    Extension = cExtension
    DireccionInicio = cDireccionInicio
    Tamaño = cTamaño
    Descripcion = cDescripcion
End Sub

Public Function nose() As Byte()
    nose = Archivo
End Function
