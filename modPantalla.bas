Attribute VB_Name = "modPantalla"
Option Explicit

Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixelV Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte
Private Type Bitmap
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type
Private Declare Function GetObject Lib "GDI32" Alias "GetObjectA" (ByVal hObject As Long, _
  ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "GDI32" (ByVal hBitmap As Long, ByVal dwCount As Long, _
  ByRef lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "GDI32" (ByVal hBitmap As Long, ByVal dwCount As Long, _
  ByRef lpBits As Any) As Long

Private Escala As Long 'relación pixel pantalla/pixel juego
Private pbPantalla As PictureBox
Private Colores(3) As Long 'colores del modo1
Private ColorBorde As Long

Public Sub SeleccionarPaleta(Paleta As Long)
Select Case Paleta
    Case Is = 0 ' paleta negra
        Colores(0) = 0 'negro
        Colores(1) = 0 'negro
        Colores(2) = 0 'negro
        Colores(3) = 0 'negro
        ColorBorde = 0 'negro
    Case Is = 1 'pergamino
        Colores(0) = &H7D7DF8 'rosa
        Colores(1) = 0 'negro
        Colores(2) = &H7D& 'rojo sangre
        Colores(3) = &HF8&     'rojo
        ColorBorde = &H7D& 'rojo sangre
    Case Is = 2 'día
        Colores(0) = &H7D7D00 'azul turquesa
        Colores(1) = &H7DF8F8 'amarillo
        Colores(2) = &H7DF8& 'naranja
        Colores(3) = 0 'negro
    Case Is = 3 'noche
        Colores(0) = &H7D0000  'azul oscuro
        Colores(1) = &H7D7D7D 'gris
        Colores(2) = &HF8007F  'morado
        Colores(3) = 0 'negro
        ColorBorde = 0 'negro
End Select
End Sub

Public Sub InicializarPantalla(ValorEscala As Long, ObjetoPantalla As PictureBox)
Escala = ValorEscala
Set pbPantalla = ObjetoPantalla
'pbPantalla.BackColor = 0
pbPantalla.BackColor = &H7D7D00 'provisional
SeleccionarPaleta 2
End Sub

Public Sub DibujarRectangulo(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Color As Long)
'coordenadas en pantalla
Dim X1p As Long
Dim Y1p As Long
Dim X2p As Long
Dim Y2p As Long
X1p = X1 * Escala
Y1p = Y1 * Escala
X2p = X1p + (X2 - X1) * Escala - 1
Y2p = Y1p + (Y2 - Y1) * Escala - 1
pbPantalla.Line (X1p, Y1p)-(X2p, Y2p), Color, BF
End Sub

Public Sub DibujarRectanguloCGA(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Color As Byte)
'usa los colores de la paleta
DibujarRectangulo X1, Y1, X2, Y2, Colores(Color)
End Sub

Public Sub DibujarPunto(X As Long, Y As Long, Color As Long)
Dim Xp As Long
Dim Yp As Long
Dim ContadorX As Long
Dim ContadorY As Long
Xp = X * Escala
Yp = Y * Escala
For ContadorY = 0 To Escala - 1
    For ContadorX = 0 To Escala - 1
        SetPixelV pbPantalla.hDC, Xp + ContadorX, Yp + ContadorY, Color
    Next
Next

End Sub

Public Sub DibujarPunto2(X As Long, Y As Long, Color As Long)
Dim Xp As Long
Dim Yp As Long
Dim ContadorX As Long
Dim ContadorY As Long
Xp = X * Escala
Yp = Y * Escala
pbPantalla.Line (Xp, Yp)-(Xp + Escala - 1, Yp + Escala - 1), Color, BF

End Sub

Public Sub PantallaCGA2PC(PunteroPantalla As Long, Color As Byte)
'convierte la información de cga para dibujar en PC
Dim Y As Long
Dim X As Long
Dim ColorRGB(3) As Long 'cada byte de cga contiene información de 4 píxeles
Dim Cociente As Long 'múltiplo de 8
Dim Resto As Long '0-7
Dim Contador As Long
Static Bloque As Byte
Cociente = Int((PunteroPantalla And &H7FF) / &H50)
Resto = shr(PunteroPantalla, 11) And &H7&
Y = Cociente * 8 + Resto
X = ((PunteroPantalla And &H7FF&) - Cociente * &H50) * 4 'posición del pixel más a la izquierda
'If X = 0 Then Stop
'Color = b7 b6 b5 b4 b3 b2 b1 b0
'Color Pixel1 = b7 b3
'Color Pixel2 = b6 b2
'Color Pixel3 = b5 b1
'Color Pixel4 = b4 b0
'pixel1
Resto = 0
If Color And &H80 Then Resto = 2
If Color And &H8 Then Resto = Resto + 1
ColorRGB(0) = Resto
'pixel2
Resto = 0
If Color And &H40 Then Resto = 2
If Color And &H4 Then Resto = Resto + 1
ColorRGB(1) = Resto
'pixel3
Resto = 0
If Color And &H20 Then Resto = 2
If Color And &H2 Then Resto = Resto + 1
ColorRGB(2) = Resto
'pixel4
Resto = 0
If Color And &H10 Then Resto = 2
If Color And &H1 Then Resto = Resto + 1
ColorRGB(3) = Resto
For Contador = 0 To 3
    DibujarPunto X + Contador, Y, Colores(ColorRGB(Contador))
Next
'Bloque = Bloque + 1
'If Bloque >= 64 Then
'    Bloque = 0
'    pbPantalla.Refresh
'    DoEvents
'End If
End Sub

Public Sub Refrescar()
    pbPantalla.Refresh
    DoEvents
End Sub

Public Sub DefinirModo(Modo As Long)
'modo0: 160x100, 16 colores
'modo1: 320x200, 4 colores
'###pendiente

End Sub
