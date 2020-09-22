Attribute VB_Name = "GroundDetection"

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long


Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type BSBITMAP
    Info As BITMAP
    Bits() As Byte
End Type

Private ImgBuffer As BSBITMAP

Public Sub LoadBitmap()
    GetObject Form1.Pic1.Image, Len(ImgBuffer.Info), ImgBuffer.Info
    ReDim ImgBuffer.Bits(0 To ImgBuffer.Info.bmWidthBytes / ImgBuffer.Info.bmWidth - 1, _
                       0 To ImgBuffer.Info.bmWidth - 1, _
                       0 To ImgBuffer.Info.bmHeight - 1) As Byte
    
    GetBitmapBits Form1.Pic1.Image, ImgBuffer.Info.bmWidthBytes * ImgBuffer.Info.bmHeight, ImgBuffer.Bits(0, 0, 0)
End Sub

Public Function GetPixelR(Xp As Single, Yp As Single) As Byte
GetPixelR = ImgBuffer.Bits(2, Xp, Yp)
End Function
Public Function GetPixelG(Xp As Single, Yp As Single) As Byte
GetPixelG = ImgBuffer.Bits(1, Xp, Yp)
End Function
Public Function GetPixelB(Xp As Single, Yp As Single) As Byte
GetPixelB = ImgBuffer.Bits(0, Xp, Yp)
End Function

Public Function GroundCol(Xp, Yp) As Boolean
On Error Resume Next
If ImgBuffer.Bits(2, Xp, Yp) < 125 Then
GroundCol = True
Else
GroundCol = False
End If
End Function

