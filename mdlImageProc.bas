Attribute VB_Name = "mdlImageProc"
'*******************************************************************************
'** File Name   : mdlImageProc.mdl                                            **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Description : A module to handle b&w image processing needed in VB        **
'**               Perceptron                                                  **
'** Dependencies: -                                                           **
'** Last modified on January 13, 2005                                         **
'*******************************************************************************

Option Explicit

'-- Windows API constants and types declaration
Public Const DIB_RGB_COLORS = 1

Public Type BITMAPINFOHEADER
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Public Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type

Public Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type

'-- Windows API functions declaration
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

' Purpose    : Adds noises to a two dimensinoal array of b&w pixels in RGBQuad
'              format
' Inputs     : * rgbPixels (the two dimensional array of pixels)
'              * sngNoisePercentage (the noises percentage;
'                                    0 <= sngNoisePercentage <= 100)
Public Sub AddNoise(rgbPixels() As RGBQUAD, sngNoisePercentage As Single)
  Dim blnInvertPixels() As Boolean
  Dim lngNNoise As Long
  Dim lngTargetNNoise As Long
  Dim X As Long, Y As Long
  Dim lngRandomX As Long, lngRandomY As Long
  
  ReDim blnInvertPixels(UBound(rgbPixels, 1), UBound(rgbPixels, 2))
  For X = 0 To UBound(blnInvertPixels, 1)
    For Y = 0 To UBound(blnInvertPixels, 2)
      blnInvertPixels(X, Y) = False
    Next
  Next
  lngTargetNNoise = CLng((UBound(blnInvertPixels, 1) + 1) * _
                         (UBound(blnInvertPixels, 2) + 1) * _
                         (sngNoisePercentage / 100))
  
  lngNNoise = 0
  Randomize Timer
  Do Until lngNNoise = lngTargetNNoise
    Do
      lngRandomX = CLng(Rnd * UBound(blnInvertPixels, 1))
      lngRandomY = CLng(Rnd * UBound(blnInvertPixels, 2))
    Loop Until Not blnInvertPixels(lngRandomX, lngRandomY)
    blnInvertPixels(lngRandomX, lngRandomY) = _
      Not blnInvertPixels(lngRandomX, lngRandomY)
    lngNNoise = lngNNoise + 1
  Loop
  
  For X = 0 To UBound(blnInvertPixels, 1)
    For Y = 0 To UBound(blnInvertPixels, 2)
      If blnInvertPixels(X, Y) Then
        If RGB(rgbPixels(X, Y).rgbRed, _
               rgbPixels(X, Y).rgbGreen, rgbPixels(X, Y).rgbBlue) = vbWhite Then
          rgbPixels(X, Y).rgbRed = 0
          rgbPixels(X, Y).rgbGreen = 0
          rgbPixels(X, Y).rgbBlue = 0
        Else
          rgbPixels(X, Y).rgbRed = 255
          rgbPixels(X, Y).rgbGreen = 255
          rgbPixels(X, Y).rgbBlue = 255
        End If
      End If
    Next
  Next
End Sub

' Purpose    : Resizes a two dimensional array of b&w pixels in RGBQuad format
'              to a given size
' Inputs     : * rgbPixels (the two dimentional array of pixels)
'              * lngWidth, lngHeight (the given size)
Public Sub AutoFitPixels(rgbPixels() As RGBQUAD, hdc As Long, _
                         Optional lngWidth As Long = 0, _
                         Optional lngHeight As Long = 0)
  Dim bminfo As BITMAPINFO
  Dim hdcTmp As Long
  Dim hdcBitmap As Long
  Dim hgdiBitmap As Long
  Dim hTrash As Long
  Dim lngPWidth As Long, lngPHeight As Long
  Dim XMin As Long, YMin As Long, XMax As Long, YMax As Long
  Dim X As Long, Y As Long
  
  lngPWidth = UBound(rgbPixels, 1) + 1
  lngPHeight = UBound(rgbPixels, 2) + 1
  If lngWidth = 0 Then lngWidth = lngPWidth
  If lngHeight = 0 Then lngHeight = lngPHeight
  
  XMin = lngWidth
  YMin = lngHeight
  XMax = 0
  YMax = 0
  For X = 0 To lngPWidth - 1
    If X >= lngWidth Then Exit For
    For Y = 0 To lngPHeight - 1
      If Y >= lngHeight Then Exit For
      If RGB(rgbPixels(X, Y).rgbRed, _
             rgbPixels(X, Y).rgbGreen, rgbPixels(X, Y).rgbBlue) = vbBlack Then
        If X < XMin Then XMin = X
        If Y < YMin Then YMin = Y
        If X > XMax Then XMax = X
        If Y > YMax Then YMax = Y
      End If
    Next
  Next
  
  With bminfo.bmiHeader
    .biSize = Len(bminfo.bmiHeader)
    .biWidth = lngPWidth
    .biHeight = -lngPHeight
    .biPlanes = 1
    .biBitCount = 32
  End With
  hdcTmp = CreateCompatibleDC(hdc)
  hdcBitmap = CreateCompatibleBitmap(hdc, lngWidth, lngHeight)
  hgdiBitmap = SelectObject(hdcTmp, hdcBitmap)
  
  SetDIBits hdcTmp, hdcBitmap, 0, lngPHeight, _
            rgbPixels(0, 0), bminfo, DIB_RGB_COLORS
  StretchBlt hdcTmp, 0, 0, lngWidth, lngHeight, _
             hdcTmp, XMin, YMin, XMax - XMin + 1, YMax - YMin + 1, vbSrcCopy
  
  ReDim rgbPixels(lngWidth - 1, lngHeight - 1)
  bminfo.bmiHeader.biWidth = lngWidth
  bminfo.bmiHeader.biHeight = -lngHeight
  GetDIBits hdcTmp, hdcBitmap, 0, lngHeight, _
            rgbPixels(0, 0), bminfo, DIB_RGB_COLORS
  
  '-- Clean up
  hTrash = SelectObject(hdcTmp, hgdiBitmap)
  DeleteObject hTrash
  DeleteDC hdcTmp
End Sub

' Purpose    : Draws a two dimensional array of pixels in RGBQuad format to a
'              device context
' Inputs     : * rgbPixels (the two dimensional array of pixels)
'              * hdcTarget (the valid device context)
Public Sub DrawPixels(rgbPixels() As RGBQUAD, hdcTarget As Long)
  Dim bminfo As BITMAPINFO
  Dim hdcTmp As Long
  Dim hdcBitmap As Long
  Dim hgdiBitmap As Long
  Dim hTrash As Long
  Dim lngWidth As Long
  Dim lngHeight As Long
  
  lngWidth = UBound(rgbPixels, 1) + 1
  lngHeight = UBound(rgbPixels, 2) + 1
  With bminfo.bmiHeader
    .biSize = Len(bminfo.bmiHeader)
    .biWidth = lngWidth
    .biHeight = -lngHeight
    .biPlanes = 1
    .biBitCount = 32
  End With
  
  hdcTmp = CreateCompatibleDC(hdcTarget)
  hdcBitmap = CreateCompatibleBitmap(hdcTarget, lngWidth, lngHeight)
  hgdiBitmap = SelectObject(hdcTmp, hdcBitmap)
  
  SetDIBits hdcTmp, hdcBitmap, 0, lngHeight, _
            rgbPixels(0, 0), bminfo, DIB_RGB_COLORS
  BitBlt hdcTarget, 0, 0, lngWidth - 1, lngHeight - 1, hdcTmp, 0, 0, vbSrcCopy
  
  ' Clean up
  hTrash = SelectObject(hdcTmp, hgdiBitmap)
  DeleteObject hTrash
  DeleteDC hdcTmp
End Sub

' Purpose    : Retrieves a two dimensional array of b&w pixels in bipolar format
'              from a picture box
' Input      : picSource (the picture box)
' Return     : lngBipolarPixels (the two dimensional array of b&w pixels)
Public Sub GetPixels(picSource As PictureBox, lngBipolarPixels() As Long)
  Dim X As Long, Y As Long

  ReDim lngBipolarPixels(picSource.ScaleWidth * picSource.ScaleHeight)
  For Y = 0 To picSource.ScaleHeight - 1
    For X = 0 To picSource.ScaleWidth - 1
      If picSource.Point(X, Y) = vbBlack Then
        lngBipolarPixels((Y * picSource.ScaleWidth) + X + 1) = 1
      Else
        lngBipolarPixels((Y * picSource.ScaleWidth) + X + 1) = -1
      End If
    Next
  Next
End Sub
