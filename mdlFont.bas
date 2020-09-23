Attribute VB_Name = "mdlFont"
'*******************************************************************************
'** File Name   : mdlFont.mdl                                                 **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Description : A module to handle functions related to font needed in VB   **
'**               Perceptron                                                  **
'** Dependencies: -                                                           **
'** Last modified on January 20, 2005                                         **
'*******************************************************************************

Option Explicit

'-- Windows API constants and types declaration
Public Const LF_FACESIZE = 32

Public Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * LF_FACESIZE
End Type

Public Type NEWTEXTMETRIC
  tmHeight As Long
  tmAscent As Long
  tmDescent As Long
  tmInternalLeading As Long
  tmExternalLeading As Long
  tmAveCharWidth As Long
  tmMaxCharWidth As Long
  tmWeight As Long
  tmOverhang As Long
  tmDigitizedAspectX As Long
  tmDigitizedAspectY As Long
  tmFirstChar As Byte
  tmLastChar As Byte
  tmDefaultChar As Byte
  tmBreakChar As Byte
  tmItalic As Byte
  tmUnderlined As Byte
  tmStruckOut As Byte
  tmPitchAndFamily As Byte
  tmCharSet As Byte
  ntmFlags As Long
  ntmSizeEM As Long
  ntmCellHeight As Long
  ntmAveWidth As Long
End Type

'-- Windows API functions declaration
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Public Declare Function GetTabbedTextExtent Lib "user32" Alias "GetTabbedTextExtentA" (ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long) As Long
Public Declare Function TabbedTextOut Lib "user32" Alias "TabbedTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long, ByVal nTabOrigin As Long) As Long

' Purpose    : Retrieves a two-dimensional array of b&w pixels in RGBQuad format
'              of a text from a font with certain size
' Inputs     : * strText (the text)
'              * strFontName (the font)
'              * lngFontSize (the font size)
'              * hdc (a temporary device context used in the retrieving process)
' Returns    : rgbPixels (the two dimensional array of b&b pixels)
Public Sub GetTextPixels(strText As String, _
                         strFontName As String, lngFontSize As Long, _
                         hdc As Long, rgbPixels() As RGBQUAD)
  Dim bminfo As BITMAPINFO
  Dim dblAlias As Double
  Dim dblColor As Double
  Dim hdcBitmap As Long
  Dim hdcFont As Long
  Dim hgdiBitmap As Long
  Dim hgdiFont As Long
  Dim hTrash As Long
  Dim lfFont As LOGFONT
  Dim lngHeight As Long, lngWidth As Long
  Dim lngSize As Long
  Dim rgbPixelsTmp() As RGBQUAD
  Dim X As Long, Y As Long
  
  '-- Get the font height and width
  hdcFont = CreateCompatibleDC(hdc)
  With lfFont
    .lfHeight = -lngFontSize
    .lfFaceName = strFontName & Chr(0)
  End With
  hgdiFont = SelectObject(hdcFont, CreateFontIndirect(lfFont))
  lngSize = GetTabbedTextExtent(hdcFont, strText, Len(strText), 0, 0)
  lngHeight = lngSize \ (CLng(256) * CLng(256))
  lngWidth = lngSize Mod (CLng(256) * CLng(256))
  
  '-- Clean up
  hTrash = SelectObject(hdcFont, hgdiFont)
  DeleteObject hTrash
  DeleteDC hdcFont
  
  '-- Check for zero width or height
  If (lngWidth = 0) Or (lngHeight = 0) Then
    ' Just return empty rgbPixels
    ReDim rgbPixels(0, 0)
    rgbPixels(0, 0).rgbRed = 255
    rgbPixels(0, 0).rgbGreen = 255
    rgbPixels(0, 0).rgbBlue = 255
    Exit Sub
  End If
    
  '-- Write the text to buffer
  hdcFont = CreateCompatibleDC(hdc)
  hdcBitmap = CreateCompatibleBitmap(hdc, lngWidth, lngHeight)
  hgdiBitmap = SelectObject(hdcFont, hdcBitmap)
  hgdiFont = SelectObject(hdcFont, CreateFontIndirect(lfFont))
  SetBkColor hdcFont, vbBlack
  SetTextColor hdcFont, vbWhite
  TabbedTextOut hdcFont, 0, 0, strText, Len(strText), 0, 0, 0
  
  '-- Resize the arrays to the same size as the buffer
  ReDim rgbPixels(lngWidth - 1, lngHeight - 1)
  ReDim rgbPixelsTmp(lngWidth - 1, lngHeight - 1)
  
  '-- Store the drawn text to array brgPixelsTmp
  With bminfo.bmiHeader
    .biSize = Len(bminfo.bmiHeader)
    .biWidth = lngWidth
    .biHeight = lngHeight
    .biPlanes = 1
    .biBitCount = 32
  End With
  GetDIBits hdcFont, hdcBitmap, 0, lngHeight, rgbPixelsTmp(0, 0), bminfo, 1
  BitBlt hdcFont, 0, 0, lngWidth, lngHeight, hdc, 0, -lngHeight / 2, vbSrcCopy
  GetDIBits hdcFont, hdcBitmap, 0, lngHeight, rgbPixels(0, 0), bminfo, 1
  
  '-- Do the anti-alias and store it to array brgPixels
  For X = 0 To lngWidth - 2 Step 2
    For Y = 0 To lngHeight - 2 Step 2
      '- Count anti-alias average
      dblAlias = CLng(rgbPixelsTmp(X, Y).rgbBlue) / 255
      
      '- Process the red part
      dblColor = rgbPixels(X / 2, Y / 2).rgbRed * (1 - dblAlias)
      If dblColor > 255 Then dblColor = 255
      rgbPixels(X / 2, Y / 2).rgbRed = dblColor
      
      '- Process the green part
      dblColor = rgbPixels(X / 2, Y / 2).rgbGreen * (1 - dblAlias)
      If dblColor > 255 Then dblColor = 255
      rgbPixels(X / 2, Y / 2).rgbGreen = dblColor
    
      '- Process the blue part
      dblColor = rgbPixels(X / 2, Y / 2).rgbBlue * (1 - dblAlias)
      If dblColor > 255 Then dblColor = 255
      rgbPixels(X / 2, Y / 2).rgbBlue = dblColor
    Next
  Next
  
  '-- Adjust the pixels info in rgbPixels
  SetDIBits hdcFont, hdcBitmap, 0, lngHeight, _
            rgbPixels(0, 0), bminfo, DIB_RGB_COLORS
  BitBlt hdcFont, 0, 0, lngWidth, lngHeight, _
         hdcFont, 0, lngHeight / 2, vbSrcCopy
  lngHeight = lngHeight \ 2
  ReDim Preserve rgbPixels(lngWidth - 1, lngHeight - 1)
  bminfo.bmiHeader.biHeight = -lngHeight
  GetDIBits hdcFont, hdcBitmap, 0, lngHeight, _
            rgbPixels(0, 0), bminfo, DIB_RGB_COLORS
  
  '-- Clean up
  hTrash = SelectObject(hdcFont, hgdiFont)
  DeleteObject hTrash
  hTrash = SelectObject(hdcFont, hgdiBitmap)
  DeleteObject hTrash
  DeleteDC hdcFont
End Sub

' Purpose    : Implements the callback function of EnumFontFamilies API function
'              to fills a list box with the names of all installed fonts
' Inputs     : * lpNLF (a pointer to an ENUMLOGFONT structure that contains
'                       information about the logical attributes of the font)
'              * lpNTM (a pointer to a NEWTEXTMETRIC structure that contains
'                       information about the physical attributes of the font,
'                       if the font is a TrueType font)
'              * FontType (the type of the font)
'              * lParam (a pointer to the list box passed by the
'                        EnumFontFamilies function)
Private Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
                                 ByVal FontType As Long, _
                                 lParam As ListBox) As Long
  Dim strFaceName As String                                       'the font name
  
  strFaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
  If lpNTM.tmCharSet <= 0 Then
    If Len(strFaceName) > 0 Then
      lParam.AddItem Left(strFaceName, InStr(strFaceName, vbNullChar) - 1)
    End If
  End If
  EnumFontFamProc = 1                                  'continue the enumeration
End Function

' Purpose    : Fills a list box with all installed fonts
' Input      : lbo (the list box)
Public Sub FillTextualFonts(lbo As ListBox)
  Dim hdc As Long
    
  lbo.Clear
  hdc = GetDC(lbo.hwnd)
  EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, lbo
  ReleaseDC lbo.hwnd, hdc
End Sub

