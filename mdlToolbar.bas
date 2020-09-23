Attribute VB_Name = "mdlToolbar"
'*******************************************************************************
'** File Name   : mdlToolbar.mdl                                              **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Description : A module to make a toolbar flat                             **
'** Dependencies: -                                                           **
'** Last modified on January 11, 2005                                         **
'*******************************************************************************

Option Explicit
  
'-- Windows API constants declaration
Private Const WM_USER = &H400
Private Const CCS_NODIVIDER = &H40
Private Const TB_GETSTYLE = (WM_USER + 57)
Private Const TB_SETSTYLE = (WM_USER + 56)
Private Const TBSTYLE_FLAT = &H800
Private Const TBSTYLE_TRANSPARENT = &H8000

'-- Windows API functions declaration
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendTBMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

' Purpose    : Makes a toolbar flat
' Input      : tlb (the toolbar)
Public Sub MakeFlat(tlb As Toolbar)
  Dim hToolbar As Long
  Dim strStyle As Long

  hToolbar = FindWindowEx(tlb.hwnd, 0&, "ToolbarWindow32", vbNullString)
  strStyle = SendTBMessage(hToolbar, TB_GETSTYLE, 0&, 0&)
  strStyle = strStyle Or TBSTYLE_FLAT Or TBSTYLE_TRANSPARENT Or CCS_NODIVIDER
  SendTBMessage hToolbar, TB_SETSTYLE, 0, strStyle
  tlb.Refresh
End Sub
