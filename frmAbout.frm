VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About VB Perceptron"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -75
      TabIndex        =   4
      Top             =   765
      WhatsThisHelpID =   10385
      Width           =   5550
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3630
      TabIndex        =   3
      Top             =   960
      WhatsThisHelpID =   10379
      Width           =   1260
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   90
      Picture         =   "frmAbout.frx":000C
      Top             =   165
      Width           =   540
   End
   Begin VB.Label label2 
      Caption         =   "Copyright (C) 2005 Theo Zacharias"
      Height          =   225
      Left            =   735
      TabIndex        =   2
      Top             =   495
      WhatsThisHelpID =   10383
      Width           =   2565
   End
   Begin VB.Label Label1 
      Caption         =   "VB Perceptron"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   135
      WhatsThisHelpID =   10382
      Width           =   3630
   End
   Begin VB.Label lblEMail 
      Caption         =   "(theo_yz@yahoo.com)"
      Height          =   255
      Left            =   3285
      MouseIcon       =   "frmAbout.frx":044E
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   495
      Width           =   1695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'** File Name   : frmAbout.frm                                                **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Description : A form to show the copyright information                    **
'** Dependencies: -                                                           **
'** Last modified on January 11, 2005                                         **
'*******************************************************************************

Option Explicit

'-- Windows API function declaration
Private Declare Sub ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

'---------------
'  Form Events
'---------------

' Purpose    : Unload the form
Private Sub cmdOK_Click()
  Unload Me
End Sub

' Purpose    : Removes the underline style of the mail address text
Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, X As Single, Y As Single)
  lblEMail.Font.Underline = False
End Sub

' Purpose    : Open the default mail client to send an e-mail to the author
Private Sub lblEMail_Click()
  ShellExecute hwnd:=Me.hwnd, lpOperation:=vbNullString, _
               lpFile:="mailto:theo_yz@yahoo.com", lpParameters:=vbNullString, _
               lpDirectory:=vbNullString, nShowCmd:=1
End Sub

' Purpose    : Applies the underline style of the mail address text
Private Sub lblEMail_MouseMove(Button As Integer, _
                               Shift As Integer, X As Single, Y As Single)
  lblEMail.Font.Underline = True
End Sub
