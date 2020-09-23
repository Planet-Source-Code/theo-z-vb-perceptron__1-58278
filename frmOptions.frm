VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Training Options"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5805
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDefault 
      Cancel          =   -1  'True
      Caption         =   "&Default"
      Height          =   390
      Left            =   4425
      TabIndex        =   8
      Top             =   1245
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   4425
      TabIndex        =   7
      Top             =   735
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   4425
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parameter"
      Height          =   1110
      Left            =   135
      TabIndex        =   11
      Top             =   150
      Width           =   4110
      Begin VB.TextBox txtLearningRate 
         Height          =   315
         Left            =   1365
         TabIndex        =   0
         Top             =   255
         Width           =   720
      End
      Begin VB.TextBox txtThreshold 
         Height          =   315
         Left            =   1365
         TabIndex        =   1
         Top             =   645
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Learning Rate:"
         Height          =   240
         Left            =   165
         TabIndex        =   13
         Top             =   300
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Threshold:"
         Height          =   240
         Left            =   135
         TabIndex        =   12
         Top             =   690
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stop Condition"
      Height          =   1080
      Left            =   150
      TabIndex        =   9
      Top             =   1365
      Width           =   4080
      Begin VB.OptionButton optStopCondition 
         Caption         =   "Train Until Reach"
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   255
         Width           =   1575
      End
      Begin VB.TextBox txtMinError 
         Height          =   315
         Left            =   1725
         TabIndex        =   3
         Top             =   255
         Width           =   720
      End
      Begin VB.TextBox txtMaxEpochs 
         Height          =   315
         Left            =   1995
         TabIndex        =   5
         Top             =   645
         Width           =   720
      End
      Begin VB.OptionButton optStopCondition 
         Caption         =   "Train No More Than"
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   645
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "% of Error (Or Less)"
         Height          =   225
         Left            =   2505
         TabIndex        =   14
         Top             =   300
         Width           =   1470
      End
      Begin VB.Label Label3 
         Caption         =   "epochs"
         Height          =   225
         Left            =   2790
         TabIndex        =   10
         Top             =   690
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'** File Name   : frmOptions.frm                                              **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Description : A form to provide user interface for training options input **
'** Dependencies: -                                                           **
'** Last modified on January 11, 2005                                         **
'*******************************************************************************

Option Explicit

Public gblnCancel As Boolean           'true if the user press the Cancel button

'---------------
'  Form Events
'---------------

' Purpose    : Marks the cancel sign and unloads the form
Private Sub cmdCancel_Click()
  gblnCancel = True
  Unload Me
End Sub

' Purpose    : Sets the options input interface with the default value
Private Sub cmdDefault_Click()
  txtLearningRate.Text = "1"
  txtThreshold.Text = "0"
  txtMinError.Text = ""
  txtMaxEpochs.Text = "100"
  optStopCondition(1).Value = True
End Sub

' Purpose    : Unmarks the cancel sign and hides the form
Private Sub cmdOK_Click()
  gblnCancel = False
  Visible = False
End Sub

' Purpose    : Validates the learning rate input
' Returns    : Cancel (true if the learning rate input is not valid)
Private Sub txtLearningRate_Validate(Cancel As Boolean)
  On Error GoTo ErrorHandler
  
  If Not IsNumeric(txtLearningRate) Then
    Cancel = True
  Else
    Cancel = (CSng(txtLearningRate) <= 0) Or (CSng(txtLearningRate) > 1)
  End If
  If Cancel Then
    MsgBox Prompt:="Please enter a numeric value greater " & _
                   "than 0 and less than or equal 1.", Buttons:=vbCritical
  End If
  Exit Sub

ErrorHandler:
  MsgBox Prompt:=Err.Description, Buttons:=vbCritical
End Sub

' Purpose    : Validates the maximum epochs input
' Returns    : Cancel (true if the maximum epochs input is not valid)
Private Sub txtMaxEpochs_Validate(Cancel As Boolean)
  On Error GoTo ErrorHandler
  
  If Not IsNumeric(txtMaxEpochs) Then
    Cancel = True
  ElseIf CInt(txtMaxEpochs) <> CSng(txtMaxEpochs) Then
    Cancel = True
  Else
    Cancel = (CSng(txtMaxEpochs) <= 0)
  End If
  If Cancel Then
    MsgBox Prompt:="Please enter an integer value greater than 0.", _
           Buttons:=vbExclamation
  End If
  Exit Sub

ErrorHandler:
  MsgBox Prompt:=Err.Description, Buttons:=vbCritical
End Sub

' Purpose    : Validates the error target input
' Returns    : Cancel (true if the error target input is not valid)
Private Sub txtMinError_Validate(Cancel As Boolean)
  On Error GoTo ErrorHandler
  
  If Not IsNumeric(txtMinError) Then
    Cancel = True
  Else
    Cancel = (CSng(txtMinError) < 0) Or (CSng(txtMinError) > 100)
  End If
  If Cancel Then
    MsgBox Prompt:="Please enter a numeric value between 0 and 100.", _
           Buttons:=vbExclamation
  End If
  Exit Sub

ErrorHandler:
  MsgBox Prompt:=Err.Description, Buttons:=vbCritical
End Sub

' Purpose    : Validates the threshold input
' Returns    : Cancel (true if the threshold input is not valid)
Private Sub txtThreshold_Validate(Cancel As Boolean)
  On Error GoTo ErrorHandler
  
  If Not IsNumeric(txtThreshold) Then
    Cancel = True
  Else
    Cancel = (CSng(txtThreshold) < 0)
  End If
  If Cancel Then
    MsgBox Prompt:="Please enter a numeric value greater than or equal 0.", _
           Buttons:=vbExclamation
  End If
  Exit Sub

ErrorHandler:
  MsgBox Prompt:=Err.Description, Buttons:=vbCritical
End Sub
