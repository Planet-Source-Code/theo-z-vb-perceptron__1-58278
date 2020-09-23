VERSION 5.00
Begin VB.Form frmOpenSaveProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saving..."
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   ControlBox      =   0   'False
   Icon            =   "frmOpenSaveProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   1605
      TabIndex        =   6
      Top             =   825
      Width           =   1215
   End
   Begin VBPerceptron.ctlProgressBar pgbWeights 
      Height          =   240
      Left            =   1230
      TabIndex        =   0
      Top             =   135
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VBPerceptron.ctlProgressBar pgbTrainingSets 
      Height          =   240
      Left            =   1230
      TabIndex        =   3
      Top             =   435
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Training Sets:"
      Height          =   225
      Left            =   135
      TabIndex        =   5
      Top             =   450
      Width           =   1035
   End
   Begin VB.Label lblOpenSaveText2 
      Caption         =   "saved"
      Height          =   225
      Left            =   3750
      TabIndex        =   4
      Top             =   450
      Width           =   540
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Weights:"
      Height          =   225
      Left            =   105
      TabIndex        =   2
      Top             =   165
      Width           =   1065
   End
   Begin VB.Label lblOpenSaveText1 
      Caption         =   "saved"
      Height          =   225
      Left            =   3750
      TabIndex        =   1
      Top             =   165
      Width           =   555
   End
End
Attribute VB_Name = "frmOpenSaveProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'** File Name   : frmOpenSaveProgress.frm                                     **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Description : A form for opening/saving process and for providing the     **
'**               progress information to the user                            **
'** Dependencies: -                                                           **
'** Last modified on January 11, 2005                                         **
'*******************************************************************************

Option Explicit

Public Enum menmOpenSaveOp
  osOpen = 0
  osSave = 1
End Enum

Public WithEvents gobjPerceptron As clsPerceptron      'the perceptron that will
Attribute gobjPerceptron.VB_VarHelpID = -1
                                                       '      be opened or saved
Public gstrFileName As String
Public gudeOperation As menmOpenSaveOp

Private mblnCancel As Boolean          'true if the user press the Cancel button
Private mblnProcessing As Boolean               'used so that the activate event
                                                '   the form triggered only once

' Purpose    : Marks the cancel sign if the user confirms it
' Note       : The cancel sign will be passed to the opening or saving event of
'              the perceptron
Private Sub cmdCancel_Click()
  Dim strConfirmation As String
  
  Select Case gudeOperation
    Case osOpen
      strConfirmation = "Are you sure you want to cancel the open operation?"
    Case osSave
      strConfirmation = "Are you sure you want to cancel the save operation?"
  End Select
  mblnCancel = _
    (MsgBox(Prompt:=strConfirmation, Buttons:=vbQuestion + vbYesNo) = vbYes)
End Sub

' Purpose    : Calls open or save method of the perceptron
Private Sub Form_Activate()
  If Not mblnProcessing Then
    Select Case gudeOperation
      Case osOpen
        Set gobjPerceptron = New clsPerceptron
        gobjPerceptron.OpenKnowledge gstrFileName
      Case osSave
        gobjPerceptron.SaveKnowledge gstrFileName
    End Select
    mblnProcessing = True
  End If
End Sub

' Purpose    : Intializes the opening or saving interface
Public Sub Form_Load()
  mblnProcessing = False
  mblnCancel = False
  Select Case gudeOperation
    Case osOpen
      Caption = "Opening..."
      lblOpenSaveText1 = "opened"
      lblOpenSaveText2 = "opened"
    Case osSave
      Caption = "Saving..."
      lblOpenSaveText1 = "saved"
      lblOpenSaveText2 = "saved"
  End Select
End Sub

' Purpose    : Unloads the form
' Note       : This is called when the user press the Cancel button and confirms
'              to cancel the opening or saving process
Private Sub gobjPerceptron_FinishOpening()
  Unload Me
End Sub

'-------------------------
'  gobjPerceptron Events
'-------------------------

' Purpose    : Unloads the form
' Note       : This is called when the user press the Cancel button and confirms
'              to cancel the opening or saving process
Private Sub gobjPerceptron_FinishSaving()
  Unload Me
End Sub

' Purpose    : Updates the opening training sets progress interface
' Inputs     : * NProcessed (number of training sets processed)
'              * NTotal (total number of training sets)
' Returns    : Cancel (setting this parameter to true triggers the FinishOpening
'                      event)
' Note       : This is called when there is a new progress on the perceptron
'              opening training sets data from a file
Private Sub gobjPerceptron_OpeningTrainingSets( _
              NProcessed As Long, NTotal As Long, Cancel As Boolean _
            )
  pgbTrainingSets.Value = CLng((NProcessed / NTotal) * 100)
  Refresh
  
  DoEvents
  Cancel = mblnCancel
End Sub

' Purpose    : Updates the opening weights progress interface
' Inputs     : * NProcessed (number of class' weights processed)
'              * NTotal (total number of classes)
' Returns    : Cancel (setting this parameter to true triggers the FinishOpening
'                      event)
' Note       : This is called when there is a new progress on the perceptron
'              opening weights data from a file
Private Sub gobjPerceptron_OpeningWeights(NProcessed As Long, _
                                          NTotal As Long, Cancel As Boolean)
  pgbWeights.Value = CLng((NProcessed / NTotal) * 100)
  Refresh
  
  DoEvents
  Cancel = mblnCancel
End Sub

' Purpose    : Updates the saving training sets progress interface
' Inputs     : * NProcessed (number of training sets processed)
'              * NTotal (total number of training sets)
' Returns    : Cancel (setting this parameter to true triggers the FinishSaving
'                      event)
' Note       : This is called when there is a new progress on the perceptron
'              saving training sets data from a file
Private Sub gobjPerceptron_SavingTrainingSets(NProcessed As Long, _
                                              NTotal As Long, Cancel As Boolean)
  pgbTrainingSets.Value = CLng((NProcessed / NTotal) * 100)
  Refresh
  
  DoEvents
  Cancel = mblnCancel
End Sub

' Purpose    : Updates the saving weights progress interface
' Inputs     : * NProcessed (number of class' weights processed)
'              * NTotal (total number of classes)
' Returns    : Cancel (setting this parameter to true triggers the FinishSaving
'                      event)
' Note       : This is called when there is a new progress on the perceptron
'              saving weights data from a file
Private Sub gobjPerceptron_SavingWeights(NProcessed As Long, _
                                         NTotal As Long, Cancel As Boolean)
  pgbWeights.Value = CLng((NProcessed / NTotal) * 100)
  Refresh
  
  DoEvents
  Cancel = mblnCancel
End Sub
