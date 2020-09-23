VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPerceptron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** File Name   : clsPerceptron.cls                                           **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Description : A class module to represent the perceptron object           **
'** Dependencies: Microsoft Scripting Runtime (scrrun.dll)                    **
'** Members     :                                                             **
'**   * Properties : Initialized (r/o), HighestEpoch (r/o), MaxEpoch (r/w),   **
'**                  NTrainingError (r/o), LearningRate (r/w),                **
'**                  MinError (r/w), Threshold (r/w), TrainingError (r/o)     **
'**   * Methods    : AddTrainingSet, GetClassNames, Init, Learn,              **
'**                  OpenKnowledge, Recognize, SaveKnowledge                  **
'**   * Events     : CountingScore, CountingTotalError, FinishTraining,       **
'**                  FinishOpening, FinishSaving, Training,                   **
'**                  OpeningTrainingSets, OpeningWeights, SavingTrainingSets  **
'**                  SavingWeights                                            **
'** Last modified on January 20, 2005                                         **
'*******************************************************************************

Option Explicit

'--- Private Type Declaration
Private Type mtypTrainingSet
  lngInputPattern() As Long
End Type

Private Type mtypClass
  udtTrainingSets() As mtypTrainingSet
  lngMaxScore As Long               'the lngMaxScore and lngMinScore are used to
  lngMinScore As Long               '       count the score of the input pattern
                                    '                 in the recognition process
  strName As String
End Type

'--- Property Variables
Private mblnInitialized As Long     'true if the perceptron has been initialized
Private mlngHighestEpoch As Long    'highest epoch reach in the training process
Private mlngMaxEpoch As Long              'maximum epoch of the training process
Private mlngNTrainingError As Long                'number of training sets error
Private msngLearningRate As Single
Private msngMinError As Single     'minimum error target of the training process
Private msngThreshold As Single

'--- Property Default Values
Private Const mconDefaultLearningRate As Single = 1
Private Const mconDefaultMaxEpoch As Long = 100
Private Const mconDefaultMinError As Single = 0
Private Const mconDefaultThreshold As Single = 0

'--- Other Variables/Constants
Private Const mconFileHeaderId As String = "VB Perceptron Knowledge"
                 'the value used to identify the validity of the perceptron file
Private mlngWeight() As Long                       'the weight of the perceptron
Private mudtClass() As mtypClass  'the class information known by the perceptron

'---------------------
' Class Custom Events
'---------------------

'Description: Occurs when there is a new progress on the perceptron counting the
'             score of all the training sets
'Argument   : * NProcessed (number of training sets processed)
'             * NTotal (total number of training sets)
Public Event CountingScore(NProcessed As Long, NTotal As Long)

'Description: Occurs when there is a new progress on the perceptron counting the
'             error of the training sets
'Argument   : * NProcessed (number of training sets processed)
'             * NTotal (total number of training sets)
Public Event CountingTotalError(NProcessed As Long, NTotal As Long)

'Description: Occurs when the perceptron is finished training
Public Event FinishTraining()

'Description: Occurs when the perceptron is finished opening a file in
'             OpenKnowledge action
Public Event FinishOpening()

'Description: Occurs when the perceptron is finished saving a file in
'             SaveKnowledge action
Public Event FinishSaving()

'Description: Occurs when there is a new progress on the perceptron training
'             from the training sets
'Argument   : * IxClass (current index of class)
'             * IxEpoch (current index of epoch)
'             * NTrainingSets (number of training sets)
'             * NTrainingError (number of training sets error for the current
'                               class)
'             * Cancel (setting this argument to true stops the perceptron from
'                       training)
Public Event Training(IxClass As Long, IxEpoch As Long, NTrainingSets As Long, _
                      NTrainingError As Long, Cancel As Boolean)
                      
'Description: Occurs when there is a new progress on the perceptron opening
'             training sets data from a file
'Argument   : * NProcessed (number of training sets processed)
'             * NTotal (total number of training sets)
'             * Cancel (setting this argument to true stops the perceptron from
'                       opening the file)
Public Event OpeningTrainingSets(NProcessed As Long, _
                                 NTotal As Long, Cancel As Boolean)

'Description: Occurs when there is a new progress on the perceptron opening
'             weights data from a file
'Argument   : * NProcessed (number of class' weights processed)
'             * NTotal (total number of classes)
'             * Cancel (setting this argument to true stops the perceptron from
'                       opening the file)
Public Event OpeningWeights(NProcessed As Long, _
                            NTotal As Long, Cancel As Boolean)

'Description: Occurs when there is a new progress on the perceptron saving
'             training sets data from a file
'Argument   : * NProcessed (number of training sets processed)
'             * NTotal (total number of training sets)
'             * Cancel (setting this argument to true stops the perceptron from
'                       saving the file)
Public Event SavingTrainingSets(NProcessed As Long, _
                                NTotal As Long, Cancel As Boolean)

'Description: Occurs when there is a new progress on the perceptron saving
'             weights data from a file
'Argument   : * NProcessed (number of class' weights processed)
'             * NTotal (total number of classes)
'             * Cancel (setting this argument to true stops the perceptron from
'                       saving the file)
Public Event SavingWeights(NProcessed As Long, _
                           NTotal As Long, Cancel As Boolean)
                      
'-------------------
' Class Constructor
'-------------------

Private Sub Class_Initialize()
  mblnInitialized = False
  mlngMaxEpoch = mconDefaultMaxEpoch
  msngLearningRate = mconDefaultLearningRate
  msngMinError = mconDefaultMinError
  msngThreshold = mconDefaultThreshold
End Sub

'------------------
' Class Properties
'------------------

' Purpose    : Returns a class name of given class index
' Assumptions: The perceptron has been initialized
' Input      : lngIxClass (the valid class index)
Public Property Get ClassName(IxClass As Long) As String
  ClassName = mudtClass(IxClass).strName
End Property

' Purpose    : Returns the highest epoch index of the training process
Public Property Get HighestEpoch() As Long
  HighestEpoch = mlngHighestEpoch
End Property

' Purpose    : Returns a value that determines whether the perceptron has been
'              initialized
Public Property Get Initialized() As Boolean
  Initialized = mblnInitialized
End Property

' Purpose    : Sets the learning rate of the perceptron
' Input      : sngLearningRate (the new LearningRate property value)
Public Property Let LearningRate(sngLearningRate As Single)
  msngLearningRate = sngLearningRate
End Property

' Purpose    : Returns the learning rate of the perceptron
Public Property Get LearningRate() As Single
  LearningRate = msngLearningRate
End Property

' Purpose    : Sets the maximum epoch to train the perceptron
' Input      : lngMaxEpoch (the new MaxEpoch property value)
Public Property Let MaxEpoch(lngMaxEpoch As Long)
  mlngMaxEpoch = lngMaxEpoch
End Property

' Purpose    : Returns the maximum epoch to train the perceptron
Public Property Get MaxEpoch() As Long
  MaxEpoch = mlngMaxEpoch
End Property

' Purpose    : Sets the minimum error to stop training the perceptron
' Input      : sngMinError (the new MinError property value)
Public Property Let MinError(sngMinError As Single)
  msngMinError = sngMinError
End Property

' Purpose    : Returns the minimum error to stop training the perceptron
Public Property Get MinError() As Single
  MinError = msngMinError
End Property

' Purpose    : Returns the number of classes known by the perceptron
' Assumption : The perceptron has been initialized
Public Property Get NClasses() As Long
  NClasses = UBound(mlngWeight, 2)
End Property

' Purpose    : Returns the number of perceptron's input neurons
' Assumption : The perceptron has been initialized
Public Property Get NNeurons() As Long
  NNeurons = UBound(mlngWeight, 1)
End Property

' Purpose    : Returns the number of training sets error
Public Property Get NTrainingError() As Long
  NTrainingError = mlngNTrainingError
End Property

' Purpose    : Returns the number of training sets error of class IxClass
' Assumption : The perceptron has been initialized
Public Property Get NTrainingSets(IxClass As Long) As Long
  NTrainingSets = UBound(mudtClass(IxClass).udtTrainingSets)
End Property

' Purpose    : Returns the number of training sets
' Assumption : The perceptron has been initialized
Public Property Get NTotalTrainingSets() As Long
  Dim lngIxClass As Long
  Dim lngNTotalTrainingSets As Long
  
  lngNTotalTrainingSets = 0
  For lngIxClass = 1 To NClasses
    lngNTotalTrainingSets = lngNTotalTrainingSets + NTrainingSets(lngIxClass)
  Next
  NTotalTrainingSets = lngNTotalTrainingSets
End Property

' Purpose    : Sets the threshold of the perceptron
Public Property Let Threshold(sngThreshold As Single)
  msngThreshold = sngThreshold
End Property

' Purpose    : Returns the threshold of the perceptron
Public Property Get Threshold() As Single
  Threshold = msngThreshold
End Property

' Purpose    : Returns the training error percentage of the perceptron
Public Property Get TrainingError() As Single
  If NTotalTrainingSets * NClasses = 0 Then
    TrainingError = 0
  Else
    TrainingError = (mlngNTrainingError / (NTotalTrainingSets * NClasses)) * 100
  End If
End Property

'---------------
' Class Methods
'---------------

' Purpose    : Adds a training pattern to the training set of the perceptron
' Assumption : The perceptron has been initialized
' Inputs     : * ClassName (the class name of the training set)
'              * InputPattern (the training pattern that will be added)
Public Sub AddTrainingSet(ClassName As String, InputPattern() As Long)
  Dim lngIxNeuron As Long
  Dim lngIxClass As Long
  
  '-- Get the class index of class ClassName
  lngIxClass = 1
  Do Until lngIxClass > NClasses
    If mudtClass(lngIxClass).strName = ClassName Then
      Exit Do
    Else
      lngIxClass = lngIxClass + 1
    End If
  Loop
  
  '-- If the class ClassName doesn't have an index, create a new class
  If lngIxClass > UBound(mudtClass) Then
    ReDim Preserve mlngWeight(NNeurons, lngIxClass)
    For lngIxNeuron = 0 To NNeurons
      mlngWeight(lngIxNeuron, lngIxClass) = 0
    Next
  
    ReDim Preserve mudtClass(lngIxClass)
    mudtClass(lngIxClass).strName = ClassName
    ReDim mudtClass(lngIxClass).udtTrainingSets(0)
  End If
  
  '-- Add the training set
  With mudtClass(lngIxClass)
    ReDim Preserve .udtTrainingSets(NTrainingSets(lngIxClass) + 1)
    InputPattern(0) = 1
    .udtTrainingSets(NTrainingSets(lngIxClass)).lngInputPattern = InputPattern
  End With
End Sub

' Purpose    : Gets the names of all classes known by the perceptron
' Assumption : The perceptron has been initialized
' Return     : ClassNames
Public Sub GetClassNames(ClassNames() As String)
  Dim lngIxClass As Long

  ReDim ClassNames(NClasses)
  For lngIxClass = 1 To NClasses
    ClassNames(lngIxClass) = mudtClass(lngIxClass).strName
  Next
End Sub

' Purpose    : Initializes the perceptron
' Input      : NINputNeuron (number of neurons)
' Note       : This method should be called after the perceptron is created
Public Sub Init(NInputNeuron As Long)
  If Not mblnInitialized Then
    ReDim mlngWeight(NInputNeuron, 0)
    ReDim mudtClass(0)
    
    mblnInitialized = True
  End If
End Sub

' Purpose    : Opens a perceptron from a file
' Effect     : The perceptron has been initizalied with the data from the file
' Input      : FileName (the file)
Public Sub OpenKnowledge(FileName As String)
  Dim blnCancel As Boolean
  Dim fso As Scripting.FileSystemObject
  Dim file As Scripting.TextStream
  Dim i As Long
  Dim lngIxClass As Long
  Dim lngIxNeuron As Long
  Dim lngIxTrainingSet As Long
  Dim lngNProcessed As Long
  Dim lngNTrainingSets As Long
  Dim strData() As String
  
  '-- Open the file
  Set fso = New Scripting.FileSystemObject
  Set file = fso.OpenTextFile(FileName:=FileName, _
                              IOMode:=ForReading, Create:=True)
                              
  With file
    '-- Read the header
    For i = 1 To 5: .SkipLine: Next
    strData = Split(.ReadLine, " "): mlngHighestEpoch = CLng(strData(1))
    strData = Split(.ReadLine, " "): msngLearningRate = CSng(strData(1))
    strData = Split(.ReadLine, " "): mlngMaxEpoch = CLng(strData(1))
    strData = Split(.ReadLine, " "): msngMinError = CSng(strData(1))
    strData = Split(.ReadLine, " "): ReDim mudtClass(CLng(strData(1)))
    strData = Split(.ReadLine, " "): ReDim mlngWeight(CLng(strData(1)), _
                                                      UBound(mudtClass))
    strData = Split(.ReadLine, " "): mlngNTrainingError = CLng(strData(1))
    strData = Split(.ReadLine, " "): lngNTrainingSets = CLng(strData(1))
    strData = Split(.ReadLine, " "): msngThreshold = CSng(strData(1))
    
    '-- Read the weights
    For i = 1 To 3: .SkipLine: Next
    For lngIxClass = 1 To NClasses
      strData = Split(.ReadLine, " ")
      For lngIxNeuron = 0 To NNeurons
        mlngWeight(lngIxNeuron, lngIxClass) = CLng(strData(lngIxNeuron))
      Next
      
      '- Raise opening weights progress event
      RaiseEvent OpeningWeights(lngIxClass, NClasses, blnCancel)
      '- If the user cancels opening the file, stop opening
      If blnCancel Then
        ReDim mlngWeight(0, 0)
        GoTo FinishOpening
      End If
    Next
    
    For i = 1 To 3: .SkipLine: Next
    lngNProcessed = 0
    For lngIxClass = 1 To NClasses
      With mudtClass(lngIxClass)
        '-- Read the class info
        strData = Split(file.ReadLine, " "): .strName = strData(1)
        strData = Split(file.ReadLine, " "): .lngMaxScore = CLng(strData(1))
        strData = Split(file.ReadLine, " "): .lngMinScore = CLng(strData(1))
        strData = Split(file.ReadLine, " "): ReDim _
                                             .udtTrainingSets(CLng(strData(1)))
                
        '-- Read the training set
        For lngIxTrainingSet = 1 To NTrainingSets(lngIxClass)
          strData = Split(file.ReadLine, " ")
          ReDim .udtTrainingSets(lngIxTrainingSet).lngInputPattern(NNeurons)
          For lngIxNeuron = 1 To NNeurons
            .udtTrainingSets(lngIxTrainingSet). _
              lngInputPattern(lngIxNeuron) = CLng(strData(lngIxNeuron))
          Next
          
          '- Raise opening training set progress event
          lngNProcessed = lngNProcessed + 1
          RaiseEvent OpeningTrainingSets(lngNProcessed, _
                                         lngNTrainingSets, blnCancel)
          '- If the user cancels opening the file, stop opening
          If blnCancel Then
            ReDim mlngWeight(0, 0)
            GoTo FinishOpening
          End If
        Next
      End With
      .SkipLine
    Next
  End With

FinishOpening:
  '-- Close the file
  file.Close
  Set file = Nothing
  Set fso = Nothing
  
  RaiseEvent FinishOpening
  Exit Sub
End Sub

' Purpose    : Recognizes a pattern
' Assumptions: The perceptron has been initialized and train with at leaas one
'              symbol
' Input      : Pattern
' Returns    : * Percentage (match percentages of the pattern compare with each
'                            of the known classes)
'              * Score (match score of the pattern compare with each of the
'                       known classes)
Public Sub Recognize(Pattern() As Long, Percentage() As Long, Score() As Long)
  Const conInfinity As Long = 2147483647
  
  Dim lngIxClass As Long
  Dim lngMaxScore As Long
  Dim lngMinScore As Long

  '-- Initialization
  Pattern(0) = 1                               'the bias always receives input 1
  ReDim Percentage(NClasses)
  ReDim Score(NClasses)
  
  '-- Count TotalScore and ScoreRange (used to count the match percentages)
  lngMaxScore = -conInfinity
  lngMinScore = conInfinity
  For lngIxClass = 1 To NClasses
    Score(lngIxClass) = TotalInputSignal(lngInputPattern:=Pattern, _
                                         lngIxClass:=lngIxClass)
    If Score(lngIxClass) > lngMaxScore Then lngMaxScore = Score(lngIxClass)
    If Score(lngIxClass) < lngMinScore Then lngMinScore = Score(lngIxClass)
  Next
  
  '-- Count match percentages
  For lngIxClass = 1 To NClasses
    With mudtClass(lngIxClass)
      Percentage(lngIxClass) = _
        (MatchPercentage(lngScore:=Score(lngIxClass), _
                         lngMinScore:=.lngMinScore, _
                         lngMaxScore:=.lngMaxScore) + _
         MatchPercentage(lngScore:=Score(lngIxClass), _
                         lngMinScore:=lngMinScore, _
                         lngMaxScore:=lngMaxScore)) / 2
    End With
  Next
End Sub

' Purpose    : Saves the perceptron to a file
' Effect     : If FileName exists, it is replaced; if not, the new the new file
'              is created
' Input      : FileName (the file)
Public Sub SaveKnowledge(FileName As String)
  Dim blnCancel As Boolean
  Dim fso As Scripting.FileSystemObject
  Dim file As Scripting.TextStream
  Dim lngIxClass As Long
  Dim lngIxNeuron As Long
  Dim lngIxTrainingSet As Long
  Dim lngNProcessed As Long
  Dim strData As String
  
  '-- Open the file
  Set fso = New Scripting.FileSystemObject
  Set file = fso.OpenTextFile(FileName:=FileName, _
                              IOMode:=ForWriting, Create:=True)
  
  With file
    '-- Write header
    .WriteLine mconFileHeaderId
    .WriteLine String(Len(mconFileHeaderId), "=")
    .WriteLine
    .WriteLine "Properties"
    .WriteLine "----------"
    .WriteLine "Highest_Epoch: " & CStr(mlngHighestEpoch)
    .WriteLine "Learning_Rate: " & CStr(msngLearningRate)
    .WriteLine "Max_Epoch: " & CStr(mlngMaxEpoch)
    .WriteLine "Min_Error: " & CStr(msngMinError)
    .WriteLine "N_Class: " & CStr(NClasses)
    .WriteLine "N_Neurons: " & CStr(NNeurons)
    .WriteLine "N_Training_Error: " & CStr(mlngNTrainingError)
    .WriteLine "N_Training_Sets: " & CStr(NTotalTrainingSets)
    .WriteLine "Threshold: " & CStr(msngThreshold)
    .WriteLine
    
    '-- Write the weights
    .WriteLine "Weight"
    .WriteLine "------"
    For lngIxClass = 1 To NClasses
      strData = ""
      For lngIxNeuron = 0 To NNeurons
        strData = strData & mlngWeight(lngIxNeuron, lngIxClass) & " "
      Next
      .WriteLine strData
      
      '- Raise saving weights progress event
      RaiseEvent SavingWeights(lngIxClass, NClasses, blnCancel)
      '- If the user cancels saving the file, stop saving
      If blnCancel Then
        .Close
        fso.DeleteFile FileSpec:=FileName
        GoTo FinishSaving
      End If
    Next
    .WriteLine
    
    '-- Write the training set
    .WriteLine "Training Set"
    .WriteLine "------------"
    lngNProcessed = 0
    For lngIxClass = 1 To NClasses
      .WriteLine "Class: " & mudtClass(lngIxClass).strName
      .WriteLine "Max_Score: " & CStr(mudtClass(lngIxClass).lngMaxScore)
      .WriteLine "Min_Score: " & CStr(mudtClass(lngIxClass).lngMinScore)
      .WriteLine "N_Training_Sets: " & CStr(NTrainingSets(IxClass:=lngIxClass))
      For lngIxTrainingSet = 1 To NTrainingSets(IxClass:=lngIxClass)
        strData = "Training_Set_" & CStr(lngIxTrainingSet) & ":"
        For lngIxNeuron = 1 To NNeurons
          strData = strData & _
                    " " & _
                    mudtClass(lngIxClass).udtTrainingSets(lngIxTrainingSet). _
                      lngInputPattern(lngIxNeuron)
        Next
        .WriteLine strData
        
        '- Raise saving training set progress event
        lngNProcessed = lngNProcessed + 1
        RaiseEvent _
          SavingTrainingSets(lngNProcessed, NTotalTrainingSets, blnCancel)
        '- If the user cancels saving the file, stop saving
        If blnCancel Then
          .Close
          fso.DeleteFile FileSpec:=FileName
          GoTo FinishSaving
        End If
      Next
      .WriteLine
    Next
    .Close
  End With
  
FinishSaving:
  Set file = Nothing
  Set fso = Nothing
  RaiseEvent FinishSaving
End Sub

' Purpose    : Trains the perceptron with the perceptron's training set
' Assumption : The perceptron has been initialized
' Effects    : * The perceptron's weight has been adjusted to match with the
'                perceptron's training set
'              * Variables lngMaxScore, lngMinScore and sngAverageScore of each
'                of the perceptron's classes have been updated
'              * Variables mlngHighestEpoch and mlngNTrainingError have been
'                updated
Public Sub Train()
  Dim blnCancel As Boolean
  Dim blnNoWeightsChanged As Boolean
  Dim lngIxEpoch As Long
  Dim lngIxClass1 As Long
  Dim lngIxClass2 As Long
  Dim lngIxNeuron As Long
  Dim lngIxTrainingSet As Long
  Dim lngNTrainingSets As Long
  Dim lngNTrainingError As Long
  Dim lngOutput As Long
  
  lngNTrainingSets = NTotalTrainingSets
  blnCancel = False
  mlngHighestEpoch = 0
  For lngIxClass1 = 1 To NClasses
    lngIxEpoch = 0
    Do
      lngIxEpoch = lngIxEpoch + 1
      lngNTrainingError = 0
      blnNoWeightsChanged = True
      For lngIxClass2 = 1 To NClasses
        lngOutput = IIf(lngIxClass1 = lngIxClass2, 1, -1)
        For lngIxTrainingSet = 1 To NTrainingSets(lngIxClass2)
          With mudtClass(lngIxClass2).udtTrainingSets(lngIxTrainingSet)
            If Activation(lngY_In:=TotalInputSignal( _
                                     lngInputPattern:=.lngInputPattern, _
                                     lngIxClass:=lngIxClass1)) _
               <> lngOutput Then
              ' Update the weights
              lngNTrainingError = lngNTrainingError + 1
              For lngIxNeuron = 0 To NNeurons
                mlngWeight(lngIxNeuron, lngIxClass1) = _
                  mlngWeight(lngIxNeuron, lngIxClass1) + _
                  (msngLearningRate * .lngInputPattern(lngIxNeuron) * lngOutput)
              Next
              blnNoWeightsChanged = False
            End If
          End With
        Next
      Next
      
      '- Raise training event
      RaiseEvent Training(lngIxClass1, lngIxEpoch, _
                          lngNTrainingSets, lngNTrainingError, blnCancel)
      '- If the user cancel the training process, stop training
      If blnCancel Then GoTo FinishTraining
    Loop Until blnNoWeightsChanged Or (lngIxEpoch = mlngMaxEpoch) Or _
               (((lngNTrainingError / lngNTrainingSets) * 100) <= msngMinError)
    If lngIxEpoch > mlngHighestEpoch Then mlngHighestEpoch = lngIxEpoch
  Next
  
FinishTraining:
  '- Raise finish training event
  RaiseEvent FinishTraining
  
  CountScore
  CountTrainingError
End Sub

'----------------------------------
' Private Functions and Procedures
'----------------------------------

' Purpose    : Returns the activation value of certain input
' Input      : lngY_In (the input value)
Private Function Activation(lngY_In As Long) As Long
  If lngY_In > msngThreshold Then
    Activation = 1
  ElseIf (-msngThreshold <= lngY_In) And (lngY_In >= msngThreshold) Then
    Activation = 0
  ElseIf lngY_In < -msngThreshold Then
    Activation = -1
  End If
End Function

' Purpose    : Count the perceptron training set score
' Note       : The perceptron training set score is used to count the match
'              percentage in the recognition process
Private Sub CountScore()
  Const conInfinity As Long = 2147483647
  
  Dim lngIxClass1 As Long
  Dim lngIxClass2 As Long
  Dim lngIxTrainingSet As Long
  Dim lngMaxScore As Long
  Dim lngMinScore As Long
  Dim lngNTotalStep As Long
  Dim lngNProcessedStep As Long
  Dim lngScore As Long
  Dim lngTotalScore As Long
  Dim sngAverageScore As Single
  
  lngNTotalStep = NTotalTrainingSets * NClasses
  lngNProcessedStep = 0
  For lngIxClass1 = 1 To NClasses
    With mudtClass(lngIxClass1)
      .lngMaxScore = conInfinity
      .lngMinScore = conInfinity
      For lngIxTrainingSet = 1 To NTrainingSets(lngIxClass1)
        lngMinScore = conInfinity
        lngTotalScore = 0
        For lngIxClass2 = 1 To NClasses
          lngScore = TotalInputSignal(lngInputPattern:= _
                                        .udtTrainingSets(lngIxTrainingSet). _
                                          lngInputPattern, _
                                      lngIxClass:=lngIxClass2)
          If lngIxClass2 = lngIxClass1 Then lngMaxScore = lngScore
          If lngScore < lngMinScore Then lngMinScore = lngScore
          lngTotalScore = lngTotalScore + lngScore
        
          lngNProcessedStep = lngNProcessedStep + 1
          RaiseEvent CountingScore(lngNProcessedStep, lngNTotalStep)
        Next
        sngAverageScore = lngTotalScore / NClasses
        If lngMaxScore < .lngMaxScore Then .lngMaxScore = lngMaxScore
        If lngMinScore < .lngMinScore Then .lngMinScore = lngMinScore
      Next
    End With
  Next
End Sub

' Purpose    : Count the perceptron training error
Private Sub CountTrainingError()
  Dim lngIxClass1 As Long
  Dim lngIxClass2 As Long
  Dim lngIxTrainingSet As Long
  Dim lngNProcessedStep As Long
  Dim lngNTotalStep As Long
  Dim lngScore As Long
  Dim lngY As Long
  
  lngNTotalStep = NTotalTrainingSets * NClasses
  mlngNTrainingError = 0
  lngNProcessedStep = 0
  For lngIxClass1 = 1 To NClasses
    For lngIxClass2 = 1 To NClasses
      For lngIxTrainingSet = 1 To NTrainingSets(lngIxClass2)
        lngScore = TotalInputSignal(lngInputPattern:= _
                     mudtClass(lngIxClass2). _
                       udtTrainingSets(lngIxTrainingSet).lngInputPattern, _
                     lngIxClass:=lngIxClass1)
        lngY = Activation(lngY_In:=lngScore)
        If ((lngIxClass2 = lngIxClass1) And (lngY <> 1)) Or _
           ((lngIxClass2 <> lngIxClass1) And (lngY <> -1)) Then
          mlngNTrainingError = mlngNTrainingError + 1
        End If
        
        lngNProcessedStep = lngNProcessedStep + 1
        RaiseEvent CountingTotalError(lngNProcessedStep, lngNTotalStep)
      Next
    Next
  Next
End Sub

' Purpose    : Returns the match percentage value of a specified score,
'              minimum score and maximum score
' Inputs     : * lngScore (the score)
'              * lngMinScore (the minimum score)
'              * lngMaxScore (the maximum score)
' Return     : The match percentage
Private Function MatchPercentage(lngScore As Long, _
                                 lngMinScore As Long, _
                                 lngMaxScore As Long) As Long
  Dim lngMatchPercentage As Long
  
  If lngMaxScore = lngMinScore Then
    If lngMinScore = 0 Then
      lngMatchPercentage = IIf(lngScore > 0, 100, 0)
    Else
      lngMatchPercentage = CLng((lngScore / lngMinScore) * 100)
    End If
  Else
    lngMatchPercentage = CLng(((lngScore - lngMinScore) / _
                              (lngMaxScore - lngMinScore)) * 100)
  End If
  If lngMatchPercentage > 100 Then
    lngMatchPercentage = 100
  ElseIf lngMatchPercentage < 0 Then
    lngMatchPercentage = 0
  End If
  MatchPercentage = lngMatchPercentage
End Function

' Purpose    : Returns the match score of a pattern compare with a class
' Inputs     : * lngInputPattern (the pattern)
'              * lngIxClass (the class)
Private Function TotalInputSignal(lngInputPattern() As Long, _
                                  lngIxClass As Long) As Long
  Dim lngIxNeuron As Long
  Dim lngY_In As Long

  lngY_In = 0
  For lngIxNeuron = 0 To NNeurons
    lngY_In = lngY_In + (lngInputPattern(lngIxNeuron) * _
                         mlngWeight(lngIxNeuron, lngIxClass))
  Next
  TotalInputSignal = lngY_In
End Function
