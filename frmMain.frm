VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perceptron for Symbol Recognition"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9015
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   459
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   601
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      ImageList       =   "imlMain"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input Patterns"
      Height          =   6270
      Left            =   150
      TabIndex        =   0
      Top             =   465
      Width           =   3945
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   600
         TabIndex        =   10
         Top             =   5385
         Width           =   2760
         Begin VB.CommandButton cmdTrain 
            Caption         =   "&Train"
            Height          =   390
            Left            =   105
            TabIndex        =   12
            Top             =   195
            Width           =   1215
         End
         Begin VB.CommandButton cmdRecognize 
            Caption         =   "&Recognize"
            Height          =   390
            Left            =   1425
            TabIndex        =   11
            Top             =   195
            Width           =   1215
         End
      End
      Begin VB.ComboBox cboSymbol 
         Height          =   315
         Left            =   2715
         TabIndex        =   8
         Top             =   4560
         Width           =   960
      End
      Begin VB.OptionButton optInputPatterns 
         Caption         =   "Handwriting"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   2820
      End
      Begin VB.OptionButton optInputPatterns 
         Caption         =   "Textual Fonts"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   2820
      End
      Begin VB.PictureBox picAdjustedSymbol 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1560
         Left            =   2100
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   4
         Top             =   2955
         Width           =   1560
      End
      Begin VB.PictureBox picOriginalSymbol 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1560
         Left            =   405
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   3
         Top             =   2940
         Width           =   1560
      End
      Begin VB.ListBox lboChars 
         Height          =   2010
         Left            =   3105
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   615
         Width           =   555
      End
      Begin VB.ListBox lboFonts 
         Height          =   2010
         Left            =   405
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   615
         Width           =   2595
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   285
         TabIndex        =   13
         Top             =   4335
         Width           =   1800
         Begin VB.OptionButton optTools 
            Height          =   360
            Index           =   1
            Left            =   480
            Picture         =   "frmMain.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Eraser"
            Top             =   165
            UseMaskColor    =   -1  'True
            WhatsThisHelpID =   10295
            Width           =   360
         End
         Begin VB.OptionButton optTools 
            Height          =   360
            Index           =   0
            Left            =   120
            Picture         =   "frmMain.frx":04C1
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Pencil"
            Top             =   165
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            WhatsThisHelpID =   10298
            Width           =   360
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "&Clear"
            Height          =   360
            Left            =   840
            TabIndex        =   14
            Top             =   165
            Width           =   840
         End
      End
      Begin ComctlLib.Slider sldNoise 
         Height          =   270
         Left            =   990
         TabIndex        =   61
         Top             =   5010
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   476
         _Version        =   327682
         LargeChange     =   10
         Max             =   100
         TickFrequency   =   10
      End
      Begin VB.Label lblNoise 
         Caption         =   "0%"
         Height          =   240
         Left            =   3225
         TabIndex        =   63
         Top             =   5025
         Width           =   555
      End
      Begin VB.Label Label17 
         Caption         =   "Noise:"
         Height          =   240
         Left            =   525
         TabIndex        =   62
         Top             =   5025
         Width           =   540
      End
      Begin VB.Label Label2 
         Caption         =   "Symbol:"
         Height          =   240
         Left            =   2100
         TabIndex        =   9
         Top             =   4605
         Width           =   750
      End
   End
   Begin VB.Frame frame5 
      Caption         =   "Result"
      Height          =   2970
      Left            =   4320
      TabIndex        =   18
      Top             =   3765
      Width           =   4545
      Begin VB.Frame fraCorrection 
         BorderStyle     =   0  'None
         Height          =   750
         Left            =   60
         TabIndex        =   19
         Top             =   2055
         Visible         =   0   'False
         Width           =   4440
         Begin VB.CommandButton cmdCorrectionTrain 
            Caption         =   "Tr&ain"
            Height          =   300
            Left            =   3660
            TabIndex        =   20
            Top             =   255
            Width           =   690
         End
         Begin VB.ComboBox cboCorrectionSymbol 
            Height          =   315
            Left            =   2670
            TabIndex        =   21
            Top             =   255
            Width           =   960
         End
         Begin VB.Line Line1 
            X1              =   -15
            X2              =   4380
            Y1              =   150
            Y2              =   150
         End
         Begin VB.Line Line2 
            X1              =   15
            X2              =   4410
            Y1              =   660
            Y2              =   660
         End
         Begin VB.Label Label1 
            Caption         =   "Should be 100% match with symbol:"
            Height          =   255
            Left            =   60
            TabIndex        =   22
            Top             =   315
            Width           =   2610
         End
      End
      Begin VB.Frame fraResultButtons 
         Height          =   705
         Left            =   375
         TabIndex        =   24
         Top             =   2085
         Visible         =   0   'False
         Width           =   3705
         Begin VB.CommandButton cmdDetailResult 
            Caption         =   "Show &Detail"
            Height          =   390
            Left            =   105
            TabIndex        =   25
            Top             =   180
            Width           =   1710
         End
         Begin VB.CommandButton cmdNotCorrect 
            Caption         =   "This is &Not Correct!"
            Height          =   390
            Left            =   1890
            TabIndex        =   26
            Top             =   180
            Width           =   1710
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdResult 
         Height          =   1785
         Left            =   210
         TabIndex        =   23
         Top             =   285
         Visible         =   0   'False
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   3149
         _Version        =   393216
         Rows            =   0
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         Enabled         =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   3
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin VB.Label lblMatchPercentage 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   60
         TabIndex        =   27
         Top             =   360
         Width           =   4245
      End
      Begin VB.Label lblMatchSymbol 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         Left            =   60
         TabIndex        =   28
         Top             =   600
         Width           =   4290
      End
   End
   Begin VB.Frame fraKnowledge 
      Caption         =   "Knowledge"
      Height          =   3150
      Left            =   4305
      TabIndex        =   30
      Top             =   465
      Width           =   4545
      Begin VB.Frame Frame6 
         Height          =   705
         Left            =   1560
         TabIndex        =   59
         Top             =   2280
         Width           =   1455
         Begin VB.CommandButton cmdRetrain 
            Caption         =   "R&etrain"
            Height          =   390
            Left            =   120
            TabIndex        =   60
            Top             =   195
            Width           =   1215
         End
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Training Sets Error:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   210
         TabIndex        =   58
         Top             =   1800
         Width           =   2835
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Highest Epoch:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   210
         TabIndex        =   57
         Top             =   1320
         Width           =   2835
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of Training Sets:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   210
         TabIndex        =   56
         Top             =   855
         Width           =   2835
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of Known Symbol:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   210
         TabIndex        =   55
         Top             =   375
         Width           =   2835
      End
      Begin VB.Label lblTrainingError 
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3180
         TabIndex        =   51
         Top             =   1815
         Width           =   1080
      End
      Begin VB.Label lblHighestEpoch 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3165
         TabIndex        =   50
         Top             =   1335
         Width           =   1140
      End
      Begin VB.Label lblNTrainingSets 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3165
         TabIndex        =   49
         Top             =   870
         Width           =   1140
      End
      Begin VB.Label lblNSymbols 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3165
         TabIndex        =   48
         Top             =   390
         Width           =   1035
      End
   End
   Begin VB.Frame fraTrainingProgress 
      Caption         =   "Training Progress"
      Height          =   3150
      Left            =   4305
      TabIndex        =   17
      Top             =   465
      Width           =   4545
      Begin VB.Frame Frame4 
         Height          =   705
         Left            =   1560
         TabIndex        =   52
         Top             =   2280
         Width           =   1455
         Begin VB.CommandButton cmdStop 
            Caption         =   "&Stop"
            Height          =   390
            Left            =   120
            TabIndex        =   53
            Top             =   195
            Width           =   1215
         End
         Begin VB.CommandButton cmdFinish 
            Caption         =   "&Finish"
            Height          =   390
            Left            =   105
            TabIndex        =   54
            Top             =   195
            Width           =   1215
         End
      End
      Begin VBPerceptron.ctlProgressBar pgbPattern 
         Height          =   240
         Left            =   1035
         TabIndex        =   29
         Top             =   315
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
      Begin VBPerceptron.ctlProgressBar pgbSymbol 
         Height          =   240
         Left            =   1035
         TabIndex        =   35
         Top             =   645
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
      Begin VBPerceptron.ctlProgressBar pgbError 
         Height          =   240
         Left            =   1035
         TabIndex        =   36
         Top             =   1305
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   423
         CaptionInteger  =   0   'False
         FillColor       =   255
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
      Begin VBPerceptron.ctlProgressBar pgbScore 
         Height          =   240
         Left            =   1035
         TabIndex        =   37
         Top             =   1620
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
      Begin VBPerceptron.ctlProgressBar pgbTotalError 
         Height          =   240
         Left            =   1035
         TabIndex        =   38
         Top             =   1935
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
      Begin VB.Label lblTotalError 
         Height          =   270
         Left            =   1035
         TabIndex        =   47
         Top             =   1980
         Width           =   885
      End
      Begin VB.Label lblTECountedText 
         Caption         =   "counted"
         Height          =   315
         Left            =   3540
         TabIndex        =   46
         Top             =   1980
         Width           =   780
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Error:"
         Height          =   315
         Left            =   75
         TabIndex        =   45
         Top             =   1980
         Width           =   885
      End
      Begin VB.Label Label12 
         Caption         =   "counted"
         Height          =   315
         Left            =   3540
         TabIndex        =   44
         Top             =   1650
         Width           =   780
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Score:"
         Height          =   315
         Left            =   180
         TabIndex        =   43
         Top             =   1650
         Width           =   780
      End
      Begin VB.Label Label10 
         Caption         =   "left"
         Height          =   315
         Left            =   3540
         TabIndex        =   42
         Top             =   1335
         Width           =   780
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Error:"
         Height          =   315
         Left            =   195
         TabIndex        =   41
         Top             =   1335
         Width           =   780
      End
      Begin VB.Label lblEpoch 
         Height          =   315
         Left            =   1035
         TabIndex        =   40
         Top             =   1005
         Width           =   2490
      End
      Begin VB.Label Label7 
         Caption         =   "trained"
         Height          =   315
         Left            =   3540
         TabIndex        =   39
         Top             =   675
         Width           =   780
      End
      Begin VB.Label Label6 
         Caption         =   "retrieved"
         Height          =   315
         Left            =   3540
         TabIndex        =   34
         Top             =   345
         Width           =   780
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Epoch:"
         Height          =   315
         Left            =   180
         TabIndex        =   33
         Top             =   1005
         Width           =   780
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Symbols:"
         Height          =   315
         Left            =   180
         TabIndex        =   32
         Top             =   675
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Patterns:"
         Height          =   315
         Left            =   180
         TabIndex        =   31
         Top             =   345
         Width           =   780
      End
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   3975
      Top             =   1695
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.ptr"
      DialogTitle     =   "Save As"
      Filter          =   "Perceptron File (*.ptr)|*.ptr"
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   3975
      Top             =   1170
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.ptr"
      DialogTitle     =   "Open"
      Filter          =   "Perceptron File (*.ptr)|*.ptr"
      Flags           =   4
   End
   Begin ComctlLib.ImageList imlMain 
      Left            =   3915
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0540
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0652
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0764
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuTrain 
         Caption         =   "&Train"
      End
      Begin VB.Menu mnuRecognize 
         Caption         =   "&Recognize"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'** File Name   : frmMain.frm                                                 **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Description : VB Perceptron main form                                     **
'** Dependencies: * Microsoft Scripting Runtime (scrrun.dll)                  **
'**               * Microsoft Common Dialog Control 6.0 (SP3) (comdlg32.ocx)  **
'**               * Microsoft FlexGrid Control 6.0 (msflxgrd.ocx)             **
'**               * Microsoft Windows Common Controls 5.0 (SP2)               **
'**                 (comctl32.ocx)                                            **
'** Version    : 1.11                                                         **
'** - 1.11 (January 20, 2005):                                                **
'**     * bug fixed on font's character with zero width or height (thanks to  **
'**       j514 for the bug info)                                              **
'** - 1.10 (January 19, 2005):                                                **
'**     * add copy, paste and cut features                                    **
'** - 1.00 (January 14, 2005):                                                **
'**     * initial creation                                                    **
'** Last modified on January 20, 2005                                         **
'*******************************************************************************

Option Explicit

Private Enum menmInputPatterns
  ipFont = 0
  ipHandwriting = 1
End Enum

Private Enum menmDrawingTools
  dtPen = 0
  dtEraser = 1
End Enum

'-- Resource File Constants
Private Const mconCurPen = 101
Private Const mconCurEraser = 102

Private mblnCancelTraining As Boolean              'true when the user press the
                                                   '      cancel training button
Private mblnCorrectionTrain As Boolean     'true while retraining the perceptron
                                           '        due to the recognition error
Private mblnRetrain As Boolean             'true while retraining the perceptron
                                           '             due to the user command
Private mstrFileName As String                'the source file of the perceptron
Private mudeDrawingTool As menmDrawingTools           'the selected drawing tool

Private WithEvents mPerceptron As clsPerceptron                  'the perceptron
Attribute mPerceptron.VB_VarHelpID = -1

'---------------
'  Form Events
'---------------

' Purpose    : Erases the content of the adjusted and original symbol picture
'              box
Private Sub cmdClear_Click()
  picAdjustedSymbol.Picture = LoadPicture()
  picOriginalSymbol.Picture = LoadPicture()
End Sub

' Purpose    : Retrains the perceptron due to the recognition error
Private Sub cmdCorrectionTrain_Click()
  mblnCorrectionTrain = True
  cmdTrain_Click
  mblnCorrectionTrain = False
End Sub

' Purpose    : Shows or hides the recognition result details
Private Sub cmdDetailResult_Click()
  Select Case cmdDetailResult.Caption
    Case "Hide &Detail"
      grdResult.Visible = False
      cmdDetailResult.Caption = "Show &Detail"
    Case "Show &Detail"
      grdResult.Visible = True
      cmdDetailResult.Caption = "Hide &Detail"
  End Select
End Sub

' Purpose    : Shows the perceptron knowledge information
Private Sub cmdFinish_Click()
  fraKnowledge.ZOrder
End Sub

' Purpose    : Shows the options to retrain the perceptron due to the
'              recognition error
Private Sub cmdNotCorrect_Click()
  fraResultButtons.Visible = False
  fraCorrection.Visible = True
  cboCorrectionSymbol.SetFocus
End Sub

' Purpose    : Recognizes the pattern on specified picture box
' Assumptions: * Picture box picOriginalSymbol already contains the pattern to
'                be recognized
'              * The perceptron has been initialized
' Effect     : The recognition result has been shown to the user
Private Sub cmdRecognize_Click()
  Dim i As Long
  Dim lngBipolarPixels() As Long
  Dim lngIxChar As Long
  Dim lngIxFont As Long
  Dim lngPercentage() As Long
  Dim lngScore() As Long
  Dim rgbPixels() As mdlImageProc.RGBQUAD
  
  If mPerceptron.NClasses = 0 Then
    MsgBox Prompt:="The perceptron has not been trained.", _
           Buttons:=vbExclamation
    Exit Sub
  End If
  
  If optInputPatterns(0).Value = True Then
    '-- The pattern is taken from the installed font list
    For lngIxChar = 0 To lboChars.ListCount - 1
      lboChars.Selected(lngIxChar) = (lngIxChar = lboChars.ListIndex)
    Next
    For lngIxFont = 0 To lboFonts.ListCount - 1
      lboFonts.Selected(lngIxFont) = (lngIxFont = lboFonts.ListIndex)
    Next
  End If
  
  '-- Autofit and add noise to the pattern in picOriginalSymbol and place the
  '   result on the picAdjustedSymbol
  mdlImageProc.GetPixels picSource:=picOriginalSymbol, _
                         lngBipolarPixels:=lngBipolarPixels
  BipolarToRGB lngBipolarPixels, rgbPixels, _
               lngWidth:=picAdjustedSymbol.ScaleWidth, _
               lngHeight:=picAdjustedSymbol.ScaleHeight
  mdlImageProc.AutoFitPixels rgbPixels:=rgbPixels, hdc:=picAdjustedSymbol.hdc, _
                             lngWidth:=picAdjustedSymbol.ScaleWidth, _
                             lngHeight:=picAdjustedSymbol.ScaleHeight
  mdlImageProc.AddNoise rgbPixels:=rgbPixels, sngNoisePercentage:=sldNoise.Value
  mdlImageProc.DrawPixels rgbPixels:=rgbPixels, hdcTarget:=picAdjustedSymbol.hdc
  picAdjustedSymbol.Refresh
  RGBToBipolar rgbPixels:=rgbPixels, lngBipolarPixels:=lngBipolarPixels
  
  '-- Recognize the pattern on picAdjustedSymbol
  mPerceptron.Recognize Pattern:=lngBipolarPixels, _
                        Percentage:=lngPercentage, Score:=lngScore
  
  '-- Show the recognition result
  With grdResult
    Do Until .Rows = 2
      .RemoveItem 1
    Loop
    For i = 1 To UBound(lngPercentage) - 1
      .AddItem ""
    Next
    For i = 1 To UBound(lngPercentage)
      .TextMatrix(i, 1) = mPerceptron.ClassName(i)
      .TextMatrix(i, 2) = CStr(lngPercentage(i)) & "%"
      .TextMatrix(i, 3) = CStr(lngScore(i))
    Next
    .Col = 2
    .Sort = flexSortGenericDescending
    For i = 1 To UBound(lngPercentage)
      .TextMatrix(i, 0) = CStr(i) & "."
    Next
  
    lblMatchPercentage = .TextMatrix(1, 2) & " match with symbol:"
    lblMatchSymbol = .TextMatrix(1, 1)
  End With
  fraCorrection.Visible = False
  fraResultButtons.Visible = True
  
  lboChars.SetFocus
End Sub

' Purpose    : Retrains the perceptron due to the user command
Private Sub cmdRetrain_Click()
  mblnRetrain = True
  cmdTrain_Click
  mblnRetrain = False
End Sub

' Purpose    : Marks the cancel training sign if the user confirms it
Private Sub cmdStop_Click()
  mblnCancelTraining = _
    (MsgBox(Prompt:="Are you sure you want to stop training the perceptron?", _
            Buttons:=vbQuestion + vbYesNo) = vbYes)
End Sub

' Purpose    : Trains the perceptron with the input patterns given by the user
' Assumption : The perceptron has been initialized
' Effect     : The training result has been shown to the user
Private Sub cmdTrain_Click()
  Dim lngBipolarPixels() As Long
  Dim lngIxChar As Long
  Dim lngIxFont As Long
  Dim lngNPatternRetrieved As Long
  Dim rgbPixels() As RGBQUAD
  
  '-- Preparing the training interface
  MousePointer = vbHourglass
  pgbTotalError.Visible = True
  lblTECountedText.Visible = True
  cmdFinish.Visible = False
  cmdStop.Visible = True
  cmdStop.Enabled = False
  pgbPattern.Value = 0
  pgbSymbol.Value = 0
  pgbError.Caption = "100%"
  pgbError.Value = 100
  pgbScore.Value = 0
  pgbTotalError.Value = 0
  fraTrainingProgress.ZOrder
  
  If mblnRetrain Then
    '-- This is retraining due to the user command, no need to retrieves the
    '   patterns from the selected fonts and symbols
    pgbPattern.Value = 100
  Else
    If optInputPatterns(ipHandwriting) Or mblnCorrectionTrain Then
      '-- The pattern is taken from a handwriting or this is retraining due to
      '   the recognition error
      '- If there is no class information of the pattern, cancel the training
      If (Not mblnCorrectionTrain And (Trim(cboSymbol) = "")) Or _
         (mblnCorrectionTrain And (Trim(cboCorrectionSymbol) = "")) Then
        fraKnowledge.ZOrder
        MousePointer = vbDefault
        If mblnCorrectionTrain Then
          cboCorrectionSymbol.SetFocus
        Else
          cboSymbol.SetFocus
        End If
        MsgBox "Please type or select the symbol name.", vbExclamation
        Exit Sub
      End If
      
      '- Autofit and add noise to the pattern on the picOriginalSymbol and place
      '  the result on the picAdjustedSymbol
      mdlImageProc.GetPixels picSource:=picOriginalSymbol, _
                             lngBipolarPixels:=lngBipolarPixels
      BipolarToRGB lngBipolarPixels, rgbPixels, _
                   lngWidth:=picAdjustedSymbol.ScaleWidth, _
                   lngHeight:=picAdjustedSymbol.ScaleHeight
      mdlImageProc.AutoFitPixels _
        rgbPixels:=rgbPixels, hdc:=picAdjustedSymbol.hdc, _
        lngWidth:=picAdjustedSymbol.ScaleWidth, _
        lngHeight:=picAdjustedSymbol.ScaleHeight
      mdlImageProc.AddNoise rgbPixels:=rgbPixels, _
                            sngNoisePercentage:=sldNoise.Value
      mdlImageProc.DrawPixels rgbPixels:=rgbPixels, _
                              hdcTarget:=picAdjustedSymbol.hdc
      picAdjustedSymbol.Refresh
      RGBToBipolar rgbPixels:=rgbPixels, lngBipolarPixels:=lngBipolarPixels
      
      '-- Add the pattern on the picAdjustedSymbol to the perceptron's training
      '   set
      mPerceptron.AddTrainingSet _
        ClassName:=IIf(mblnCorrectionTrain, _
                       cboCorrectionSymbol.Text, cboSymbol.Text), _
        InputPattern:=lngBipolarPixels
    ElseIf optInputPatterns(ipFont) Then
      '-- Retrieve the patterns from the selected fonts and symbols
      lngNPatternRetrieved = 0
      For lngIxChar = 0 To lboChars.ListCount - 1
        If lboChars.Selected(lngIxChar) Then
          For lngIxFont = 0 To lboFonts.ListCount - 1
            If lboFonts.Selected(lngIxFont) Then
              picOriginalSymbol.Picture = LoadPicture()
              picAdjustedSymbol.Picture = LoadPicture()
              
              '- Draw the selected symbol to the picOriginalSymbol
              mdlFont.GetTextPixels _
                strText:=lboChars.List(lngIxChar), _
                strFontName:=lboFonts.List(lngIxFont), lngFontSize:=170, _
                hdc:=picOriginalSymbol.hdc, rgbPixels:=rgbPixels
                
              '- Autofit and add noise to the pattern on the picOriginalSymbol
              '  and place it on the picAdjustedSymbol
              mdlImageProc.AutoFitPixels _
                rgbPixels:=rgbPixels, hdc:=picAdjustedSymbol.hdc, _
                lngWidth:=picAdjustedSymbol.ScaleWidth, _
                lngHeight:=picAdjustedSymbol.ScaleHeight
              mdlImageProc.AddNoise rgbPixels:=rgbPixels, _
                                    sngNoisePercentage:=sldNoise.Value
              RGBToBipolar rgbPixels:=rgbPixels, _
                           lngBipolarPixels:=lngBipolarPixels
              
              '- Add the pattern on the picAdjustedSymbol to the perceptron's
              '  training set
              mPerceptron.AddTrainingSet ClassName:=lboChars.List(lngIxChar), _
                                         InputPattern:=lngBipolarPixels
            
              '- Show the progress
              lngNPatternRetrieved = lngNPatternRetrieved + 1
              pgbPattern.Value = _
                CLng((lngNPatternRetrieved / _
                     (lboChars.SelCount * lboFonts.SelCount)) * 100)
              Refresh
            End If
          Next
        End If
      Next
    End If
  End If
  
  MousePointer = vbArrowHourglass
  mblnCancelTraining = False
  cmdStop.Enabled = True
  
  '-- Train the perceptron with its training set
  mPerceptron.Train
  
  '-- Show the training result
  lblTotalError.Caption = Format(mPerceptron.TrainingError, "0.00;;0") & "%"
  pgbTotalError.Visible = False
  lblTECountedText.Visible = False
    
  UpdateInterface
  mnuSave.Enabled = True
  cmdStop.Visible = False
  cmdFinish.Visible = True
  cmdRetrain.Enabled = True
  MousePointer = vbDefault
End Sub

' Purpose    : Initializes the interface
Private Sub Form_Load()
  Dim i As Long
  
  '-- Fill the list box font with textual symbols from installed fonts
  mdlFont.FillTextualFonts lboFonts
  For i = 48 To 122
    If i = 58 Then i = 65
    If i = 91 Then i = 97
    lboChars.AddItem Chr(i)
  Next
  lboChars.ListIndex = 0
  lboChars.Selected(0) = True
  lboFonts.ListIndex = 0
  lboFonts.Selected(0) = True
  lboFonts_Click
  
  '-- Initialize the drawing area for handwriting input pattern
  mudeDrawingTool = dtPen
  picOriginalSymbol.MouseIcon = LoadResPicture(mconCurPen, vbResCursor)
  picOriginalSymbol.DrawWidth = 8
  
  '-- Initialize the grid used to show the recognition result
  With grdResult
    .ColWidth(0) = 545
    .ColAlignment(0) = flexAlignCenterCenter
    For i = 1 To 3
      .ColWidth(i) = 1150
      .ColAlignment(i) = flexAlignCenterCenter
    Next

    .AddItem "No." & vbTab & "Symbol" & vbTab & "Match" & vbTab & "Score"
    .AddItem ""
    .FixedRows = 1
    
    For i = 0 To 1
      .RowHeight(i) = 300
    Next
  End With
 
  '-- Initialize the open/save dialog
  dlgSave.Flags = cdlOFNHideReadOnly Or _
                  cdlOFNOverwritePrompt Or cdlOFNPathMustExist
  dlgOpen.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
  
  
  '-- Initialize other interfaces
  mdlToolbar.MakeFlat tlbMain
  fraKnowledge.ZOrder
  mnuSave.Enabled = False
  
  '-- Create a new perceptron
  mnuNew_Click
End Sub

' Purpose    : Saves the perceptron if the user confirms to
Private Sub Form_Unload(Cancel As Integer)
  Dim lngSave As Long

  If mnuSave.Enabled Then
    lngSave = MsgBox("Do you want to save the changes?", _
                     vbYesNoCancel + vbExclamation)
    Select Case lngSave
      Case vbYes
        mnuSave_Click
        Cancel = mnuSave.Enabled
      Case vbCancel
        Cancel = True
    End Select
  End If
  Set mPerceptron = Nothing
End Sub

' Purpose    : Shows the pattern of the selected symbol of the selected font to
'              the user
Private Sub lboChars_Click()
  lboFonts_Click
End Sub

' Purpose    : Shows the pattern of the selected symbol of the selected font to
'              the user
Private Sub lboFonts_Click()
  Dim rgbPixels() As RGBQUAD

  optInputPatterns(0).Value = True
  
  '-- Draw the selected symbol to the picOriginalSymbol
  picOriginalSymbol.Picture = LoadPicture()
  mdlFont.GetTextPixels strText:=lboChars.List(lboChars.ListIndex), _
                        strFontName:=lboFonts.List(lboFonts.ListIndex), _
                        lngFontSize:=170, hdc:=picOriginalSymbol.hdc, _
                        rgbPixels:=rgbPixels
  mdlImageProc.DrawPixels rgbPixels:=rgbPixels, hdcTarget:=picOriginalSymbol.hdc
  
  '-- Autofit and add noise to the pattern on the picOriginalSymbol and place it
  '   on the picAdjustedSymbol
  picAdjustedSymbol.Picture = LoadPicture()
  mdlImageProc.AutoFitPixels rgbPixels:=rgbPixels, hdc:=picAdjustedSymbol.hdc, _
                             lngWidth:=picAdjustedSymbol.ScaleWidth, _
                             lngHeight:=picAdjustedSymbol.ScaleHeight
  mdlImageProc.AddNoise rgbPixels:=rgbPixels, sngNoisePercentage:=sldNoise.Value
  mdlImageProc.DrawPixels rgbPixels:=rgbPixels, hdcTarget:=picAdjustedSymbol.hdc
End Sub

' Purpose    : Shows the about dialog box
Private Sub mnuAbout_Click()
  frmAbout.Show Modal:=vbModal, OwnerForm:=Me
End Sub

' Purpose    : Copies the drawing on picOriginalSymbol picture box to the
'              clipboard
Private Sub mnuCopy_Click()
  Clipboard.Clear
  Clipboard.SetData picOriginalSymbol.Image
End Sub

' Purpose    : Copies the drawing on picOriginalSymbol picture box to the
'              clipboard and clear the drawing on picOriginalSymbol
Private Sub mnuCut_Click()
  mnuCopy_Click
  cmdClear_Click
End Sub

' Purpose    : Arranges the paste menu enabled property
Private Sub mnuEdit_Click()
  mnuPaste.Enabled = Clipboard.GetFormat(vbCFBitmap)
End Sub

' Purpose    : Exits from VB Perceptron
Private Sub mnuExit_Click()
  Unload Me
End Sub

' Purpose    : Creates a new perceptron and saves the old one if the user
'              confirms to
Private Sub mnuNew_Click()
  Dim lngSave As Long

  If mnuSave.Enabled = True Then
    lngSave = MsgBox("Do you want to save the changes?", _
                     vbYesNoCancel + vbExclamation)
  Else
    lngSave = vbNo
  End If
  If lngSave = vbYes Then mnuSave_Click
  If lngSave <> vbCancel Then
    Set mPerceptron = New clsPerceptron
    mPerceptron.Init NInputNeuron:=(picAdjustedSymbol.ScaleWidth * _
                                    picAdjustedSymbol.ScaleHeight)
    
    lblTotalError.Caption = "0%"
    fraKnowledge.ZOrder
    UpdateInterface
    mstrFileName = ""
    mnuSave.Enabled = False
    mnuSaveAs.Enabled = False
    cmdRetrain.Enabled = False
  End If
End Sub

' Purpose    : Opens a perceptron from a file
Private Sub mnuOpen_Click()
  Const conErrCancel = 32755
  
  Dim lngSave As Long

  On Error GoTo ErrorHandler

  If mnuSave.Enabled = True Then
    lngSave = MsgBox("Do you want to save the changes?", _
                     vbYesNoCancel + vbExclamation)
  Else
    lngSave = vbNo
  End If
  If lngSave = vbYes Then mnuSave_Click
  If lngSave <> vbCancel Then
    dlgOpen.ShowOpen
    If dlgOpen.FileName <> "" Then
      mstrFileName = dlgOpen.FileName
      With frmOpenSaveProgress
        .gudeOperation = osOpen
        .gstrFileName = mstrFileName
        .Show Modal:=vbModal, OwnerForm:=Me
        Set mPerceptron = .gobjPerceptron
      End With
      UpdateInterface
      mnuSave.Enabled = False
      mnuSaveAs.Enabled = True
      cmdRetrain.Enabled = True
    End If
  End If
  Exit Sub
  
ErrorHandler:
  If Err.Number <> conErrCancel Then
    MsgBox Prompt:=Err.Description, Buttons:=vbCritical
  End If
End Sub

' Purpose    : Intializes and shows the options form
Private Sub mnuOptions_Click()
  Const conInfinity As Long = 2147483647

  With frmOptions
    .txtLearningRate.Text = CStr(mPerceptron.LearningRate)
    .txtThreshold.Text = CStr(mPerceptron.Threshold)
    If mPerceptron.MaxEpoch = conInfinity Then
      .txtMaxEpochs.Text = ""
      .txtMinError.Text = CStr(mPerceptron.MinError)
      .optStopCondition(0).Value = True
    Else
      .txtMinError.Text = ""
      .txtMaxEpochs.Text = CStr(mPerceptron.MaxEpoch)
      .optStopCondition(1).Value = True
    End If
    .Show Modal:=vbModal, OwnerForm:=Me
    If Not .gblnCancel Then
      mPerceptron.LearningRate = CSng(.txtLearningRate.Text)
      mPerceptron.Threshold = CSng(.txtThreshold.Text)
      If .optStopCondition(0).Value = True Then
        mPerceptron.MaxEpoch = conInfinity
        mPerceptron.MinError = CSng(.txtMinError.Text)
      Else
        mPerceptron.MaxEpoch = CLng(.txtMaxEpochs.Text)
        mPerceptron.MinError = 0
      End If
      Unload frmOptions
    End If
  End With
End Sub

' Purpose    : Pastes the image on the clipboard to picOriginalSymbol picture
'              box
Private Sub mnuPaste_Click()
  picOriginalSymbol.Picture = Clipboard.GetData
  optInputPatterns(1).Value = True
End Sub

' Purpose    : Recognizes the pattern on specified picture box
' Assumptions: * Picture box picOriginalSymbol already contains the pattern to
'                be recognized
'              * The perceptron has been initialized
' Effect     : The recognition result has been shown to the user
Private Sub mnuRecognize_Click()
  cmdRecognize_Click
End Sub

' Purpose    : Saves the perceptron to a file
Private Sub mnuSave_Click()
  If mstrFileName = "" Then
    mnuSaveAs_Click
  Else
    MousePointer = vbArrowHourglass
    With frmOpenSaveProgress
      Set .gobjPerceptron = mPerceptron
      .gudeOperation = osSave
      .gstrFileName = mstrFileName
      .Show Modal:=vbModal, OwnerForm:=Me
    End With
    MousePointer = vbDefault
    mnuSave.Enabled = False
  End If
End Sub

' Purpose    : Saves the perceptron to a new file
Private Sub mnuSaveAs_Click()
  Const conErrCancel = 32755

  On Error GoTo ErrorHandler

  dlgSave.ShowSave
  If dlgSave.FileName <> "" Then
    mstrFileName = dlgSave.FileName
    mnuSave_Click
  End If
  Exit Sub
  
ErrorHandler:
  If Err.Number <> conErrCancel Then
    MsgBox Prompt:=Err.Description, Buttons:=vbCritical
  End If
End Sub

' Purpose    : Trains the perceptron with the input patterns given by the user
' Assumption : The perceptron has been initialized
' Effect     : The training result has been shown to the user
Private Sub mnuTrain_Click()
  cmdTrain_Click
End Sub

' Purpose    : Change the drawing area mouse icon based on the selected tool
Private Sub optTools_Click(Index As Integer)
  mudeDrawingTool = Index
  With picOriginalSymbol
    Select Case mudeDrawingTool
      Case dtPen
        .MouseIcon = LoadResPicture(mconCurPen, vbResCursor)
      Case dtEraser
        .MouseIcon = LoadResPicture(mconCurEraser, vbResCursor)
    End Select
  End With
  picOriginalSymbol.SetFocus
End Sub

' Purpose    : Draws a pixel on the drawing area
Private Sub picOriginalSymbol_MouseDown(Button As Integer, Shift As Integer, _
                                        X As Single, Y As Single)
  If Button = vbLeftButton Then
    optInputPatterns(1).Value = True
    Select Case mudeDrawingTool
      Case dtPen
        picOriginalSymbol.PSet (X, Y), vbBlack
      Case dtEraser
        picOriginalSymbol.PSet (X, Y), vbWhite
    End Select
  End If
End Sub

' Purpose    : Draws a pixel on the drawing area
Private Sub picOriginalSymbol_MouseMove(Button As Integer, Shift As Integer, _
                                        X As Single, Y As Single)
  If Button = vbLeftButton Then
    Select Case mudeDrawingTool
      Case dtPen
        picOriginalSymbol.PSet (X, Y), vbBlack
      Case dtEraser
        picOriginalSymbol.PSet (X, Y), vbWhite
    End Select
  End If
End Sub

' Purpose    : Draws an auto-fit and noise-added pattern to the adjusted symbol
Private Sub sldNoise_MouseMove(Button As Integer, _
                               Shift As Integer, X As Single, Y As Single)
  Dim lngBipolarPixels() As Long
  Dim rgbPixels() As RGBQUAD
  
  If Button = vbLeftButton Then
    lblNoise.Caption = CStr(sldNoise.Value) & "%"
    picAdjustedSymbol.Picture = LoadPicture()
    mdlImageProc.GetPixels picSource:=picOriginalSymbol, _
                           lngBipolarPixels:=lngBipolarPixels
    BipolarToRGB lngBipolarPixels, rgbPixels, _
                 lngWidth:=picAdjustedSymbol.ScaleWidth, _
                 lngHeight:=picAdjustedSymbol.ScaleHeight
    mdlImageProc.AutoFitPixels _
      rgbPixels:=rgbPixels, hdc:=picAdjustedSymbol.hdc, _
      lngWidth:=picAdjustedSymbol.ScaleWidth, _
      lngHeight:=picAdjustedSymbol.ScaleHeight
    mdlImageProc.AddNoise rgbPixels:=rgbPixels, _
                          sngNoisePercentage:=sldNoise.Value
    mdlImageProc.DrawPixels rgbPixels:=rgbPixels, _
                            hdcTarget:=picAdjustedSymbol.hdc
    picAdjustedSymbol.Refresh
  End If
End Sub

' Purpose    : Processes the toolbar button clicked by the user
Private Sub tlbMain_ButtonClick(ByVal Button As ComctlLib.Button)
  Select Case Button.Key
    Case "New"
      mnuNew_Click
    Case "Open"
      mnuOpen_Click
    Case "Save"
      mnuSave_Click
  End Select
End Sub

'----------------------
'  mPerceptron Events
'----------------------

' Purpose    : Updates the counting score process progress
Private Sub mPerceptron_CountingScore(NProcessed As Long, NTotal As Long)
  cmdStop.Enabled = False
  pgbScore.Value = CLng((NProcessed / NTotal) * 100)
  Refresh
End Sub

' Purpose    : Updates the counting total error process progress
Private Sub mPerceptron_CountingTotalError(NProcessed As Long, NTotal As Long)
  pgbTotalError.Value = CLng((NProcessed / NTotal) * 100)
  Refresh
End Sub

' Purpose    : Updates the counting score process progress
Private Sub mPerceptron_FinishTraining()
  MousePointer = vbHourglass
End Sub

' Purpose    : Updates the training process progress
Private Sub mPerceptron_Training(IxClass As Long, _
                                 IxEpoch As Long, NTrainingSets As Long, _
                                 NTrainingError As Long, Cancel As Boolean)
  pgbSymbol.Value = CLng((IxClass / mPerceptron.NClasses) * 100)
  lblEpoch.Caption = CStr(IxEpoch) & " processed"
  pgbError.Value = (NTrainingError / NTrainingSets) * 100
  Refresh
  
  DoEvents
  Cancel = mblnCancelTraining
End Sub

'----------------------
'  Private Procedures
'----------------------

' Purpose    : Converts a one dimensional b&w pixels in bipolar format to a two
'              dimensional b&w pixels in rgb format
' Inputs     : * lngBipolarPixels (the one dimensional b&w pixels)
'              * lngWidth, lngHeight (the size of the two dimensional b&w
'                                     pixels)
' Output     : rgbPixels (the two dimensional b&w pixels)
Private Sub BipolarToRGB(lngBipolarPixels() As Long, rgbPixels() As RGBQUAD, _
                         lngWidth As Long, lngHeight As Long)
  Dim X As Long, Y As Long
  Dim i As Long

  ReDim rgbPixels(lngWidth - 1, lngHeight - 1)
  i = 0
  
  For Y = 0 To UBound(rgbPixels, 2)
    For X = 0 To UBound(rgbPixels, 1)
      i = i + 1
      If lngBipolarPixels(i) = -1 Then
        rgbPixels(X, Y).rgbRed = 255
        rgbPixels(X, Y).rgbGreen = 255
        rgbPixels(X, Y).rgbBlue = 255
      Else
        rgbPixels(X, Y).rgbRed = 0
        rgbPixels(X, Y).rgbGreen = 0
        rgbPixels(X, Y).rgbBlue = 0
      End If
    Next
  Next
End Sub

' Purpose    : Converts a two dimensional b&w pixels in rgb format to a one
'              dimensional b&w pixels in bipolar format
' Input      : rgbPixels (the two dimensional b&w pixels)
' Output     : lngBipolarPixels (the one dimensional b&w pixels)
Private Sub RGBToBipolar(rgbPixels() As RGBQUAD, lngBipolarPixels() As Long)
  Dim X As Long, Y As Long

  ReDim lngBipolarPixels((UBound(rgbPixels, 1) + 1) * _
                         (UBound(rgbPixels, 2) + 1))
  For Y = 0 To UBound(rgbPixels, 2)
    For X = 0 To UBound(rgbPixels, 1)
      If RGB(rgbPixels(X, Y).rgbRed, _
             rgbPixels(X, Y).rgbGreen, rgbPixels(X, Y).rgbBlue) = vbWhite Then
        lngBipolarPixels((Y * (UBound(rgbPixels, 1) + 1)) + X + 1) = -1
      Else
        lngBipolarPixels((Y * (UBound(rgbPixels, 1) + 1)) + X + 1) = 1
      End If
    Next
  Next
End Sub

' Purpose    : Updates the form interface with the new data gained after
'              training the perceptron
Private Sub UpdateInterface()
  Dim i As Long
  Dim strClassNames() As String
  
  mPerceptron.GetClassNames ClassNames:=strClassNames
  With cboSymbol
    .Tag = cboSymbol.Text
    .Clear
    For i = 1 To UBound(strClassNames)
      .AddItem Item:=strClassNames(i)
    Next
    .Text = cboSymbol.Tag
  End With
  With cboCorrectionSymbol
    .Tag = cboSymbol.Text
    .Clear
    For i = 1 To UBound(strClassNames)
      .AddItem Item:=strClassNames(i)
    Next
    .Text = cboSymbol.Tag
  End With
  
  lblNSymbols.Caption = CStr(mPerceptron.NClasses)
  lblNTrainingSets.Caption = CStr(mPerceptron.NTotalTrainingSets)
  lblHighestEpoch.Caption = CStr(mPerceptron.HighestEpoch)
  lblTrainingError.Caption = Format(mPerceptron.TrainingError, "0.00;;0") & "%"

End Sub

