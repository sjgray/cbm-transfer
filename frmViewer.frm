VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmViewer 
   Caption         =   "Viewer:"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11925
   Icon            =   "frmViewer.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   8100
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLA 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6900
      TabIndex        =   171
      Text            =   "0000"
      ToolTipText     =   "Load Address from File, or Entered manually"
      Top             =   30
      Width           =   495
   End
   Begin VB.CheckBox cbLA 
      Caption         =   "LA:"
      Height          =   255
      Left            =   6330
      TabIndex        =   170
      ToolTipText     =   "File includes Load Address at start"
      Top             =   60
      Value           =   1  'Checked
      Width           =   555
   End
   Begin VB.Frame frSEQ 
      Caption         =   "SEQ Viewer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   5580
      TabIndex        =   8
      Top             =   4380
      Visible         =   0   'False
      Width           =   3510
      Begin VB.CheckBox cbIgnoreLF 
         Caption         =   "&Ignore LF"
         Height          =   195
         Left            =   2280
         TabIndex        =   142
         Top             =   300
         Value           =   1  'Checked
         Width           =   1005
      End
      Begin VB.CheckBox cbSeqFont 
         Caption         =   "&Use C64 Font"
         Height          =   195
         Left            =   870
         TabIndex        =   10
         Top             =   300
         Width           =   1455
      End
      Begin VB.ListBox lstSEQ 
         BackColor       =   &H00008080&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   120
         TabIndex        =   9
         Top             =   585
         Width           =   1245
      End
      Begin VB.Label lblSEQTheme 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   450
         TabIndex        =   161
         Top             =   240
         Width           =   285
      End
      Begin VB.Label lblSEQTheme 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   158
         Top             =   240
         Width           =   285
      End
   End
   Begin VB.Frame frBasic 
      Caption         =   "BASIC Lister"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   4050
      TabIndex        =   2
      Top             =   6960
      Visible         =   0   'False
      Width           =   8070
      Begin VB.ListBox lstBAS 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8383&
         Height          =   645
         Left            =   105
         TabIndex        =   3
         Top             =   930
         Width           =   1635
      End
      Begin VB.Frame frBOpts 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   270
         TabIndex        =   127
         Top             =   180
         Width           =   7995
         Begin VB.ComboBox cboMode 
            Height          =   315
            ItemData        =   "frmViewer.frx":0442
            Left            =   930
            List            =   "frmViewer.frx":0458
            Style           =   2  'Dropdown List
            TabIndex        =   136
            Top             =   0
            Width           =   1920
         End
         Begin VB.CommandButton cmdCpyClip 
            Caption         =   "To &Clipboard"
            Height          =   315
            Left            =   5970
            TabIndex        =   135
            ToolTipText     =   "Export current view text to clipboard"
            Top             =   0
            Width           =   1215
         End
         Begin VB.CheckBox cbRev 
            Caption         =   "&Reverse Text"
            Height          =   240
            Left            =   2940
            TabIndex        =   134
            ToolTipText     =   "Reverse display of Text"
            Top             =   0
            Width           =   1425
         End
         Begin VB.CheckBox cbUseFont 
            Caption         =   "Use CBM &Font"
            Height          =   240
            Left            =   2940
            TabIndex        =   133
            ToolTipText     =   "Use special C64 Font"
            Top             =   240
            Width           =   1425
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "E&xport"
            Height          =   315
            Left            =   5955
            TabIndex        =   132
            ToolTipText     =   "Save current view text to file"
            Top             =   360
            Width           =   1230
         End
         Begin VB.CheckBox cbExp 
            Caption         =   "Expand &Special ("
            Height          =   240
            Left            =   2940
            TabIndex        =   131
            ToolTipText     =   "Expand special characters (ie {RVS} )"
            Top             =   480
            Value           =   1  'Checked
            Width           =   1530
         End
         Begin VB.CheckBox cbOneLine 
            Caption         =   "&Break Multi"
            Height          =   240
            Left            =   4470
            TabIndex        =   130
            ToolTipText     =   "Break multi-statement lines (list one statement per line)"
            Top             =   0
            Width           =   1200
         End
         Begin VB.CheckBox cbPad 
            Caption         =   "Pad &Tokens"
            Height          =   240
            Left            =   4470
            TabIndex        =   129
            ToolTipText     =   "Append SPACE to tokens"
            Top             =   225
            Width           =   1215
         End
         Begin VB.CheckBox cbUC 
            Caption         =   "UCase)"
            Height          =   240
            Left            =   4470
            TabIndex        =   128
            ToolTipText     =   "Special characters printed UpperCase"
            Top             =   480
            Value           =   1  'Checked
            Width           =   900
         End
         Begin VB.Label lblBASTheme 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   0
            TabIndex        =   160
            Top             =   390
            Width           =   285
         End
         Begin VB.Label lblBASTheme 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8383&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   159
            Top             =   90
            Width           =   285
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "BASIC:"
            Height          =   195
            Left            =   360
            TabIndex        =   140
            Top             =   60
            Width           =   510
         End
         Begin VB.Label lblGuess 
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   1605
            TabIndex        =   139
            ToolTipText     =   "Computer model"
            Top             =   405
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "LOAD:"
            Height          =   195
            Left            =   390
            TabIndex        =   138
            Top             =   435
            Width           =   480
         End
         Begin VB.Label lblLoadAdr 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            Height          =   285
            Left            =   930
            TabIndex        =   137
            Top             =   405
            Width           =   600
         End
      End
      Begin VB.Label lblBView 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   60
         TabIndex        =   126
         ToolTipText     =   "Toggle Options pane"
         Top             =   180
         Width           =   255
      End
   End
   Begin VB.Frame frBIN 
      Caption         =   "Binary Viewer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   5580
      TabIndex        =   6
      Top             =   5550
      Visible         =   0   'False
      Width           =   6870
      Begin VB.CheckBox cbShowCBM 
         Caption         =   "Show CBM"
         Height          =   195
         Left            =   4020
         TabIndex        =   157
         ToolTipText     =   "Show CBM screen codes"
         Top             =   240
         Width           =   1155
      End
      Begin VB.CheckBox cbHexSync 
         Caption         =   "Sync with ASM"
         Height          =   195
         Left            =   5190
         TabIndex        =   156
         ToolTipText     =   "File includes Load Address at start"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox cbWide 
         Caption         =   "Wide"
         Height          =   195
         Left            =   420
         TabIndex        =   125
         ToolTipText     =   "File includes Load Address at start"
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox cb7bit 
         Caption         =   "7-bit View"
         Height          =   195
         Left            =   2820
         TabIndex        =   41
         ToolTipText     =   "Enable 7-bit View"
         Top             =   240
         Width           =   1035
      End
      Begin VB.CheckBox cbShowP 
         Caption         =   "Show Printable"
         Height          =   195
         Left            =   1290
         TabIndex        =   16
         Top             =   240
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.ListBox lstBIN 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   120
         TabIndex        =   7
         Top             =   540
         Width           =   1275
      End
      Begin VB.Image imgBWH 
         Height          =   255
         Left            =   90
         Picture         =   "frmViewer.frx":04C9
         Top             =   210
         Width           =   255
      End
   End
   Begin VB.Frame frFont 
      Caption         =   "Font Viewer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4650
      Left            =   6000
      TabIndex        =   28
      Top             =   4200
      Visible         =   0   'False
      Width           =   11685
      Begin VB.CheckBox cbMC 
         Caption         =   "Multi-color"
         Height          =   255
         Left            =   6600
         TabIndex        =   164
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmdSB 
         Caption         =   ">>"
         Height          =   270
         Index           =   5
         Left            =   10770
         TabIndex        =   148
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmdSB 
         Caption         =   ">"
         Height          =   270
         Index           =   4
         Left            =   10440
         TabIndex        =   147
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmdSB 
         Caption         =   "+"
         Height          =   270
         Index           =   3
         Left            =   10110
         TabIndex        =   146
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmdSB 
         Caption         =   "-"
         Height          =   270
         Index           =   2
         Left            =   9840
         TabIndex        =   145
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdSB 
         Caption         =   "<"
         Height          =   270
         Index           =   1
         Left            =   9570
         TabIndex        =   144
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdSB 
         Caption         =   "<<"
         Height          =   270
         Index           =   0
         Left            =   9240
         TabIndex        =   143
         Top             =   240
         Width           =   315
      End
      Begin VB.CheckBox cbFCols 
         Caption         =   "Wide"
         Height          =   255
         Left            =   4980
         TabIndex        =   52
         Top             =   240
         Value           =   1  'Checked
         Width           =   675
      End
      Begin VB.TextBox txtCSkip 
         Height          =   285
         Left            =   8520
         TabIndex        =   50
         Text            =   "0"
         ToolTipText     =   "Set number of bytes to skip (decimal)"
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton cmdSaveCSet 
         Caption         =   "Save BMP..."
         Height          =   315
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "Save as BMP file"
         Top             =   4200
         Width           =   1245
      End
      Begin VB.OptionButton optChrH 
         Caption         =   "8x8"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   3900
         Value           =   -1  'True
         Width           =   555
      End
      Begin VB.OptionButton optChrH 
         Caption         =   "8x16"
         Height          =   195
         Index           =   1
         Left            =   750
         TabIndex        =   35
         Top             =   3930
         Width           =   675
      End
      Begin VB.CheckBox cbBorder 
         Caption         =   "Border"
         Height          =   255
         Left            =   5760
         TabIndex        =   34
         Top             =   240
         Value           =   1  'Checked
         Width           =   885
      End
      Begin VB.ComboBox cboTheme 
         Height          =   315
         ItemData        =   "frmViewer.frx":087F
         Left            =   1440
         List            =   "frmViewer.frx":0898
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   240
         Width           =   1575
      End
      Begin VB.PictureBox picV 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   1440
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   32
         Top             =   570
         Width           =   465
      End
      Begin VB.PictureBox picChr 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2490
         Left            =   120
         ScaleHeight     =   162
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   31
         Top             =   1380
         Width           =   1280
      End
      Begin VB.CommandButton cmdCSPrev 
         Caption         =   "<"
         Height          =   255
         Left            =   660
         TabIndex        =   30
         ToolTipText     =   "Previous character"
         Top             =   570
         Width           =   330
      End
      Begin VB.CommandButton cmdCSNxt 
         Caption         =   ">"
         Height          =   255
         Left            =   990
         TabIndex        =   29
         ToolTipText     =   "Next character"
         Top             =   570
         Width           =   360
      End
      Begin VB.Label lblTheme 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   4
         Left            =   1050
         TabIndex        =   163
         Top             =   390
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lblTheme 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   3
         Left            =   1050
         TabIndex        =   162
         Top             =   240
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lblZoom 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5x"
         Height          =   270
         Index           =   4
         Left            =   4530
         TabIndex        =   155
         Top             =   240
         Width           =   345
      End
      Begin VB.Label lblZoom 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4x"
         Height          =   270
         Index           =   3
         Left            =   4170
         TabIndex        =   154
         Top             =   240
         Width           =   345
      End
      Begin VB.Label lblZoom 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3x"
         Height          =   270
         Index           =   2
         Left            =   3810
         TabIndex        =   153
         Top             =   240
         Width           =   345
      End
      Begin VB.Label lblZoom 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2x"
         Height          =   270
         Index           =   1
         Left            =   3450
         TabIndex        =   152
         Top             =   240
         Width           =   345
      End
      Begin VB.Label lblZoom 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1x"
         Height          =   270
         Index           =   0
         Left            =   3090
         TabIndex        =   151
         Top             =   240
         Width           =   345
      End
      Begin VB.Label lblTheme 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   150
         Top             =   240
         Width           =   285
      End
      Begin VB.Label lblTheme 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   420
         TabIndex        =   149
         Top             =   240
         Width           =   285
      End
      Begin VB.Label lblChrX 
         BackColor       =   &H0080C0FF&
         Height          =   465
         Left            =   120
         TabIndex        =   141
         Top             =   870
         Width           =   1245
      End
      Begin VB.Label lblEndRange 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   11220
         TabIndex        =   51
         Top             =   270
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Skip Bytes:"
         Height          =   195
         Left            =   7710
         TabIndex        =   49
         Top             =   270
         Width           =   795
      End
      Begin VB.Label lblTheme 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   285
      End
      Begin VB.Label lblChrNum 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "000"
         Height          =   225
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Frame frBlank 
      Height          =   855
      Left            =   9390
      TabIndex        =   44
      Top             =   1200
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Select Viewer with button above..."
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   2430
      End
   End
   Begin VB.CheckBox cbLockView 
      Caption         =   "Lock View"
      Height          =   315
      Left            =   10770
      TabIndex        =   43
      ToolTipText     =   "Lock to Current View"
      Top             =   30
      Width           =   1215
   End
   Begin VB.PictureBox Pix 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFF00&
      Height          =   3840
      Left            =   18960
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   120
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   11460
      Top             =   -30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frBMP 
      Caption         =   "Bitmap Viewer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   5700
      TabIndex        =   18
      Top             =   1140
      Visible         =   0   'False
      Width           =   3555
      Begin VB.CommandButton cmdBSave 
         Caption         =   "Save..."
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   300
         Width           =   975
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H00000000&
         Height          =   3000
         Left            =   120
         ScaleHeight     =   200
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   213
         TabIndex        =   19
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label lblMoment 
         Caption         =   "One moment... loading BMP"
         Height          =   435
         Left            =   90
         TabIndex        =   174
         Top             =   810
         Width           =   2835
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Comment:"
         Height          =   195
         Left            =   1260
         TabIndex        =   24
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Format:"
         Height          =   195
         Left            =   1260
         TabIndex        =   23
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lblBType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2040
         TabIndex        =   21
         Top             =   240
         Width           =   45
      End
      Begin VB.Label lblBComment 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2040
         TabIndex        =   20
         Top             =   480
         Width           =   45
      End
   End
   Begin VB.Frame frML 
      Caption         =   "Machine Language Disassembler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9345
      Left            =   30
      TabIndex        =   4
      Top             =   540
      Visible         =   0   'False
      Width           =   12900
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "X"
         Height          =   315
         Index           =   6
         Left            =   7590
         TabIndex        =   178
         ToolTipText     =   "Make Hidden Block"
         Top             =   210
         Width           =   375
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "*C*"
         Height          =   315
         Index           =   4
         Left            =   9960
         TabIndex        =   114
         ToolTipText     =   "Add Comment with * Separator"
         Top             =   210
         Width           =   435
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "***"
         Height          =   315
         HelpContextID   =   7
         Index           =   7
         Left            =   11310
         TabIndex        =   113
         ToolTipText     =   "Add * Separator"
         Top             =   210
         Width           =   435
      End
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "W"
         Height          =   315
         Index           =   5
         Left            =   7200
         TabIndex        =   112
         ToolTipText     =   "Make Word Block"
         Top             =   210
         Width           =   375
      End
      Begin VB.CommandButton cmdAddLabel 
         Caption         =   "Label"
         Height          =   315
         Left            =   4500
         TabIndex        =   111
         ToolTipText     =   "Add Label"
         Top             =   210
         Width           =   585
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "==="
         Height          =   315
         Index           =   6
         Left            =   10860
         TabIndex        =   110
         ToolTipText     =   "Add = Separator"
         Top             =   210
         Width           =   435
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "----"
         Height          =   315
         Index           =   5
         Left            =   10410
         TabIndex        =   109
         ToolTipText     =   "Add - Separator"
         Top             =   210
         Width           =   435
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "=C="
         Height          =   315
         Index           =   3
         Left            =   9510
         TabIndex        =   108
         ToolTipText     =   "Add Comment with = Separator"
         Top             =   210
         Width           =   435
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "--C--"
         Height          =   315
         Index           =   2
         Left            =   9060
         TabIndex        =   107
         ToolTipText     =   "Add Comment with - Separator"
         Top             =   210
         Width           =   435
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "C"
         Height          =   315
         Index           =   1
         Left            =   8610
         TabIndex        =   106
         ToolTipText     =   "Add Standalone Comment"
         Top             =   210
         Width           =   435
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   ";C"
         Height          =   315
         Index           =   0
         Left            =   8160
         TabIndex        =   105
         ToolTipText     =   "Add Inline Comment"
         Top             =   210
         Width           =   435
      End
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "V"
         Height          =   315
         Index           =   4
         Left            =   6810
         TabIndex        =   104
         ToolTipText     =   "Make Vector Block"
         Top             =   210
         Width           =   375
      End
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "R"
         Height          =   315
         Index           =   3
         Left            =   6420
         TabIndex        =   103
         ToolTipText     =   "Make RTS vector block"
         Top             =   210
         Width           =   375
      End
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "T"
         Height          =   315
         Index           =   2
         Left            =   6030
         TabIndex        =   102
         ToolTipText     =   "Make Text Block"
         Top             =   210
         Width           =   375
      End
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "H"
         Height          =   315
         Index           =   1
         Left            =   5640
         TabIndex        =   101
         ToolTipText     =   "Make Hex Block"
         Top             =   210
         Width           =   375
      End
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "D"
         Height          =   315
         Index           =   0
         Left            =   5250
         TabIndex        =   100
         ToolTipText     =   "Make Dec Byte Block"
         Top             =   210
         Width           =   375
      End
      Begin VB.CommandButton cmdFindAll 
         Caption         =   "Find All"
         Height          =   315
         Left            =   3030
         TabIndex        =   54
         ToolTipText     =   "Find all occurences"
         Top             =   210
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   315
         Left            =   3750
         TabIndex        =   40
         ToolTipText     =   "Jump to Next"
         Top             =   210
         Width           =   495
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   315
         Left            =   2460
         TabIndex        =   39
         ToolTipText     =   "Find Text"
         Top             =   210
         Width           =   555
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   315
         Left            =   930
         TabIndex        =   27
         Top             =   210
         Width           =   705
      End
      Begin VB.ListBox lstML 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3990
         MultiSelect     =   2  'Extended
         TabIndex        =   5
         Top             =   660
         Width           =   1275
      End
      Begin VB.Frame frTView 
         Height          =   8625
         Left            =   90
         TabIndex        =   55
         Top             =   570
         Width           =   3825
         Begin VB.Frame frMLSettings 
            Height          =   6045
            Left            =   450
            TabIndex        =   69
            Top             =   2280
            Width           =   3615
            Begin VB.CommandButton cmdImport 
               Caption         =   "Import"
               Height          =   345
               Left            =   2040
               TabIndex        =   120
               ToolTipText     =   "Import Symbols"
               Top             =   4890
               Width           =   1455
            End
            Begin VB.CheckBox cbIncSym 
               Caption         =   "Include Symbol comments"
               Height          =   375
               Left            =   150
               TabIndex        =   99
               Top             =   4110
               Value           =   1  'Checked
               Width           =   3255
            End
            Begin VB.ComboBox cboCPUFile 
               BackColor       =   &H00FFFFFF&
               CausesValidation=   0   'False
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "frmViewer.frx":08D6
               Left            =   2220
               List            =   "frmViewer.frx":08D8
               Style           =   2  'Dropdown List
               TabIndex        =   98
               Top             =   1740
               Visible         =   0   'False
               Width           =   1305
            End
            Begin VB.ComboBox cboCPU 
               BackColor       =   &H00000080&
               CausesValidation=   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmViewer.frx":08DA
               Left            =   810
               List            =   "frmViewer.frx":08E1
               Style           =   2  'Dropdown List
               TabIndex        =   96
               Top             =   1740
               Width           =   2715
            End
            Begin VB.TextBox txtDivLen 
               BackColor       =   &H00000080&
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   1920
               MaxLength       =   2
               TabIndex        =   95
               Text            =   "80"
               Top             =   3060
               Width           =   345
            End
            Begin VB.ComboBox cboPlatFile 
               BackColor       =   &H00FFFFFF&
               CausesValidation=   0   'False
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "frmViewer.frx":08F3
               Left            =   2220
               List            =   "frmViewer.frx":08F5
               Style           =   2  'Dropdown List
               TabIndex        =   93
               Top             =   1380
               Visible         =   0   'False
               Width           =   1305
            End
            Begin VB.ComboBox cboPlatform 
               BackColor       =   &H00000080&
               CausesValidation=   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmViewer.frx":08F7
               Left            =   810
               List            =   "frmViewer.frx":08FE
               Style           =   2  'Dropdown List
               TabIndex        =   92
               Top             =   1410
               Width           =   2715
            End
            Begin VB.CommandButton cmdMLHelp 
               Caption         =   "Help"
               Height          =   465
               Left            =   600
               TabIndex        =   90
               ToolTipText     =   "Display HELP file"
               Top             =   5400
               Width           =   2385
            End
            Begin VB.CheckBox cbLabelBlanks 
               Caption         =   "Add blank line before Labels"
               Height          =   375
               Left            =   150
               TabIndex        =   89
               Top             =   3840
               Value           =   1  'Checked
               Width           =   3285
            End
            Begin VB.ComboBox cboPrefix 
               BackColor       =   &H00000080&
               CausesValidation=   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmViewer.frx":0910
               Left            =   1110
               List            =   "frmViewer.frx":0917
               Style           =   2  'Dropdown List
               TabIndex        =   87
               Top             =   2730
               Width           =   2415
            End
            Begin VB.ComboBox cboTarget 
               BackColor       =   &H00000080&
               CausesValidation=   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmViewer.frx":0929
               Left            =   810
               List            =   "frmViewer.frx":0936
               Style           =   2  'Dropdown List
               TabIndex        =   85
               Top             =   2400
               Width           =   2715
            End
            Begin VB.CommandButton cmdSaveASM 
               Caption         =   "Save..."
               Height          =   375
               Left            =   1050
               TabIndex        =   81
               ToolTipText     =   "Save disassembly to file"
               Top             =   4470
               Width           =   915
            End
            Begin VB.CheckBox cbSpaceRTS 
               Caption         =   "Add blank line after RTS/RTI instructions"
               Height          =   375
               Left            =   150
               TabIndex        =   80
               Top             =   3570
               Value           =   1  'Checked
               Width           =   3285
            End
            Begin VB.CommandButton cmdPurge 
               Caption         =   "Purge"
               Height          =   345
               Left            =   1050
               TabIndex        =   79
               ToolTipText     =   "Purge unselected symbol entries"
               Top             =   4890
               Width           =   915
            End
            Begin VB.CommandButton cmdClrTables 
               Caption         =   "New Project"
               Height          =   315
               Left            =   2340
               TabIndex        =   78
               ToolTipText     =   "Clear Lists and start a new project"
               Top             =   600
               Width           =   1185
            End
            Begin VB.CheckBox cbClearOnLoad 
               Caption         =   "Clear Lists on Load"
               Height          =   375
               Left            =   120
               TabIndex        =   77
               ToolTipText     =   "Uncheck if you want to keep existing entries when loading"
               Top             =   570
               Value           =   1  'Checked
               Width           =   2055
            End
            Begin VB.CommandButton cmdProjSave 
               Caption         =   "Save Project..."
               Height          =   315
               Left            =   2100
               TabIndex        =   76
               ToolTipText     =   "Save Lists to file"
               Top             =   210
               Width           =   1410
            End
            Begin VB.CommandButton cmdProjLoad 
               Caption         =   "Load Project..."
               Height          =   315
               Left            =   120
               TabIndex        =   75
               ToolTipText     =   "Load Lists from a file"
               Top             =   210
               Width           =   1410
            End
            Begin VB.CommandButton cmdCopyClip2 
               Caption         =   "Copy To &Clipboard"
               Height          =   375
               Left            =   2040
               TabIndex        =   73
               ToolTipText     =   "Paste disassembly to clipboard"
               Top             =   4470
               Width           =   1455
            End
            Begin VB.CheckBox cbEquates 
               Caption         =   "Show Equates"
               Height          =   195
               Left            =   150
               TabIndex        =   72
               ToolTipText     =   "Include Equates in output"
               Top             =   3390
               Width           =   1515
            End
            Begin VB.ComboBox cboMLFmt 
               BackColor       =   &H00000080&
               CausesValidation=   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmViewer.frx":0951
               Left            =   810
               List            =   "frmViewer.frx":0964
               Style           =   2  'Dropdown List
               TabIndex        =   70
               Top             =   2070
               Width           =   2715
            End
            Begin VB.Label lblChanged 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   225
               Left            =   1710
               TabIndex        =   121
               ToolTipText     =   "Project Status (Green=OK, Red=Changed, White=No Project Loaded)"
               Top             =   270
               Width           =   225
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Symbols:"
               Height          =   195
               Left            =   330
               TabIndex        =   119
               Top             =   4920
               Width           =   630
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "CPU:"
               Height          =   195
               Left            =   390
               TabIndex        =   97
               Top             =   1800
               Width           =   375
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Comment Divider length:"
               Height          =   195
               Left            =   120
               TabIndex        =   94
               Top             =   3090
               Width           =   1725
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Platform:"
               Height          =   195
               Left            =   150
               TabIndex        =   91
               Top             =   1470
               Width           =   615
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Disassembly:"
               Height          =   195
               Left            =   90
               TabIndex        =   88
               Top             =   4530
               Width           =   915
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Label Prefix:"
               Height          =   195
               Left            =   120
               TabIndex        =   86
               Top             =   2790
               Width           =   870
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Target:"
               Height          =   195
               Left            =   270
               TabIndex        =   84
               Top             =   2460
               Width           =   510
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "View Fmt:"
               Height          =   195
               Left            =   90
               TabIndex        =   71
               Top             =   2130
               Width           =   690
            End
         End
         Begin VB.ListBox lstEntryPt 
            BackColor       =   &H00000080&
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            ItemData        =   "frmViewer.frx":09B9
            Left            =   90
            List            =   "frmViewer.frx":09BB
            TabIndex        =   176
            Top             =   1590
            Width           =   705
         End
         Begin VB.Frame frTrace 
            Height          =   4425
            Left            =   150
            TabIndex        =   166
            Top             =   2040
            Width           =   3675
            Begin VB.CheckBox cbMLAddLabels 
               Caption         =   " Add Labels"
               Height          =   255
               Left            =   150
               TabIndex        =   177
               Top             =   2130
               Value           =   1  'Checked
               Width           =   1155
            End
            Begin VB.CommandButton cmdAddTables 
               Caption         =   "Add To Tables"
               Height          =   645
               Left            =   120
               TabIndex        =   169
               Top             =   1230
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.CommandButton cmdTrace 
               Caption         =   "START"
               Height          =   795
               Left            =   90
               TabIndex        =   168
               Top             =   270
               Width           =   1155
            End
            Begin VB.ListBox lstEP 
               BackColor       =   &H00FFFFFF&
               CausesValidation=   0   'False
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   3960
               ItemData        =   "frmViewer.frx":09BD
               Left            =   1380
               List            =   "frmViewer.frx":09BF
               Sorted          =   -1  'True
               TabIndex        =   167
               Top             =   240
               Width           =   2160
            End
         End
         Begin VB.ListBox lstJSR 
            BackColor       =   &H00C0C000&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            ItemData        =   "frmViewer.frx":09C1
            Left            =   2940
            List            =   "frmViewer.frx":09C3
            Sorted          =   -1  'True
            TabIndex        =   116
            Top             =   1320
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.ListBox lstLabels 
            BackColor       =   &H00C0C000&
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   255
            ItemData        =   "frmViewer.frx":09C5
            Left            =   1980
            List            =   "frmViewer.frx":09C7
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   82
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.ListBox lstCmnt 
            BackColor       =   &H00000080&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            ItemData        =   "frmViewer.frx":09C9
            Left            =   2940
            List            =   "frmViewer.frx":09CB
            Sorted          =   -1  'True
            TabIndex        =   68
            Top             =   1590
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.ListBox lstULabels 
            BackColor       =   &H00000080&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            ItemData        =   "frmViewer.frx":09CD
            Left            =   2190
            List            =   "frmViewer.frx":09CF
            Sorted          =   -1  'True
            TabIndex        =   66
            Top             =   1590
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.ListBox lstDT 
            BackColor       =   &H00000080&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            ItemData        =   "frmViewer.frx":09D1
            Left            =   1530
            List            =   "frmViewer.frx":09D3
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   65
            Top             =   1590
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.ListBox lstSYM 
            BackColor       =   &H00000080&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            ItemData        =   "frmViewer.frx":09D5
            Left            =   840
            List            =   "frmViewer.frx":09D7
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   64
            Top             =   1590
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.CommandButton cmdSymAdd 
            Caption         =   "Add"
            Height          =   315
            Left            =   2070
            TabIndex        =   63
            ToolTipText     =   "Add an entry"
            Top             =   930
            Width           =   495
         End
         Begin VB.CommandButton cmdSymDel 
            Caption         =   "Del"
            Height          =   315
            Left            =   2610
            TabIndex        =   62
            ToolTipText     =   "Delete current entry"
            Top             =   930
            Width           =   495
         End
         Begin VB.CommandButton cmdSYMGoto 
            Caption         =   "Find"
            Height          =   315
            Left            =   3180
            TabIndex        =   61
            ToolTipText     =   "Find Selected"
            Top             =   930
            Width           =   555
         End
         Begin VB.CommandButton cmdSymSave 
            Caption         =   "Save"
            Height          =   315
            Left            =   690
            TabIndex        =   60
            ToolTipText     =   "Save file"
            Top             =   930
            Width           =   555
         End
         Begin VB.CommandButton cmdSymLoad 
            Caption         =   "Load"
            Height          =   315
            Left            =   90
            TabIndex        =   59
            ToolTipText     =   "Load a file"
            Top             =   930
            Width           =   555
         End
         Begin VB.CommandButton cmdRemDupLbls 
            Caption         =   "Remove Duplicates"
            Height          =   315
            Left            =   90
            TabIndex        =   118
            ToolTipText     =   "Remove Duplicate Entries"
            Top             =   930
            Width           =   1845
         End
         Begin VB.CommandButton cmdRemDupJSR 
            Caption         =   "Remove Duplicates"
            Height          =   315
            Left            =   90
            TabIndex        =   117
            ToolTipText     =   "Remove Duplicate Entries"
            Top             =   930
            Width           =   1845
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Entry Pt"
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   2
            Left            =   60
            TabIndex        =   175
            Top             =   540
            Width           =   720
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            BackColor       =   &H00008000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TRACER"
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   1
            Left            =   1050
            TabIndex        =   165
            Top             =   180
            Width           =   870
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ext JSR"
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   8
            Left            =   2910
            TabIndex        =   115
            Top             =   180
            Width           =   840
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Gen Labels"
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   7
            Left            =   1950
            TabIndex        =   83
            Top             =   180
            Width           =   930
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PROJECT"
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   74
            Top             =   180
            Width           =   960
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Comments"
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   6
            Left            =   2910
            TabIndex        =   67
            Top             =   540
            Width           =   840
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tables"
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   4
            Left            =   1560
            TabIndex        =   58
            Top             =   540
            Width           =   630
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Labels"
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   5
            Left            =   2220
            TabIndex        =   57
            Top             =   540
            Width           =   660
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Symbols"
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   3
            Left            =   810
            TabIndex        =   56
            Top             =   540
            Width           =   720
         End
      End
      Begin VB.CheckBox cbAuto 
         Caption         =   "Auto"
         Height          =   195
         Left            =   1680
         TabIndex        =   53
         ToolTipText     =   "Automatically Refresh"
         Top             =   270
         Value           =   1  'Checked
         Width           =   675
      End
      Begin VB.Image imgBW 
         Height          =   255
         Left            =   330
         Picture         =   "frmViewer.frx":09D9
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblShw 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   48
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblGood 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   630
         TabIndex        =   26
         ToolTipText     =   "Disassembly Status (Green=OK, Red=Problems)"
         Top             =   270
         Width           =   225
      End
      Begin VB.Label lblEA 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   195
         Left            =   11880
         TabIndex        =   25
         ToolTipText     =   "Address range"
         Top             =   270
         Width           =   105
      End
   End
   Begin VB.Shape shOverflow 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   8400
      Shape           =   3  'Circle
      Top             =   60
      Width           =   225
   End
   Begin VB.Label lblVSize 
      Alignment       =   1  'Right Justify
      Caption         =   "00000"
      Height          =   225
      Left            =   7920
      LinkTimeout     =   0
      TabIndex        =   173
      Top             =   75
      Width           =   450
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Size:"
      Height          =   225
      Left            =   7530
      TabIndex        =   172
      Top             =   75
      Width           =   345
   End
   Begin VB.Label lblSSize 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "||"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   2
      Left            =   10020
      TabIndex        =   124
      ToolTipText     =   "Return split to CENTRE"
      Top             =   30
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblSSize 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ">>"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   10350
      TabIndex        =   123
      ToolTipText     =   "Move Split RIGHT"
      Top             =   30
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblSSize 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<<"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   9720
      TabIndex        =   122
      ToolTipText     =   "Move Split LEFT"
      Top             =   30
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblSelect 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ">"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9090
      TabIndex        =   47
      ToolTipText     =   "Select LEFT/RIGHT View"
      Top             =   30
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   8700
      TabIndex        =   46
      ToolTipText     =   "Toggle Dual View Mode"
      Top             =   30
      Width           =   345
   End
   Begin VB.Label lblView 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BITMAP"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   5
      Left            =   5370
      TabIndex        =   17
      Top             =   30
      Width           =   915
   End
   Begin VB.Label lblView 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ASM"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   4
      Left            =   4440
      TabIndex        =   15
      Top             =   30
      Width           =   915
   End
   Begin VB.Label lblView 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FONT"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   3510
      TabIndex        =   14
      Top             =   30
      Width           =   915
   End
   Begin VB.Label lblView 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HEX"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   2
      Left            =   2580
      TabIndex        =   13
      Top             =   30
      Width           =   915
   End
   Begin VB.Label lblView 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SEQ"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   1650
      TabIndex        =   12
      Top             =   30
      Width           =   915
   End
   Begin VB.Label lblView 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BASIC"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   720
      TabIndex        =   11
      Top             =   30
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "View As:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   615
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2017 Steve J. Gray
' ====================================================
'
' frmViewer - Multi-view File Viewer
'
' Supports:
' BAS    - Most dialects of CBM BASIC from 2.0 to 10
' SEQ    - Sequential Files
' HEX    - Binary Files
' FONT   - Commodore character sets, or any 8-pixel storage format
' ASM    - Machine Language Interactive Symbolic Disassembler with Flow Tracer
' BITMAP - Common CBM Bitmap file formats, including GeoPaint

'==== Common shared variables
Public ViewerReady As Boolean

Public VBuf As String                                          'ViewFile Buffer - All viewers share this buffer
Public VFileName As String, VName As String, VExt As String    'ViewFile Info
Public VLen As Long, VLA As Long                               'ViewFile Length, Load Address
Public VP00Buf As String, VP00Flag As Boolean                  'ViewFile P00 buffer, and flag
Public ViewReady As Boolean, ViewBusy As Boolean               'Flag when processing
Public ViewMode As Integer, ViewMode2 As Integer               'Which tabs are displayed
Public LockV1 As Integer, LockV2 As Integer                    'Which tabs are locked
Public SplitMode As Boolean, SplitSize As Integer              'Dual-view split


'==== Bitmap Viewer
Const NUMB = 20, GEO = -1, HRBW = 0, HR = 1, MC = 2

Dim PBuf As String                                              'Picture Buffer
Dim PicName As String
Dim CBMColor(15) As Long                                        'VIC-II colour values for bitmaps and character
Dim ImageType As Integer
Dim PFIO As Integer                                             'Picture file#, shared with multiple subs (needs re-writing)

Dim xInit As String, xFile As String

Dim p_name(0 To NUMB)   As String
Dim p_sa(1 To NUMB)     As Long
Dim p_len(1 To NUMB)    As Long
Dim p_bitmap(1 To NUMB) As Long
Dim p_screen(1 To NUMB) As Long
Dim p_colour(1 To NUMB) As Long
Dim p_back(1 To NUMB)   As Long
Dim p_type(0 To NUMB)   As Integer

Dim Pow(7) 'binary powers array

'==== BASIC Viewer
Dim Token(358) As String

'==== FONT Viewer
Dim SelChr As Integer, FontH As Integer, ChrZoom As Integer

'==== ML Viewer

Dim OP(255) As String                                           '6502 Opcodes
Dim OpModeLen As String                                         'Opcode Addresing Mode Lengths (number of bytes for specified addressing mode)
Dim OpB As String, OpJ As String, OpZ As String                 'Tracer opcode groups: Branches, Jumps, Stops
Dim OpDesc As String                                            'Opcode Description from file

Dim LastFile As String, LastComment As String, LastSymPos As Integer
Public ProjFlag As Boolean, MLCFlag As Boolean, ChangeFlag As Boolean
Public MLTabNum As Integer
Public OpCodeFlag As Boolean, ShowTables As Boolean
Public DOTORG As String, DOTWORD As String, DOTBYTE As String, DOTTEXT As String
Public LPrefix As String, ProjFilename As String

'---- Load the Form
Private Sub Form_Load()
    Dim i As Integer
    
    On Error Resume Next
    
    ViewerReady = False                     'Make sure changing drop-down menus doesn't cause other code to run
    
    ViewMode = 0: ViewMode2 = -1            'Default View Modes
    SplitMode = False: SplitSize = 50       'Dual-view mode
        
    cboMode.ListIndex = 0                   'MLView
    
    cboMLFmt.ListIndex = 0                  'MLView output format combo
    cboTarget.ListIndex = 0                 'MLView targe assembler combo
    
    cboPrefix.ListIndex = 0                 'MLView label prefix combo
    cboPlatform.ListIndex = 0               'MLView platform combo
    cboCPU.ListIndex = 0                    'MLView CPU combo
    
    ProjFlag = False                        'ML Viewer
    MLCFlag = False                         'ML Viewer
    ShowTables = False                      'ML Viewer
    MLTabNum = 0                            'ML Viewer
    SetTarget 0                             'Target Assembler
    SetPrefix 0                             'Label Prefix
            
    ChrZoom = 1: SelChr = 0                 'Chr Viewer
    
    For i = 0 To 7: Pow(i) = 2 ^ i: Next    'Set Powers of 2
    
    Call SetColor                           'Setup C64 colours
    
    Me.Show: DoEvents
    ViewerReady = True
    
End Sub

'---- Process the Dropped File
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.GetFormat(vbCFFiles) Then
     Dim vFn As Variant
     For Each vFn In Data.Files
       ViewIt ViewMode, vFn, "", ""     'vFn is name of file dropped
       Exit For                         'only process the first dropped file!
     Next
  End If
End Sub

'-- Unload the Form? - Check if ASM Project needs saving before Exiting
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If OverwriteProject = False Then Cancel = True
End Sub

'--- Resize the Form
Private Sub Form_Resize()
    DrawVLayout
End Sub

'---- ViewIt
' This is called from the other form to start things off!
' We call this when the file needs loading, or if the Load Address checkbox is changed.
' The file is opened and the contents (minus P00 header and load address bytes) are loaded to the buffer
'
' Mode..... Default View# (TAB#)
' SrcFile.. full filename with path
' SrcName.. Name for Titlebar
' SrcExt... Extension of file if known
Sub ViewIt(ByVal Mode As Integer, ByVal SrcFile As String, ByVal SrcName As String, Optional SrcExt As String)
    Dim Lo As Integer, Hi As Integer, FIO As Integer, Tmp As String

    If SrcFile = "" Then Exit Sub
    
    If Exists(SrcFile) = False Then MyMsg "Viewer: File '" & SrcFile & "' not found!": Exit Sub
    
    VFileName = SrcFile: VName = SrcName: VExt = SrcExt     'ViewFile Details
    VP00Flag = False                                        'Assume normal file
    If FileExtU(VFileName) = "P00" Then VP00Flag = True     'P00 file found!
    
    cbLA.Enabled = False                                                        'Don't allow changing LA checkbox
    
    '-- Load the file to the buffer, update and display file details
    FIO = FreeFile
    Open VFileName For Binary As FIO: VLen = intLOF(FIO)
        If VP00Flag = True Then VP00Buf = Input(26, FIO): VLen = VLen - 26      'Skip over header
        If cbLA.value = vbChecked Then
            VBuf = Input(2, FIO): VLen = VLen - 2                               'Read the Load address
            Lo = Asc(Mid(VBuf, 1, 1))                                           'Use first two bytes as load address
            Hi = Asc(Mid(VBuf, 2, 1))
            VLA = Hi * 256& + Lo                                                'Load Address
            txtLA.Enabled = False
        Else
            VLA = 0                                                             'No Load Address
            txtLA.Enabled = True
        End If
        
        shOverflow.Visible = False
        If VLen > 32760 Then VLen = 32760: shOverflow.Visible = True            'Max size we can load! Overflow indicator
        VBuf = Input(VLen, FIO)                                                 'Read contents to buffer
       
    Close FIO
    
    txtLA.Text = MyHex(VLA, 4)                                                  'Load Address in Hex
    lblVSize.Caption = Format(VLen)                                             'File Size
    cbLA.Enabled = True                                                         'Re-enable LA checkbox
    
    Tmp = "Viewer: " & FileNameOnly(VName)                                      'Titlebar string
    If VP00Flag = True Then Tmp = Tmp & " (Contained inside P00)"               'Add P00 note
    Me.Caption = Tmp                                                            'Set Window Titlebar
        
    SetBestMatch Mode                                                           'Determine Best Initial view
    SelectNewTab ViewMode
    UpdateViews                                                                 'Update Tab Views

End Sub

'---- SetBestMatch
' Called when view options change (ie require content refresh) but file options stay the same (LA)
Sub SetBestMatch(ByVal Mode As Integer)
    Dim i As Integer, Flag As Boolean
        
    ViewMode = Mode                                'default for unkown types
    
    If cbLockView.value = vbUnchecked Then
        Select Case UCase(VExt)
            Case "BAS": ViewMode = 0
            Case "SEQ", ",S": ViewMode = 1
            Case "BIN", "ROM"
                ViewMode = 2
                If Exists(FileBase(VFileName) & ".asm-proj") = True Then ViewMode = 4 'if there is an ASM Project file use ASM mode
            Case "ASM": ViewMode = 4
            Case "ART", "CDU", "GEO", "KOA": ViewMode = 5
        End Select
    Else
        ViewMode = LockV1
        ViewMode2 = LockV2
    End If

End Sub

'-- SelectNewTab - Handle clicking view tab
Private Sub SelectNewTab(ByVal NewTabNum As Integer)
        
    If (lblSelect.Caption = "<") Or (SplitMode = False) Then
        If NewTabNum <> ViewMode Then
            If NewTabNum = ViewMode2 Then ViewMode2 = ViewMode
            ViewMode = NewTabNum: LockV = NewTabNum
            RefreshContent ViewMode
        End If
    Else
        If NewTabNum <> ViewMode2 Then
            If NewTabNum = ViewMode Then ViewMode = ViewMode2
            ViewMode2 = NewTabNum: LockV2 = NewTabNum
            RefreshContent ViewMode2
        End If
    End If
    
    DrawVLayout
    
End Sub

'---- UpdateViews - Update contents of left and/or right
Private Sub UpdateViews()
    RefreshContent ViewMode                             'Update content in left view
    If SplitMode = True Then RefreshContent ViewMode2   'Update content in right view
End Sub

'---- Refresh Content
' Re-draw content of selected View
Private Sub RefreshContent(ByVal Mode As Integer)

    Select Case Mode
        Case 0: BASView
        Case 1: SEQView
        Case 2: HEXView
        Case 3: FONTView
        Case 4: MLView
        Case 5: BMPView
    End Select

    ViewBusy = False
    
End Sub
'---- Draw View Layout
' This is called when the form is resized. It positions the tab frames and updates the tab highlighting.
' It makes the frames visible. SplitMode=TRUE enables dual view.
' NOTE: This does NOT update the contents of the frames!
Private Sub DrawVLayout()
    Dim W As Single, h As Single                        'Original Window Size
    Dim W1 As Single, H1 As Single, L1 As Single        'Scaled Window Size LEFT frame
    Dim T1 As Single                                    'Top offset
    Dim W2 As Single, L2 As Single                      'Scaled Width and LeftPosition for RIGHT frame
    Dim i As Integer
    
    If ViewerReady = False Then Exit Sub
    
    '-- Hide all the frames
    frBasic.Visible = False
    frFont.Visible = False
    frML.Visible = False
    frBIN.Visible = False
    frSEQ.Visible = False
    frBMP.Visible = False
    frBlank.Visible = False
    For i = 0 To 2: lblSSize(i).Visible = False: Next
    
    DoEvents
        
    '-- Calculate window sizes
    W = Me.Width - 390:   If W < 4400 Then W = 4400         'Window Width - enforce minimum size for elements
    h = Me.Height - 1000:  If h < 3700 Then h = 3700        'Window Height - enforce min size for elements
    L1 = 75: T1 = 375                                       'Left/Top Margins
    W1 = W: W2 = W: H1 = h: L2 = L1                         'Set for single-view mode
    
    '-- Calculate Split mode sizes
    If SplitMode = True Then
        For i = 0 To 2: lblSSize(i).Visible = True: Next    'Show the split re-sizers
        W1 = W * (SplitSize / 100)                          'Calc new width of LEFT frame
        W2 = W * ((100 - SplitSize) / 100) - L1             'Calc new width pf RIGHT frame
        L2 = L1 * 2 + W1                                    'Calculate Left offset
    End If
    
    '-- Position the frames
    SetFrame ViewMode, L1, T1, W1, H1, True                 'Position and Show Frame on LEFT
    SetFrame ViewMode2, L2, T1, W2, H1, SplitMode           'Position and Show Frame on RIGHT (if SplitMode=TRUE)
    DoEvents
    
    '-- Update top line buttons
    For i = 0 To 5
        If (i = ViewMode) Or ((i = ViewMode2) And (SplitMode = True)) Then
            lblView(i).Font.Bold = True
            lblView(i).ForeColor = vbWhite
        Else
            lblView(i).Font.Bold = False
            lblView(i).ForeColor = vbBlack
        End If
    Next
    
    DoEvents
End Sub

'---- SetFrame
' Arrange View Elements
' N=Frame#, Size: L=Left,T=Top,W=Width,H=Height, FLAG=Frame Visible?
' In Dual-View Mode FLAG=TRUE
Sub SetFrame(ByVal n As Integer, ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal h As Single, ByVal Flag As Boolean)
    Dim W2 As Single, H2 As Single, W3 As Single, H3 As Single, W4 As Single, H4 As Single
    Dim WW As Single, HH As Single, LL As Single, TT As Single
    
    W2 = W: H2 = h: W3 = W - 200: H3 = h - 600
   
    Select Case n
        Case -1 '-- Blank frame with message
            frBlank.Visible = Flag
            frBlank.Move L, T, W2, H2
            
        Case 0  '-- Adjust BASIC Viewer Size
            TT = 930: HH = h - 1100: frBOpts.Visible = True
            If lblBView.Caption = ">>" Then TT = 390: HH = h - 440: frBOpts.Visible = False
            frBasic.Visible = Flag
            frBasic.Move L, T, W2, H2
            lstBAS.Move 105, TT, W3, HH
            'lstBAS.Height = H3 - 500
    
        Case 1  '-- Adjust SEQ Viewer Size
            frSEQ.Visible = Flag
            frSEQ.Move L, T, W2, H2
            lstSEQ.Width = W3
            lstSEQ.Height = H3
        
        Case 2  '-- Adjust BIN Viewer Size
            frBIN.Visible = Flag
            frBIN.Move L, T, W2, H2
            lstBIN.Width = W3
            lstBIN.Height = H3
            
        Case 3  '-- Adjust ChrSet Viewer Size
            frFont.Visible = Flag
            frFont.Move L, T, W2, H2
    
        Case 4  '-- Adjust ML Viewer Size
            frML.Visible = Flag
            frML.Move L, T, W2, H2
                        
            If ShowTables = False Then
                lblShw.Caption = ">>"
                lstML.Move 120, 600, W - 210, H3
            Else
                lblShw.Caption = "<<"
                
                If W < 4500 Then W = 4500
                LL = 60: TT = 1320: WW = 3825: HH = H3 - TT - 60: W4 = WW - 120
                
                frTView.Move 120, 520, WW, H3
                frMLSettings.Move LL, 800, W4, HH + 480         'The Settings Frame
                frTrace.Move LL, 800, W4, HH + 480              'The Tracer frame
                
                lstEntryPt.Move LL, TT, W4, HH                  'The Entry Points list
                lstSYM.Move LL, TT, W4, HH                      'The Symbols list
                lstDT.Move LL, TT, W4, HH                       'The Data Tables list
                lstULabels.Move LL, TT, W4, HH                  'The Generated Labels list
                lstCmnt.Move LL, TT, W4, HH                     'The Comment list
                lstLabels.Move LL, TT, W4, HH                   'The Labels list
                lstJSR.Move LL, TT, W4, HH                      'The External JSR list
                lstML.Move WW + 240, 600, W - WW - 330, H3      'The output list
                
                lstEP.Height = HH                               'The Tracer Entry Point List

                DrawMLTabs
            End If
    
            frTView.Visible = ShowTables

        Case 5  '-- Adjust IMG Viewer Size
            frBMP.Visible = Flag
            frBMP.Move L, T, W2, H2
            
    End Select
    DoEvents
    
End Sub

'---- Adjust Dual-View Split Sizing
Private Sub lblSSize_Click(Index As Integer)
    SetSplit Index, False   'Normal step size
End Sub
Private Sub lblSSize_DblClick(Index As Integer)
    SetSplit Index, True    'Doubles the step size when user clicks too fast and generates Double-click
End Sub

'---- Adjust Dual-View Split proportions
Private Sub SetSplit(ByVal Index As Integer, ByVal Flag As Boolean)
    Dim n As Integer
    
    n = 5: If Flag = True Then n = 10 'Step Size
    
    Select Case Index
        Case 0: SplitSize = SplitSize - n: If SplitSize < 20 Then SplitSize = 20    'Move split LEFT
        Case 1: SplitSize = SplitSize + n: If SplitSize > 80 Then SplitSize = 80    'Move split RIGHT
        Case 2: SplitSize = 50                                                      'Return to MIDDLE
    End Select
    DrawVLayout
End Sub

'---- View Tab was clicked
Private Sub lblView_Click(Index As Integer)
    SelectNewTab Index
End Sub

'----- ML Viewer Project/Table Buttons
Private Sub lblTView_Click(Index As Integer)
    MLTabNum = Index
    DrawMLTabs
End Sub

'---- Draw ML Viewer Side-panel elements
Private Sub DrawMLTabs()
    Dim i As Integer, VV As Boolean, V2 As Boolean
    
    VV = True: If (MLTabNum < 2) Or (MLTabNum > 6) Then VV = False
    V2 = VV: If (MLTabNum > 6) Then V2 = True
    
    cmdSymLoad.Visible = VV                 'Show or Hide Symbol buttons
    cmdSymSave.Visible = VV
    cmdSymAdd.Visible = VV
    cmdSymDel.Visible = VV
    cmdSYMGoto.Visible = V2
    
    frMLSettings.Visible = (MLTabNum = 0)   'Makes visible if MLTabNum matches View mode#
    frTrace.Visible = (MLTabNum = 1)
    lstEntryPt.Visible = (MLTabNum = 2)
    lstSYM.Visible = (MLTabNum = 3)
    lstDT.Visible = (MLTabNum = 4)
    lstULabels.Visible = (MLTabNum = 5)
    lstCmnt.Visible = (MLTabNum = 6)
    lstLabels.Visible = (MLTabNum = 7): cmdRemDupLbls.Visible = (MLTabNum = 7)
    lstJSR.Visible = (MLTabNum = 8):    cmdRemDupJSR.Visible = (MLTabNum = 8)
    
        
    For i = 0 To 8
        lblTView(i).ForeColor = vbBlack: lblTView(i).Font.Bold = False  'De-select all View buttons
    Next i
    
    lblTView(MLTabNum).ForeColor = vbWhite                              'Hi-light the currently Selected View button
    lblTView(MLTabNum).Font.Bold = True
    DoEvents

End Sub

'---- Lock the current view
Private Sub cbLockView_Click()
    LockV1 = ViewMode
    LockV2 = ViewMode2
End Sub

'---- Set the side select indicator
Private Sub lblSelect_Click()
    If lblSelect.Caption = ">" Then
        lblSelect.Caption = "<"
    Else
        lblSelect.Caption = ">"
    End If
End Sub

'---- Toggle Single or Dual-View Mode
Private Sub lblSplit_Click()
    If lblSplit.Caption = "+" Then
        lblSplit.Caption = "-"
        lblSelect.Visible = True
        SplitMode = True
    Else
        lblSplit.Caption = "+"
        lblSelect.Visible = False
        SplitMode = False
    End If
    DrawVLayout

End Sub

'============
'BASIC Viewer
'============
Sub BASView()
    Dim pLo As Integer, pHi As Integer                  'Program Line Links
    Dim lLo As Integer, lHi As Integer, LNum As Long    'Line numbers
    Dim i As Integer                                    'bufer position counter
    Dim First As Boolean, Quote As Boolean
    Dim C As Integer, C2 As Integer                     'character value
    Dim Ch As String                                    'character string
    Dim Tmp As String
    Dim RevText As Boolean, oneLine As Boolean, ExpFlag As Boolean, UCFlag As Boolean
    Dim BGuess As String
    Dim UnK As String                                   'Unknown Token string
    Dim Pad As String, TLine As String
    
    Me.Show: DoEvents
    
    If Token(0) = "" Then LoadTokens                    'Load Tokens if first run
    
    UnK = "{unknown}"
    RevText = (cbRev.value = 1)                         'Reverse text case
    ExpFlag = (cbExp.value = 1)                         'Expand special characters
    UCFlag = (cbUC.value = 1)                           'Uppercase special characters
    oneLine = (cbOneLine.value = 1)                     'One statement per line mode
    Pad = "": If cbPad.value = 1 Then Pad = " "         'Padding of tokens
    
    Mode = cboMode.ListIndex 'Basic Mode dropdown
    lblGuess.Caption = ""
    
    '-- Set Font Option
    If cbUseFont.value = 1 Then
        lstBAS.Font.Size = 10: lstBAS.Font.Name = "C64 User Mono": lstBAS.Font.Size = 10  'Try C64/Style font
    Else
        lstBAS.Font.Size = 10: lstBAS.Font.Name = "courier new": lstBAS.Font.Size = 10
    End If
    
    i = 1 'position in bufer
    If VLen < 2 Then Exit Sub
    
    lblLoadAdr.Caption = "$" & MyHex(VLA, 4)
    
    If Mode = 0 Then
        Select Case VLA
            Case 3: Mode = 2: BGuess = "CBM2"
            Case 1024, 1025: Mode = 1: BGuess = "PET"
            Case 2049: Mode = 1: BGuess = "C64"
            Case 4097, 4609: Mode = 1: BGuess = "Vic20"
            Case 4096, 8192: Mode = 3: BGuess = "C16/Plus4"
            Case 7169: Mode = 4: BGuess = "C128 Basic 7"
            Case Else: Mode = 1: BGuess = "Unknown"
        End Select
        lblGuess.Caption = BGuess
    End If
    
    lstBAS.Clear
    
    i = 1 'Start of BASIC DATA

    Do
        If i > VLen Then Exit Do
        
        pLo = Asc(Mid(VBuf, i, 1)):     If i + 1 > VLen Then Exit Do
        pHi = Asc(Mid(VBuf, i + 1, 1)): If (pHi + pLo) = 0 Then Exit Do 'program link=0 means end

        If (i + 3) > VLen Then Exit Do
        lLo = Asc(Mid(VBuf, i + 2, 1))
        lHi = Asc(Mid(VBuf, i + 3, 1))
        LNum = lHi * 256! + lLo                                         'Line number
        TLine = Format(LNum) & " "                                      'Text of entire line
        i = i + 4

        Quote = False                                                   'Flags

        Do
            C = Asc(Mid(VBuf, i, 1)): i = i + 1: Tmp = ""

            If (i >= VLen) Then Exit Do                                 'End of file
            If (C = 0) Then
                lstBAS.AddItem TLine                                    'NUL=End of line. Add it to the listbox
                Exit Do
            End If

            If (Quote = True) Or (C < 128) Then
                'Handle Non-Tokens or Characters inside Quotes
                Select Case C
                    Case 1 To 31                                        'Special keys (curB0Hr etc)
                         If ExpFlag = True Then
                                 Tmp = Token(297 + C - 1)
                                If UCFlag Then Tmp = UCase(Tmp)
                        End If
                        
                    Case 32, 160
                        Tmp = " "                          'Space
                    
                    Case 34
                        Tmp = Qu: Quote = Not Quote        'Quote
                    
                    Case 33 To 64
                        Tmp = Chr(C)
                        If Tmp = ":" Then
                            If oneLine = True Then Tmp = "": lstBAS.AddItem TLine: TLine = Space$(Len(Format(LNum)) + 1)
                        End If
                        
                    Case 65 To 90
                        If RevText Then C = Reverse(C)
                        Tmp = Chr(C)
                    
                    Case 97 To 122
                        If RevText Then C = Reverse(C)
                        Tmp = Chr(C)
                        
                    Case 129 To 159                                     'Special keys (colours,curB0Hr etc)
                        If ExpFlag = True Then
                            Tmp = Token(328 + C - 129)
                            If UCFlag Then Tmp = UCase(Tmp)
                        End If
                        
                    Case 193 To 218 'a to z
                        C = C - 96: If RevText Then C = Reverse(C)
                        Tmp = Chr(C)
                        
                    Case Else
                        Tmp = "{" & Hex(C) & "}"                        'Hex code for Graphic character
                        
                End Select
                TLine = TLine & Tmp
            Else
                '-----------------Convert to Tokens
                Select Case Mode
                    Case 1 '-- BASIC 1/2
                        Select Case C
                            Case 128 To 203, 255: Tmp = Token(C - 128) 'Common Tokens
                            Case 254 'Expansion C64 Tokens
                                C2 = Asc(Mid(VBuf, i, 1)): i = i + 1    'Get second Token byte
                                If (C2 > 127) And (C2 < 159) Then Tmp = Token(266 + C2 - 128): lblGuess.Caption = "C64 Exp"
                        End Select

                    Case 2 '-- BASIC 4/4+
                        Select Case C
                            Case 128 To 203, 255: Tmp = Token(C - 128)  'Common Tokens
                            Case 204 To 232: Tmp = Token(128 + C - 204) 'Basic4/4+ Tokens
                        End Select
                        
                    Case 3 '-- BASIC 3.5
                            Tmp = Token(C - 128) 'Common Tokens/Basic3.5

                    Case 4 '-- BASIC 7
                        Select Case C
                            Case 128 To 205, 207 To 253, 255: Tmp = Token(C - 128) 'Common Tokens/Basic3.5
                            Case 206 'CE Tokens; CE02 to CE0A
                                C2 = Asc(Mid(buf, i, 1)): i = i + 1
                                If C2 > 1 And C2 < 11 Then Tmp = Token(194 + C2 - 2)
                            Case 254 'FE Tokens; FE02 to FE26
                                C2 = Asc(Mid(buf, i, 1)): i = i + 1
                                If C2 > 1 And C2 < 39 Then Tmp = Token(157 + C2 - 2)
                        End Select

                    Case 5 '-- BASIC 10
                       Select Case C
                            Case 128 To 205, 207 To 253, 255: Tmp = Token(C - 128) 'Common Tokens/Basic3.5
                            Case 206 'CE Tokens; CE02 to CE0A
                                C2 = Asc(Mid(VBuf, i, 1)): i = i + 1
                                If C2 > 1 And C2 < 11 Then Tmp = Token(194 + C2 - 2)
                            Case 254 'FE Tokens; FE02 to FE3D
                                C2 = Asc(Mid(VBuf, i, 1)): i = i + 1
                                If C2 > 1 And C2 < 64 Then Tmp = Token(206 + C2 - 2)
                        End Select
                End Select
                
                If Tmp = "" Then Tmp = UnK
                TLine = TLine & Tmp & Pad
            End If
        Loop
    Loop
    
    If i < (VLen - 1) Then
        lstBAS.AddItem " "
        lstBAS.AddItem ">>>> NOTE: There are " & Format(VLen - i - 1) & " additional bytes following BASIC end!"
    End If

End Sub

'---- Toggle Options pane
Private Sub lblBView_Click()
    If lblBView.Caption = ">>" Then
        lblBView.Caption = "<<"
    Else
        lblBView.Caption = ">>"
    End If
    DrawVLayout
End Sub

'---- Save Listing to File
Private Sub cmdSave_Click()
    Dim FIO As Integer, Filename As String
    
    Filename = FileOpenSave(FileBase(LastFile), 1, 5, "Save Listing as Text")
    If Filename = "" Then Exit Sub
    
    FIO = FreeFile
    Open Filename For Output As FIO
    For j = 0 To lstBAS.ListCount - 1
        Print #FIO, lstBAS.List(j)
    Next
    Close FIO
    ChDir Exepath
NoFile:

End Sub

'---- Copy BASIC listing to clipboard
Private Sub cmdCpyClip_Click()
    Dim j As Integer, Tmp As String
    
    For j = 0 To lstBAS.ListCount - 1
        Tmp = Tmp & lstBAS.List(j) & vbCrLf
    Next j
    
    Clipboard.Clear
    Clipboard.SetText Tmp

End Sub

'---- Set BAS background colour
Private Sub lblBASTheme_Click(Index As Integer)
    
    frmColourPicker.Show vbModal
    If PickedColour < 0 Then Exit Sub
    
    lblBASTheme(Index).BackColor = PickedColour
    
    Select Case Index
        Case 0: lstBAS.ForeColor = PickedColour
        Case 1: lstBAS.BackColor = PickedColour
    End Select
    
End Sub

' Load Token strings into array
' Offsets for token groups
' 0   ;--COMMON TOKENS;basic1/2 (PET,VIC,C64)
' 75  ;--BASIC 3.5/7/10 (single byte tokens)
' 127 ;--BASIC4 (PET)
' 142 ;--BASIC 4+ (CBM2)
' 156 ;--BASIC7-fe (double-byte tokens)
' 193 ;--BASIC7-ce (double-byte tokens) Shared with BASIC 10
' 202 ;--BASIC10 (single differences) 'e3-e5;these differ from v7
' 205 ;--BASIC10-fe (double-byte tokens)
' 265 ;--C64 EXPANSION
' 296 ;--Quotemode strings
' 327 ;--Keys 129-159
' 358 ;DONE!
'
Sub LoadTokens()
    Dim Filename As String, FIO As Integer
    Dim Tmp As String, C As Integer

    C = 0
    
    Filename = AddSlash(App.Path) & "tokens.dat"
    If Exists(Filename) = False Then MsgBox "Can't load Token file!": Exit Sub
    
    FIO = FreeFile: Open Filename For Input As FIO
    
    Do
        If EOF(FIO) Then Exit Do
        Line Input #FIO, Tmp
        If Left(Tmp, 1) <> ";" Then Token(C) = Tmp: C = C + 1
    Loop
    Close FIO
End Sub

'=============================
' FONT VIEWER
'=============================
Public Sub FONTView()
    Dim i As Integer
    
    If cboTheme.ListIndex = -1 Then cboTheme.ListIndex = 0
    If cbMC.value = vbChecked Then CreatePixels (True) Else CreatePixels (False)    'Create pixel multicolour or normal font pixels
    
    For i = 0 To 4
        lblZoom(i).BackColor = vbWhite
    Next i
    lblZoom(ChrZoom - 1).BackColor = &H80C0FF    'orange
    
    If optChrH(0).value = True Then
        picChr.Height = 1270
        ViewFont 8
    Else
        picChr.Height = 2490
        ViewFont 16
    End If
End Sub
 
 Public Sub ViewFont(ByVal FH As Integer)
    Dim j As Integer, K As Integer, X As Integer, Y As Integer, V As Integer, TopX As Integer, TopY As Integer
    Dim R As Integer, C As Integer, MaxR As Integer, MaxC As Integer, MaxH As Integer
    Dim CZ As Integer, RZ As Integer, PZ As Integer 'zoomed size
    Dim Offset As Long
    
    
    C = 0: R = 0: X = 0: Y = 0
    MaxR = 32                                               'Max Row was 16 - changed feb'2015
    TopX = 0: TopY = 0                                      'Top-Left Offset
    MaxC = 16: If cbFCols.value = vbChecked Then MaxC = 32  'How many characters wide?
    CW = 8: RW = FH                                         'Chr width
    PZ = CW * ChrZoom                                       'Scale factor for drawing one line of pixels
    
    Offset = Val(txtCSkip.Text): If Offset < 1 Then Offset = 1
    If Offset > 32767 Then Offset = 32767
    
    If cbBorder.value = vbChecked Then
        CW = CW + 1: RW = RW + 1
        TopX = ChrZoom: TopY = ChrZoom
    End If
    
    CZ = CW * ChrZoom: RZ = RW * ChrZoom                'Size of one character including borders
    FontH = FH                                          'Set for calculating chr when clicked
            
    picV.Width = (CZ * MaxC + TopY) * Screen.TwipsPerPixelX
    picV.Height = (RZ * MaxR + TopX) * Screen.TwipsPerPixelY
    picV.BackColor = lblTheme(2).BackColor
    picV.Cls
    picV.Visible = False
    DoEvents
    
    For j = Offset To VLen
        V = Asc(Mid(VBuf, j, 1))
        '----paintpicture {srceimg},destX,destY,destW,destH ,srcX,srcY,srcW,srcH,mode
        picV.PaintPicture Pix.Image, TopX + C * CZ, TopY + R * RZ + Y * ChrZoom, PZ, ChrZoom, 0, V, 8, 1 'blit the pixel representation to the view window
        Y = Y + 1
        If Y = FH Then Y = Y - FH: C = C + 1: If C >= MaxC Then C = 0: R = R + 1
        If R > MaxR Then Exit For
    Next j
    If R < MaxR Then picV.Height = (RZ * R + TopX) * Screen.TwipsPerPixelY
    
    lblEndRange.Caption = "to" & Str(j)
    picV.Visible = True
    DoEvents
    
    ShowChr
    
End Sub

'==============
'Font View Subs
'==============

Private Sub cmdSaveCSet_Click()
    Dim Filename As String
    
    Filename = FileOpenSave(FileBase(VFileName), 1, 3, "Save as BMP")
    picV.Picture = picV.Image 'crop to visible
    If Filename <> "" Then SavePicture picV.Image, Filename

End Sub

'-- Toggle Multicolour mode
Private Sub cbMC_click()
    If cbMC.value = vbChecked Then
        lblTheme(3).Visible = True
        lblTheme(4).Visible = True
    Else
        lblTheme(3).Visible = False
        lblTheme(4).Visible = False
    End If
    FONTView
    
End Sub

'-- Change Zoom Factor
Private Sub lblZoom_Click(Index As Integer)
    ChrZoom = Index + 1
    FONTView 'draw character set
End Sub

'-- Set Colour Theme
Private Sub cboTheme_Click()
    Dim n As Integer, FG As Long, BG As Long, BO As Long
    
    n = cboTheme.ListIndex: If n < 0 Then n = 0
    BO = CBMColor(0)                                                    'assume black border
    Select Case n
        Case 0: FG = CBMColor(14): BG = CBMColor(6)                     '-- C64
        Case 1: FG = CBMColor(6): BG = CBMColor(1)                      '-- SX-64
        Case 2: FG = CBMColor(6): BG = CBMColor(1)                      '-- VIC-20
        Case 3: FG = CBMColor(0): BG = CBMColor(1): BO = CBMColor(12)   '-- TED
        Case 4: FG = CBMColor(1): BG = CBMColor(0): BO = CBMColor(12)   '-- PET White
        Case 5: FG = CBMColor(5): BG = CBMColor(0): BO = CBMColor(12)   '-- PET Green
        Case 6: FG = CBMColor(7): BG = CBMColor(0): BO = CBMColor(12)   '-- PET Amber
    End Select
    
    lblTheme(0).BackColor = FG: lblTheme(1).BackColor = BG: lblTheme(2).BackColor = BO
    DoEvents
    FONTView

End Sub

Private Sub Label4_Click()
    lstLabels.Visible = Not lstLabels.Visible
End Sub

Private Sub lblTheme_Click(Index As Integer)
    
    frmColourPicker.Show vbModal
    If PickedColour >= 0 Then lblTheme(Index).BackColor = PickedColour: FONTView

End Sub

Private Sub lstLabels_DblClick()
    Dim Tmp As String, Tmp2 As String
    
    Tmp = lstLabels.List(lstLabels.ListIndex) & ",name,-"              'Make default text entry string
    Tmp2 = InputBox("HHHH,LABELNAME,DESCRIPTION", "Add Label from [GEN] label", Tmp)
    If Len(Tmp2) > 12 Then lstULabels.AddItem Tmp2: MLReViewA

End Sub

Private Sub txtCSkip_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then FONTView
End Sub

Public Sub CreatePixels(ByVal MultiFlag As Boolean)
    Dim j As Integer, K As Integer, Power(7) As Integer, CI As Integer
    Dim MC(3) As Long 'Array to hold multicolour values
    
    MC(0) = lblTheme(1).BackColor   'Background colour
    MC(1) = lblTheme(3).BackColor   'Register colour #1
    MC(2) = lblTheme(4).BackColor   'Register colour #2
    MC(3) = lblTheme(0).BackColor   'Foreground Colour
    
    For j = 0 To 7: Power(j) = 2 ^ j: Next  'Init Powers of 2 array
    
    Pix.ForeColor = lblTheme(0).BackColor
    Pix.BackColor = lblTheme(1).BackColor
    Pix.Cls
    
    If MultiFlag = True Then
        '-- Create a 4-colour bitmap with pixels to match binary representation of pixlex pairs (row=value,cols 0 to 7=pixel)
        For j = 0 To 255
            For K = 0 To 7 Step 2
                CI = 0                                      'Colour Index
                If (j And Power(K)) Then CI = CI + 2        'Check first bit
                If (j And Power(K + 1)) Then CI = CI + 1    'Check second bit
                Pix.ForeColor = MC(CI)                      'Set the colour of the pixel to draw
                Pix.PSet (7 - K, j)                         'Set the first pixel
                Pix.PSet (6 - K, j)                         'Set the second pixel
            Next K
        Next j
    Else
        '-- Create a 2-colour bitmap with pixels to match binary representation of value (row=value,cols 0 to 7=pixel)
        For j = 0 To 255
            For K = 0 To 7
                If (j And Power(K)) Then Pix.PSet (7 - K, j)
            Next K
        Next j
    End If
End Sub

'---- Jump to Next Character
Private Sub cmdCSNxt_Click()
    SelChr = SelChr + 1: If SelChr > 255 Then SelChr = 255
    ShowChr
End Sub

'---- Jump to Previous Character
Private Sub cmdCSPrev_Click()
    SelChr = SelChr - 1: If SelChr < 0 Then SelChr = 0
    ShowChr
End Sub

'---- Show the Selected Character
Public Sub ShowChr()
    Dim R As Integer, C As Integer, X As Integer, Y As Integer, XYOff As Integer
    Dim RW As Integer, CW As Integer, CMax As Integer
    Dim SetNum As Integer, ChrNum As Integer
    
    CMax = 16: If cbFCols.value = vbChecked Then CMax = 32                          'Max# chars per line
    RW = FontH: CW = 8: XYOff = 0                                                   'Pixels in one char
    If cbBorder.value = vbChecked Then RW = RW + 1: CW = CW + 1: XYOff = ChrZoom    'Adjust for border
    
    SetNum = SelChr \ 128: ChrNum = SelChr Mod 128                                  'Set based on 128 char font
    
    '-- Show Info
    lblChrNum.Caption = Format(SelChr, "000")
    lblChrX.Caption = "Set# " & Format(SetNum) & Cr & " Chr# " & Format(ChrNum) & " ($" & MyHex(ChrNum, 2) & ")"
    
    '-- Set the Selected chr colours to match theme
    picChr.BackColor = lblTheme(1).BackColor
    picChr.ForeColor = lblTheme(2).BackColor
    picChr.Cls
    
    '-- Calc position
    R = Int(SelChr / CMax)
    C = SelChr - R * CMax
    X = C * CW * ChrZoom + XYOff: Y = R * RW * ChrZoom + XYOff
        
    If picV.Height >= FontH * ChrZoom * 15 Then
        picChr.PaintPicture picV.Image, 0, 0, 80, 10 * FontH, X, Y, 8 * ChrZoom, FontH * ChrZoom    'Draw the Character
        
        If cbBorder.value = vbChecked Then
            For i = 0 To 16: picChr.Line (0, i * 10)-Step(160, 0): Next i   'Draw Horizontal Lines
            For i = 0 To 8: picChr.Line (i * 10, 0)-Step(0, 160): Next i    'Draw Vertical Lines
        End If
    End If
    
End Sub

'---- Select a character
Private Sub picV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim R As Integer, C As Integer, RW As Integer, CW As Integer, CMax As Integer
    
    CMax = 16: If cbFCols.value = vbChecked Then CMax = 32
    RW = FontH: CW = 8
    If cbBorder.value = vbChecked Then RW = RW + 1: CW = CW + 1
        
    R = Int(Y / (RW * ChrZoom)): If R > 32 Then R = 32
    C = Int(X / (CW * ChrZoom)): If C > CMax Then C = CMax
    SelChr = R * CMax + C
    ShowChr
End Sub

'---- Change Skip-bytes
Private Sub cmdSB_Click(Index As Integer)
    Dim Offset As Integer
    
    Offset = Val(txtCSkip.Text)
    Select Case Index
        Case 0: Offset = Offset - 256
        Case 1: Offset = Offset - 8
        Case 2: Offset = Offset - 1
        Case 3: Offset = Offset + 1
        Case 4: Offset = Offset + 8
        Case 5: Offset = Offset + 256
    End Select
    If Offset < 0 Then Offset = 0
    txtCSkip.Text = Format(Offset)
    FONTView
End Sub

'==================================
' ASM Machine Language Disassembler
'==================================
Sub MLView()
    Dim GoodFlag As Boolean
    Dim j As Integer
    Dim C As Integer                                                    'Counter
    Dim Tmp As String, TmpB As String                                   'Temp strings
    
    Dim B0A As Long, B1A As Long, B2A As Long                           'Byte 0-2 ASCII values
    Dim B0H As String, B0C As String                                    'Byte 0 hex, chr
    
    Dim SH As String, SL As String, SHL As String                       'Address strings
    Dim Lo As Integer, Hi As Integer
    Dim Address As Double, TAddress As Double                           'Address and Target Address
    Dim JAddress As String, RAddress As String                          'Jump addresses
    Dim StartAddress As String, EndAddress As String                    'Range Addresses
    
    Dim OpLen As String                                                 'Opcode Length
    Dim NM As String, MD As Integer, NB As Integer                      'Opcode parameters
    
    Dim T0 As String, T1 As String, T2 As String                        'ASM Output variables
    Dim T3 As String, T4 As String, T5 As String
    Dim OutFmt As Integer, ALabel As String, UComment As String
    Dim Padd As String
    
    Dim LNum As Long, LInc As Integer                                   'Line Numbers
    Dim a As Integer, p As Integer
    
    Dim DTMode As Boolean, DTCount As Integer, DTType As String         'Data Table variables
    Dim DTCountMax As Integer, DTMax As Integer, DTPos As Integer       'Data Table variables
    Dim DTStart As Long, DTEnd As Long, DTAscMode As Integer            'Data Table variables
    Dim DTComment As String, DTAddress As String, DTOutStr As String    'Data Table variables
    
    Dim Pass As Integer
    Dim RTSOption As Boolean, SymComment As Boolean, DivLen As Integer  'options
    
    Padd = Space(50)            'spaces for padding byte lists
    LInc = 10                   'Line# Increment
        
    '---- Options
    RTSOption = False: If cbSpaceRTS.value = vbChecked Then RTSOption = True
    SymComment = False: If cbIncSym.value = vbChecked Then SymComment = True
    DivLen = Val(txtDivLen.Text)
     
    
    '============================================
    ' Load Support Files and Config settings etc
    '============================================
    
    '---- Load ML Config File
    If MLCFlag = False Then LoadMLConfig
    If MLCFlag = False Then MyMsg "ML Config file is missing!": Exit Sub
    
    '---- Load project file that has same base name if present
    If ProjFlag = False Then
        Tmp = FileBase(VFileName) & ".asm-proj"
        If Exists(Tmp) = True Then
            ProjFlag = True
            LoadProjFile Tmp
            ShowTables = True
        End If
    End If
        
    '---- Read Opcode file into array
    If OpCodeFlag = False Then
        If cboCPU.ListCount > 0 Then
            Tmp = ExeDir & cboCPUFile.List(0)
            If Exists(Tmp) = False Then MsgBox "Missing file:" & Tmp, vbCritical: Exit Sub
            LoadOpcodes Tmp
        End If
    End If
        
    If ShowTables = True Then DrawVLayout
    
    '---- Set initial modes etc
    DTMax = lstDT.ListCount
    OutFmt = cboMLFmt.ListIndex
    
    lstML.Visible = False
    lstLabels.Clear                                 'Clear [GEN] labels list
    lstJSR.Clear                                    'Clear [JSR] list
    lblGood.BackColor = vbYellow: GoodFlag = True   'Set status box colour
    
    DoEvents
    
    '=========================================================================================
    ' This is the PASS loop. In PASS 1 labels are generated. In PASS 2 the output is generated
    '=========================================================================================
    
    For Pass = 1 To 2
        C = StartC                                          'Start position 1 or 3 depending if load address is skipped
        lblEA.Caption = "Disassembling... PASS#" & Str(Pass)
        lblEA.BackColor = vbYellow
        DoEvents
        
        lstML.Clear                                         'Clear the output
        DTMode = False
        DTCount = 0: DTPos = -1: DTStart = 0: DTEnd = 0     'Reset Data Table pointer
        LNum = 1000
        C = 1
        Address = VLA: If cbLA.value = vbChecked Then Address = MyDec(txtLA.Text)
        txtLA.Text = MyHex(Address, 4)
        StartAddress = MyHex(Address, 4)
        EndAddress = MyHex(Address + VLen - 3, 4)
        
        '---- PASS 2 - Add Equates
        
        If (Pass = 2) And (cbEquates.value = vbChecked) Then
            If OutFmt = 2 Then
                lstML.AddItem Format(LNum) & " ; Disassembly of: " & FileNameOnly(VName): LNum = LNum + LInc
                lstML.AddItem Format(LNum) & " ;": LNum = LNum + LInc
                lstML.AddItem Format(LNum) & " ; ----- Equates": LNum = LNum + LInc
                lstML.AddItem Format(LNum) & " ;": LNum = LNum + LInc
            Else
                lstML.AddItem "; Disassembly of: " & FileNameOnly(VName)
                lstML.AddItem ";"
                lstML.AddItem "; ---- Equates"
                lstML.AddItem ";"
            End If
            
            For j = 0 To lstSYM.ListCount - 1
                If lstSYM.Selected(j) = True Then
                    Tmp = lstSYM.List(j)
                    T1 = "": If OutFmt = 2 Then T1 = Format(LNum) & " ": LNum = LNum + LInc
                    lstML.AddItem T1 & GetField(Tmp, 2) & " = " & GetField(Tmp, 1) & "   ;" & GetField(Tmp, 3)
                End If
            Next j
            If OutFmt = 2 Then lstML.AddItem Format(LNum) & " ;" Else lstML.AddItem ";"
        End If
        
        '---- PASS 2 - Add Code Origin
        
        If (Pass = 2) Then
            If OutFmt = 2 Then
                lstML.AddItem Format(LNum) & " " & DOTORG & "$" & StartAddress: LNum = LNum + LInc
                lstML.AddItem Format(LNum) & " ;": LNum = LNum + LInc
                lstML.AddItem Format(LNum) & " ; ---- Code": LNum = LNum + LInc
                lstML.AddItem Format(LNum) & " ;": LNum = LNum + LInc
            Else
                lstML.AddItem DOTORG & "$" & StartAddress
                lstML.AddItem ";"
                lstML.AddItem "; ---- Code"
                lstML.AddItem ";"
            End If
        End If
                
        '====================================
        ' Process!
        '====================================
        
        Do
            '---- Process the Address
            T0 = MyHex(Address, 4)                                  'Current Hex Address XXXX
            T1 = T0 & ": "                                          'Current Hex Address Label XXXX:
            B0C = Mid(VBuf, C, 1)                                   'Byte 0 Char
            B0A = Asc(B0C)                                          'Byte 0 Value
            B0H = MyHex(B0A, 2)                                     'Byte 0 Hex
            T2 = B0H & "        "                                   'Default to opcode byte and spacing
            T4 = ""                                                 'Formatted code
            T5 = ""                                                 'Comment area
            LastComment = ""                                        'Clear Last Comment
            SH = ""                                                 'HI
            SL = ""                                                 'LO
            SHL = ""                                                'Word
            DTMode = False                                          'Clear Data Table Mode
                        
            '---- PASS 2 only. Handle Symbols, Labels, and Comments
            
            If Pass = 2 Then
                '---- Handle Comments
                UComment = FindComment(T0)                          'Check for a comment here
                If UComment > "" Then
                    TmpB = UCase(Left(UComment, 1))                 'Check comment type (I,S or divider)
                    UComment = Mid(UComment, 3)                     'Strip away comment type
                    If TmpB <> "I" Then
                        '---- add standalone comment
                        Select Case OutFmt
                            Case 2
                                If TmpB <> "S" Then lstML.AddItem Format(LNum) & " ; " & String(DivLen, TmpB): LNum = LNum + LInc
                                If UComment > "" Then lstML.AddItem Format(LNum) & " ; " & UComment: LNum = LNum + LInc
                                If UComment > "" Then If TmpB <> "S" Then lstML.AddItem Format(LNum) & " ; " & String(DivLen, TmpB): LNum = LNum + LInc
                            Case Else
                                If TmpB <> "S" Then lstML.AddItem ";" & String(DivLen, TmpB): LNum = LNum + LInc
                                If UComment > "" Then lstML.AddItem "; " & UComment
                                If UComment > "" Then If TmpB <> "S" Then lstML.AddItem ";" & String(DivLen, TmpB): LNum = LNum + LInc
                        End Select
                        UComment = "" 'clear it since it's been used. if type is "i" (inline) then we'll add it later
                    End If
                End If
                
                '---- PASS2 - Handle Labels
                
                Tmp = FindUL(T0)    'Find User Label or Generated Label
                If Tmp > "" Then
                    ALabel = Tmp & ":"
                
                    Select Case OutFmt
                        Case 0, 1, 3
                            If cbLabelBlanks.value = vbChecked Then lstML.AddItem ";"
                            lstML.AddItem ALabel
                        Case 2
                            If cbLabelBlanks.value = vbChecked Then lstML.AddItem Format(LNum) & " ;": LNum = LNum + LInc
                            lstML.AddItem Format(LNum) & " " & ALabel: LNum = LNum + LInc
                        Case 4 'label cmd param
                    End Select
                End If
            End If
            
            '===================================================
            ' PASS 1 and 2 - Handle Stepping through Data Tables
            '===================================================
        
            '-- Check if New Data Table Range
            If DTStart = 0 Then
                '---- Get the Next SELECTED Data Table Entry
                Do
                    DTPos = DTPos + 1                               'Go to next position
                    If DTPos >= DTMax Then Exit Do
                    If lstDT.Selected(DTPos) = True Then Exit Do    'If it is selected then ust it
                Loop
                
                If DTPos < DTMax Then
                    '---- Look at the current range entry.       Format: HHHH,HHHH,T,Comment
                    Tmp = lstDT.List(DTPos)                     'Get the line from the list
                    DTStart = MyDec(Mid(Tmp, 1, 4))             'Get Range Start
                    DTEnd = MyDec(Mid(Tmp, 6, 4))               'Get Range End
                    DTType = UCase(Mid(Tmp, 11, 1))             'Get Type (Asc,Byte,Word,Vector,RVector)
                    If Pass = 2 Then DTComment = Mid(Tmp, 13)   'Get Comment
                Else
                    '---- No more ranges, B0H set to highest byte $FFFF
                    DTStart = CLng(65536): DTEnd = CLng(65536): DTComment = "end"
                End If
            End If
                            
            If Address >= DTStart Then
                MD = 0
                '---- The address could be a valid table range
                '     Check if Table also has a symbol or label. If not add one
                If Address = DTStart Then
                    '---- This is the first byte of the range
                    DTAscMode = 0 'Reset Asc mode
                    '---- It should have a label
                    Tmp = FindSym(DTStart)

                    If Tmp = "" Then
                       Tmp = FindLabel(DTStart)
                       If Tmp = "" Then
                            lstLabels.AddItem MyHex(DTStart, 4) 'make it a label
                       End If
                    End If
                End If
                
                If Address <= DTEnd Then
                    '---- We are inside a data range!
                    ' In PASS 1 we can generally skip over everything.. except for
                    ' "V" and "R" modes, which need to add labels for the target addresses.

                    DTMode = True 'Set the Flag
                    T2 = "          "
                    
                    If DTCount = 0 Then DTAddress = T1: DTOutStr = "" 'Initialize line string
                    If (DTCount > 0) And ((DTType <> "S") And (DTType <> "T")) Then DTOutStr = DTOutStr & ","    'Add a comma between entries unless String mode
                    
                    Select Case DTType
                        Case "S", "T" '---- String/Text Directive
                            If Pass = 2 Then
                                '---- We now need to build the output string, handling printable and non-printable bytes
                                DTCountMax = 99
                                T3 = DOTTEXT
                                '---- DTAscMode: 0=initial state, 1=non-printable, 2=printable (inside quotes)
                                Select Case B0C
                                    Case Qu 'quote
                                        If DTAscMode = 2 Then DTOutStr = DTOutStr & Qu & "," 'finish off quote mode then comma
                                        DTOutStr = DTOutStr & "$22"
                                        DTAscMode = 1 'non-printable mode (hex values)
                                        
                                    Case " " To "z"
                                        If DTAscMode = 0 Then DTOutStr = DTOutStr & Qu       'quote
                                        If DTAscMode = 1 Then DTOutStr = DTOutStr & "," & Qu 'comma+quote
                                        DTOutStr = DTOutStr & B0C
                                        DTAscMode = 2 'printable mode (inside quotes)
                                        
                                    Case Else
                                        If DTAscMode = 2 Then DTOutStr = DTOutStr & Qu & "," 'end quote + comma
                                        If DTAscMode = 1 Then DTOutStr = DTOutStr & ","      'comma
                                        DTOutStr = DTOutStr & "$" & MyHex(B0A, 2)
                                        DTAscMode = 1 'non-printable mode (hex values)
                                End Select
                            End If
                            
                        Case "B", "H" '---- Byte Directive (Hex)
                            If Pass = 2 Then
                                T3 = DOTBYTE
                                DTCountMax = 8
                                DTOutStr = DTOutStr & "$" & B0H
                            End If
                            
                        Case "D"  '---- Byte Directive (Dec)
                            If Pass = 2 Then
                                T3 = DOTBYTE
                                DTCountMax = 8
                                DTOutStr = DTOutStr & B0A
                            End If
                            
                        Case "W"  '---- Word Directive (Hex)
                            If Pass = 2 Then
                                T3 = DOTWORD
                                DTCountMax = 6
                                Address = Address + 1: C = C + 1    'Increment address
                                B1A = Asc(Mid(VBuf, C, 1))           'Get next byte
                                SL = B0H                            'Lo Byte
                                SH = MyHex(B1A, 2)                  'HI Byte
                                DTOutStr = DTOutStr & "$" & SH & SL 'Add to output list
                            End If
                            
                        Case "V" '----Word, Vector address
                            '---- Take the next byte and generate an address.
                            DTCountMax = 6
                            Address = Address + 1: C = C + 1
                            B1A = Asc(Mid(VBuf, C, 1))           'Value of byte
                            TAddress = B1A * 256 + B0A          'Calculate Target Address (decimal)
                            JAddress = MyHex(TAddress, 4)       'make it a string
                            SHL = "$" & JAddress                'Make string for output

                            If Pass = 1 Then
                                If (JAddress >= StartAddress) And (JAddress <= EndAddress) Then
                                    lstLabels.AddItem JAddress  'target is inside code range so make it a label
                                End If
                            Else
                                '---- PASS 2
                                T3 = DOTWORD
                                Tmp = FindSL(JAddress)
                                If Tmp = "" Then Tmp = SHL
                                DTOutStr = DTOutStr & Tmp
                            End If
                            
                        Case "R" '--Word, RTS Vector address
                            '---- Take the next byte and generate an address
                            DTCountMax = 6
                            Address = Address + 1: C = C + 1
                            B1A = Asc(Mid(VBuf, C, 1))           'Value of byte
                            TAddress = B1A * 256 + B0A + 1      'Calculate Target Address (decimal) with offsett
                            JAddress = MyHex(TAddress, 4)       'Make it a string
                            SHL = "$" & JAddress                'Make string for output
                            
                            If Pass = 1 Then
                                If (JAddress >= StartAddress) And (JAddress <= EndAddress) Then
                                    lstLabels.AddItem JAddress  'Target is inside code range so make it a label
                                End If
                            Else
                                '---- PASS 2
                                T3 = DOTWORD
                                Tmp = FindSL(JAddress)
                                If Tmp = "" Then Tmp = SHL
                                DTOutStr = DTOutStr & Tmp & "-1"    'Add to output with "-1" offset
                            End If
                        
                        Case "X" '---- Hide the entire range
                            T3 = ""
                    
                    End Select
                    
                    C = C + 1
                    Address = Address + 1 'Increment address for each byte
                    DTCount = DTCount + 1 'Store up x bytes
                    
                    If Pass = 2 Then
                        If (DTCount >= DTCountMax) Or (Len(DTOutStr) >= 44) Or (Address > DTEnd) Or (C > VLen) Then
                            '---- We've done 'DTCountMax' entries, or we've reached the end of the table or file
                            If (DTType = "S") Or (DTType = "T") Then
                                '-- we need to finish off the string properly
                                If DTAscMode = 2 Then DTOutStr = DTOutStr & Qu 'add ending quote
                                DTAscMode = 0
                            End If

                            If T3 > "" Then
                                '----  padd DTOutStr here!
                                Tmp = Left(DTOutStr & Padd, 50)
                                '---- Add a line according to selected format
                                Select Case OutFmt
                                    Case 0: lstML.AddItem DTAddress & T2 & T3 & Tmp & " ;" & DTComment      'addr bb bb bb cmd param
                                    Case 1: lstML.AddItem DTAddress & T3 & Tmp & " ;" & DTComment           'addr cmd param
                                    Case 2: lstML.AddItem Format(LNum) & " " & T3 & Tmp & " ;" & DTComment     'nnnn cmd param
                                    Case 3: lstML.AddItem T2 & T3 & Tmp & " ;" & DTComment                  'cmd param
                                    Case 4:
                                        ALabel = Left(ALabel & Padd, 15)
                                        lstML.AddItem ALabel & T3 & Tmp & " ;" & DTComment                  'label cmd param
                                        ALabel = ""                                                         'blank it for multi-line tables
                                End Select
        
                                LNum = LNum + LInc
                            End If
                            T4 = ""
                            DTCount = 0
                        End If
                    End If
                    
                    '---- Check if we are finished with the current table
                    If Address > DTEnd Then
                        If RTSOption = True Then
                            Select Case OutFmt
                                Case 2: lstML.AddItem Format(LNum) & " ;": LNum = LNum + LInc
                                Case Else: lstML.AddItem " " 'add a blank line
                            End Select
                        End If
                        DTStart = 0
                    End If
                    
                End If
            End If
                        
            '===============================================
            ' If not in a Table then process as valid opcode
            '===============================================
            
            If DTMode = False Then
                NM = Left(OP(B0A), Len(OP(BOA)) - 1)                'Mneumonic string (eg: JSR or BBR0)
                If Pass = 2 Then If Left(NM, 1) = "?" Then GoodFlag = False    'Found an unknown opcode
                MD = Asc(Right(OP(B0A), 1)) - 96                    'Addressing mode (A-M,N-P)
                NB = Val(Mid(OpModeLen, MD, 1))                     'How many bytes for this opcode? (1 to 3)

                '---- All modes >2 use one or two-byte address
                If MD > 1 Then
                    '---- Opcode+Byte
                    If NB > 1 Then
                        If C + 1 <= VLen Then
                            B1A = Asc(Mid(VBuf, C + 1, 1))
                            SL = MyHex(B1A, 2): Mid(T2, 4, 2) = SL  'Set second byte
                        End If
                        
                        '---- Opcode+Word
                        If NB > 2 Then
                            If C + 2 <= VLen Then
                                B2A = Asc(Mid(VBuf, C + 2, 1))
                                SH = MyHex(B2A, 2): Mid(T2, 7, 2) = SH  'Set third byte
                            End If
                        Else
                            SH = "00"                               'Set third byte as $00 (zero page)
                        End If
                        
                        JAddress = SH & SL                          'Absolute Jump address
                        SHL = "$" & JAddress                        'Add the $ to HI string
                        SL = "$" & SL                               'Add the $ to LO string
                    End If
                    
                    '---- Now look up the address
                    If (MD > 2) And (NB > 1) Then
                        Tmp = FindSL(JAddress)                          'Look for a SYMBOL, ULABEL, or LABEL for this address
                        If Tmp > "" Then
                            SL = Tmp                                    'Substitute Symbol for single-byte address
                            SHL = Tmp                                   'Substitute Symbol for two-byte address
                            If (SymComment = True) And (LastComment > "") Then T5 = " ;" & LastComment
                        End If
                    End If
                End If
                                                
                '---- Handle JMP/JSR addresses
                
                Select Case B0H
                    Case "20", "4C", "6C"
                        If Pass = 1 Then
                            If (JAddress >= StartAddress) And (JAddress <= EndAddress) Then
                                lstLabels.AddItem JAddress  'target is inside code range so make it a label
                            Else
                                If B0H = "20" Then
                                    Tmp = FindSym(JAddress)
                                    lstJSR.AddItem JAddress & " " & LastComment
                                End If
                            End If
                        End If
                End Select

                '---- Handle Relative Branch
                If Pass = 1 And MD = 10 Then
                    If B1A > 127 Then B1A = B1A - 256       'Calculate backwards branch
                    RAddress = MyHex(Address + B1A + 2, 4)  'Make $HHHH string
                    lstLabels.AddItem RAddress              'Add to labels
                End If
                
                '---- PASS 2 - Build output string
                
                If Pass = 2 Then
                    T3 = UCase(NM)                          'The mneumonic string
                                        
                    '-- Handle Opcode Addressing Mode
                    
                    Select Case MD
                        Case 1: T4 = ""                     'a-Accumulator Adressing
                        Case 2: T4 = " #" & SL              'b-Immediate Addressing
                        Case 3: T4 = " " & SL               'c-Zero Page
                        Case 4: T4 = " " & SL & ",X"        'd-Indexed Zero page with X
                        Case 5: T4 = " " & SL & ",Y"        'e-Indexed Zero page with Y
                        Case 6: T4 = " " & SHL              'f-Absolute Addressing
                        Case 7: T4 = " " & SHL & ",X"       'g-Indexed Absolute with X
                        Case 8: T4 = " " & SHL & ",Y"       'h-Indexed Absolute with Y
                        Case 9                              'i-Implied
                        Case 10                             'j-Relative Addressing
                            If B1A > 127 Then B1A = B1A - 256                       'Calculate backwards branch
                            RAddress = MyHex(Address + B1A + 2, 4)                  'Make HHHH string
                            Tmp = FindSL(RAddress): If Tmp > "" Then T4 = " " & Tmp 'Lookup Relative Address in Symbols and Labels lists
                            T5 = ""
                        Case 11: T4 = " (" & SL & ",X)"     'k-Indexed Indirect Addressing with X
                        Case 12: T4 = " (" & SL & "),Y"     'l-Indexed Indirect Addressing with Y
                        Case 13: T4 = " (" & SHL & ")"      'm-Absolute Indirect
                        Case 14: T4 = " (" & SHL & ",X)"    'n-iax (65c02)
                        Case 15: T4 = " " & SL & "," & SH   'o-zpr (65c02) ***** need to convert SH to HHHH relative address
                        Case 16: T4 = " (" & SL & ")"       'p-izp (65c02)
                    End Select
                                    
                    '---- Handle inline comments
                    
                    If UComment > "" Then T5 = " ; " & UComment                     'Use user comment string
                                        
                    '---- Output line in specified format
                    
                    Select Case OutFmt
                        Case 0: lstML.AddItem T1 & T2 & T3 & T4 & T5                'addr bytes cmd param
                        Case 1: lstML.AddItem T1 & T3 & T4 & T5                     'addr cmd param
                        Case 2: lstML.AddItem Format(LNum) & " " & T3 & T4 & T5     'nnnn cmd param
                        Case 3: lstML.AddItem "          " & T3 & T4 & T5           'cmd param
                        Case 4:
                            ALabel = Left(ALabel & "               ", 15)
                            lstML.AddItem ALabel & T3 & T4 & T5                     'label cmd param
                    End Select
                                        
                    If MD = 9 Then
                        '-- Space after RTS/RTI option
                        If (T3 = "RTS") Or (T3 = "RTI") Then
                            If RTSOption = True Then
                                If OutFmt = 2 Then
                                    LNum = LNum + LInc                          'next line number
                                    lstML.AddItem Format(LNum) & " ;"           'add line number and semicolon
                                Else
                                    lstML.AddItem ";"                           'add a blank line after RTI or RTS
                                End If
                            End If
                        End If
                    End If
                End If
                                
                C = C + NB
                Address = Address + NB      'increment address according to bytes used by opcode
                LNum = LNum + LInc          'next line number
                ALabel = ""                 'clear out label
                '---- end of opcode mode
                DoEvents                    'added for long files
            End If
            
        Loop While C <= VLen
    Next Pass
    
    '---- Disassembly is complete
    
    lblEA.Caption = "Code from $" & StartAddress & " to $" & EndAddress & " (" & Format(C - 3) & " bytes)"
    lblEA.BackColor = &H8000000F
    
    lstML.Visible = True: DoEvents
    If lstML.Visible = True Then lstML.SetFocus
    If GoodFlag = True Then lblGood.BackColor = vbGreen
    
End Sub

Private Sub cmdTrace_Click()
    If lstEntryPt.ListCount = 0 Then MyMsg "You must add Entrypoints first!": Exit Sub
    TraceIt
End Sub

'---- Flow Tracing
Private Sub TraceIt()

    Dim i As Long, j As Integer                     'Loop counters
    Dim Tmp As String                               'Temp strings
    Dim PC  As Long, EA As Long                     'Program Counter, Effective Address
    Dim CodeOffset As Long                          'For calculating position in buffer
    Dim StartAddr As Long, EndAddr As Long          'Start and end addresses of code range
    Dim TargetAddr As Long, TargetH As String       'Target address dec and hex
    
    Dim OpByte As Integer                           'Opcode byte            (0 to 255)
    Dim OpDef  As String                            'Opcode definition      (ie: BRKi = BRK immediate)
    Dim OpStr  As String                            'Opcode Mnuemonic       (ie: BRK)
    Dim OpMode As Integer                           'Opcode Addressing Mode (ie: immediate)
    Dim StopFlag As Boolean                         'Flag to indicate flow is stopped (end of branch)
    Dim RangeS As Integer, RangeE As Integer        'Range boundaries
    
    ReDim Addr(32767) As Boolean                    'This is the address space (FALSE=data, TRUE=code)
    
    '-- Initialize
    lstML.Clear
    lstEP.Clear
    
    For i = 0 To 32767: Addr(i) = False: Next                   'Mark entire address space as data
    StopFlag = True                                             'Set to True so first EP is removed from list

    Address = VLA: If VLA = 0 Then Address = MyDec(txtLA.Text)  'VLA or User Specified Address
    StartAddr = Address
    EndAddr = Address + VLen - 1
    CodeOffset = Address - 1
    PC = StartAddr:  EA = 1
    
    lstML.AddItem "Target CPU: " & OpDesc
    lstML.AddItem "Jumps.....: " & OpJ
    lstML.AddItem "Branches..: " & OpB
    lstML.AddItem "Stops.....: " & OpZ

    lstML.AddItem "Loading entry points..."
    For i = 0 To lstEntryPt.ListCount - 1
        lstEP.AddItem Left(lstEntryPt.List(i), 4)               'Add Hex adress to tracer list
    Next i
    cmdAddTables.Visible = False                                'Reset visibility
       
    lstML.AddItem "Starting Trace. Code from $" & MyHex(StartAddr, 4) & " to $" & MyHex(EndAddr, 4) & " (" & Format(StartAddr) & "-" & Format(EndAddr) & ") = " & Format(VLen) & " bytes."
        
    Do
        '-- Are we STOPPED?
        If StopFlag = True Then
            i = lstEP.ListCount - 1                             'How many entry points?
            If i < 0 Then lstML.AddItem "Finished trace!": Exit Do
            
            Tmp = Left(lstEP.List(i), 4)                        'Read new EP (hex format HHHH)
            lstML.AddItem "Tracing from $" & Tmp & " (" & Format(PC) & ")"
            lstEP.RemoveItem i                                  'Remove EP from bottom of list
            PC = MyDec(Tmp)                                     'Set Program Counter address (convert to decimal)
            StopFlag = False
            DoEvents
        End If
                
        '-- Get instruction
        EA = PC - CodeOffset                                    'Calculate buffer pos
        
        If Addr(EA) = False Then
            OpByte = Asc(Mid(VBuf, EA, 1))                      'Get the opcode
            OpDef = UCase(OP(OpByte))                           'Opcode definition (OP array is global and should be loaded at ASM init)
            OpStr = Left(OpDef, Len(OpDef) - 1)                 'Opcode string
            OpMode = Asc(Right(UCase(OpDef), 1)) - 64           'Opcode addressing mode (a-z)
            OpLen = Val(Mid(OpModeLen, OpMode, 1))              'Get instruction length
            Addr(EA) = True                                     'Mark byte as code (opcode)
            
            If OpLen > 1 Then
                OpByte = Asc(Mid(VBuf, EA + 1, 1))              'Get the first parameter byte
                Addr(EA + 1) = True                             'Mark it as code
                If OpMode = 10 Then
                    If OpByte > 127 Then OpByte = OpByte - 256  'Calculate backwards branch
                    TargetAddr = PC + OpByte + 2                'Relative offset for branch
                Else
                    TargetAddr = OpByte                         'LO byte of target
                End If
            End If
            
            If OpLen = 3 Then
                OpByte = Asc(Mid(VBuf, EA + 2, 1))              'Get the second parameter byte
                Addr(EA + 2) = True                             'Mark it as code
                TargetAddr = TargetAddr + 256& * OpByte         'HI byte of target gets combined with LO
            End If
            
            TargetH = MyHex(TargetAddr, 4)
            
            '-- Is opcode a flow change?
            If InStr(1, OpJ, OpStr) > 0 Then
                lstML.AddItem "Flow change at $" & TargetH
                If (TargetAddr >= StartAddr) And (TargetAddr <= EndAddr) Then
                    lstEP.AddItem TargetH                       'Add target to EP then STOP
                End If
                StopFlag = True                                 'STOP
            End If
            
            '-- Is opcode a flow split?
            If InStr(1, OpB, OpStr) > 0 Then
                lstML.AddItem "found flow split at $" & TargetH
                If (TargetAddr >= StartAddr) And (TargetAddr <= EndAddr) Then
                    lstEP.AddItem TargetH                       'Add target to EP
                    lstML.AddItem "Adding " & TargetH & " to EP list."
                End If
                
                Tmp = MyHex(PC + OpLen, 4)
                lstEP.AddItem Tmp                               'Add next opcode to EP
                lstML.AddItem "Adding " & Tmp & " to EP list."
                
                StopFlag = True                                 'STOP
            End If
            
            '-- Is opcode a flow stop?
            If InStr(1, OpZ, OpStr) > 0 Then
                lstML.AddItem "Flow stop at" & Str(PC) & "."
                StopFlag = True                                 'STOP
            End If
        Else
            lstML.AddItem "Found marked code - skipping."
            StopFlag = True                                     'Instruction already marked as code, treat as STOP
        End If
                
        PC = PC + OpLen                                         'Increment Pc
        EA = EA + OpLen                                         'Increment buffer pointer
    Loop
    
    '-- Completed Trace
    StopFlag = False                                            'Flag to indicate we are in a data area
    lstEP.Clear                                                 'List should be empty but clear just in case
    
    lstML.AddItem "Building data table list..."
    For i = 1 To VLen
        If Addr(i) = False Then
            If StopFlag = False Then RangeS = i: StopFlag = True
        Else
            If StopFlag = True Then
                RangeE = i - 1: StopFlag = False                'If we are in a data range then this is the end
                lstEP.AddItem MyHex(RangeS + CodeOffset, 4) & "," & MyHex(RangeE + CodeOffset, 4) & ",b,trace data" 'Add it to the list
            End If
        End If
    Next i
    
    '-- Handle data block at the end of the code range
    If StopFlag = True Then
        RangeE = i - 1: StopFlag = False                        'If we are in a data range then this is the end
        lstEP.AddItem MyHex(RangeS + CodeOffset, 4) & "," & MyHex(RangeE + CodeOffset, 4) & ",b,trace data" 'Add it to the list
    End If
    lstML.AddItem "Done!"
    
    If lstEP.ListCount > 0 Then cmdAddTables.Visible = True     'Make ADD button visible
End Sub

'---- Add Results of Trace to Data Tables list
Private Sub cmdAddTables_Click()
    Dim i As Integer, Tmp As String
    
    For i = 0 To lstEP.ListCount - 1
        Tmp = lstEP.List(i)                     'Get from Trace List
        lstDT.AddItem Tmp                       'Add to Data Tables list
        lstDT.Selected(lstDT.NewIndex) = True
        If cbMLAddLabels.value = vbChecked Then
            lstULabels.AddItem Left(Tmp, 4) & ",Trace" & Left(Tmp, 4)
        End If
    Next
    lstEP.Clear
    cmdAddTables.Visible = False
    MLReViewC
End Sub

'---- Save ASM ouput to file
Private Sub cmdSaveASM_Click()
    Dim j As Integer, Filename As String, FIO As Integer
    
    Filename = FileOpenSave(FileBase(VFileName), 1, 4, "Save ASM code")
    If Filename = "" Then Exit Sub
    If Overwrite(Filename) = False Then Exit Sub
    
    FIO = FreeFile
    Open Filename For Output As FIO
    
    For j = 0 To lstML.ListCount - 1
        Print #FIO, lstML.List(j)
    Next j
    Close FIO

End Sub

'---- Copy ASM ouput to Clipboard
Private Sub cmdCopyClip2_Click()
    Dim j As Integer, Tmp As String
    
    For j = 0 To lstML.ListCount - 1
        Tmp = Tmp & lstML.List(j) & vbCrLf
    Next j
    
    Clipboard.Clear
    Clipboard.SetText Tmp

End Sub
'---- Show Project Changed Status
Sub ShowMLChange()
    If frML.Visible = True Then
        If ProjFilename = "" Then
            lblChanged.BackColor = vbWhite
        Else
            If ChangeFlag = True Then lblChanged.BackColor = vbRed Else lblChanged.BackColor = vbGreen
        End If
    End If
End Sub

'---- This Re-Views the file when options have changed but only if autorefresh is true
Sub MLReViewA()
    If cbAuto.value = vbChecked Then MLReView
    ShowMLChange
End Sub

'---- This Re-Views the file as above, but also sets the Changes Flag=True
Sub MLReViewC()
    ChangeFlag = True
    If cbAuto.value = vbChecked Then MLReView
    ShowMLChange
End Sub

'---- This Re-Views the file when options have changed
Sub MLReView()
    Dim TopPos As Integer
    
    TopPos = lstML.TopIndex                             'Remember the position
    If ViewerReady = True Then MLView
    ShowMLChange                                        'ML Project status
    lstML.TopIndex = TopPos                             'Restore the position
End Sub


'---- Jump to selected entry in currently visible table
Private Sub cmdSYMGoto_Click()
    Dim i As Integer, Tmp As String
        
    Select Case MLTabNum
        Case 2
            If lstEntryPt.ListCount = -1 Then Exit Sub
            i = lstEntryPt.ListIndex: If i < 0 Then i = 0
            Tmp = Left(lstEntryPt.List(i), 4)
    
        Case 3
            If lstSYM.ListCount = -1 Then Exit Sub
            i = lstSYM.ListIndex: If i < 0 Then i = 0
            Tmp = GetField(lstSYM.List(i), 2)
        
        Case 4
            If lstDT.ListCount = -1 Then Exit Sub
            i = lstDT.ListIndex: If i < 0 Then i = 0
            Tmp = Left(lstDT.List(i), 4)
            
        Case 5
            If lstULabels.ListCount = -1 Then Exit Sub
            i = lstULabels.ListIndex: If i < 0 Then i = 0
            Tmp = Left(lstULabels.List(i), 4)
        Case 6
            If lstCmnt.ListCount = -1 Then Exit Sub
            i = lstCmnt.ListIndex: If i < 0 Then i = 0
            Tmp = Left(lstCmnt.List(i), 4)
        Case 7
            If lstLabels.ListCount = -1 Then Exit Sub
            i = lstLabels.ListIndex: If i < 0 Then i = 0
            Tmp = Left(lstLabels.List(i), 4)
        Case 8
            If lstJSR.ListCount = -1 Then Exit Sub
            i = lstJSR.ListIndex: If i < 0 Then i = 0
            Tmp = Left(lstJSR.List(i), 4)
    End Select
    
    JumpList Tmp, True

End Sub

'---- Find and jump to the next undefined opcode
Private Sub lblGood_Click()
    JumpList "???", False
End Sub

'---- Find specified string
Private Sub cmdFind_Click()
    Dim Tmp As String
    
    Tmp = InputBox("Enter String to find:", "Find")
    If Tmp <> "" Then
        cmdFindAll.ToolTipText = ""
        JumpList Tmp, False
    End If
    
End Sub

'---- Find ALL occurances of last search string
Private Sub cmdFindAll_Click()
    JumpList "", True
End Sub

'---- Jump to next occurance of search string
Private Sub cmdNext_Click()
    JumpList "", False
End Sub

'---- Search for string
' Blank string to search for last. Set flag true to start from top, otherwise start from current position
Sub JumpList(ByVal Txt As String, ByVal Flag As Boolean)
    Static LastTxt As String, Count As Integer 'These values are retained between calls
    
    Dim i As Integer, j As Integer, Max As Integer
    
    If Txt = "" Then Txt = LastTxt
    If Txt = "" Then Exit Sub
       
    Max = lstML.ListCount - 1
    
    If Flag = False Then
        i = lstML.ListIndex + 1             'FLAG=false - start at next index position
    Else
        i = 0                               'FLAG=true - start at top
        Count = 0
    End If
   
    Do
        If InStr(1, lstML.List(i), Txt, vbTextCompare) > 0 Then
            j = i - 5: If j < 0 Or j > Max Then j = i
            lstML.TopIndex = j          'move top of list to near found line
            lstML.ListIndex = i         'move to selected line
            lstML.Selected(i) = True    'hilight it
            Count = Count + 1
            If Flag = False Then Exit Do
        End If
        i = i + 1                       'next line
    Loop While i < Max
    
    If Count > 0 Then
        cmdFindAll.ToolTipText = "Found" & Str(Count) & " line(s)"
        cmdNext.ToolTipText = "Find: " & Txt
    End If
    DoEvents
    LastTxt = Txt
    
End Sub

'---- Find an Address in the following order: SYMBOL, ULABEL, LABEL.
'     Return SYMBOL name, ULABEL name, or "L_xxxx" LABEL string
Private Function FindSL(ByVal Addr As String) As String
    Dim Tmp As String
    
    Tmp = FindSym(Addr)
    If Tmp > "" Then
        FindSL = Tmp
    Else
        Tmp = FindUL(Addr)
        If Tmp > "" Then FindSL = Tmp
    End If
End Function

'---- Find a User Label or Generated Label in the following order: ULABEL, LABEL.
'     Return ULABEL name, or LABEL with Prefix string
Private Function FindUL(ByVal Addr As String) As String
    Dim Tmp As String
    
    Tmp = FindULabel(Addr)
    If Tmp > "" Then
        FindUL = Tmp
    Else
        Tmp = FindLabel(Addr)
        If Tmp > "" Then FindUL = LPrefix & Tmp
    End If
    
End Function

'---- Lookup SYMBOL and return string. Also Set LastComment
' FORMAT of SYMBOL list entry: HHHH,symbolstring,comment
Private Function FindSym(ByVal Addr As String) As String
    Dim Tmp As String, Tmp2 As String, Tmp3 As String
    Dim R1 As Integer, R2 As Integer, R3 As Integer 'binary search range
        
    R3 = lstSYM.ListCount - 1                   'Range End position
    If R3 < 0 Then Exit Function                'Exit if no entries
    R1 = 0                                      'Range Start position
    LastComment = ""                            'Clear Last Comment string
    LastSymPos = 0                              'Clear Last SYM position
    
    Do
        R2 = (R1 + R3) \ 2                      'Calculate middle of range
        Tmp = lstSYM.List(R2)                   'Check array at middle position
        Tmp2 = Left(Tmp, 4)                     'Extract address part
        
        If Tmp2 = Addr Then
            Tmp3 = GetField(Tmp, 2)
            If Tmp3 = "" Then Tmp3 = "$" & Addr 'If not symbol then just use address
            FindSym = MyTrim(Tmp3)              'Return string
            LastComment = GetField(Tmp, 3)      'Get the comment
            LastSymPos = R2                     'Remember it's position
            lstSYM.Selected(R2) = True
            Exit Do
        Else
            If Tmp2 > Addr Then R3 = R2 - 1 Else R1 = R2 + 1 'Adjust range end points depending on comparison
        End If
        If R1 > R3 Then FindSym = "": Exit Do  'No more in range, so exit with NULL string
    Loop

End Function

'---- Find ULABEL Address and return Symbol string
' FORMAT of ULABEL list entry: HHHH,symbolstring
Private Function FindULabel(ByVal Addr As String) As String
    Dim Tmp As String, Tmp2 As String
    Dim R1 As Integer, R2 As Integer, R3 As Integer 'binary search range
        
    R3 = lstULabels.ListCount - 1               'Range End position
    If R3 < 0 Then Exit Function                'Exit if no entries
    R1 = 0                                      'Range Start position
        
    Do
        R2 = (R1 + R3) \ 2                      'Calculate middle of range
        Tmp = lstULabels.List(R2)               'Check array at middle position
        Tmp2 = Left(Tmp, 4)                     'Extract Address
        
        If Tmp2 = Addr Then
            FindULabel = GetField(Tmp, 2)       'Substitute label
            Exit Do
        Else
            If Tmp2 > Addr Then R3 = R2 - 1 Else R1 = R2 + 1 'Adjust range end points depending on comparison
        End If
        If R1 > R3 Then FindULabel = "": Exit Do 'No more in range, so exit with NULL string
        DoEvents
    Loop

End Function

'---- Find LABEL Address and return Address string
' FORMAT of LABEL list entry: HHHH
Private Function FindLabel(ByVal Addr As String) As String
    Dim Tmp As String, Tmp2 As String
    Dim R1 As Integer, R2 As Integer, R3 As Integer 'binary search range
        
    R3 = lstLabels.ListCount - 1                'Range End position
    If R3 < 0 Then Exit Function                'Exit if no entries
    R1 = 0                                      'Range Start position

    Do
        R2 = (R1 + R3) \ 2                      'Calculate middle of range
        Tmp = lstLabels.List(R2)                'Check array at middle position
        Tmp2 = Left(Tmp, 4)                     'Extract Address
                
        If Tmp2 = Addr Then
            FindLabel = Tmp2                     'Return the label
            Exit Do
        Else
            If Tmp2 > Addr Then R3 = R2 - 1 Else R1 = R2 + 1 'Adjust range end points depending on comparison
        End If
        If R1 > R3 Then FindLabel = "": Exit Do 'No more in range, so exit with NULL string
        DoEvents
    Loop

End Function

'---- Lookup comment for specified address and return "type,commentstring"
' FORMAT of COMMENT list entry: HHHH,type,commentstring
'
Private Function FindComment(ByVal Addr As String) As String
    Dim Tmp As String, Tmp2 As String
    Dim R1 As Integer, R2 As Integer, R3 As Integer 'binary search range
        
    R3 = lstCmnt.ListCount - 1                  'Range End position
    If R3 < 0 Then Exit Function                'Exit if no entries
    R1 = 0                                      'Range Start position
    
    Do
        R2 = (R1 + R3) \ 2                      'Calculate middle of range
        Tmp = lstCmnt.List(R2)                  'Check array at middle position
        Tmp2 = Left(Tmp, 4)                     'Extract Address
        
        If Tmp2 = Addr Then
            FindComment = Mid(Tmp, 6)           'Return the type and commentstring
            Exit Do
        Else
            If Tmp2 > Addr Then R3 = R2 - 1 Else R1 = R2 + 1 'Adjust range end points depending on comparison
        End If
        If R1 > R3 Then FindComment = "": Exit Do 'No more in range, so exit with NULL string
    Loop

End Function

'---- Quick Add Label
Private Sub cmdAddLabel_Click()
    Dim RS As String, Tmp As String, Tmp2 As String, i As Integer
    
    Tmp = "Please select a line with an address first!"
    
    i = lstML.ListIndex: If i < 0 Then MsgBox Tmp: Exit Sub                 'Ooops, no line selected!
    RS = ExtractAddr(lstML.List(i)): If RS = "" Then MsgBox Tmp: Exit Sub   'Ooops, line didn't have an address!
 
    Tmp2 = InputBox("Add label at " & RS & Cr & Cr & "Enter LABEL Name:", "Add Label", "")
    If Tmp2 > "" Then lstULabels.AddItem RS & "," & Tmp2: MLReViewC
    
End Sub

'---- Quick Add Comment / Separator ( ;C / C / -C- / =C= / - / = )
Private Sub cmdAddComment_Click(Index As Integer)
    Dim RS As String, Tmp As String, Tmp2 As String, i As Integer
    
    Tmp = "Please select a line with an address first!"
    
    i = lstML.ListIndex: If i < 0 Then MsgBox Tmp: Exit Sub     'Oops, no line selected!
    RS = ExtractAddr(lstML.List(i)): If RS = "" Then MsgBox Tmp: Exit Sub   'Opps, line didn't have an address!
        
    Tmp = Mid("is-=*-=*", Index + 1, 1): Tmp2 = ""
    
    '---- 0 to 4 need a Comment, 5 to 8 are dividers
    If Index < 5 Then Tmp2 = InputBox("Enter a comment at position " & RS & ":", "Enter Comment", ""): If Tmp2 = "" Then Exit Sub
    lstCmnt.AddItem RS & "," & Tmp & "," & Tmp2     'Add it
    MLReViewC
End Sub

'---- Quick Add Data Table (DHSRVW)
Private Sub cmdDTAdd_Click(Index As Integer)
    Dim Tmp As String, Tmp2 As String
    Dim Flag As Boolean, p As Integer, RS As String, RE As String

    Flag = False
    
    'Check if there is a range selected
    For i = 0 To lstML.ListCount - 1
        If lstML.Selected(i) = True Then
            If Flag = False Then RS = ExtractAddr(lstML.List(i)): Flag = True   'Found first selected line
            p = i                                                               'remember it
        Else
            If Flag = True Then RE = ExtractAddr(lstML.List(p)): Exit For       'Not selected so use last remembered line for end
        End If
    Next i
         
    If Flag = True Then
        If RE = "" Then RE = RS
        Select Case Index 'DHSRVW
            Case 0: Tmp = "D": Tmp2 = "Decimal Byte Table"
            Case 1: Tmp = "H": Tmp2 = "Hex Byte Table"
            Case 2: Tmp = "S": Tmp2 = "Text/String Table"
            Case 3: Tmp = "R": Tmp2 = "RTS Address Table (Generates Labels)"
            Case 4: Tmp = "V": Tmp2 = "Address Table (Generates Labels)"
            Case 5: Tmp = "W": Tmp2 = "Word Table"
            Case 6: Tmp = "X": Tmp2 = "Hidden Table"
        End Select
                   
        Tmp2 = InputBox("Type : " & Tmp2 & Cr & "Range: " & RS & " to " & RE & Cr & Cr & "Enter a description:", "Add Table", "")
        If Tmp2 <> "" Then
            lstDT.AddItem RS & "," & RE & "," & Tmp & "," & Tmp2    'Add it
            lstDT.Selected(lstDT.NewIndex) = True                   'Make it selected
            MLReViewC
        End If
    Else
        MsgBox "Please select a range first!"
    End If
    
End Sub

'---- Edit a Data Table Entry
Private Sub lstDT_DblClick()
    Dim i As Integer, Tmp As String, Tmp2 As String
    
    i = lstDT.ListIndex
    If i >= 0 Then
        Tmp = lstDT.List(i)
        Tmp2 = InputBox("Edit Data Table:", "Edit Data Table", Tmp)
        If Tmp2 > "" Then
            lstDT.RemoveItem i
            lstDT.AddItem Tmp2
            lstDT.Selected(lstDT.NewIndex) = True
        End If
    End If

End Sub
Private Sub lstML_Click()
    Dim Tmp As String, Tmp2 As String, Tmp3 As String, Addr As String
    Dim R1 As Integer, R2 As Integer, R3 As Integer 'binary search range
    
    
    If frBIN.Visible = True Then
        'DualView with Hex visible - Try to find matching hex listing line
        Addr = Left(lstML.List(lstML.ListIndex), 4) 'Address in ASM listing
        If Len(Addr) <> 4 Then Exit Sub
        Tmp = Right(Addr, 1): Tmp2 = "0"             'Last digit and replacement default
        If cbWide.value = vbUnchecked Then
            If Tmp < "8" Then Tmp2 = "0" Else Tmp2 = "8"
        End If
        Mid(Addr, 4, 1) = Tmp2                      'Replace the last digit
        
        R3 = lstBIN.ListCount - 1                   'Range End position
        If R3 < 0 Then Exit Sub                     'Exit if no entries
        R1 = 0                                      'Range Start position
        
        Do
            R2 = (R1 + R3) \ 2                      'Calculate middle of range
            Tmp = lstBIN.List(R2)                   'Check array at middle position
            Tmp2 = Left(Tmp, 4)                     'Extract address part
            
            If Tmp2 = Addr Then
                lstBIN.ListIndex = R2: Exit Do      'Highlight the BIN line
            Else
                If Tmp2 > Addr Then R3 = R2 - 1 Else R1 = R2 + 1 'Adjust range end points depending on comparison
            End If
            If R1 > R3 Then Exit Do  'No more in range, so exit with NULL string
        Loop
    End If
        
End Sub

'---- Add a Symbol by Doubleclick of ML line
Private Sub lstML_DblClick()
    Call cmdSymAdd_Click
End Sub

'---- Edit Symbol Table Entry
Private Sub lstSYM_dblClick()
    Dim i As Integer, Tmp As String, Tmp2 As String
    
    i = lstSYM.ListIndex
    If i >= 0 Then
        Tmp = lstSYM.List(i)
        Tmp2 = InputBox("Edit Symbol:", "Edit Symbol", Tmp)
        If Tmp2 > "" Then
            lstSYM.RemoveItem i
            lstSYM.AddItem Tmp2
            MLReViewC
        End If
    End If
    
End Sub

'---- Edit User Label Table Entry
Private Sub lstULabels_dblClick()
    Dim i As Integer, Tmp As String, Tmp2 As String
    
    i = lstULabels.ListIndex
    If i >= 0 Then
        Tmp = lstULabels.List(i)
        Tmp2 = InputBox("Edit Label:", "Edit Label", Tmp)
        If Tmp2 > "" Then
            lstULabels.RemoveItem i
            lstULabels.AddItem Tmp2
            MLReViewC
        End If
    End If
    
End Sub

'---- Edit USer Comment Table Entry
Private Sub lstCmnt_dblClick()
    Dim i As Integer, Tmp As String, Tmp2 As String
    
    i = lstCmnt.ListIndex
    If i >= 0 Then
        Tmp = lstCmnt.List(i)
        Tmp2 = InputBox("Edit Comment:", "Edit Comment", Tmp)
        If Tmp2 > "" Then
            lstCmnt.RemoveItem i
            lstCmnt.AddItem Tmp2
            MLReViewC
        End If
    End If
    
End Sub

'---- Toggle displaying of Data and Symbol Table frames
Private Sub lblShw_Click()
    ShowTables = Not ShowTables
    DrawVLayout
End Sub

'---- Prompt to Save Symbol Table to file
Private Sub cmdSymSave_Click()
    Dim FIO As Integer, Filename As String, i As Integer, Msg As String
    
    Select Case MLTabNum
        Case 2: Msg = "Entry Point"
        Case 3: Msg = "Symbol Tables"
        Case 4: Msg = "Data Tables"
        Case 5: Msg = "Labels"
        Case 6: Msg = "Comments"
    End Select
    
    Filename = FileOpenSave("", 1, 1, "Save " & Msg & " file"): If Filename = "" Then Exit Sub
    If Overwrite(Filename) = False Then Exit Sub
    
    FIO = FreeFile
    Open Filename For Output As FIO
    
    Select Case MLTabNum
    
        Case 2
            For i = 0 To lstEntryPt.ListCount - 1
                Print #FIO, lstEntryPt.List(i)
            Next i
    
        Case 3
            For i = 0 To lstSYM.ListCount - 1
                Print #FIO, lstSYM.List(i)
            Next i
            
        Case 4
            For i = 0 To lstDT.ListCount - 1
                Print #FIO, lstDT.List(i)
            Next i
            
        Case 5
            For i = 0 To lstULabels.ListCount - 1
                Print #FIO, lstULabels.List(i)
            Next i
            
        Case 6
            For i = 0 To lstCmnt.ListCount - 1
                Print #FIO, lstCmnt.List(i)
            Next i
    End Select
    
    Close FIO

End Sub

'---- Prompt for Loading a new Symbol Table File
Private Sub cmdSymLoad_Click()
    Dim Filename As String, Msg As String
    
    Select Case MLTabNum
        Case 2: Msg = "Entry Points"
        Case 3: Msg = "Symbol Tables"
        Case 4: Msg = "Data Tables"
        Case 5: Msg = "Labels"
        Case 6: Msg = "Comments"
    End Select
    
    Filename = FileOpenSave("", 0, 1, "Load " & Msg & " file"): If Filename = "" Then Exit Sub
    If Exists(Filename) = False Then Exit Sub
    
    LoadSymFile Filename, MLTabNum
    MLReViewA
    
End Sub

'---- Process selection of a new Platform from the list
Private Sub cboPlatform_Click()
    Dim Filename As String, i As Integer
    
    If MLCFlag = False Then Exit Sub
    If ViewerReady = False Then Exit Sub
    
    i = cboPlatform.ListIndex: If i = 0 Then Exit Sub
    
    Filename = ExeDir & cboPlatFile.List(i)
    If Exists(Filename) = False Then MsgBox "Sorry, Platform file not found! " & Filename: Exit Sub
    If OverwriteProject = True Then LoadSymFile Filename, 3
    MLReView
    
End Sub

'---- Process selection of a new CPU from the list
Private Sub cboCPU_Click()
    Dim Filename As String
    If MLCFlag = False Then Exit Sub
    If ViewerReady = False Then Exit Sub
    
    Filename = ExeDir & cboCPUFile.List(cboCPU.ListIndex)
    If Exists(Filename) = False Then MsgBox "Sorry, CPU file not found! " & Filename: Exit Sub
    LoadOpcodes Filename
    MLReView

End Sub

'---- Check Project Changed status and Prompt for Saving Project if there is a change
' Returns TRUE if:
'   - project has not changed, or there is no project file
'   - project has changed and YES or NO is selected. If YES is selected then project will be saved first
' Returns FALSE if CANCEL is selected
Private Function OverwriteProject() As Boolean
    Dim Result As VbMsgBoxResult
    
    OverwriteProject = False 'Assume NOT ok to continue
    If (ProjFilename <> "") And (ChangeFlag = True) Then
        Result = MsgBox("Project has changed. Save Changes first?", vbYesNoCancel)
        If Result = vbCancel Then Exit Function
        If Result = vbYes Then SaveProjFile ProjFilename 'YES=save project, NO=loose changes
    End If
    OverwriteProject = True
End Function

'---- Prompt for project filename to load
Private Sub cmdProjLoad_Click()
    If OverwriteProject = True Then
        Filename = FileOpenSave("", 1, 2, "Load ASM Project File"): If Filename = "" Then Exit Sub
        LoadProjFile Filename
        MLReView
    End If
End Sub

'---- Load specified Project File
' A Project file contains lines to be loaded into the tabels
' Each table group must be proceeded by a selection marker:
' [SYMBOLS] [TABLES] [LABELS] [COMMENTS]

Private Sub LoadProjFile(ByVal Filename As String)
    Dim FIO As Integer, Tmp As String, Tmp2 As String, TMode As Integer
    Dim LA As String, LAFlag As Boolean
        
    If Exists(Filename) = False Then Exit Sub
        
    FIO = FreeFile
    Open Filename For Input As FIO
    TMode = 0: LAFlag = False: ProjFlag = True: ProjFilename = Filename
    
    If cbClearOnLoad.value = vbChecked Then ClearTables
    
    While Not EOF(FIO)
        Line Input #FIO, Tmp
        If Left(Tmp, 1) = "[" Then
            '---- Check for section marker
            Select Case Tmp
                Case "[PROJECT]":  TMode = 1
                Case "[SYMBOLS]":  TMode = 2
                Case "[TABLES]":   TMode = 3
                Case "[LABELS]":   TMode = 4
                Case "[COMMENTS]": TMode = 5
                Case "[ENTRYPT]":  TMode = 6
                Case Else: TMode = 0
            End Select
        Else
            If (Left(Tmp, 1) <> ";") And (Tmp <> "") Then
                '---- Process according to current section marker
                Select Case TMode
                    Case 1: If Len(Tmp) = 4 Then LAFlag = True: LA = Tmp 'Load Address for project found
                    Case 2: lstSYM.AddItem Tmp
                    Case 3
                        lstDT.AddItem Tmp
                        lstDT.Selected(lstDT.NewIndex) = True
                        
                    Case 4: lstULabels.AddItem Tmp
                    Case 5: lstCmnt.AddItem Tmp
                    Case 6: lstEntryPt.AddItem Tmp
                End Select
            End If
        End If
    Wend

    Close FIO
    If LAFlag = True Then
        cbLA.value = vbUnchecked: txtLA.Text = LA   'Use Load address from Project file
    Else
        cbLA.value = vbChecked                      'Use Load address from source file
    End If
    cboPlatform.ListIndex = 0                       'Display "from project file"
    ChangeFlag = False
    MLReViewA
    
End Sub

'---- Prompt for Filename then save Project
Private Sub cmdProjSave_Click()
    Dim Tmp As String
    
    Tmp = ProjFilename: If Tmp = "" Then Tmp = FileBase(VFileName) 'Use Project Filename as default, otherwise use name of view file
    Filename = FileOpenSave(Tmp, 1, 2, "Save ASM Project File"): If Filename = "" Then Exit Sub
    SaveProjFile Filename
    ShowMLChange
End Sub

'---- Save specified Project File
' A Project file contains lines to be loaded into the tabels
' Each table group must be proceeded by a selection marker:
' [SYMBOLS] [TABLES] [LABELS] [COMMENTS]

Private Sub SaveProjFile(ByVal Filename As String)

    Dim FIO As Integer, Tmp As String, j As Integer
        
    If Overwrite(Filename) = False Then Exit Sub
        
    FIO = FreeFile
    Open Filename For Output As FIO
    TMode = 0
    
    '-- [PROJECT]
    Print #FIO, "[PROJECT]"
    If cbLA.value = vbUnchecked Then Print #FIO, txtLA.Text     'Save the specified Load Address
    
    '-- [ENTRY POINTS]
    If lstSYM.ListCount > 0 Then
        Print #FIO, "[ENTRYPT]"
        For j = 0 To lstEntryPt.ListCount - 1
            Print #FIO, lstEntryPt.List(j)
        Next j
    End If
    
    '-- [SYMBOLS]
    If lstSYM.ListCount > 0 Then
        Print #FIO, "[SYMBOLS]"
        For j = 0 To lstSYM.ListCount - 1
            Print #FIO, lstSYM.List(j)
        Next j
    End If
      
    '-- [TABLES]
    If lstDT.ListCount > 0 Then
        Print #FIO, "[TABLES]"
        For j = 0 To lstDT.ListCount - 1
            Print #FIO, lstDT.List(j)
        Next j
    End If
    
    '-- [LABELS]
    If lstULabels.ListCount > 0 Then
        Print #FIO, "[LABELS]"
        For j = 0 To lstULabels.ListCount - 1
            Print #FIO, lstULabels.List(j)
        Next j
    End If
    
    '-- [COMMENTS]
    If lstCmnt.ListCount > 0 Then
        Print #FIO, "[COMMENTS]"
        For j = 0 To lstCmnt.ListCount - 1
            Print #FIO, lstCmnt.List(j)
        Next j
    End If

    Close FIO
    ProjFilename = Filename 'Remember the project file
    ChangeFlag = False      'Clear Changes flag
    
End Sub

Private Sub cmdClrTables_Click()
    If OverwriteProject = True Then
        ClearTables
        ProjFilename = ""
        MLReViewA
    End If
End Sub

'---- Clear All Tables
Private Sub ClearTables()
    lstEntryPt.Clear
    lstSYM.Clear
    lstDT.Clear
    lstULabels.Clear
    lstCmnt.Clear
End Sub

'---- Load specified List from File
Private Sub LoadSymFile(ByVal Filename As String, ByVal TabNum As Integer)
    Dim FIO As Integer, Tmp As String, Tmp2 As String, Mode As Integer
    Dim Addr As String, Sym As String, Cmnt As String, Flag As Boolean
    
    If Exists(Filename) = False Then Exit Sub
    
    Mode = 0: Tmp = FileExtU(Filename): If Tmp = "SY4" Then Mode = 1    'Check for 'SYM4' file
    
    If TabNum = 3 Then
        'Do extra check for ReGenerator Symbol import
        Tmp = LCase(FileNameOnly(Filename))                                 'Get the filename without path etc
        If Tmp = "labels.txt" Or Tmp = "comments.txt" Then
            If MsgBox("Is this a ReGenerator file?", vbYesNo) = vbYes Then
                Mode = 3: If Tmp = "labels.txt" Then Mode = 2               'Found and Confirmed ReGenerator file
            End If
        End If
    End If
    
    Flag = (cbClearOnLoad.value = vbChecked)                                'Get Clear On Load option
    
    FIO = FreeFile
    Open Filename For Input As FIO
    
    Select Case TabNum
        Case 2
            If Flag = True Then lstEntryPt.Clear 'Clear Data Tables
            While Not EOF(FIO)
                Line Input #FIO, Tmp
                If Left(Tmp, 1) <> ";" Then lstEntryPt.AddItem Tmp
            Wend

        Case 3
            If (Mode < 2) And (Flag = True) Then lstSYM.Clear 'Clear Symbols
            While Not EOF(FIO)
                Line Input #FIO, Tmp
                If (Left(Tmp, 1) <> ";") And (Left(Tmp, 1) <> ":") And (MyTrim(Tmp) <> "") Then
                    Select Case Mode
                        Case 0 '--Standard format input
                            If Left(Tmp, 1) <> ";" Then lstSYM.AddItem Tmp
                        Case 1 '--SYM4 format
                            If Mid(Tmp, 2, 1) <> " " Then
                                Tmp2 = Mid(Tmp, 13, 4) & "," & Mid(Tmp, 2, 6) & "," & Mid(Tmp, 37)
                                lstSYM.AddItem Tmp2
                            End If
                        Case 2 '--Regenerator Label format: HHHH SYMBOL
                            Addr = Left(Tmp, 4)         'Save Address
                            Sym = MyTrim(Mid(Tmp, 6))   'Save Symbol
                            Tmp = FindSym(Addr)         'Check if there is an existing symbol
                            If Tmp = "" Then
                                lstSYM.AddItem Addr & "," & Sym & ","
                            End If
    
                        Case 3 '--ReGenerator Comment format: HHHH Comment
                            Addr = Left(Tmp, 4)         'Save Address
                            Cmnt = MyTrim(Mid(Tmp, 6))  'Save Symbol
                            Tmp = FindSym(Addr)         'Check if there is an existing symbol (LastSymPos will point to it)
                            If Tmp = "" Then
                                lstSYM.AddItem Addr & ",," & Cmnt
                            Else
                                'Symbol was found, so update data
                                If LastComment = "" Then
                                    'Only update if the symbol has no existing comment
                                    lstSYM.RemoveItem LastSymPos                    'Remove it
                                    lstSYM.AddItem Addr & "," & Tmp & "," & Cmnt    'Add replacement
                                End If
                            End If
                    End Select
                End If
            Wend

        Case 4
            If Flag = True Then lstDT.Clear 'Clear Data Tables
            While Not EOF(FIO)
                Line Input #FIO, Tmp
                If Left(Tmp, 1) <> ";" Then
                    lstDT.AddItem Tmp
                    lstDT.Selected(lstDT.NewIndex) = True
                End If
            Wend
                        
        Case 5
            If Flag = True Then lstULabels.Clear 'Clear User Labels
            While Not EOF(FIO)
                Line Input #FIO, Tmp: If Left(Tmp, 1) <> ";" Then lstULabels.AddItem Tmp
            Wend
        Case 6
            If Flag = True Then lstCmnt.Clear 'Clear Comment
            While Not EOF(FIO)
                Line Input #FIO, Tmp: If Left(Tmp, 1) <> ";" Then lstCmnt.AddItem Tmp
            Wend

    End Select
    Close FIO

End Sub

'---- Add a new List Entry
Private Sub cmdSymAdd_Click()
    Dim i As Integer, p As Integer, Flag As Boolean
    Dim RS As String, RE As String, Tmp As String, Tmp2 As String
    
    i = lstML.ListIndex
    Tmp2 = ""
    If i > 0 Then Tmp2 = ExtractAddr(lstML.List(lstML.ListIndex))                       'Find the address on selected line
    
    Select Case MLTabNum
        Case 0, 1
            MsgBox "Select the TAB for the type of entry you want first, or use the quick-add buttons at the top of the window!"
            
        Case 2 'Entry Points
            Tmp = Tmp2 & ",-"                                                           'Make default text entry string
            Tmp2 = InputBox("HHHH,DESCRIPTION", "Add Entry Pointl", Tmp)
            If Len(Tmp2) > 3 Then lstEntryPt.AddItem Tmp2: MLReViewC                    'Review plus set changeflag=true
        
        Case 3 'Symbols
            Tmp = Tmp2 & ",symbol,-"                                                    'Make default text entry string
            Tmp2 = InputBox("HHHH,SYMBOL,DESCRIPTION", "Add Symbol", Tmp)
            If Len(Tmp2) > 12 Then lstSYM.AddItem Tmp2: MLReViewC                       'Review plus set changeflag=true
        
        Case 4 'Data Tables
            'Check if there is a range selected
            For i = 0 To lstML.ListCount - 1
                If lstML.Selected(i) = True Then
                    If Flag = False Then RS = ExtractAddr(lstML.List(i)): Flag = True   'Found first selected line
                    p = i                                                               'remember it
                Else
                    If Flag = True Then RE = ExtractAddr(lstML.List(p)): Exit For       'Not selected so use last remembered line for end
                End If
            Next i
            
            If Flag = True Then Tmp = RS & "," & RE & ",b,-"
            Tmp2 = InputBox("Types: A/T=Text,B/H=Hex Bytes,D=Dec Bytes,W=Word,R=RTS,V=Vect" & Cr & Cr & "HHHH,HHHH,TYPE,DESCRIPTION", "Add Table", Tmp)
            If Len(Tmp2) > 12 Then
                lstDT.AddItem Tmp2
                lstDT.Selected(lstDT.NewIndex) = True
                MLReViewC
            End If
            
        Case 5 'User Labels
            Tmp = Tmp2 & ",name,-"                                                      'Make default text entry string
            Tmp2 = InputBox("HHHH,LABELNAME,DESCRIPTION", "Add Label", Tmp)
            If Len(Tmp2) > 12 Then lstULabels.AddItem Tmp2: MLReViewC
        
        Case 6 'Comments
            Tmp = Tmp2 & ",s,-"                                                         'Make default text entry string
            Tmp2 = InputBox("Types: I=In-line,S=Single,OTHER=Double Divider Chr" & Cr & "(For Single Divider leave comment empty)" & Cr & Cr & "HHHH,TYPE,COMMENT", "Add Comment", Tmp)
            If Len(Tmp2) > 6 Then lstCmnt.AddItem Tmp2: MLReViewC
    End Select
    
End Sub

'---- Extracts the HEX Address from the string using current PREFIX
' If PREFIX is not found then look at start of line
Private Function ExtractAddr(ByVal Str As String) As String
    Dim p As Integer, Tmp As String, Tmp2 As String, L As Integer
    
    L = Len(LPrefix)
    p = 1
    If Left(Str, L) = LPrefix Then p = L + 1          'Skip over prefix
    Tmp = UCase(Mid(Str, p, 4))                                             'Extract the hex address
    Tmp2 = Left(Tmp, 1)                                                     'Get first character
    If (Tmp2 < "0") Or (Tmp2 > "F") Then Exit Function                      'Exit if not 0-F
    If (Tmp2 <= "9") Or (Tmp2 >= "A") Then ExtractAddr = Tmp                'Check for valid 0-9 or A-F

End Function

'---- Remove the Current List Entry
' Uses global variable MLTabNum to determine list. If item is removed ChangeFlag is set true
Private Sub cmdSymDel_Click()
    Dim i As Integer
    
    Select Case MLTabNum
        Case 2
            i = lstEntryPt.ListIndex
            If i >= 0 Then lstEntryPt.RemoveItem (i)
        Case 3
           i = lstSYM.ListIndex
           If i >= 0 Then lstSYM.RemoveItem (i): MLReViewC
           
        Case 4
           i = lstDT.ListIndex
           If i >= 0 Then lstDT.RemoveItem (i): MLReViewC
        
        Case 5
           i = lstULabels.ListIndex
           If i >= 0 Then lstULabels.RemoveItem (i): MLReViewC
           
        Case 6
           i = lstCmnt.ListIndex
           If i >= 0 Then lstCmnt.RemoveItem (i): MLReViewC
    End Select
     
End Sub
Private Sub cboPrefix_Click()
    SetPrefix cboPrefix.ListIndex
    MLReViewA
End Sub

Private Sub SetPrefix(ByVal n As Integer)
    LPrefix = cboPrefix.List(n)
End Sub

Private Sub cboTarget_Click()
    SetTarget cboTarget.ListIndex
    MLReViewA
End Sub

'---- Sets Target Assembler Directives
Private Sub SetTarget(ByVal n As Integer)
    Select Case n
        Case 0: DOTORG = "*=":    DOTWORD = "!WORD ": DOTBYTE = "!BYTE ": DOTTEXT = "!TEXT "
        Case 1: DOTORG = "*=":    DOTWORD = ".WORD ": DOTBYTE = ".BYTE ": DOTTEXT = ".TEXT "
        Case 2: DOTORG = ".ORG ": DOTWORD = ".WOR ": DOTBYTE = ".BYT ":   DOTTEXT = ".TXT "
    End Select
End Sub

'---- Load opcodes from specified file
Private Sub LoadOpcodes(ByVal Filename As String)
    Dim Tmp As String, j As Integer
    
    If Exists(Filename) = False Then Exit Sub
    FIO = FreeFile
    Open Filename For Input As FIO
    
    Line Input #FIO, Tmp                                'CBM-Transfer header line
    Line Input #FIO, OpDesc                             'CPU description string
    Line Input #FIO, Tmp                                'Divider line
    
    For j = 0 To 255
        Input #FIO, Tmp: OP(j) = Tmp
    Next j
    
    Line Input #FIO, Tmp                                'Divider line
    Line Input #FIO, OpModeLen                          'Opcode lengths
    Line Input #FIO, OpJ                                'Tracer Jumps    (unconditional - single flow)
    Line Input #FIO, OpB                                'Tracer Branches (conditional - two flows)
    Line Input #FIO, OpZ                                'Tracer Stops    (end flow)
    'The rest of the file is ignored
    
    Close FIO
    OpCodeFlag = True
End Sub

'---- Import Symbol Entries
' Supports Fixed, Comma and Tab-delimited files using parameters entered by user
Private Sub cmdImport_Click()
    Dim Tmp As String, Tmp2 As String, Meth As String
    Dim Filename As String, FIO As Integer
    Dim Par(6) As Integer, Out As String, Flag As Boolean
    Dim C As Integer, i As Integer, p1 As Integer, p2 As Integer
    
    Tmp2 = "Enter Import control string: TYPE,n,n,n..." & Cr & Cr _
         & "Where 'TYPE' is: 'C','T', or 'F'" & Cr _
         & "(See Docs for parameter usage!)"
    Tmp = InputBox(Tmp2, "Enter Import Control String", "T,1,2,3"): If Tmp = "" Then Exit Sub

    Flag = True: Meth = UCase(GetField(Tmp, 1)): C = 3: If Meth = "F" Then C = 6
    
    For i = 1 To C
      Par(i) = Val(GetField(Tmp, i + 1))
      If Par(i) < 1 Then Flag = False: Exit For
    Next
    
    If Flag = False Then MsgBox "All numbers must be >0!": Exit Sub
    
    Filename = FileOpenSave("", 0, 0, "Import Symbol file"): If Filename = "" Then Exit Sub
    
    C = 0 'count of symbols imported
    
    FIO = FreeFile
    Open Filename For Input As FIO
    
    While Not EOF(FIO)
        Line Input #FIO, Tmp
        If Left(Tmp, 1) <> ";" Then
            Out = ""
            For i = 1 To 3
                Tmp2 = ""                                       'Clear Tmp2
                Select Case Meth
                    Case "C": Tmp2 = GetField(Tmp, Par(i))      'Extract field from Comma-delimited line
                    Case "T": Tmp2 = GetDField(Tmp, "", Par(i))  'Extract field from delimited line (Null Delimiter defaults to TAB)
                    Case "F"
                        p1 = Par(i * 2 - 1)                     'Start Position
                        p2 = Par(i * 2)                         'Length
                        If p2 > 0 Then Tmp2 = MyTrim(Mid(Tmp, p1, p2)) 'Extract the field at position p1 with length p2 and trim it
                 End Select
                 If (i = 1) And (Left(Tmp2, 1) = "$") Then Tmp2 = Mid(Tmp2, 2)  'If Addr begins with "$" remove it!
                 Out = Out & Tmp2                               'Build the string
                 If i < 3 Then Out = Out & ","                  'Add a comma
            Next
            
            Tmp = Left(Out, 4)                                  'Check Hex
            If Tmp >= "0000" And Tmp <= "FFFF" Then
                C = C + 1: lstSYM.AddItem Out                   'Add it to the symbol list
            End If
        End If
    Wend
    Close FIO
    
    MsgBox "File imported! " & Str(C) & " symbols loaded."
    MLReViewC
    
End Sub

'---- Purge Un-selected entries from SYMBOL table
Private Sub cmdPurge_Click()
    Dim i As Integer
    
    For i = lstSYM.ListCount - 1 To 0 Step -1
        If lstSYM.Selected(i) = False Then lstSYM.RemoveItem (i)
    Next i
    MLReViewC
End Sub

'---- Remove Duplicate Generated Label entries
Private Sub cmdRemDupLbls_Click()
    Dim i As Integer
    
    For i = lstLabels.ListCount - 1 To 1 Step -1
       If lstLabels.List(i) = lstLabels.List(i - 1) Then lstLabels.RemoveItem (i)
    Next i
    
End Sub

'---- Remove Duplicate External JSR entries
Private Sub cmdRemDupJSR_Click()
    Dim i As Integer
    
    For i = lstJSR.ListCount - 1 To 1 Step -1
       If lstJSR.List(i) = lstJSR.List(i - 1) Then lstJSR.RemoveItem (i)
    Next i
    
End Sub

'---- Toggle listing colours
Private Sub imgBW_Click()
    If lstML.BackColor = vbWhite Then
        lstML.BackColor = vbBlack: lstML.ForeColor = vbWhite
    Else
        lstML.BackColor = vbWhite: lstML.ForeColor = vbBlack
    End If
End Sub

'---- Display HELP file
Private Sub cmdMLHelp_Click()
    ViewFile ExeDir & "\ml-help.txt"
End Sub

'---- Load ML Config File
' The ML Config file contains lines to be loaded into the drop-down menus along with the specified file resource
' Each table group must be proceeded by a selection marker:
' [PLATFORM] [CPU] [PREFIX]

Private Sub LoadMLConfig()
    Dim FIO As Integer, Tmp As String, Tmp2 As String, TMode As Integer, Filename As String
    Dim c1 As Integer, C2 As Integer
    
    Filename = ExeDir & "ml-config.txt"
    If Exists(Filename) = False Then MsgBox "ML Config file is missing!": Exit Sub
        
    FIO = FreeFile
    Open Filename For Input As FIO
    
    TMode = 0: c1 = 0: C2 = 0
    ViewerReady = False
        
    While Not EOF(FIO)
        Line Input #FIO, Tmp
        If Left(Tmp, 1) = "[" Then
            '---- Check for section marker
            Select Case Tmp
                Case "[PLATFORM]": TMode = 1: cboPlatform.Clear
                Case "[CPU]":      TMode = 2: cboCPU.Clear
                Case "[PREFIX]":   TMode = 3: cboPrefix.Clear
                Case Else: TMode = 0
            End Select
        Else
            If (Left(Tmp, 1) <> ";") And (Tmp <> "") Then
                p = InStr(1, Tmp, ",") 'look for comma separator
                '---- Process according to current section marker
                Select Case TMode
                    Case 1 '-- PLATFORM
                        If p > 0 Then
                            Tmp2 = Left(Tmp, p - 1)
                            cboPlatform.List(c1) = Tmp2
                            cboPlatFile.List(c1) = Mid(Tmp, p + 1)
                            c1 = c1 + 1
                        End If

                    Case 2 '-- CPU
                        If p > 0 Then
                            Tmp2 = Left(Tmp, p - 1)
                            cboCPU.List(C2) = Tmp2
                            cboCPUFile.List(C2) = Mid(Tmp, p + 1)
                            C2 = C2 + 1
                        End If
                        
                    Case 3 '-- Prefix
                        cboPrefix.AddItem Tmp
                End Select
            End If
        End If
    Wend
    cboPlatform.ListIndex = 0
    cboCPU.ListIndex = 0
    cboPrefix.ListIndex = 0
    Close FIO
    MLCFlag = True
    ViewerReady = True
End Sub

'=================
'HEX/Binary Viewer
'=================
Sub HEXView()

    Dim C As Single, W As Integer, h As Integer
    Dim Tmp As String, TLine As String, ALine As String, LCount As Integer
    Dim Flag As Boolean, MaxW As Integer
    Dim Lo As Integer, Hi As Integer, Address As Long, BMASK As Integer, CBMFlag As Boolean
   
    BMASK = 255: If cb7bit.value = vbChecked Then BMASK = 127 'Enable 7-bit view
    lstBIN.Clear
    
    If cbWide.value = vbChecked Then MaxW = 15 Else MaxW = 7
    Flag = False: If cbShowP.value = vbUnchecked Then Flag = True           'Show Printable
    CBMFlag = False: If cbShowCBM.value = vbChecked Then CBMFlag = True     'Show CBM
    
    C = 0: W = 0: Tmp = "": TLine = "": ALine = "": LCount = 0  'Initialize
    
    If cbHexSync.value = vbChecked Then
        Address = MyDec(txtLA.Text)                             'Use Address specified in ASM project
    Else
        Address = VLA                                           'Use Load Address from file
    End If
    
    '-- Loop through buffer
    Do
        If W > MaxW Then
            If Flag = False Then lstBIN.AddItem TLine & ALine
            If Flag = True Then lstBIN.AddItem TLine
            W = 0: LCount = LCount + 1
        End If
        
        W = W + 1: If W = 1 Then TLine = MyHex(Address, 4) & ": ": ALine = "> "
        C = C + 1: Address = Address + 1
        Tmp = Mid(VBuf, C, 1): h = Asc(Tmp)
        
        TLine = TLine & MyHex(h, 2) & " "
        
        Select Case (h And BMASK)
            Case 0 To 31
                If CBMFlag = True Then
                    ALine = ALine & Chr((h And Mask) + 64)      'Converts CTRL chrs to Letter range
                Else
                    ALine = ALine & "."                         'Un-Printable
                End If
            Case 32 To 127: ALine = ALine & Chr(h And BMASK)    'Printable
            Case Else: ALine = ALine & "."                      'Un-Printable
        End Select
        
    Loop While (C < VLen) 'And (LCount < 32766)
    
    If TLine <> "" Then
        If Flag = False Then lstBIN.AddItem TLine & ALine
        If Flag = True Then lstBIN.AddItem TLine
    End If
    
    If lstBIN.Visible = True Then lstBIN.SetFocus
    
End Sub

'---- Toggle HEX listing colours
Private Sub imgBWH_Click()
    If lstBIN.BackColor = vbWhite Then
        lstBIN.BackColor = vbBlack: lstBIN.ForeColor = vbWhite
    Else
        lstBIN.BackColor = vbWhite: lstBIN.ForeColor = vbBlack
    End If

End Sub

'---- Sync HEX view with FONT Offset
Private Sub lstBIN_Click()
    Dim Tmp As String
    
    If frFont.Visible = True Then
        Tmp = Left(lstBIN.List(lstBIN.ListIndex), 4)    'Get the HEX address
        txtCSkip.Text = Format(MyDec(Tmp))              'Convert it to decimal and store it in the offset field
        FONTView                                        'Re-display font
    End If
End Sub

'============
'SEQ Viewer
'============
Sub SEQView()
    Dim FIO As Integer, C As Integer, Tmp As String, TLine As String, h As Integer

    lstSEQ.Clear
    
    C = 1: Tmp = "": TLine = ""
    Do
        If Len(TLine) > 80 Then lstSEQ.AddItem TLine: TLine = ""
        Tmp = Mid(VBuf, C, 1): h = Asc(Tmp)
        Select Case h
            Case 32 To 127: TLine = TLine & Tmp
            Case 192 To 218: TLine = TLine & Chr(h And 127)
            Case 10: If cbIgnoreLF.value <> vbChecked Then lstSEQ.AddItem TLine: TLine = ""
            Case 13: lstSEQ.AddItem TLine: TLine = ""
            Case Else
        End Select
        C = C + 1
    Loop While (C < VLen) And (C < 32767)
    
    If TLine <> "" Then lstSEQ.AddItem TLine
    
End Sub

'---- Set SEQ Background Colour
Private Sub lblSEQTheme_Click(Index As Integer)
    Dim C As Long
    
    frmColourPicker.Show vbModal
    If PickedColour < 0 Then Exit Sub
    
    lblSEQTheme(Index).BackColor = PickedColour
    
    Select Case Index
        Case 0: lstSEQ.ForeColor = PickedColour
        Case 1: lstSEQ.BackColor = PickedColour
    End Select

End Sub

'---- Change SEQ View Font
Private Sub cbSeqFont_Click()
    If cbSeqFont.value = 1 Then
        lstSEQ.Font = "SEQ VIEWER"
    Else
        lstSEQ.Font = "MS Sans Serif"
    End If
End Sub


'===============
' Bitmap Viewer
'===============

Private Sub BMPView()
    Dim Comment As String, i As Integer, X As Integer, FLen As Long
    Dim TwipX As Integer, TwipY As Integer, LAOffset As Integer
        
    TwipX = Screen.TwipsPerPixelX
    TwipY = Screen.TwipsPerPixelY
    
    If p_name(0) = "" Then Call LoadPicFormats                      'Load picture formats if needed
        
    lblBComment.Caption = "None"                                    'Clear comment
    LAOffset = 0: If cbLA.value = vbChecked Then LAOffset = 2       'Compensate for LA if selected
    ImageType = HRBW                                                'Default to hi-res mono
    
    Picture1.Visible = False                                        'Hide the picture
    lblMoment.Visible = True                                        'Display loading message
    DoEvents
    
    '-- Read shared buffer and determin what type of bitmap file it is
    If Mid(VBuf, 22 - LAOffset, 2) = Chr(1) & Chr(7) Then
        If Mid(VBuf, 330 - LAOffset, 11) = "Paint Image" Then
            '-- GeoPaint Image
            Comment = Mid(VBuf, 413 - LAOffset, 256)
            X = InStr(1, Comment, Nu): If X > 0 Then Comment = Left(Comment, X - 1)
            lblBType.Caption = "GeoPaint Image"
            lblBComment.Caption = Comment
            ImageType = GEO
            Picture1.Width = 640 * TwipX
            Picture1.Height = 720 * TwipY
            Read_GeoPaint VFileName
        End If
    Else
        '-- Search for matching image parameters
        FLen = VLen: If cbLA.value = vbChecked Then FLen = FLen + 2
        For i = 1 To NUMB
            If (p_sa(i) = VLA) And (p_len(i) = FLen) Then
                ImageType = i
                Exit For
            End If
        Next i
    
        lblBType.Caption = p_name(ImageType)
        
        Picture1.Width = 320 * TwipX                        'Standard 320x200 bitmap
        Picture1.Height = 200 * TwipY
        Read_Bitmap VFileName
    End If
    
    Picture1.Visible = True                                 'Hide the picture
    lblMoment.Visible = False                               'Display loading message
    DoEvents
    
End Sub

'---- Read GeoPaint Image
Private Sub Read_GeoPaint(ByVal Filename As String)
    Dim Dat As String, M As String, FIO As Integer
    Dim c0 As Long, c1 As Long 'Pixel on and off colours - new May 2017
    
    ReDim blocks(0 To 44, 1 To 2)
    ReDim pat(0 To 7)
    
    Close PFIO
        
    PFIO = FreeFile
    Open Filename For Binary As PFIO
    
    PBuf = ReadBlock()                                      'First sector
    PBuf = ReadBlock()                                      'Second sector
    PBuf = ReadBlock()                                      'Third sector
    
    validsectors = 0: sector = 0

    Picture1.Cls
    
    c0 = CBMColor(1)                                        'White Background Colour - new 2017
    c1 = CBMColor(0)                                        'Black Foreground Colour - new 2017
    Picture1.BackColor = c0                                 'Default to white background
    Picture1.Cls                                            'Clear to background colour
    
    For i = 0 To 44
      M = Left(PBuf, 2)
      blocks(i, 1) = Asc(M)
      If blocks(i, 1) <> 0 Then
        blocks(i, 2) = Asc(Right(M, 1))
        validsectors = validsectors + 1
      End If
      PBuf = Mid(PBuf, 3)
    Next i
    
    '-- Display loop
    For i = 0 To 44
        If blocks(i, 1) > 0 Then
            Dat = ""
            For j = 1 To blocks(i, 1)
                PBuf = ReadBlock()
                If j = blocks(i, 1) Then PBuf = Left(PBuf, blocks(i, 2))
                Dat = Dat & PBuf
            Next j
            
            bitposh = 0:  bitposv = 0
            
            dpos = 1
            ldat = Len(Dat)
            
            DoEvents
            
            Do While (bitposv < 16) And (ldat >= dpos)
                nxt = Asc(Mid(Dat, dpos, 1) & Nu): dpos = dpos + 1
                
                Select Case nxt
                  Case 1 To 63
                    For K = 1 To nxt
                      Pel = Asc(Mid(Dat, dpos, 1) & Nu): dpos = dpos + 1
                      GoSub PaintBit
                    Next K
                    
                  Case 65 To 127
                    For K = 0 To 7
                      pat(K) = Asc(Mid(Dat, dpos, 1) & Nu): dpos = dpos + 1
                    Next K
                    
                    For L = 1 To (nxt And 63)
                      For K = 0 To 7
                        Pel = pat(K): GoSub PaintBit
                      Next K
                    Next L
                    
                  Case 129 To 255
                    DT = Asc(Mid(Dat, dpos, 1) & Nu): dpos = dpos + 1
                    For K = 1 To (nxt - 128)
                      Pel = DT
                      GoSub PaintBit
                    Next K
                End Select
            Loop
            
            sector = sector + 1
        End If
    Next i
    
    Close PFIO
Exit Sub

'---- Paint Bits
PaintBit:
    XX = bitposh * 8 + 7
    YY = i * 16 + bitposv
    
    For k2 = 0 To 7
        If (Pel And Pow(k2)) Then Picture1.PSet (XX - k2, YY), c1 'Set Black dot
    Next k2
    
    bitposv = bitposv + 1
    
    If bitposv = 8 Or bitposv = 16 Then
        bitposh = bitposh + 1: bitposv = bitposv - 8
        If bitposh > 79 Then bitposh = bitposh - 80: bitposv = bitposv + 8
    End If
    
    Return

End Sub

Private Sub Read_Bitmap(ByVal Filename As String)
    Dim Bitmap As String, Scrn As String, Col As String, Bk As String, Pel As Integer, BG As Integer
        
    '-- Allocate space for data
    Bitmap = Space(8000)
    Scrn = Space(1000)
    Col = Space(1000)
    Bk = Chr(1)
    
    PFIO = FreeFile
    Open Filename For Binary As PFIO
    
    '-- Read data from File. ImageType is from comparing load address and file size
    '   The arrays hold the offset to the data inside the file. The variable lengths determine how much data is loaded.
    Select Case p_type(ImageType)
        Case HRBW
            Get #PFIO, 3, Bitmap
        Case HR
            Get #PFIO, p_bitmap(ImageType) + 3, Bitmap
            Get #PFIO, p_screen(ImageType) + 3, Scrn
        Case MC
            Get #PFIO, p_bitmap(ImageType) + 3, Bitmap
            Get #PFIO, p_screen(ImageType) + 3, Scrn
            Get #PFIO, p_colour(ImageType) + 3, Col
            Get #PFIO, p_back(ImageType) + 3, Bk
    End Select
    
    Close PFIO
    
    bitposh = 0: bitposv = 0: dpos = 1: CPos = 1
    BG = Asc(Bk)
    
    Picture1.Cls
    DoEvents
        
    Do While (bitposv < 200)
    
        Pel = Asc(Mid(Bitmap, dpos, 1))
        dpos = dpos + 1
        XX = bitposh * 8 + 7
        YY = 0
        Select Case p_type(ImageType)
            Case HRBW 'High-res Mono Mode
                For k2 = 0 To 7
                    Picture1.PSet (XX - k2, bitposv), IIf(Pel And Pow(k2), CBMColor(0), CBMColor(1))
                Next k2
                
            Case HR 'High-res Colour Mode
                S = Asc(Mid(Scrn, CPos, 1))
                For k2 = 0 To 7
                    Picture1.PSet (XX - k2, bitposv), IIf(Pel And Pow(k2), CBMColor((S And 240) / 16), CBMColor(S And 15))
                Next k2
                
            Case MC 'Multi-colour Mode
                S = Asc(Mid(Scrn, CPos, 1))
                C = Asc(Mid(Col, CPos, 1))
                For k2 = 0 To 6 Step 2
                    k3 = Pow(k2)
                    Bit$ = IIf(Pel And (k3 * 2), "1", "0")
                    Bit$ = Bit$ & IIf(Pel And k3, "1", "0")
                    Select Case Bit$
                        Case "00": colput& = CBMColor(BG)
                        Case "10": colput& = CBMColor(S And 15)
                        Case "01": colput& = CBMColor((S And 240) / 16)
                        Case "11": colput& = CBMColor(C And 15)
                    End Select
                    Picture1.PSet (XX - k2, bitposv), colput&
                    Picture1.PSet (XX - k2 - 1, bitposv), colput&
                Next k2
        End Select
    
        bitposv = bitposv + 1
        If bitposv / 8 = bitposv \ 8 Then
            bitposh = bitposh + 1: bitposv = bitposv - 8
            CPos = CPos + 1
            If bitposh > 39 Then bitposh = bitposh - 40: bitposv = bitposv + 8
        End If
    Loop

End Sub

Private Sub LoadPicFormats()
    Dim Filename As String, Tmp As String, FIO As Integer
       
    Num = 0
    p_name(0) = "Hi-Res B/W Image"
    p_type(0) = HRBW
    
    Filename = ExeDir & "picformats.txt"
    If Exists(Filename) = False Then MsgBox "Picture formats file missing!!!": Exit Sub
    
    FIO = FreeFile
    Open Filename For Input As FIO
        
    Do Until (Num >= NUMB) Or (EOF(FIO) = True)
        Line Input #FIO, Tmp
        If Left(Tmp, 1) >= "A" Then
            Num = Num + 1
            p_name(Num) = Mid(Tmp, 1, 21)
            p_sa(Num) = MyDec(Mid(Tmp, 24, 4))
            p_len(Num) = Val(Mid(Tmp, 31, 5))
            p_bitmap(Num) = Val(Mid(Tmp, 39, 4))
            p_screen(Num) = Val(Mid(Tmp, 47, 4))
            p_type(Num) = MC: If Mid(Tmp, 55, 1) = "-" Then p_type(Num) = HR  'Multicolour or Hires?
            p_colour(Num) = Val(Mid(Tmp, 55, 5))
            p_back(Num) = Val(Mid(Tmp, 63, 5))
        End If
    Loop
    
    Close FIO
End Sub

'---- Save the Bitmap
Private Sub cmdBSave_Click()
    Dim Filename As String
    
    Filename = FileOpenSave(FileBase(VFileName), 1, 3, "Save Image as BMP")
    If Filename <> "" Then SavePicture Picture1.Image, Filename
End Sub

'---- Reads a chunk of 256 bytes from the bitmap file
Private Function ReadBlock() As String
    Dim buf As String
    
    buf = Space(254)
    Get #PFIO, , buf
    ReadBlock = buf
End Function

'---- Set Default colours to VIC-II palette
Private Sub SetColor()
    Dim i As Integer
    
    For i = 0 To 15
        CBMColor(i) = C64Colour(i)
    Next i
End Sub

'---- Select a Colour from Colour Dialog
Private Function PickColor() As Long
    On Local Error GoTo NoPick
    
    CommonDialog.CancelError = True
    CommonDialog.ShowColor
    PickColor = CommonDialog.Color
    Exit Function
    
NoPick:
    PickColor = -1
End Function

'--- Common File Open or Save Dialog
' You can specify a default filename, a File Filter list index (0-5), and Window Title
' MODE: 0=Open, 1=Save
' Returns a filename with full path. If cancelled will return null string
Private Function FileOpenSave(ByVal DefFile As String, ByVal Mode As Integer, FiltSet As Integer, DTitle As String) As String
    Dim Filename As String
    
    CommonDialog.CancelError = True
    On Local Error GoTo NoFile
        
    CommonDialog.InitDir = PathOnly(DefFile)
    CommonDialog.DialogTitle = DTitle
    CommonDialog.Flags = cdlOFNHideReadOnly
    CommonDialog.Filename = DefFile
    Select Case FiltSet
        Case 0: CommonDialog.Filter = "All files (*.*)|*.*"
        Case 1: CommonDialog.Filter = "Symbol Table Files (*.SYM,*.DT,*.TXT,*.SY4)|*.SYM;*.DT;*.TXT;*.SY4"
        Case 2: CommonDialog.Filter = "ASM Project Files (*.ASM-PROJ)|*.ASM-PROJ"
        Case 3: CommonDialog.Filter = "Bitmap Files(*.BMP)|*.BMP"
        Case 4: CommonDialog.Filter = "ASM Files(*.ASM,*.TXT)|*.ASM;*.TXT"
        Case 5: CommonDialog.Filter = "Text Files(*.TXT)|*.TXT"
    End Select
    
    If Mode = 0 Then CommonDialog.ShowOpen Else CommonDialog.ShowSave   'MODE: 0=Open, 1=Save
        
    If CommonDialog.Filename = "" Then Exit Function
    
    FileOpenSave = CommonDialog.Filename
    Exit Function
NoFile:

End Function

'==========================================================
' Controls that cause a Re-load of file and refresh of output (From any Viewer)
'==========================================================
Private Sub cbLA_Click()
    ViewIt ViewMode, VFileName, VName, VExt 're-load the file
End Sub

'==========================================================
' Controls that cause a refresh of output (From any Viewer)
'==========================================================

'---- ML Updates
Private Sub cbLabelBlanks_Click()
    MLReViewA
End Sub
Private Sub cbSpaceRTS_Click()
    MLReViewA
End Sub
Private Sub cbEquates_Click()
    MLReViewA
End Sub
Private Sub cbIncSym_Click()
    MLReViewA
End Sub
Private Sub cmdRefresh_Click()
    MLReView
End Sub
Private Sub cbBytes_Click()
    MLReView
End Sub
Private Sub cboMLFmt_Click()
    MLReView
End Sub
Private Sub txtLA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then MLView
End Sub

'BASIC Updates
Private Sub cboMode_Click()
    If ViewReady Then BASView
End Sub
Private Sub cbRev_Click()
    BASView
End Sub
Private Sub cbUseFont_Click()
    BASView
End Sub
Private Sub cbExp_Click()
    BASView
End Sub
Private Sub cbUC_Click()
    BASView
End Sub
Private Sub cbOneLine_Click()
    BASView
End Sub
Private Sub cbPad_Click()
    BASView
End Sub

'---- HEX updates
Private Sub cb7bit_Click()
    HEXView
End Sub
Private Sub cbWide_Click()
    HEXView
End Sub
Private Sub cbHexSync_Click()
    HEXView
End Sub
Private Sub cbShowCBM_Click()
    HEXView
End Sub
Private Sub cbShowP_Click()
    HEXView
End Sub

'---- Font Updates
Private Sub optChrH_Click(Index As Integer)
    FONTView
End Sub
Private Sub cbFCols_Click()
    FONTView
End Sub

Private Sub cbBorder_Click()
    FONTView
End Sub

'---- SEQ Updates
Private Sub cbIgnoreLF_Click()
    SEQView
End Sub

