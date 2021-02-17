VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmViewer 
   Caption         =   "Viewer:"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11925
   Icon            =   "frmViewer.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   8400
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
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
      Height          =   1185
      Left            =   60
      TabIndex        =   6
      Top             =   420
      Visible         =   0   'False
      Width           =   14790
      Begin VB.CheckBox cbCmpShow 
         Caption         =   "Show"
         Height          =   285
         Left            =   7920
         TabIndex        =   244
         ToolTipText     =   "File includes Load Address at start"
         Top             =   480
         Width           =   885
      End
      Begin VB.CommandButton cmdCompare 
         Caption         =   "Compare..."
         Height          =   315
         Left            =   7920
         TabIndex        =   243
         ToolTipText     =   "Find NEXT occurance"
         Top             =   150
         Width           =   915
      End
      Begin VB.CheckBox cbHexFmt 
         Caption         =   "ASM Fmt"
         Height          =   285
         Left            =   5850
         TabIndex        =   242
         ToolTipText     =   "File includes Load Address at start"
         Top             =   540
         Width           =   945
      End
      Begin VB.CommandButton cmdHSave 
         Caption         =   "Save"
         Height          =   315
         Left            =   60
         TabIndex        =   241
         ToolTipText     =   "Save Current view as TXT"
         Top             =   210
         Width           =   675
      End
      Begin VB.CommandButton cmdHNext 
         Caption         =   "Next"
         Height          =   315
         Left            =   5040
         TabIndex        =   235
         ToolTipText     =   "Find NEXT occurance"
         Top             =   150
         Width           =   555
      End
      Begin VB.CommandButton cmdHexFind 
         Caption         =   "Find"
         Height          =   315
         Left            =   4350
         TabIndex        =   234
         Top             =   150
         Width           =   645
      End
      Begin VB.TextBox txtHSS 
         Height          =   315
         Left            =   1380
         TabIndex        =   232
         ToolTipText     =   "Enter text string to search for, or start with ""$""  to search for HEX value(s)"
         Top             =   180
         Width           =   2895
      End
      Begin VB.CheckBox cbShowCBM 
         Caption         =   "Show CBM"
         Height          =   195
         Left            =   3540
         TabIndex        =   129
         ToolTipText     =   "Show CBM screen codes"
         Top             =   570
         Width           =   1095
      End
      Begin VB.CheckBox cbHexSync 
         Caption         =   "ASM Sync"
         Height          =   195
         Left            =   4740
         TabIndex        =   128
         ToolTipText     =   "File includes Load Address at start"
         Top             =   570
         Width           =   1095
      End
      Begin VB.CheckBox cbWide 
         Caption         =   "Wide"
         Height          =   195
         Left            =   420
         TabIndex        =   111
         ToolTipText     =   "File includes Load Address at start"
         Top             =   570
         Value           =   1  'Checked
         Width           =   705
      End
      Begin VB.CheckBox cb7bit 
         Caption         =   "7-bit"
         Height          =   195
         Left            =   2760
         TabIndex        =   32
         ToolTipText     =   "Enable 7-bit View"
         Top             =   570
         Width           =   615
      End
      Begin VB.CheckBox cbShowP 
         Caption         =   "Show Printable"
         Height          =   195
         Left            =   1260
         TabIndex        =   16
         Top             =   570
         Value           =   1  'Checked
         Width           =   1365
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
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label lblCFile 
         BackColor       =   &H80000018&
         Caption         =   "no file"
         Height          =   225
         Left            =   8880
         TabIndex        =   246
         Top             =   210
         Width           =   2895
      End
      Begin VB.Label lblDifTxt 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Height          =   195
         Left            =   8850
         TabIndex        =   245
         Top             =   510
         Width           =   2925
      End
      Begin VB.Label lblSResults 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "No search set"
         Height          =   195
         Left            =   5700
         TabIndex        =   236
         Top             =   210
         Width           =   990
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "FIND:"
         Height          =   195
         Left            =   900
         TabIndex        =   233
         Top             =   270
         Width           =   420
      End
      Begin VB.Image imgBWH 
         Height          =   255
         Left            =   90
         Picture         =   "frmViewer.frx":0442
         Top             =   540
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
      Height          =   7440
      Left            =   60
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   14925
      Begin VB.Frame frTools 
         Height          =   6585
         Left            =   90
         TabIndex        =   192
         Top             =   720
         Width           =   1515
         Begin VB.CommandButton cmdTool 
            Caption         =   "Set R Point"
            Height          =   285
            Index           =   30
            Left            =   60
            TabIndex        =   230
            Top             =   5610
            Width           =   1365
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Ins R"
            Height          =   285
            Index           =   26
            Left            =   60
            TabIndex        =   229
            ToolTipText     =   "Insert ROW below crosshair"
            Top             =   2820
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Del R"
            Height          =   285
            Index           =   27
            Left            =   750
            TabIndex        =   228
            ToolTipText     =   "Delete ROW below crosshair"
            Top             =   2820
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Ins C"
            Height          =   285
            Index           =   28
            Left            =   60
            TabIndex        =   227
            ToolTipText     =   "Insert COL to right of crosshair"
            Top             =   3150
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Del C"
            Height          =   285
            Index           =   29
            Left            =   750
            TabIndex        =   226
            ToolTipText     =   "Delete COL to right of crosshair"
            Top             =   3150
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Restore Original"
            Height          =   285
            Index           =   25
            Left            =   60
            TabIndex        =   221
            Top             =   6210
            Width           =   1365
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Restore"
            Height          =   285
            Index           =   24
            Left            =   60
            TabIndex        =   219
            ToolTipText     =   "Restore character from original loaded font"
            Top             =   5910
            Width           =   1365
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "PASTE"
            Height          =   285
            Index           =   23
            Left            =   750
            TabIndex        =   217
            ToolTipText     =   "Paste clipboard to current selected position"
            Top             =   5190
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "COPY"
            Height          =   285
            Index           =   22
            Left            =   60
            TabIndex        =   216
            ToolTipText     =   "Copy CHR or RANGE to clipboard"
            Top             =   5190
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Sel ALL"
            Height          =   255
            Index           =   21
            Left            =   750
            TabIndex        =   215
            ToolTipText     =   "Select ALL"
            Top             =   4770
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "SWAP"
            Height          =   255
            Index           =   20
            Left            =   60
            TabIndex        =   214
            ToolTipText     =   "Swap character sets"
            Top             =   4770
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Und"
            Height          =   285
            Index           =   7
            Left            =   750
            TabIndex        =   213
            ToolTipText     =   "Create Underlined character (below crosshair)"
            Top             =   1830
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Rot R"
            Height          =   285
            Index           =   9
            Left            =   750
            TabIndex        =   212
            ToolTipText     =   "Rotate character RIGHT"
            Top             =   2160
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Rot L"
            Height          =   285
            Index           =   8
            Left            =   60
            TabIndex        =   211
            ToolTipText     =   "Rotate character LEFT"
            Top             =   2160
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "2x BR"
            Height          =   285
            Index           =   19
            Left            =   750
            TabIndex        =   210
            ToolTipText     =   "Create 2x BOTTOM-RIGHT"
            Top             =   4440
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "2x BL"
            Height          =   285
            Index           =   18
            Left            =   60
            TabIndex        =   209
            ToolTipText     =   "Create 2x BOTTOM-LEFT"
            Top             =   4440
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "2x TR"
            Height          =   285
            Index           =   17
            Left            =   750
            TabIndex        =   208
            ToolTipText     =   "Create 2x TOP-RIGHT"
            Top             =   4110
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "2x TL"
            Height          =   285
            Index           =   16
            Left            =   60
            TabIndex        =   207
            ToolTipText     =   "Create 2x TOP-LEFT"
            Top             =   4110
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Wide R"
            Height          =   285
            Index           =   15
            Left            =   750
            TabIndex        =   206
            ToolTipText     =   "Create Double-Wide RIGHT "
            Top             =   3780
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Wide L"
            Height          =   285
            Index           =   14
            Left            =   60
            TabIndex        =   205
            ToolTipText     =   "Create Double-Wide LEFT"
            Top             =   3780
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Tall B"
            Height          =   285
            Index           =   13
            Left            =   750
            TabIndex        =   204
            ToolTipText     =   "Create Double-Tall BOTTOM"
            Top             =   3450
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Tall T"
            Height          =   285
            Index           =   12
            Left            =   60
            TabIndex        =   203
            ToolTipText     =   "Create Double-Tall TOP"
            Top             =   3450
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Flip V"
            Height          =   285
            Index           =   11
            Left            =   750
            TabIndex        =   202
            ToolTipText     =   "Flip character left to right"
            Top             =   2490
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Flip H"
            Height          =   285
            Index           =   10
            Left            =   60
            TabIndex        =   201
            ToolTipText     =   "Flip character top to bottom"
            Top             =   2490
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Bold"
            Height          =   285
            Index           =   6
            Left            =   60
            TabIndex        =   200
            ToolTipText     =   "Create Bold character"
            Top             =   1830
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "RVS"
            Height          =   285
            HelpContextID   =   5
            Index           =   5
            Left            =   750
            TabIndex        =   198
            ToolTipText     =   "Invert pixels"
            Top             =   1500
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Clear"
            Height          =   285
            Index           =   4
            Left            =   60
            TabIndex        =   197
            ToolTipText     =   "Clear Character to Bg"
            Top             =   1500
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Height          =   435
            Index           =   1
            Left            =   540
            Picture         =   "frmViewer.frx":07F8
            Style           =   1  'Graphical
            TabIndex        =   196
            ToolTipText     =   "Shift DOWN"
            Top             =   630
            Width           =   435
         End
         Begin VB.CommandButton cmdTool 
            Height          =   435
            Index           =   3
            Left            =   990
            Picture         =   "frmViewer.frx":0EFA
            Style           =   1  'Graphical
            TabIndex        =   195
            ToolTipText     =   "Shift RIGHT"
            Top             =   390
            Width           =   435
         End
         Begin VB.CommandButton cmdTool 
            Height          =   435
            Index           =   2
            Left            =   90
            Picture         =   "frmViewer.frx":15FC
            Style           =   1  'Graphical
            TabIndex        =   194
            ToolTipText     =   "Shift LEFT"
            Top             =   390
            Width           =   435
         End
         Begin VB.CommandButton cmdTool 
            Height          =   435
            Index           =   0
            Left            =   540
            Picture         =   "frmViewer.frx":1CFE
            Style           =   1  'Graphical
            TabIndex        =   193
            ToolTipText     =   "Shift UP"
            Top             =   180
            Width           =   435
         End
         Begin VB.CheckBox cbShiftMode 
            Height          =   255
            Left            =   1110
            TabIndex        =   199
            ToolTipText     =   "When checked pixels wrap to opposite side. When unset pixels are LOST!"
            Top             =   840
            Value           =   1  'Checked
            Width           =   195
         End
         Begin VB.Label lblPixelMode 
            Alignment       =   2  'Center
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Xor"
            Height          =   270
            Index           =   2
            Left            =   990
            TabIndex        =   225
            ToolTipText     =   "Toggle pixel colour"
            Top             =   1140
            Width           =   405
         End
         Begin VB.Label lblPixelMode 
            Alignment       =   2  'Center
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "FG"
            Height          =   270
            Index           =   1
            Left            =   540
            TabIndex        =   224
            ToolTipText     =   "Draw using Foreground colour"
            Top             =   1140
            Width           =   405
         End
         Begin VB.Label lblPixelMode 
            Alignment       =   2  'Center
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "BG"
            Height          =   270
            Index           =   0
            Left            =   90
            TabIndex        =   223
            ToolTipText     =   "Draw using Background colour"
            Top             =   1140
            Width           =   405
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Shift:"
            Height          =   195
            Left            =   60
            TabIndex        =   218
            Top             =   210
            Width           =   360
         End
      End
      Begin VB.Frame frChr 
         Height          =   6585
         Left            =   1680
         TabIndex        =   185
         Top             =   720
         Width           =   2205
         Begin VB.PictureBox picChr 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   3960
            Left            =   90
            ScaleHeight     =   260
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   131
            TabIndex        =   188
            Top             =   750
            Width           =   2025
         End
         Begin VB.CommandButton cmdCSNxt 
            Caption         =   ">"
            Height          =   255
            Left            =   1770
            TabIndex        =   187
            ToolTipText     =   "Next character"
            Top             =   180
            Width           =   300
         End
         Begin VB.CommandButton cmdCSPrev 
            Caption         =   "<"
            Height          =   255
            Left            =   1470
            TabIndex        =   186
            ToolTipText     =   "Previous character"
            Top             =   180
            Width           =   270
         End
         Begin VB.Label lblFStat 
            BackColor       =   &H80000003&
            Caption         =   "M"
            Height          =   1695
            Left            =   90
            TabIndex        =   222
            Top             =   4770
            Width           =   1995
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblRange 
            BackColor       =   &H0080C0FF&
            Height          =   3885
            Left            =   90
            TabIndex        =   191
            Top             =   780
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.Label lblChrNum 
            Alignment       =   2  'Center
            BackColor       =   &H0080C0FF&
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1470
            TabIndex        =   190
            Top             =   450
            Width           =   615
         End
         Begin VB.Label lblChrX 
            BackColor       =   &H0080C0FF&
            Height          =   525
            Left            =   90
            TabIndex        =   189
            Top             =   180
            Width           =   1305
         End
      End
      Begin VB.Frame frControls 
         Height          =   555
         Left            =   450
         TabIndex        =   157
         Top             =   180
         Width           =   14355
         Begin VB.CheckBox cbOutline 
            Caption         =   "Outline"
            Height          =   225
            Left            =   13380
            TabIndex        =   240
            Top             =   210
            Width           =   795
         End
         Begin VB.HScrollBar HScroll1 
            CausesValidation=   0   'False
            Height          =   285
            LargeChange     =   2
            Left            =   12540
            Max             =   16
            Min             =   1
            TabIndex        =   239
            Top             =   180
            Value           =   1
            Width           =   735
         End
         Begin VB.ComboBox cboTheme 
            Height          =   315
            ItemData        =   "frmViewer.frx":2400
            Left            =   2700
            List            =   "frmViewer.frx":241F
            Style           =   2  'Dropdown List
            TabIndex        =   165
            Top             =   180
            Width           =   1245
         End
         Begin VB.TextBox txtCSkip 
            Height          =   285
            Left            =   8070
            TabIndex        =   164
            Text            =   "0"
            ToolTipText     =   "Set number of bytes to skip (decimal)"
            Top             =   180
            Width           =   675
         End
         Begin VB.CommandButton cmdSB 
            Caption         =   "<<"
            Height          =   270
            Index           =   0
            Left            =   8790
            TabIndex        =   163
            Top             =   180
            Width           =   315
         End
         Begin VB.CommandButton cmdSB 
            Caption         =   "<"
            Height          =   270
            Index           =   1
            Left            =   9120
            TabIndex        =   162
            Top             =   180
            Width           =   255
         End
         Begin VB.CommandButton cmdSB 
            Caption         =   "-"
            Height          =   270
            Index           =   2
            Left            =   9390
            TabIndex        =   161
            Top             =   180
            Width           =   255
         End
         Begin VB.CommandButton cmdSB 
            Caption         =   "+"
            Height          =   270
            Index           =   3
            Left            =   9660
            TabIndex        =   160
            Top             =   180
            Width           =   315
         End
         Begin VB.CommandButton cmdSB 
            Caption         =   ">"
            Height          =   270
            Index           =   4
            Left            =   9990
            TabIndex        =   159
            Top             =   180
            Width           =   315
         End
         Begin VB.CommandButton cmdSB 
            Caption         =   ">>"
            Height          =   270
            Index           =   5
            Left            =   10320
            TabIndex        =   158
            Top             =   180
            Width           =   315
         End
         Begin VB.Label lblBorderSize 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   12210
            TabIndex        =   238
            Top             =   180
            Width           =   285
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Border:"
            Height          =   195
            Left            =   11670
            TabIndex        =   237
            Top             =   210
            Width           =   510
         End
         Begin VB.Label lblWidth 
            Alignment       =   2  'Center
            BackColor       =   &H00008000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "8"
            Height          =   300
            Index           =   0
            Left            =   5940
            TabIndex        =   220
            Top             =   180
            Width           =   285
         End
         Begin VB.Label lblTheme 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   1410
            TabIndex        =   184
            Top             =   180
            Width           =   285
         End
         Begin VB.Label lblTheme 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   1710
            TabIndex        =   183
            Top             =   180
            Width           =   285
         End
         Begin VB.Label lblTheme 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   2010
            TabIndex        =   182
            Top             =   180
            Width           =   285
         End
         Begin VB.Label lblTheme 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   3
            Left            =   2340
            TabIndex        =   181
            Top             =   180
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblTheme 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   4
            Left            =   2340
            TabIndex        =   180
            Top             =   330
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Offset:"
            Height          =   195
            Left            =   7590
            TabIndex        =   179
            Top             =   210
            Width           =   465
         End
         Begin VB.Label lblEndRange 
            AutoSize        =   -1  'True
            Caption         =   "-"
            Height          =   195
            Left            =   10770
            TabIndex        =   178
            Top             =   210
            Width           =   45
         End
         Begin VB.Label lblZoom 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1x"
            Height          =   300
            Index           =   0
            Left            =   4050
            TabIndex        =   177
            Top             =   180
            Width           =   285
         End
         Begin VB.Label lblZoom 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2x"
            Height          =   300
            Index           =   1
            Left            =   4350
            TabIndex        =   176
            Top             =   180
            Width           =   285
         End
         Begin VB.Label lblZoom 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "3x"
            Height          =   300
            Index           =   2
            Left            =   4650
            TabIndex        =   175
            Top             =   180
            Width           =   285
         End
         Begin VB.Label lblZoom 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "4x"
            Height          =   300
            Index           =   3
            Left            =   4950
            TabIndex        =   174
            Top             =   180
            Width           =   285
         End
         Begin VB.Label lblZoom 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "5x"
            Height          =   300
            Index           =   4
            Left            =   5250
            TabIndex        =   173
            Top             =   180
            Width           =   285
         End
         Begin VB.Label lblWidth 
            Alignment       =   2  'Center
            BackColor       =   &H00008000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "16"
            Height          =   300
            Index           =   1
            Left            =   6240
            TabIndex        =   172
            Top             =   180
            Width           =   285
         End
         Begin VB.Label lblWidth 
            Alignment       =   2  'Center
            BackColor       =   &H00008000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "32"
            Height          =   300
            Index           =   2
            Left            =   6540
            TabIndex        =   171
            Top             =   180
            Width           =   285
         End
         Begin VB.Label lblWidth 
            Alignment       =   2  'Center
            BackColor       =   &H00008000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "64"
            Height          =   300
            Index           =   3
            Left            =   6840
            TabIndex        =   170
            Top             =   180
            Width           =   285
         End
         Begin VB.Label lblWidth 
            Alignment       =   2  'Center
            BackColor       =   &H00008000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "128"
            Height          =   300
            Index           =   4
            Left            =   7140
            TabIndex        =   169
            Top             =   180
            Width           =   345
         End
         Begin VB.Label lblZoom 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "6x"
            Height          =   300
            Index           =   5
            Left            =   5550
            TabIndex        =   168
            Top             =   180
            Width           =   285
         End
         Begin VB.Label lblChrHeight 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "8x8"
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   167
            Top             =   180
            Width           =   615
         End
         Begin VB.Label lblChrHeight 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "8x16"
            Height          =   300
            Index           =   1
            Left            =   690
            TabIndex        =   166
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.PictureBox picV 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   6540
         Left            =   3960
         ScaleHeight     =   436
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   721
         TabIndex        =   29
         Top             =   840
         Width           =   10815
      End
      Begin VB.Image cmdFontMenu 
         Height          =   255
         Left            =   120
         Picture         =   "frmViewer.frx":247D
         ToolTipText     =   "Font Menu"
         Top             =   390
         Width           =   255
      End
   End
   Begin VB.TextBox txtLA 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6900
      TabIndex        =   140
      Text            =   "0000"
      ToolTipText     =   "Load Address from File, or Entered manually"
      Top             =   30
      Width           =   495
   End
   Begin VB.CheckBox cbLA 
      Caption         =   "LA:"
      Height          =   255
      Left            =   6360
      TabIndex        =   139
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
      Left            =   4110
      TabIndex        =   8
      Top             =   3390
      Visible         =   0   'False
      Width           =   3510
      Begin VB.CheckBox cbIgnoreLF 
         Caption         =   "&Ignore LF"
         Height          =   195
         Left            =   2280
         TabIndex        =   127
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
         TabIndex        =   133
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
         TabIndex        =   130
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
      Left            =   60
      TabIndex        =   2
      Top             =   1650
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
         TabIndex        =   113
         Top             =   180
         Width           =   7995
         Begin VB.ComboBox cboMode 
            Height          =   315
            ItemData        =   "frmViewer.frx":2833
            Left            =   930
            List            =   "frmViewer.frx":2849
            Style           =   2  'Dropdown List
            TabIndex        =   122
            Top             =   0
            Width           =   1920
         End
         Begin VB.CommandButton cmdCpyClip 
            Caption         =   "To &Clipboard"
            Height          =   315
            Left            =   5970
            TabIndex        =   121
            ToolTipText     =   "Export current view text to clipboard"
            Top             =   0
            Width           =   1215
         End
         Begin VB.CheckBox cbRev 
            Caption         =   "&Reverse Text"
            Height          =   240
            Left            =   2940
            TabIndex        =   120
            ToolTipText     =   "Reverse display of Text"
            Top             =   0
            Width           =   1425
         End
         Begin VB.CheckBox cbUseFont 
            Caption         =   "Use CBM &Font"
            Height          =   240
            Left            =   2940
            TabIndex        =   119
            ToolTipText     =   "Use special C64 Font"
            Top             =   240
            Width           =   1425
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "E&xport"
            Height          =   315
            Left            =   5955
            TabIndex        =   118
            ToolTipText     =   "Save current view text to file"
            Top             =   360
            Width           =   1230
         End
         Begin VB.CheckBox cbExp 
            Caption         =   "Expand &Special ("
            Height          =   240
            Left            =   2940
            TabIndex        =   117
            ToolTipText     =   "Expand special characters (ie {RVS} )"
            Top             =   480
            Value           =   1  'Checked
            Width           =   1530
         End
         Begin VB.CheckBox cbOneLine 
            Caption         =   "&Break Multi"
            Height          =   240
            Left            =   4470
            TabIndex        =   116
            ToolTipText     =   "Break multi-statement lines (list one statement per line)"
            Top             =   0
            Width           =   1200
         End
         Begin VB.CheckBox cbPad 
            Caption         =   "Pad &Tokens"
            Height          =   240
            Left            =   4470
            TabIndex        =   115
            ToolTipText     =   "Append SPACE to tokens"
            Top             =   225
            Width           =   1215
         End
         Begin VB.CheckBox cbUC 
            Caption         =   "UCase)"
            Height          =   240
            Left            =   4470
            TabIndex        =   114
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
            TabIndex        =   132
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
            TabIndex        =   131
            Top             =   90
            Width           =   285
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "BASIC:"
            Height          =   195
            Left            =   360
            TabIndex        =   126
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
            TabIndex        =   125
            ToolTipText     =   "Computer model"
            Top             =   405
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "LOAD:"
            Height          =   195
            Left            =   390
            TabIndex        =   124
            Top             =   435
            Width           =   480
         End
         Begin VB.Label lblLoadAdr 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            Height          =   285
            Left            =   930
            TabIndex        =   123
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
         TabIndex        =   112
         ToolTipText     =   "Toggle Options pane"
         Top             =   180
         Width           =   255
      End
   End
   Begin VB.Frame frBlank 
      Height          =   855
      Left            =   9390
      TabIndex        =   34
      Top             =   2430
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Select Viewer with button above..."
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   2430
      End
   End
   Begin VB.CheckBox cbLockView 
      Caption         =   "Lock View"
      Height          =   315
      Left            =   10770
      TabIndex        =   33
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
      Height          =   1245
      Left            =   60
      TabIndex        =   18
      Top             =   3390
      Visible         =   0   'False
      Width           =   3945
      Begin VB.CommandButton cmdLoadVPL 
         Caption         =   "Load VPL..."
         Height          =   315
         Left            =   1140
         TabIndex        =   231
         ToolTipText     =   "Load VICE Palette file"
         Top             =   300
         Width           =   1305
      End
      Begin VB.CommandButton cmdBSave 
         Caption         =   "Save..."
         Height          =   315
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Save to BMP file"
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
         TabIndex        =   143
         Top             =   810
         Width           =   2835
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Comment:"
         Height          =   195
         Left            =   2610
         TabIndex        =   24
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Format:"
         Height          =   195
         Left            =   2610
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
         Left            =   3390
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
         Left            =   3390
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
      Left            =   120
      TabIndex        =   4
      Top             =   7260
      Visible         =   0   'False
      Width           =   12900
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "Z"
         Height          =   315
         Index           =   7
         Left            =   8610
         TabIndex        =   156
         ToolTipText     =   "Make Binary Byte Block"
         Top             =   210
         Width           =   315
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "< ]"
         Height          =   315
         Index           =   3
         Left            =   3930
         TabIndex        =   155
         ToolTipText     =   "Bottom Up "
         Top             =   210
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "<"
         Height          =   315
         Index           =   2
         Left            =   3630
         TabIndex        =   154
         ToolTipText     =   "Next Up"
         Top             =   210
         Width           =   285
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   315
         Index           =   1
         Left            =   3330
         TabIndex        =   153
         ToolTipText     =   "Next Down"
         Top             =   210
         Width           =   285
      End
      Begin VB.Frame frInfo 
         Height          =   525
         Left            =   3990
         TabIndex        =   151
         Top             =   570
         Width           =   8715
         Begin VB.Label lblInfo 
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "Click table entry for info"
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
            Height          =   225
            Left            =   120
            TabIndex        =   152
            Top             =   180
            Width           =   8430
         End
      End
      Begin VB.CommandButton cmdAddEP 
         Caption         =   "EntryPt"
         Height          =   315
         Left            =   5520
         TabIndex        =   150
         ToolTipText     =   "Add Label"
         Top             =   210
         Width           =   675
      End
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "X"
         Height          =   315
         Index           =   6
         Left            =   8280
         TabIndex        =   147
         ToolTipText     =   "Make Hidden Block"
         Top             =   210
         Width           =   315
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "*C*"
         Height          =   315
         Index           =   4
         Left            =   10920
         TabIndex        =   100
         ToolTipText     =   "Add Comment with * Separator"
         Top             =   210
         Width           =   405
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "***"
         Height          =   315
         HelpContextID   =   7
         Index           =   7
         Left            =   12180
         TabIndex        =   99
         ToolTipText     =   "Add * Separator"
         Top             =   210
         Width           =   405
      End
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "W"
         Height          =   315
         Index           =   5
         Left            =   7950
         TabIndex        =   98
         ToolTipText     =   "Make Word Block"
         Top             =   210
         Width           =   315
      End
      Begin VB.CommandButton cmdAddLabel 
         Caption         =   "Label"
         Height          =   315
         Left            =   4830
         TabIndex        =   97
         ToolTipText     =   "Add Label"
         Top             =   210
         Width           =   645
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "==="
         Height          =   315
         Index           =   6
         Left            =   11760
         TabIndex        =   96
         ToolTipText     =   "Add = Separator"
         Top             =   210
         Width           =   405
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "----"
         Height          =   315
         Index           =   5
         Left            =   11340
         TabIndex        =   95
         ToolTipText     =   "Add - Separator"
         Top             =   210
         Width           =   405
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "=C="
         Height          =   315
         Index           =   3
         Left            =   10500
         TabIndex        =   94
         ToolTipText     =   "Add Comment with = Separator"
         Top             =   210
         Width           =   405
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "--C--"
         Height          =   315
         Index           =   2
         Left            =   10080
         TabIndex        =   93
         ToolTipText     =   "Add Comment with - Separator"
         Top             =   210
         Width           =   405
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "C"
         Height          =   315
         Index           =   1
         Left            =   9660
         TabIndex        =   92
         ToolTipText     =   "Add Standalone Comment"
         Top             =   210
         Width           =   405
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   ";C"
         Height          =   315
         Index           =   0
         Left            =   9240
         TabIndex        =   91
         ToolTipText     =   "Add Inline Comment"
         Top             =   210
         Width           =   405
      End
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "V"
         Height          =   315
         Index           =   4
         Left            =   7620
         TabIndex        =   90
         ToolTipText     =   "Make Vector Block"
         Top             =   210
         Width           =   315
      End
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "R"
         Height          =   315
         Index           =   3
         Left            =   7290
         TabIndex        =   89
         ToolTipText     =   "Make RTS vector block"
         Top             =   210
         Width           =   315
      End
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "T"
         Height          =   315
         Index           =   2
         Left            =   6960
         TabIndex        =   88
         ToolTipText     =   "Make Text Block"
         Top             =   210
         Width           =   315
      End
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "H"
         Height          =   315
         Index           =   1
         Left            =   6630
         TabIndex        =   87
         ToolTipText     =   "Make Hex Block"
         Top             =   210
         Width           =   315
      End
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "D"
         Height          =   315
         Index           =   0
         Left            =   6300
         TabIndex        =   86
         ToolTipText     =   "Make Dec Byte Block"
         Top             =   210
         Width           =   315
      End
      Begin VB.CommandButton cmdFindAll 
         Caption         =   "All"
         Height          =   315
         Left            =   2580
         TabIndex        =   40
         ToolTipText     =   "Find all occurences"
         Top             =   210
         Width           =   375
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "[ >"
         Height          =   315
         Index           =   0
         Left            =   2970
         TabIndex        =   31
         ToolTipText     =   "Top Down"
         Top             =   210
         Width           =   345
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   315
         Left            =   1980
         TabIndex        =   30
         ToolTipText     =   "Find Text"
         Top             =   210
         Width           =   555
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   315
         Left            =   1140
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
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Frame frTView 
         Height          =   8625
         Left            =   90
         TabIndex        =   41
         Top             =   570
         Width           =   3825
         Begin VB.Frame frMLSettings 
            Height          =   6345
            Left            =   1680
            TabIndex        =   55
            Top             =   2430
            Width           =   3615
            Begin VB.TextBox txtInlineCol 
               BackColor       =   &H00000080&
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   1920
               MaxLength       =   2
               TabIndex        =   149
               Text            =   "50"
               Top             =   2850
               Width           =   345
            End
            Begin VB.CommandButton cmdImport 
               Caption         =   "Import"
               Height          =   345
               Left            =   2040
               TabIndex        =   106
               ToolTipText     =   "Import Symbols"
               Top             =   4680
               Width           =   1455
            End
            Begin VB.CheckBox cbIncSym 
               Caption         =   "Include Symbol comments"
               Height          =   375
               Left            =   150
               TabIndex        =   85
               Top             =   3900
               Value           =   1  'Checked
               Width           =   3255
            End
            Begin VB.ComboBox cboCPUFile 
               BackColor       =   &H00FFFFFF&
               CausesValidation=   0   'False
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "frmViewer.frx":28BA
               Left            =   2370
               List            =   "frmViewer.frx":28BC
               Style           =   2  'Dropdown List
               TabIndex        =   84
               Top             =   2820
               Visible         =   0   'False
               Width           =   1305
            End
            Begin VB.ComboBox cboCPU 
               BackColor       =   &H00000080&
               CausesValidation=   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmViewer.frx":28BE
               Left            =   810
               List            =   "frmViewer.frx":28C5
               Style           =   2  'Dropdown List
               TabIndex        =   82
               Top             =   1200
               Width           =   2715
            End
            Begin VB.TextBox txtDivLen 
               BackColor       =   &H00000080&
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   1920
               MaxLength       =   2
               TabIndex        =   81
               Text            =   "80"
               Top             =   2520
               Width           =   345
            End
            Begin VB.ComboBox cboPlatFile 
               BackColor       =   &H00FFFFFF&
               CausesValidation=   0   'False
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "frmViewer.frx":28D7
               Left            =   2370
               List            =   "frmViewer.frx":28D9
               Style           =   2  'Dropdown List
               TabIndex        =   79
               Top             =   2520
               Visible         =   0   'False
               Width           =   1305
            End
            Begin VB.ComboBox cboPlatform 
               BackColor       =   &H00000080&
               CausesValidation=   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmViewer.frx":28DB
               Left            =   810
               List            =   "frmViewer.frx":28E2
               Style           =   2  'Dropdown List
               TabIndex        =   78
               Top             =   870
               Width           =   2715
            End
            Begin VB.CommandButton cmdMLHelp 
               Caption         =   "Help"
               Height          =   465
               Left            =   600
               TabIndex        =   76
               ToolTipText     =   "Display HELP file"
               Top             =   5190
               Width           =   2385
            End
            Begin VB.CheckBox cbLabelBlanks 
               Caption         =   "Add blank line before Labels"
               Height          =   375
               Left            =   150
               TabIndex        =   75
               Top             =   3630
               Value           =   1  'Checked
               Width           =   3285
            End
            Begin VB.ComboBox cboPrefix 
               BackColor       =   &H00000080&
               CausesValidation=   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmViewer.frx":28F4
               Left            =   1110
               List            =   "frmViewer.frx":28FB
               Style           =   2  'Dropdown List
               TabIndex        =   73
               Top             =   2190
               Width           =   2415
            End
            Begin VB.ComboBox cboTarget 
               BackColor       =   &H00000080&
               CausesValidation=   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmViewer.frx":290D
               Left            =   810
               List            =   "frmViewer.frx":291A
               Style           =   2  'Dropdown List
               TabIndex        =   71
               Top             =   1860
               Width           =   2715
            End
            Begin VB.CommandButton cmdSaveASM 
               Caption         =   "Save..."
               Height          =   375
               Left            =   1050
               TabIndex        =   67
               ToolTipText     =   "Save disassembly to file"
               Top             =   4260
               Width           =   915
            End
            Begin VB.CheckBox cbSpaceRTS 
               Caption         =   "Add blank line after RTS/RTI instructions"
               Height          =   375
               Left            =   150
               TabIndex        =   66
               Top             =   3360
               Value           =   1  'Checked
               Width           =   3285
            End
            Begin VB.CommandButton cmdPurge 
               Caption         =   "Purge"
               Height          =   345
               Left            =   1050
               TabIndex        =   65
               ToolTipText     =   "Purge unselected symbol entries"
               Top             =   4680
               Width           =   915
            End
            Begin VB.CommandButton cmdClrTables 
               Caption         =   "New"
               Height          =   315
               Left            =   2580
               TabIndex        =   64
               ToolTipText     =   "Clear Lists and start a new project"
               Top             =   180
               Width           =   915
            End
            Begin VB.CheckBox cbClearOnLoad 
               Caption         =   "Clear Lists on Load"
               Height          =   375
               Left            =   120
               TabIndex        =   63
               ToolTipText     =   "Uncheck if you want to keep existing entries when loading"
               Top             =   510
               Value           =   1  'Checked
               Width           =   1815
            End
            Begin VB.CommandButton cmdProjSave 
               Caption         =   "Save..."
               Height          =   315
               Left            =   1050
               TabIndex        =   62
               ToolTipText     =   "Save Lists to file"
               Top             =   180
               Width           =   915
            End
            Begin VB.CommandButton cmdProjLoad 
               Caption         =   "Load..."
               Height          =   315
               Left            =   90
               TabIndex        =   61
               ToolTipText     =   "Load Lists from a file"
               Top             =   180
               Width           =   915
            End
            Begin VB.CommandButton cmdCopyClip2 
               Caption         =   "Copy To &Clipboard"
               Height          =   375
               Left            =   2040
               TabIndex        =   59
               ToolTipText     =   "Paste disassembly to clipboard"
               Top             =   4260
               Width           =   1455
            End
            Begin VB.CheckBox cbEquates 
               Caption         =   "Show Equates"
               Height          =   195
               Left            =   150
               TabIndex        =   58
               ToolTipText     =   "Include Equates in output"
               Top             =   3180
               Width           =   1515
            End
            Begin VB.ComboBox cboMLFmt 
               BackColor       =   &H00000080&
               CausesValidation=   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmViewer.frx":2935
               Left            =   810
               List            =   "frmViewer.frx":2948
               Style           =   2  'Dropdown List
               TabIndex        =   56
               Top             =   1530
               Width           =   2715
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Inline Comment column:"
               Height          =   195
               Left            =   150
               TabIndex        =   148
               Top             =   2880
               Width           =   1680
            End
            Begin VB.Label lblChanged 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   225
               Left            =   2160
               TabIndex        =   107
               ToolTipText     =   "Project Status (Green=OK, Red=Changed, White=No Project Loaded)"
               Top             =   240
               Width           =   225
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Symbols:"
               Height          =   195
               Left            =   330
               TabIndex        =   105
               Top             =   4710
               Width           =   630
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "CPU:"
               Height          =   195
               Left            =   390
               TabIndex        =   83
               Top             =   1260
               Width           =   375
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Comment Divider length:"
               Height          =   195
               Left            =   120
               TabIndex        =   80
               Top             =   2550
               Width           =   1725
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Platform:"
               Height          =   195
               Left            =   150
               TabIndex        =   77
               Top             =   930
               Width           =   615
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Disassembly:"
               Height          =   195
               Left            =   90
               TabIndex        =   74
               Top             =   4320
               Width           =   915
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Label Prefix:"
               Height          =   195
               Left            =   120
               TabIndex        =   72
               Top             =   2250
               Width           =   870
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Target:"
               Height          =   195
               Left            =   270
               TabIndex        =   70
               Top             =   1920
               Width           =   510
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "View Fmt:"
               Height          =   195
               Left            =   90
               TabIndex        =   57
               Top             =   1590
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
            ItemData        =   "frmViewer.frx":299D
            Left            =   90
            List            =   "frmViewer.frx":299F
            TabIndex        =   145
            Top             =   1590
            Width           =   705
         End
         Begin VB.Frame frTrace 
            Height          =   4425
            Left            =   150
            TabIndex        =   135
            Top             =   1980
            Width           =   3675
            Begin VB.CheckBox cbMLAddLabels 
               Caption         =   " Add Labels"
               Height          =   255
               Left            =   150
               TabIndex        =   146
               Top             =   2130
               Value           =   1  'Checked
               Width           =   1155
            End
            Begin VB.CommandButton cmdAddTables 
               Caption         =   "Add To Tables"
               Height          =   645
               Left            =   120
               TabIndex        =   138
               Top             =   1230
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.CommandButton cmdTrace 
               Caption         =   "START"
               Height          =   795
               Left            =   90
               TabIndex        =   137
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
               ItemData        =   "frmViewer.frx":29A1
               Left            =   1380
               List            =   "frmViewer.frx":29A3
               Sorted          =   -1  'True
               TabIndex        =   136
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
            ItemData        =   "frmViewer.frx":29A5
            Left            =   2940
            List            =   "frmViewer.frx":29A7
            Sorted          =   -1  'True
            TabIndex        =   102
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
            ItemData        =   "frmViewer.frx":29A9
            Left            =   1980
            List            =   "frmViewer.frx":29AB
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   68
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
            ItemData        =   "frmViewer.frx":29AD
            Left            =   2940
            List            =   "frmViewer.frx":29AF
            Sorted          =   -1  'True
            TabIndex        =   54
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
            ItemData        =   "frmViewer.frx":29B1
            Left            =   2190
            List            =   "frmViewer.frx":29B3
            Sorted          =   -1  'True
            TabIndex        =   52
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
            ItemData        =   "frmViewer.frx":29B5
            Left            =   1530
            List            =   "frmViewer.frx":29B7
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   51
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
            ItemData        =   "frmViewer.frx":29B9
            Left            =   840
            List            =   "frmViewer.frx":29BB
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   50
            Top             =   1590
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.CommandButton cmdSymAdd 
            Caption         =   "Add"
            Height          =   315
            Left            =   2070
            TabIndex        =   49
            ToolTipText     =   "Add an entry"
            Top             =   930
            Width           =   495
         End
         Begin VB.CommandButton cmdSymDel 
            Caption         =   "Del"
            Height          =   315
            Left            =   2610
            TabIndex        =   48
            ToolTipText     =   "Delete current entry"
            Top             =   930
            Width           =   495
         End
         Begin VB.CommandButton cmdSYMGoto 
            Caption         =   "Find"
            Height          =   315
            Left            =   3180
            TabIndex        =   47
            ToolTipText     =   "Find Selected"
            Top             =   930
            Width           =   555
         End
         Begin VB.CommandButton cmdSymSave 
            Caption         =   "Save"
            Height          =   315
            Left            =   690
            TabIndex        =   46
            ToolTipText     =   "Save file"
            Top             =   930
            Width           =   555
         End
         Begin VB.CommandButton cmdSymLoad 
            Caption         =   "Load"
            Height          =   315
            Left            =   90
            TabIndex        =   45
            ToolTipText     =   "Load a file"
            Top             =   930
            Width           =   555
         End
         Begin VB.CommandButton cmdRemDupLbls 
            Caption         =   "Remove Duplicates"
            Height          =   315
            Left            =   90
            TabIndex        =   104
            ToolTipText     =   "Remove Duplicate Entries"
            Top             =   930
            Width           =   1845
         End
         Begin VB.CommandButton cmdRemDupJSR 
            Caption         =   "Remove Duplicates"
            Height          =   315
            Left            =   90
            TabIndex        =   103
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
            TabIndex        =   144
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
            TabIndex        =   134
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
            TabIndex        =   101
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
            TabIndex        =   69
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
            TabIndex        =   60
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
            TabIndex        =   53
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   42
            Top             =   540
            Width           =   720
         End
      End
      Begin VB.CheckBox cbAuto 
         Height          =   195
         Left            =   900
         TabIndex        =   39
         ToolTipText     =   "Automatically Refresh"
         Top             =   300
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.Image imgShowInfo 
         Height          =   255
         Left            =   4410
         Picture         =   "frmViewer.frx":29BD
         ToolTipText     =   "Toggle Info box"
         Top             =   270
         Width           =   255
      End
      Begin VB.Image imgBW 
         Height          =   255
         Left            =   330
         Picture         =   "frmViewer.frx":2D73
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
         TabIndex        =   38
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
         Left            =   12750
         TabIndex        =   25
         ToolTipText     =   "Address range"
         Top             =   270
         Width           =   105
      End
   End
   Begin VB.Shape shOverflow 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   8400
      Shape           =   3  'Circle
      Top             =   90
      Width           =   195
   End
   Begin VB.Label lblVSize 
      Alignment       =   1  'Right Justify
      Caption         =   "00000"
      Height          =   225
      Left            =   7920
      LinkTimeout     =   0
      TabIndex        =   142
      Top             =   75
      Width           =   450
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Size:"
      Height          =   225
      Left            =   7530
      TabIndex        =   141
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
      TabIndex        =   110
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
      TabIndex        =   109
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
      TabIndex        =   108
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
      TabIndex        =   37
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
      TabIndex        =   36
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
Public VBuf2 As String                                         'ViewFile Secondary Buffer (font editor)
Public VBufAlt As String                                       'Alternate Character Set buffer (font editor)
Public VClip As String                                         'ViewFile Clipboard Buffer (font editor)
Public VRestore As String                                      'ViewFile Restore Point (font editor)

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

Dim Pow(7)              As Integer                              'binary powers array

'==== BASIC Viewer
Dim Token(358) As String

'==== FONT Viewer
Dim SelChr              As Integer                              'Selected Character (start of range?)
Dim SelChr2             As Integer                              'End of Range
Dim RangeFlag           As Boolean                              'True, when Range is valid
Dim FontH               As Integer                              'Font Height (8 or 16)
Dim ChrZoom             As Integer                              'Zoom Factor
Dim SelZoom             As Integer
Dim ChrWIndex           As Integer                              'Number of characters per line
Dim ChrHIndex           As Integer
Dim ChrHeight           As Integer
Dim BorderFlag          As Boolean                              'Display border between characters
Dim OutlineFlag         As Boolean                              'Outline each character (experimental)
Dim BorderSize          As Integer                              'Border size
Dim MCFlag              As Boolean                              'Multi-Colour Mode
Dim BitFlag             As Boolean                              'Update pixel bits?

Dim ChrEditMode         As Boolean                              'Edit Mode Flag
Dim ChrPos              As Integer                              'Current Edit Chr Byte Offset Position
Dim ChrPosEnd           As Integer                              'Current Range End Pos
Dim ChrTop              As Integer                              'Pointer to current character
Dim ChrPixelR           As Integer                              'Pixel Row Marker
Dim ChrPixelC           As Integer                              'Pixel Col Marker
Dim PixelMode           As Integer
    
Dim CMat(15)            As String * 1                           'array for one character
Dim Tr(15)              As Integer                              'translation array
Dim SBit(7, 7) As Integer, DBit(7, 7) As Integer                'source/dest bit arrays for rotation
Dim RedrawFlag          As Boolean

'==== ML Viewer
Dim OP(255) As String                                           '6502 Opcodes
Dim OpModeLen As String                                         'Opcode Addresing Mode Lengths (number of bytes for specified addressing mode)
Dim OpB As String, OpJ As String, OpZ As String                 'Tracer opcode groups: Branches, Jumps, Stops
Dim OpDesc As String                                            'Opcode Description from file

Dim LastFile As String, LastComment As String, LastSymPos As Integer
Public ProjFlag As Boolean, MLCFlag As Boolean, InfoFlag As Boolean
Public ChangeFlag As Boolean
Public MLTabNum As Integer
Public OpCodeFlag As Boolean, ShowTables As Boolean
Public DOTORG As String, DOTWORD As String, DOTBYTE As String, DOTTEXT As String
Public LPrefix As String, ProjFilename As String


'---- Load the Form
Private Sub Form_Load()
    Dim i As Integer, Filename As String
    
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
    InfoFlag = False
    ShowTables = False                      'ML Viewer
    MLTabNum = 0                            'ML Viewer
    SetTarget 0                             'Target Assembler
    SetPrefix 0                             'Label Prefix
            
   
    SelChr = 0: SelChr2 = 0                 'Chr Viewer - Selected character(s)
    RangeFlag = False                       'Chr Viewer - Valid Range selected flag
    ChrWIndex = 1                           'Chr Viewer - Width Index
    ChrHeight = 8                           'Chr Viewer - Character Height (8 or 16)
    ChrHIndex = 0                           'Chr Viewer - Height selection index 0 or 1
    ChrZoom = 4                             'Chr Viewer - Character Set Zoom Index
    ChrPos = 1: ChrPosEnd = 8               'Chr Viewer - Start/end positions into buffer
    OutlineFlag = False                     'Chr Viewer - Outline each character
    BorderFlag = True                       'Chr Viewer - Borders on
    BorderSize = 1                          'Chr Viewer - Border size
    BitFlag = True                          'Chr Viewer - Update pixel set flag
    ChrPixelR = 0: ChrPixelC = 0            'Chr Viewer - Selected Chr Pixel Markers
    ChrEditMode = False                     'Chr Viewer - Edit Mode Flag
    SelZoom = 16                            'Chr Viewer - Selected Character Zoom Factor
    PixelMode = 2                           'Chr Viewer - Pixel Drawing Mode (0=BG,1=FG,2=XOR)
    RedrawFlag = True
    
    txtBorder.ListIndex = 0
    
    For i = 0 To 7: Pow(i) = 2 ^ i: Next    'Set Powers of 2
    
    Call SetColor                           'Setup C64 colours
    Filename = ExeDir & "cbmxfer.vpl"       'Check for default VPL file
    If Exists(Filename) = True Then LoadVPL Filename
    
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
            txtLA.Text = MyHex(VLA, 4)                                          'Load Address in Hex
        Else
            VLA = MyDec(txtLA.Text)                                             'No Load Address
            txtLA.Enabled = True
        End If
        
        shOverflow.Visible = False
        If VLen > 32760 Then VLen = 32760: ChrEditMode = False: shOverflow.Visible = True          'Max size we can load! Overflow indicator
        VBuf = Input(VLen, FIO)                                                 'Read contents to buffer
    Close FIO
    
    VBuf2 = VBuf                                                                'Backup buffer for Font Editor
    
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
    Dim W As Single, H As Single                        'Original Window Size
    Dim W1 As Single, H1 As Single, L1 As Single        'Scaled Window Size LEFT frame
    Dim T1 As Single                                    'Top offset
    Dim W2 As Single, L2 As Single                      'Scaled Width and LeftPosition for RIGHT frame
    Dim i As Integer
    
    If ViewerReady = False Then Exit Sub
    
    '-- Hide all the frames
    frBasic.Visible = False
    frFont.Visible = False
    frML.Visible = False: frInfo.Visible = False
    frBIN.Visible = False
    frSEQ.Visible = False
    frBMP.Visible = False
    frBlank.Visible = False
    For i = 0 To 2: lblSSize(i).Visible = False: Next
    
    DoEvents
        
    '-- Calculate window sizes
    ' NOTE: There seems to be a difference between width and height returned when running in the IDE vs
    '       when compiled. The values -390 and -1000 look good when compiled. This could be Windows revision dependent.
    W = Me.Width - 390:   If W < 4400 Then W = 4400         'Window Width - enforce minimum size for elements
    H = Me.Height - 1000: If H < 3700 Then H = 3700         'Window Height - enforce min size for elements
    
    L1 = 75: T1 = 375                                       'Left/Top Margins
    W1 = W: W2 = W: H1 = H: L2 = L1                         'Set for single-view mode
    
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
' N=Frame#, Size: L=Left,T=Top,W=Width,H=Height, VisFlag=Frame Visible?
' In Dual-View Mode FLAG=TRUE
Sub SetFrame(ByVal n As Integer, ByVal l As Single, ByVal T As Single, ByVal W As Single, ByVal H As Single, ByVal VisFlag As Boolean)
    Dim L2 As Single, T2 As Single, W2 As Single, H2 As Single  'Second copy for modification
    Dim LL As Single, TT As Single, WW As Single, HH As Single
    Dim W3 As Single, H3 As Single
    Dim W4 As Single, H4 As Single
    
    L2 = l: T2 = T: W2 = W: H2 = H                              'Copy of original size requested
    LL = 105: TT = 420: HH = H - 440: WW = W - 200              'Adjust top and height to give a little border area
    
    W3 = W - 200: H3 = H - 600
    
    Select Case n
        Case -1 '-- Blank frame with message
            frBlank.Move l, T, W2, H2
            frBlank.Visible = VisFlag
            
        Case 0  '-- Adjust BASIC Viewer Size
            frBOpts.Visible = False
            If lblBView.Caption = "<<" Then TT = 930: HH = H - 1000: frBOpts.Visible = True 'show options
            frBasic.Move l, T, W2, H2
            lstBAS.Move LL, TT, W3, HH
            frBasic.Visible = VisFlag
    
        Case 1  '-- Adjust SEQ Viewer Size
            TT = 600: HH = H - 660                              'Adjust for Options
            frSEQ.Move l, T, W2, H2
            lstSEQ.Move LL, TT, W3, H3
            frSEQ.Visible = VisFlag
            
        Case 2  '-- Adjust BIN Viewer Size
            TT = 840: HH = H - 840                              'Adjust for Options
            frBIN.Move l, T, W2, H2
            lstBIN.Move LL, TT, W3, H3 - 300
            frBIN.Visible = VisFlag
            
        Case 3  '-- Adjust ChrSet Viewer Size
            frFont.Move l, T, W2, H2
            frFont.Visible = VisFlag
            
        Case 4  '-- Adjust ML Viewer Size
            frML.Move l, T, W2, H2                              'Move and size the MAIN frame
            frML.Visible = VisFlag
            frInfo.Visible = False                              'Hide info frame
            lblShw.Caption = ">>"                               'Assume no project tab
            
            TT = 600: HH = H - 660                              'Adjust for Options
            If W < 4600 Then W = 4600                           'Make sure frame elements have room
            
            If InfoFlag = True Then
                TT = 1090: HH = H - 1200                        'Reduce Height of ML area
                frInfo.Visible = True                           'Show info frame
            End If
            
            If ShowTables = True Then
                LL = l + 3960: WW = W - 4130 'Reduce Width
                lblShw.Caption = "<<"
            End If
            
            frInfo.Move LL, T + 145, WW                         'Position Info frame
            lblInfo.Width = WW - 240                            'Size the Info text area
            lstML.Move LL, TT, WW, HH                           'Position The output list
            
            If ShowTables = True Then
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
                                
                lstEP.Height = HH                               'The Tracer Entry Point List

                DrawMLTabs
            End If
    
            frTView.Visible = ShowTables

        Case 5  '-- Adjust IMG Viewer Size
            frBMP.Visible = VisFlag
            frBMP.Move l, T, W2, H2
            
    End Select
    DoEvents
    
End Sub

Private Sub lblPixelMode_Click(Index As Integer)
    PixelMode = Index
    FONTView
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
    If Exists(Filename) = False Then MyMsg "Can't load Token file!": Exit Sub
    
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
    
    If BitFlag = True Then CreatePixels 'Create pixel multicolour or normal font pixels
        
    For i = 0 To 1
        lblChrHeight(i).Font.Bold = False: lblChrHeight(i).ForeColor = vbBlack
    Next i
    lblChrHeight(ChrHIndex).Font.Bold = True: lblChrHeight(ChrHIndex).ForeColor = vbWhite
        
    For i = 0 To 5
        lblZoom(i).Font.Bold = False: lblZoom(i).ForeColor = vbBlack
    Next i
    lblZoom(ChrZoom - 1).Font.Bold = True: lblZoom(ChrZoom - 1).ForeColor = vbWhite
       
    For i = 0 To 4
        lblWidth(i).Font.Bold = False: lblWidth(i).ForeColor = vbBlack
    Next i
    lblWidth(ChrWIndex).Font.Bold = True: lblWidth(ChrWIndex).ForeColor = vbWhite

    For i = 0 To 2
        lblPixelMode(i).Font.Bold = False: lblPixelMode(i).ForeColor = vbBlack
    Next i
    lblWidth(PixelMode).Font.Bold = True: lblPixelMode(PixelMode).ForeColor = vbWhite


    If MCFlag = True Then
        lblTheme(3).Visible = True: lblTheme(4).Visible = True
    Else
        lblTheme(3).Visible = False: lblTheme(4).Visible = False
    End If

    If RangeFlag = True Then
        picChr.Visible = False: lblRange.Visible = True
    Else
        picChr.Visible = True: lblRange.Visible = False
    End If
    
    If ChrHIndex = 0 Then
        picChr.Height = 2030 'Height for 8x8 characters
    Else
        picChr.Height = 3950 'Height for 8x16 characters
    End If
        
    DrawChrSet
    SetEditMode
    
    DoEvents
    
End Sub
 
'---- Draws the Complete Character Set
' Uses offset, BorderFlag,OutlineFlag,Zoom and selected colours
Public Sub DrawChrSet()
    Dim j As Integer, k As Integer, X As Integer, Y As Integer, V As Integer, TopX As Integer, TopY As Integer
    Dim R As Integer, C As Integer, MaxR As Integer, MaxC As Integer, MaxH As Integer
    Dim CZ As Integer, RZ As Integer, PZ As Integer 'zoomed size
    Dim Offset As Long, ChrNum As Integer, OutFlag As Boolean
    Dim C1 As Long, Thick As Integer                        'Outline Colour
    Dim CCZ As Integer, RRZ As Integer                      'to help speed up drawing
    
    FH = ChrHeight                                          'Chr Height in pixels
    ChrNum = 0
    C = 0: R = 0: X = 0: Y = 0
    TopX = 0: TopY = 0                                      'Top-Left Offset
    MaxR = 64                                               'Max Row
    MaxC = GetCharWidth(ChrWIndex)                          'How many characters wide?
    CW = 8: RW = FH                                         'Chr width, Row width
    PZ = CW * ChrZoom                                       'Scale factor for drawing one line of pixels
    Thick = ChrZoom \ 2                                     'Outline thickness
    
    Offset = Val(txtCSkip.Text): If Offset < 1 Then Offset = 1
    If Offset > 32767 Then Offset = 32767
    
    If BorderFlag = True Then
        CW = CW + BorderSize
        RW = RW + BorderSize
        TopX = BorderSize + ChrZoom: TopY = TopX
        If (OutlineFlag = True) And (BorderSize > 2) And (ChrZoom > 1) Then
            OutFlag = True
            picV.DrawWidth = Thick
        End If
    End If
    
    C1 = vbWhite                                            'Outline Colour
    
    CZ = CW * ChrZoom                                       'Size of one character including borders
    RZ = RW * ChrZoom                                       'Size of one character including borders
    FontH = FH                                              'Set for calculating chr when clicked
            
    If RedrawFlag = True Then
        picV.Width = (CZ * MaxC + TopY) * Screen.TwipsPerPixelX
        picV.Height = (RZ * MaxR + TopX) * Screen.TwipsPerPixelY
        picV.BackColor = lblTheme(2).BackColor
        picV.Cls
        DoEvents
        picV.Visible = False
        DoEvents
    End If
    
    CCZ = TopX: RRZ = TopY
    
    For j = Offset To VLen
        V = Asc(Mid(VBuf, j, 1))
        '----paintpicture {srceimg},destX,destY,destW,destH ,srcX,srcY,srcW,srcH,mode
        If (RangeFlag = True) And (ChrNum >= SelChr) And (ChrNum <= SelChr2) Then
            picV.PaintPicture Pix.Image, CCZ, RRZ + Y * ChrZoom, PZ, ChrZoom, 0, V, 8, 1, vbNotSrcCopy  'blit the pixels - Selected character
        Else
            picV.PaintPicture Pix.Image, CCZ, RRZ + Y * ChrZoom, PZ, ChrZoom, 0, V, 8, 1                'blit the pixels - Un-selected character
        End If
        
        If OutFlag = True Then
            picV.Line (CCZ - Thick, RRZ - Thick)-Step(8 * ChrZoom + Thick * 2, FontH * ChrZoom + Thick * 2), C1, B 'Draw the outline
        End If

        Y = Y + 1
        If Y = FH Then
            Y = Y - FH: ChrNum = ChrNum + 1: C = C + 1: If C >= MaxC Then C = 0: R = R + 1
            CCZ = TopX + C * CZ 'speed up
            RRZ = TopY + R * RZ 'speed up
        End If
        If R > MaxR Then Exit For
    Next j
    
    If R < MaxR Then
        If R = 0 Then R = 1                                     'Fix if single row
        picV.Height = (RZ * R + TopX) * Screen.TwipsPerPixelY
    End If
    
    DoEvents
    lblEndRange.Caption = "to" & Str(j)
    picV.Visible = True
    DoEvents
    
    ShowChr
    RedrawFlag = False
    
End Sub

'==============
'Font View Subs
'==============

Private Sub HScroll1_Change()
    BorderSize = Int(HScroll1.value)
    If BorderSize < 1 Then BorderSize = 1
    lblBorderSize.Caption = Format(BorderSize)
    RedrawFlag = True
    FONTView 'draw character set
End Sub

Private Sub cmdFontMenu_Click()
    PopupMenu frmMenu.mnuFont
End Sub


'-- Dispatch Font Menu Selection
Public Sub DoFMenu(ByVal Index As Integer)
    Select Case Index
        Case 1: ToggleMC
        Case 2: ToggleBorder
        Case 3: SaveBMP
        Case 4: ToggleEdit
        Case 5: SaveFont 0                      'Save entire font
        Case 6: SaveFont 1                      'Save Range
        Case 100 To 104                         'Convert Font
            ConvertFont Index - 100
            SelChr = 0: SelChr2 = 0
            SetSelect
            FONTView 'draw character set
    End Select
    
End Sub

Private Sub SaveBMP()
    Dim Filename As String
    
    Filename = FileOpenSave(FileBase(VFileName), 1, 3, "Save as BMP")
    picV.Picture = picV.Image 'crop to visible
    If Filename <> "" Then SavePicture picV.Image, Filename

End Sub

'-- Change Zoom Factor
Private Sub lblZoom_Click(Index As Integer)
    ChrZoom = Index + 1
    RedrawFlag = True
    FONTView 'draw character set
End Sub

'-- Change Width
Private Sub lblWidth_Click(Index As Integer)
    ChrWIndex = Index
    RedrawFlag = True
    FONTView 'draw character set
End Sub

'-- Change Height
Private Sub lblChrHeight_Click(Index As Integer)
    Dim Tmp As String
    
    If Index = ChrHIndex Then Exit Sub 'ignore click if already in same mode
    ChrHIndex = Index
    
    If Index = 0 Then
        ChrHeight = 8: Tmp = "16 to 8"
    Else
        ChrHeight = 16: Tmp = "8 to 16"
    End If
    
    If ChrEditMode = True Then
        If MsgBox("Do you want to convert this font from " & Tmp & " pixel format?", vbYesNo, "Convert Font") = vbYes Then
            ConvertFont Index               'Convert font from 8 to 16 (index=1) or 16 to 8 (index=0)
        End If
    End If
    
    RedrawFlag = True
    SetSelect
    FONTView 'draw character set
End Sub

'-- Toggle Multicolour mode
Private Sub ToggleMC()
    MCFlag = Not MCFlag
    BitFlag = True
    RedrawFlag = True
    FONTView 'draw character set
End Sub

'-- Toggle Border
Private Sub ToggleBorder()
    BorderFlag = Not BorderFlag
    RedrawFlag = True
    FONTView 'draw character set
End Sub

'-- Toggle Outline
Private Sub cbOutline_Click()
    OutlineFlag = False: If cbOutline.value = vbChecked Then OutlineFlag = True
    RedrawFlag = True
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
        Case 7: FG = CBMColor(0): BG = CBMColor(1): BO = CBMColor(0)    '-- Black on White
        Case 8: FG = CBMColor(1): BG = CBMColor(0): BO = CBMColor(1)    '-- White on Black
    End Select
    
    lblTheme(0).BackColor = FG: lblTheme(1).BackColor = BG: lblTheme(2).BackColor = BO
    DoEvents
    RedrawFlag = True
    BitFlag = True
    FONTView

End Sub

Private Sub txtCSkip_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then FONTView
End Sub

Public Sub CreatePixels()
    Dim j As Integer, k As Integer, CI As Integer
    Dim MC(3) As Long               'Array to hold multicolour values
    
    MC(0) = lblTheme(1).BackColor   'Background colour
    MC(1) = lblTheme(3).BackColor   'Register colour #1
    MC(2) = lblTheme(4).BackColor   'Register colour #2
    MC(3) = lblTheme(0).BackColor   'Foreground Colour
    
    Pix.ForeColor = lblTheme(0).BackColor
    Pix.BackColor = lblTheme(1).BackColor
    Pix.Cls
        
    If MCFlag = True Then
        '-- Create a 4-colour bitmap with pixels to match binary representation of pixel pairs (row=value,cols 0 to 7=pixel)
        For j = 0 To 255
            For k = 0 To 7 Step 2
                CI = 0                                      'Colour Index
                If (j And Pow(k)) Then CI = CI + 2          'Check first bit
                If (j And Pow(k + 1)) Then CI = CI + 1      'Check second bit
                Pix.ForeColor = MC(CI)                      'Set the colour of the pixel to draw
                Pix.PSet (7 - k, j)                         'Set the first pixel
                Pix.PSet (6 - k, j)                         'Set the second pixel
            Next k
        Next j
    Else
        '-- Create a 2-colour bitmap with pixels to match binary representation of value (row=value,cols 0 to 7=pixel)
        For j = 0 To 255
            For k = 0 To 7
                If (j And Pow(k)) Then Pix.PSet (7 - k, j)
            Next k
        Next j
    End If
    
    BitFlag = False                                         'Bitmaps are created
End Sub

'---- Jump to Next Character
Private Sub cmdCSNxt_Click()
    SelChr = SelChr + 1: If SelChr > 255 Then SelChr = 255
    SelChr2 = SelChr
    SetSelect
    FONTView
    ShowChr
End Sub

'---- Jump to Previous Character
Private Sub cmdCSPrev_Click()
    SelChr = SelChr - 1: If SelChr < 0 Then SelChr = 0
    SelChr2 = SelChr
    SetSelect
    ShowChr
End Sub

'---- Show the Selected Character
Public Sub ShowChr()
    Dim R As Integer, C As Integer, X As Integer, Y As Integer, XYOff As Integer
    Dim RW As Integer, CW As Integer, CMax As Integer
    Dim SetNum As Integer, ChrNum As Integer
    Dim C1 As Long, C2 As Long, C3 As Long
    Dim Tmp As String
    
    CMax = GetCharWidth(ChrWIndex)
    OutFlag = False
    RW = FontH: CW = 8: XYOff = 0                                                   'Pixels in one char
    If BorderFlag = True Then
        RW = RW + BorderSize: CW = CW + BorderSize
        XYOff = BorderSize + ChrZoom            'Adjust for border
    End If
            
    SetNum = SelChr \ 128: ChrNum = SelChr Mod 128                                  'Set based on 128 char font
    If ChrPixelR >= ChrHeight Then ChrPixelR = ChrHeight - 1                        'When switching from 16 to 8 pixel tall
    
    '-- Show Info
    lblChrNum.Caption = Format(SelChr, "000")
    lblChrX.Caption = "Set# " & Format(SetNum) & Cr & " Chr# " & Format(ChrNum) & " ($" & MyHex(ChrNum, 2) & ")"
    
    lblFStat.Caption = "Crosshairs: Row=" & Format(ChrPixelR) & ", Col=" & Format(ChrPixelC) _
        & Cr & Cr & "Chr: RIGHT-CLICK to set crosshairs." & Cr & Cr & "Chr Set: CLICK on first chr, RIGHT-CLICK on last to set RANGE."
    
    Tmp = "Range:" & Cr & Cr & "From: " & Format(SelChr) & Cr & "To..: " & Format(SelChr2) & Cr & Cr & "(" & Format(SelChr2 - SelChr + 1) & " selected)"
    If Len(VClip) > 0 Then Tmp = Tmp & Cr & Cr & Format(Len(VClip)) & " bytes in clipboard"
    lblRange.Caption = Tmp
    
    '-- Calc position
    R = Int(SelChr / CMax)
    C = SelChr - R * CMax
    X = C * CW * ChrZoom + XYOff: Y = R * RW * ChrZoom + XYOff
    
    C1 = lblTheme(2).BackColor
    C2 = vbWhite
    
    If picV.Height >= FontH * ChrZoom * 15 Then
        picChr.PaintPicture picV.Image, 0, 0, SelZoom * 8, SelZoom * FontH, X, Y, 8 * ChrZoom, FontH * ChrZoom 'Draw the Character
        
        
        If BorderFlag = True Then
            For i = 0 To 16
                picChr.Line (0, i * SelZoom)-Step(160, 2), C1, BF 'Draw Horizontal Lines
            Next i
            
            For i = 0 To 8
                picChr.Line (i * SelZoom, 0)-Step(2, 320), C1, BF 'Draw Vertical Lines
            Next i
        End If
        
        picChr.Line (0, ChrPixelR * SelZoom + 1)-Step(160, 0), C2 'Draw Horizontal Crosshair
        picChr.Line (ChrPixelC * SelZoom + 1, 0)-Step(0, 320), C2 'Draw Vertical Crosshair
    End If
    
End Sub

'---- Select a character
' LEFT BUTTON=Select character, RIGHT BUTTON=Select Range End
'
Private Sub picV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim T As Integer, R As Integer, C As Integer, RW As Integer, CW As Integer, CMax As Integer
    Dim UpdateFlag As Boolean
    
    CMax = GetCharWidth(ChrWIndex)
    UpdateFlag = False
    
    RW = FontH: CW = 8
    If BorderFlag = True Then RW = RW + BorderSize: CW = CW + BorderSize
    
    R = Int(Y / (RW * ChrZoom)): If R > 32 Then R = 32
    C = Int(X / (CW * ChrZoom)): If C > CMax Then C = CMax
    T = R * CMax + C
    
    If (Shift > 0) Or (Button = 2) Then
       SelChr2 = T: RangeFlag = True: UpdateFlag = True
    Else
       SelChr = T: SelChr2 = T: If RangeFlag = True Then RangeFlag = False: UpdateFlag = True
    End If
    
    'If SelChr = SelChr2 Then RangeFlag = False  'Same chr selected - Removed to allow single character range
    
    If RangeFlag = True Then
        If SelChr > SelChr2 Then T = SelChr: SelChr = SelChr2: SelChr2 = T  'Swap endpoints
    End If
        
    SetSelect
    If UpdateFlag = True Then FONTView
    ShowChr
End Sub

Private Sub SetSelect()
       ChrPos = SelChr * ChrHeight + 1
       ChrPosEnd = SelChr2 * ChrHeight + ChrHeight
       'If RangeFlag = True Then If SelChr = SelChr2 Then RangeFlag = False: RedrawFlag = True 'Removed to allow single character range
End Sub

'---- Change Skip-bytes
Private Sub cmdSB_Click(Index As Integer)
    Dim Offset As Integer
    
    Offset = Val(txtCSkip.Text)
    Select Case Index
        Case 0: Offset = Offset - 256
        Case 1: Offset = Offset - ChrHeight
        Case 2: Offset = Offset - 1
        Case 3: Offset = Offset + 1
        Case 4: Offset = Offset + ChrHeight
        Case 5: Offset = Offset + 256
    End Select
    If Offset < 0 Then Offset = 0
    txtCSkip.Text = Format(Offset)
    FONTView
End Sub

'-------------------
'  font editing subs
'-------------------

Private Sub ToggleEdit()
    If ChrEditMode = False Then
        If shOverflow.Visible = True Then MyMsg "Sorry, font is too big to edit!": Exit Sub
    End If
    ChrEditMode = Not ChrEditMode
    SetEditMode
End Sub

'---- Set elements according to Edit Mode
Private Sub SetEditMode()
    If ChrEditMode = False Then
        frChr.Left = 90
        picV.Left = 2390
        frTools.Visible = False
        lblFStat.Visible = False
    Else
        frChr.Left = 1680
        picV.Left = 3970
        frTools.Visible = True
        lblFStat.Visible = True
    End If
End Sub

'---- Save Font to File
Private Sub SaveFont(ByVal Mode As Integer)
    Dim Filename As String, FFIO As Integer
    Dim Tmp As String
    
    Tmp = "Save Font": If Mode = 1 Then Tmp = "Save Font Range"
    
    Filename = FileOpenSave(FileBase(VFileName), 1, 6, Tmp)
    If Filename = "" Then Exit Sub
    If Overwrite(Filename) = False Then Exit Sub
    
    FFIO = FreeFile
    Open Filename For Output As FFIO
    If Mode = 0 Then
        Print #FFIO, VBuf;                                          'Write entire font
    Else
        Print #FFIO, Mid(VBuf, ChrPos, ChrPosEnd - ChrPos + 1);     'Write RANGE
    End If
    Close FFIO
    
End Sub
'---- Edit the Selected Character
Private Sub picChr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim R As Integer, C As Integer, bv As Integer, nv As Integer, TV As Integer, p As Integer
    Dim PP As Integer
    
    If ChrEditMode = False Then Exit Sub
    
    '-- convert x/y to row/col
    R = Y \ SelZoom: If R > ChrHeight - 1 Then R = ChrHeight - 1
    C = X \ SelZoom: If C > 7 Then C = 7
    
    '-- Set Markers (Shift-click or Right-click)
    If (Shift > 0) Or (Button = 2) Then
        ChrPixelR = R: ChrPixelC = C
        ShowChr
        Exit Sub                                'Exit
    End If
    
    '-- Edit Pixel
    PP = ChrPos + R                             'Position of byte to update
    bv = Asc(Mid(VBuf, PP, 1))                  'Get byte for row
    p = Pow(7 - C)                              'Get pixel bit value
    
    Select Case PixelMode
        Case 0: nv = bv And (255 - p)           'Set to Background
        Case 1: nv = bv Or p                    'Set to Foreground
        Case 2: nv = bv Xor p                   'XOR
    End Select
    
    Mid(VBuf, PP, 1) = Chr(nv)                  'update the pixel
    FONTView
    
End Sub

'---- Perform Tool Operation
Private Sub cmdTool_Click(Index As Integer)
    Dim a As Integer, B As Integer, C As Integer, cc As Integer, cStart As Integer
    Dim Row As Integer, Col As Integer
    Dim V As Integer, nv As Integer, nv2 As Integer, nv3 As Integer
    Dim Tmp As String, Tmp2 As String
    Dim Flag As Boolean, Bit As Integer
        
    Flag = False: If cbShiftMode.value = vbChecked Then Flag = True
    
    Select Case Index
        Case 0: GoSub ShiftUp                                           'Shift Up
        Case 1: GoSub ShiftDown                                         'Shift Down
        Case 2: GoSub ShiftLeft                                         'Shift Left
        Case 3: GoSub ShiftRight                                        'Shift Right
        Case 4: GoSub Clear                                             'Clear
        Case 5: GoSub RVS                                               'Reverse
        Case 6: GoSub BoldFont                                          'Make Bold
        Case 7: GoSub Underlined                                        'Make Underlined
        Case 8: GoSub RotateLeft                                        'Rotate bits Left (left-most pixel goes to end)
        Case 9: GoSub RotateRight                                       'Rotate bits Right (right-most pixel goed to beginning)
        Case 10: GoSub MirrorH                                          'Mirror Horizontal
        Case 11: GoSub MirrorV                                          'Mirror Vertical
        Case 12: cStart = 0: GoSub DoubleTall                           'Double Tall - Top
        Case 13: cStart = ChrHeight \ 2: GoSub DoubleTall               'Double Tall - Bottom
        Case 14: cc = 0: GoSub DoubleWide                               'Double Wide - Left side
        Case 15: cc = 1: GoSub DoubleWide                               'Double Wide - Right side
        Case 16: cStart = 0: cc = 0: GoSub DoubleSize                   'Double Size - Top Left
        Case 17: cStart = 0: cc = 1: GoSub DoubleSize                   'Double Size - Top Right
        Case 18: cStart = ChrHeight \ 2: cc = 0: GoSub DoubleSize       'Double Size - Bottom Left
        Case 19: cStart = ChrHeight \ 2: cc = 1: GoSub DoubleSize       'Double Size - Bottom Right
        Case 20: GoSub SwapSets                                         'Swap sets
        Case 21: GoSub SelectAll                                        'Select Entire font for Range
        Case 22: GoSub CopyClip                                         'Copy to clipboard
        Case 23: GoSub PasteClip                                        'Paste from clipboard
        Case 24: GoSub Restore1                                         'Restore character(s)
        Case 25: GoSub Restore2
        Case 26: GoSub InsRow                                           'Insert blank Row below crosshair
        Case 27: GoSub DelRow                                           'Delete Row below crosshair
        Case 28: GoSub InsCol                                           'Insert blank Col to right of crosshair
        Case 29: GoSub DelCol                                           'Delete column to right of crosshair
        Case 30: GoSub SetRestorePoint                                  'Set a restore point
        
        'Case ?: GoSub CompactFont          'Compact 8x16 font to 8x8 pixels
        'Case ?: GoSub SquishFont           'Squish 8x16 font to 8x8 pixels
        
    End Select
    
    FONTView
    Exit Sub
    
'--------------------
ShiftUp:
    For j = ChrPos To ChrPosEnd Step ChrHeight
        DoEvents
        Tmp = Mid(VBuf, j, 1)
        For k = 1 To ChrHeight - 1
            Mid(VBuf, j + k - 1, 1) = Mid(VBuf, j + k, 1)   'copy to byte above
        Next k
        If Flag = True Then
            Mid(VBuf, j + ChrHeight - 1, 1) = Tmp           'wrap to bottom line
        Else
            Mid(VBuf, j + ChrHeight - 1, 1) = Nu            'clear bottom line
        End If
    Next j
    Return

ShiftDown:
    For j = ChrPos To ChrPosEnd Step ChrHeight
        DoEvents
        Tmp = Mid(VBuf, j + ChrHeight - 1, 1)
        For k = ChrHeight - 2 To 0 Step -1
            Mid(VBuf, j + k + 1, 1) = Mid(VBuf, j + k, 1)   'copy to byte above
        Next k
        If Flag = True Then
            Mid(VBuf, j, 1) = Tmp                           'wrap to top line
        Else
            Mid(VBuf, j, 1) = Nu                            'clear top line
        End If
    Next j
    Return

ShiftLeft:
    For j = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, j, 1))                    'Read a byte
        Bit = 0: If (V And 128) > 0 Then Bit = 1
        nv = (V * 2) Mod 256                        'Shift the pixels
        If Flag = True Then nv = nv + Bit
        Mid(VBuf, j, 1) = Chr(nv)                   'Write it
    Next j
    Return

ShiftRight:
    For j = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, j, 1))                    'Read a byte
        Bit = 0: If (V And 1) > 0 Then Bit = 128
        nv = V \ 2                                  'Shift the pixels
        If Flag = True Then nv = nv + Bit
        Mid(VBuf, j, 1) = Chr(nv)                   'Write it
    Next j
    Return
  
Clear:
    For j = ChrPos To ChrPosEnd
        Mid(VBuf, j, 1) = Nu
    Next j
    Return
    
RVS:
    For j = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, j, 1))
        Mid(VBuf, j, 1) = Chr(255 - V)
    Next j
    Return
    
BoldFont:
    For j = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, j, 1))
        nv2 = Int(V / 2)                            'Shift the pixels
        nv = V Or nv2                               'Merge them
        Mid(VBuf, j, 1) = Chr(nv)                   'write it
    Next j
    Return
    
Underlined:
    For j = ChrPos To ChrPosEnd Step ChrHeight
        Mid(VBuf, j + ChrPixelR, 1) = Chr(255)
    Next j
    Return
    
RotateRight:
    If ChrHeight = 16 Then MyMsg "Rotation only supported on 8x8 characters!": Return
    C = 0

    For j = ChrPos To ChrPosEnd Step ChrHeight
        DoEvents
        ChrTop = j                              'Set current character position
        GoSub ClearBitArrays                    'Clear arrays for next character (all bits to zero)
        GoSub ReadChr                           'Get bytes and fill Source Bit array
        '---- Do Rotation 90
        For Row = 0 To 7
            For Col = 0 To 7
                DBit(7 - Col, Row) = SBit(Row, Col)
            Next Col
        Next Row
        GoSub WriteChr                          'Write the Dest Bit Array back as bytes
    Next j
    Return

RotateLeft:
    If ChrHeight = 16 Then MyMsg "Rotation only supported on 8x8 characters!": Return
    
    For j = ChrPos To ChrPosEnd Step ChrHeight
        DoEvents
        ChrTop = j                              'Set current character position
        GoSub ClearBitArrays                    'Clear arrays for next character (all bits to zero)
        GoSub ReadChr                           'Read 8 bytes and fill Source Bit array
        '---- Do Rotation 270
        For Row = 0 To 7
            For Col = 0 To 7
                DBit(Col, 7 - Row) = SBit(Row, Col)
            Next Col
        Next Row
        GoSub WriteChr                          'Write the Dest Bit Array back as bytes
    Next j
    Return

MirrorH:
    For j = ChrPos To ChrPosEnd Step ChrHeight
        DoEvents
        For k = 0 To ChrHeight - 1
            CMat(k) = Mid(VBuf, j + k, 1)                       'Read to array in order
        Next k
       
        For k = 0 To ChrHeight - 1
            Mid(VBuf, j + k, 1) = CMat(ChrHeight - k - 1)         'Write to output in reverse order
        Next k
    Next j
    Return

MirrorV:
    GoSub SetupMirrorArray

    For j = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, j, 1))                                'Read to array in order
        a = Int(V / 16): B = V Mod 16                           'Calculate HI and LO nibbles
        nv = Tr(B) * 16 + Tr(a)                                 'Reverse the bits
        Mid(VBuf, j, 1) = Chr(nv)                               'Write to output
    Next j
    Return

DoubleTall:
    For j = ChrPos To ChrPosEnd Step ChrHeight
        DoEvents
        For k = 0 To ChrHeight - 1
            CMat(k) = Mid(VBuf, j + k, 1)
        Next k
        C = cStart
        For k = 1 To ChrHeight - 1 Step 2
            Mid(VBuf, j + k - 1, 1) = CMat(C)
            Mid(VBuf, j + k, 1) = CMat(C)
            C = C + 1
        Next k
    Next j
    Return

DoubleWide:
    GoSub Setup2XArray
    For j = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, j, 1))                                'Read byte, convert to ascii
        a = Int(V / 16): B = V Mod 16                           'Calculate HI/LO nibbles
        If cc = 0 Then
            nv = Tr(a)                                          'Translate HI
        Else
            nv = Tr(B)                                          'Translate LO
        End If
        Mid(VBuf, j, 1) = Chr(nv)                               'Write
    Next j
    Return

DoubleSize:
    GoSub Setup2XArray
    For j = ChrPos To ChrPosEnd Step ChrHeight
        For k = 0 To ChrHeight - 1
            CMat(k) = Mid(VBuf, j + k, 1)                           'Read byte, convert to ascii
        Next k
        C = cStart
        For k = 1 To ChrHeight Step 2
            V = Asc(CMat(C))                                        'Get row byte
            a = Int(V / 16): B = V Mod 16                           'Calculate HI/LO nibbles
            If cc = 0 Then
                nv = Tr(a)                                          'Translate HI
            Else
                nv = Tr(B)                                          'Translate LO
            End If
            Mid(VBuf, j + k - 1, 1) = Chr(nv)                       'Write
            Mid(VBuf, j + k, 1) = Chr(nv)                           'Write
            C = C + 1
        Next k
    Next j
    Return

InsRow:
    For j = ChrPos To ChrPosEnd Step ChrHeight
        For k = ChrHeight - 2 To ChrPixelR Step -1
            Mid(VBuf, j + k + 1, 1) = Mid(VBuf, j + k, 1)
        Next k
        Mid(VBuf, j + ChrPixelR, 1) = Nu
    Next j
    Return
    
DelRow:
    For j = ChrPos To ChrPosEnd Step ChrHeight
        For k = ChrPixelR To ChrHeight - 2
            Mid(VBuf, j + k, 1) = Mid(VBuf, j + k + 1, 1)
        Next k
        Mid(VBuf, j + ChrHeight - 1, 1) = Nu
    Next j
    Return
    
InsCol:
    '-- calculate pixel masks
    a = 0: For j = 7 To (8 - ChrPixelC) Step -1: a = a + Pow(j): Next j     'LEFT side mask
    B = 0: For j = (7 - ChrPixelC) To 0 Step -1: B = B + Pow(j): Next j     'RIGHT side mask
    
    '-- insert
    For j = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, j, 1))                                            'Get byte value
        nv2 = V And a                                                       'mask left side
        nv3 = (V And B) \ 2                                                 'mask right side and shift
        Mid(VBuf, j, 1) = Chr(nv2 + nv3)                                    'recombine and write
    Next j
    Return

DelCol:
    '-- calculate pixel masks
    a = 0: For j = 7 To (8 - ChrPixelC) Step -1: a = a + Pow(j): Next j     'LEFT side mask
    B = 0: For j = (6 - ChrPixelC) To 0 Step -1: B = B + Pow(j): Next j     'RIGHT side mask
    
    '-- delete
    For j = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, j, 1))                                            'Get byte value
        nv2 = V And a                                                       'mask left side
        nv3 = (V And B) * 2                                                 'mask right side and shift
        Mid(VBuf, j, 1) = Chr(nv2 + nv3)                                    'recombine and write
    Next j
    Return

    
'---------------------------- Clipboard subs
CopyClip:
    VClip = Mid(VBuf, ChrPos, ChrPosEnd - ChrPos + 1)
    Return

PasteClip:
    V = Len(VClip): If V = 0 Then Return
    For j = ChrPos To ChrPosEnd Step V
        Mid(VBuf, j, V) = VClip                                             'Paste it once
    Next j
    Return
    
Restore1:
    If VRestore = "" Then VRestore = VBuf2
    For j = ChrPos To ChrPosEnd
        Mid(VBuf, j, 1) = Mid(VRestore, j, 1)
    Next j
    Return
    
Restore2:
    For j = ChrPos To ChrPosEnd
        Mid(VBuf, j, 1) = Mid(VBuf2, j, 1)
    Next j
    Return
    
SetRestorePoint:
    VRestore = VBuf
    Return
    
SwapSets:
    If VBufAlt = "" Then VBufAlt = VBuf                         'Only one set loaded so copy it
    Tmp = VBuf                                                  'Remember set 1
    VBuf = VBufAlt                                              'Swap set 1 and 2
    VBufAlt = Tmp
    VLen = Len(VBuf):  lblVSize.Caption = Format(VLen)          'Set buffer length
    DrawChrSet                                                    'redraw
    Return

SelectAll:
    SelChr = 0: SelChr2 = (VLen \ ChrHeight) - 1
    RangeFlag = True
    RedrawFlag = True
    SetSelect
    Return
    
'================================================================================== Manipulation Routines

SetupMirrorArray:
    Tr(0) = 0: Tr(1) = 8: Tr(2) = 4: Tr(3) = 12
    Tr(4) = 2: Tr(5) = 10: Tr(6) = 6: Tr(7) = 14
    Tr(8) = 1: Tr(9) = 2: Tr(10) = 5: Tr(11) = 13
    Tr(12) = 3: Tr(13) = 11: Tr(14) = 7: Tr(15) = 15
    Return

Setup2XArray:
    Tr(0) = 0: Tr(1) = 3: Tr(2) = 12: Tr(3) = 15
    Tr(4) = 48: Tr(5) = 51: Tr(6) = 60: Tr(7) = 63
    Tr(8) = 192: Tr(9) = 195: Tr(10) = 204: Tr(11) = 207
    Tr(12) = 240: Tr(13) = 243: Tr(14) = 252: Tr(15) = 255
    Return
    
ClearBitArrays:
    For Row = 0 To 7
        For Col = 0 To 7
            SBit(Row, Col) = 0: DBit(Row, Col) = 0
        Next Col
    Next Row
    Return
    
'---------- Read 8 bytes and fill the Source Bit array with 0's and 1's
ReadChr:
    For Row = 0 To 7
        V = Asc(Mid(VBuf, ChrTop + Row, 1))                                 'Get a byte/value
        If V > 0 Then                                                       'Only do bits if non-zero
            For Col = 0 To 7
                If (V And Pow(Col)) <> 0 Then SBit(Row, Col) = 1          'Set the bit array
            Next Col
        End If
    Next Row
    Return
    
'---------- Write DBit Array out as 8 bytes
WriteChr:
    For Row = 0 To 7
            V = 0                                                           'Reset to zero
            For Col = 0 To 7
                If DBit(Row, Col) = 1 Then V = V + Pow(Col)                  'Add the value of the bit position
            Next Col
        Mid(VBuf, ChrTop + Row, 1) = Chr(V)                                 'Store bytes
    Next Row
    Return
    
End Sub

'---- Convert Font
' n=0 --> 16 to 8 = truncate character, no padding
' n=1 --> 8 to 16 = pad with 8 blank rows
Private Sub ConvertFont(ByVal n As Integer)
    Dim j As Integer, k As Integer, l As Integer, H As Integer, B As Integer
    Dim Tmp As String, Pad As String
    
    If VLen > 16300 Then MyMsg "Sorry, font is too large to convert!": Exit Sub
    
    Select Case n
        Case 0: B = 8: H = 16: Pad = ""                         'Read 16 bytes, write 8               - 8x8 font
        Case 1: B = 8: H = 8: Pad = String(8, Nu)               'Read 8 bytes, write 8 plus 8 padding - 8x16 font
        Case 2: B = 5: H = 5: Pad = String(3, Nu)               'Read 5 bytes, write 5 plus 3 padding - 5x7 sideways font
        Case 3: B = 7: H = 7: Pad = Nu                          'Read 7 bytes, write 7 plus 1 padding - 5x7 upright font
        Case 4: B = 14: H = 14: Pad = String(2, Nu)             'Read 14 bytes, write 14 plus 2 padding - 8x14 EGA font
    End Select
    
    VBuf2 = ""                                                  'Converted font built here
    
    For j = 1 To Len(VBuf) Step H
        Tmp = Mid(VBuf, j, B)                                   'Get 8 bytes
        VBuf2 = VBuf2 & Tmp & Pad                               'Copy them plus padding if needed
    Next j

    VBuf = VBuf2                                                'Update main buffer
    VClip = ""                                                  'Clear clipboard
    VLen = Len(VBuf): lblVSize.Caption = Format(VLen)           'Update buffer length
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
    Dim Padd As String, CommentCol As Integer
    
    Dim LNum As Long, LInc As Integer                                   'Line Numbers
    Dim a As Integer, p As Integer
    
    Dim DTMode As Boolean, DTCount As Integer, DTType As String         'Data Table variables
    Dim DTCountMax As Integer, DTMax As Integer, DTPos As Integer       'Data Table variables
    Dim DTStart As Long, DTEnd As Long, DTAscMode As Integer            'Data Table variables
    Dim DTComment As String, DTAddress As String, DTOutStr As String    'Data Table variables
    
    Dim Pass As Integer
    Dim RTSOption As Boolean, SymComment As Boolean, DivLen As Integer  'options
    
    Padd = Space(80)            'spaces for padding byte lists
    LInc = 10                   'Line# Increment
        
    '---- Options
    RTSOption = False: If cbSpaceRTS.value = vbChecked Then RTSOption = True
    SymComment = False: If cbIncSym.value = vbChecked Then SymComment = True
    DivLen = Val(txtDivLen.Text)
    CommentCol = Val(txtInlineCol.Text)
    
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
        Address = VLA: If cbLA.value = vbUnchecked Then Address = MyDec(txtLA.Text)
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
                    lstML.AddItem Left(T1 & GetField(Tmp, 2) & " = $" & GetField(Tmp, 1) & Padd, CommentCol) & ";" & GetField(Tmp, 3)
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
                                If UComment > "" Then
                                    lstML.AddItem Format(LNum) & " ; " & UComment: LNum = LNum + LInc
                                    If TmpB <> "S" Then lstML.AddItem Format(LNum) & " ; " & String(DivLen, TmpB): LNum = LNum + LInc
                                End If
                            Case Else
                                If TmpB <> "S" Then lstML.AddItem ";" & String(DivLen, TmpB): LNum = LNum + LInc
                                If UComment > "" Then
                                    lstML.AddItem "; " & UComment
                                    If TmpB <> "S" Then lstML.AddItem ";" & String(DivLen, TmpB): LNum = LNum + LInc
                                End If
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
                    '---- Look at the current range entry.       Format: HHHH,HHHH,T{num},Comment
                    Tmp = lstDT.List(DTPos)                      'Get the line from the list
                    DTStart = MyDec(Mid(Tmp, 1, 4))              'Get Range Start
                    DTEnd = MyDec(Mid(Tmp, 6, 4))                'Get Range End
                    Tmp = Mid(Tmp, 11)                           'Get just the Type and Comment
                    p = InStr(Tmp, ","): If p = 0 Then p = 1     'Check for comma
                    DTType = UCase(Left(Tmp, 1))                 'Get Type (Asc,Byte,Word,Vector,RVector)
                    
                    DTCountMax = 8                                 'Default Items per line
                    If p > 1 Then DTCountMax = Val(Mid(Tmp, 2, p - 2)) 'If specified, use {num} entries. Num must be single digit
                    If DTCountMax < 1 Then DTCountMax = 8         'If Num=0 then use default
                    
                    If Pass = 2 Then DTComment = Mid(Tmp, p + 1)  'Get Comment
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
                            
                        Case "B", "H", "$" '---- Byte Directive (Hex)
                            If Pass = 2 Then
                                T3 = DOTBYTE
                                DTOutStr = DTOutStr & "$" & B0H
                            End If
                            
                        Case "D"  '---- Byte Directive (Dec)
                            If Pass = 2 Then
                                T3 = DOTBYTE
                                DTOutStr = DTOutStr & B0A
                            End If
                            
                        Case "Z", "%" '---- Byte Directive (Binary)
                            If Pass = 2 Then
                                T3 = DOTBYTE
                                DTOutStr = DTOutStr & "%" & MyBin(B0A)
                            End If
                        
                            
                        Case "W"  '---- Word Directive (Hex)
                            If Pass = 2 Then
                                T3 = DOTWORD
                                DTCountMax = 6
                                Address = Address + 1: C = C + 1    'Increment address
                                B1A = Asc(Mid(VBuf, C, 1))          'Get next byte
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
                            B1A = Asc(Mid(VBuf, C, 1))              'Value of byte
                            TAddress = B1A * 256 + B0A + 1          'Calculate Target Address (decimal) with offsett
                            JAddress = MyHex(TAddress, 4)           'Make it a string
                            SHL = "$" & JAddress                    'Make string for output
                            
                            If Pass = 1 Then
                                If (JAddress >= StartAddress) And (JAddress <= EndAddress) Then
                                    lstLabels.AddItem JAddress      'Target is inside code range so make it a label
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
                                '---- Add a line according to selected format
                                Select Case OutFmt
                                    Case 0: Tmp = DTAddress & T2 & T3 & DTOutStr        'addr bb bb bb cmd param
                                    Case 1: Tmp = DTAddress & T3 & DTOutStr             'addr cmd param
                                    Case 2: Tmp = Format(LNum) & " " & T3 & DTOutStr    'nnnn cmd param
                                    Case 3: Tmp = T2 & T3 & DTOutStr                    'cmd param
                                    Case 4:
                                        ALabel = Left(ALabel & Padd, 15)
                                        Tmp = ALabel & T3 & DTOutStr                    'label cmd param
                                        ALabel = ""                                     'blank it for multi-line tables
                                End Select
                                
                                j = Len(Tmp) + 1: If CommentCol > j Then j = CommentCol     'Comment position
                                lstML.AddItem Left(Tmp + Padd, j) & ";" & DTComment     'Add it to output
                                
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
                NM = Left(OP(B0A), Len(OP(BOA)) - 1)                            'Mneumonic string (eg: JSR or BBR0)
                If Pass = 2 Then If Left(NM, 1) = "?" Then GoodFlag = False     'Found an unknown opcode
                MD = Asc(Right(OP(B0A), 1)) - 96                                'Addressing mode (A-M,N-P)
                NB = Val(Mid(OpModeLen, MD, 1))                                 'How many bytes for this opcode? (1 to 3)

                '---- All modes >2 use one or two-byte address
                If MD > 1 Then
                    '---- Opcode+Byte
                    If NB > 1 Then
                        If C + 1 <= VLen Then
                            B1A = Asc(Mid(VBuf, C + 1, 1))
                            SL = MyHex(B1A, 2): Mid(T2, 4, 2) = SL              'Set second byte
                        End If
                        
                        '---- Opcode+Word
                        If NB > 2 Then
                            If C + 2 <= VLen Then
                                B2A = Asc(Mid(VBuf, C + 2, 1))
                                SH = MyHex(B2A, 2): Mid(T2, 7, 2) = SH          'Set third byte
                            End If
                        Else
                            SH = "00"                                           'Set third byte as $00 (zero page)
                        End If
                        
                        JAddress = SH & SL                                      'Absolute Jump address
                        SHL = "$" & JAddress                                    'Add the $ to HI string
                        SL = "$" & SL                                           'Add the $ to LO string
                    End If
                    
                    '---- Now look up the address
                    If (MD > 2) And (NB > 1) Then
                        Tmp = FindSL(JAddress)                                  'Look for a SYMBOL, ULABEL, or LABEL for this address
                        If Tmp > "" Then
                            SL = Tmp                                            'Substitute Symbol for single-byte address
                            SHL = Tmp                                           'Substitute Symbol for two-byte address
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
                    
                    If UComment > "" Then T5 = "; " & UComment             'Use user comment string
                                        
                    '---- Output line in specified format
                    
                    Select Case OutFmt
                        Case 0: Tmp = T1 & T2 & T3 & T4                     'addr bytes cmd param
                        Case 1: Tmp = T1 & T3 & T4                          'addr cmd param
                        Case 2: Tmp = Format(LNum) & " " & T3 & T4          'nnnn cmd param
                        Case 3: Tmp = "          " & T3 & T4                'cmd param
                        Case 4:
                            ALabel = Left(ALabel & Padd, 15)
                            Tmp = ALabel & T3 & T4                          'label cmd param
                    End Select
                    
                    j = Len(Tmp) + 1: If CommentCol > j Then j = CommentCol     'position for comment
                    lstML.AddItem Left(Tmp + Padd, j) & T5                  'Add to output
                                        
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
Private Sub Label4_Click()
    lstLabels.Visible = Not lstLabels.Visible
End Sub

Private Sub lblTheme_Click(Index As Integer)
    
    frmColourPicker.Show vbModal
    If PickedColour >= 0 Then
        lblTheme(Index).BackColor = PickedColour
        BitFlag = True
        RedrawFlag = True
        FONTView
    End If
    
End Sub

'---- Single clicking on one of the Lists
Private Sub lstCmnt_Click()
    lblInfo.Caption = lstCmnt.List(lstCmnt.ListIndex)
End Sub

Private Sub lstEntryPt_Click()
    lblInfo.Caption = lstEntryPt.List(lstEntryPt.ListIndex)
End Sub

Private Sub lstJSR_Click()
    lblInfo.Caption = lstJSR.List(lstJSR.ListIndex)
End Sub

Private Sub lstLabels_Click()
    lblInfo.Caption = lstLabels.List(lstLabels.ListIndex)
End Sub

Private Sub lstLabels_DblClick()
    Dim Tmp As String, Tmp2 As String
    
    Tmp = lstLabels.List(lstLabels.ListIndex) & ",name,-"              'Make default text entry string
    Tmp2 = InputBox("HHHH,LABELNAME,DESCRIPTION", "Add Label from [GEN] label", Tmp)
    If Len(Tmp2) > 12 Then lstULabels.AddItem Tmp2: MLReViewA

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

    Address = VLA: If cbLA.value = Checked Then Address = MyDec(txtLA.Text) 'VLA or User Specified Address
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
    If TopPos > lstML.ListCount Then TopPos = 0         'FIX: Large data block additions can make TopPos be past end
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
    
    JumpList Tmp, 1, False               'Find next match

End Sub

'---- Find and jump to the next undefined opcode
Private Sub lblGood_Click()
    JumpList "???", 0, False
End Sub

'---- Find specified string
Private Sub cmdFind_Click()
    Dim Tmp As String
    
    Tmp = InputBox("Enter String to find:", "Find")
    If Tmp <> "" Then
        cmdFindAll.ToolTipText = ""
        JumpList Tmp, 0, False
    End If
    
End Sub

'---- Find ALL occurances of last search string
Private Sub cmdFindAll_Click()
    JumpList "", 0, True
End Sub

'---- Jump to next occurance of search string
Private Sub cmdNext_Click(Index As Integer)
    JumpList "", Index, False
End Sub

'---- Search for string
' Blank string searches with same string
' MODE - Search method: 0=Top Down, 1=Current Down, 2=Current UP, 3=Bottom UP
' FLAG - TRUE = ALL matches
Sub JumpList(ByVal Txt As String, Mode As Integer, ByVal Flag As Boolean)
    Static LastTxt As String, Count As Integer 'These values are retained between calls
    
    Dim i As Integer, j As Integer, Max As Integer, Direction As Integer
    Dim Tmp As String
    
    If Txt = "" Then Txt = LastTxt
    If Txt = "" Then Exit Sub
    
    Max = lstML.ListCount - 1                           'Max entries
    Count = 0
    i = lstML.ListIndex                                 'Assume current position
    
    Select Case Mode
        Case 1: Direction = 1: Tmp = "Down"
        Case 2: Direction = -1: Tmp = "Up"
        Case 3: Direction = -1: Tmp = "Bottom Up": i = lstML.ListCount - 1 'Start at END
        Case Else: Direction = 1: Tmp = "Top Down": i = 0           'Start at TOP
    End Select
    
    lblInfo.Caption = "Search (" & Tmp & "): " & Txt
   
    Do
        i = i + Direction: If (i < 0) Or (i > Max) Then Exit Do
        If InStr(1, lstML.List(i), Txt, vbTextCompare) > 0 Then
            lstML.Selected(i) = True                                'Hilight it
            Count = Count + 1                                       'Count it
            
            If Flag = False Then
                j = i - 5: If j < 0 Or j > Max Then j = i
                lstML.TopIndex = j                                  'Move top of list to near found line
                lstML.ListIndex = i                                 'Move to selected line
                Exit Do                                             'Do only one search
            End If
        Else
            lstML.Selected(i) = False
        End If
    Loop
    
    If (Count > 0) And (Flag = True) Then lblInfo.Caption = "Found" & Str(Count) & " line(s) containing: " & Txt
    If (Count = 0) Then lblInfo.Caption = "String '" & Txt & "' was not found."
    
    DoEvents
    LastTxt = Txt                                                   'Remember Search string for next time
    
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
    
    i = lstML.ListIndex: If i < 0 Then MyMsg Tmp: Exit Sub                 'Ooops, no line selected!
    RS = ExtractAddr(lstML.List(i)): If RS = "" Then MyMsg Tmp: Exit Sub   'Ooops, line didn't have an address!
 
    Tmp2 = InputBox("Add LABEL at " & RS & Cr & Cr & "Enter LABEL Name:", "Add Label", "")
    If Tmp2 > "" Then lstULabels.AddItem RS & "," & Tmp2: MLReViewC
    
End Sub

'---- Quick Add Entry Point
Private Sub cmdAddEP_Click()
    Dim RS As String, Tmp As String, Tmp2 As String, i As Integer
    
    Tmp = "Please select a line with an address first!"
    
    i = lstML.ListIndex: If i < 0 Then MyMsg Tmp: Exit Sub                 'Ooops, no line selected!
    RS = ExtractAddr(lstML.List(i)): If RS = "" Then MyMsg Tmp: Exit Sub   'Ooops, line didn't have an address!
 
    Tmp2 = InputBox("Add ENTRY POINT at " & RS & Cr & Cr & "Enter ENTRY POINT Name:", "Add Entry Point", "")
    If Tmp2 > "" Then lstEntryPt.AddItem RS & "," & Tmp2: MLReViewC
    
End Sub

'---- Quick Add Comment / Separator ( ;C / C / -C- / =C= / - / = )
Private Sub cmdAddComment_Click(Index As Integer)
    Dim RS As String, Tmp As String, Tmp2 As String, i As Integer
    
    Tmp = "Please select a line with an address first!"
    
    i = lstML.ListIndex: If i < 0 Then MyMsg Tmp: Exit Sub     'Oops, no line selected!
    RS = ExtractAddr(lstML.List(i)): If RS = "" Then MyMsg Tmp: Exit Sub   'Opps, line didn't have an address!
        
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
        Select Case Index 'DHSRVWXZ
            Case 0: Tmp = "D": Tmp2 = "Decimal Byte Table"
            Case 1: Tmp = "H": Tmp2 = "Hex Byte Table"
            Case 2: Tmp = "S": Tmp2 = "Text/String Table"
            Case 3: Tmp = "R": Tmp2 = "RTS Address Table (Generates Labels)"
            Case 4: Tmp = "V": Tmp2 = "Address Table (Generates Labels)"
            Case 5: Tmp = "W": Tmp2 = "Word Table"
            Case 6: Tmp = "X": Tmp2 = "Hidden Table"
            Case 7: Tmp = "Z": Tmp2 = "Binary Byte Table"
        End Select
                   
        Tmp2 = InputBox("Type : " & Tmp2 & Cr & "Range: " & RS & " to " & RE & Cr & Cr & "Enter a description:", "Add Table", "")
        If Tmp2 <> "" Then
            lstDT.AddItem RS & "," & RE & "," & Tmp & "," & Tmp2    'Add it
            lstDT.Selected(lstDT.NewIndex) = True                   'Make it selected
            MLReViewC
        End If
    Else
        MyMsg "Please select a range first!"
    End If
    
End Sub

Private Sub lstDT_Click()
    lblInfo.Caption = lstDT.List(lstDT.ListIndex)
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

Private Sub lstSYM_Click()
    lblInfo.Caption = lstSYM.List(lstSYM.ListIndex)
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

Private Sub lstULabels_Click()
    lblInfo.Caption = lstULabels.List(lstULabels.ListIndex)
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

'---- Toggle Info frame
Private Sub imgShowInfo_Click()
    InfoFlag = Not InfoFlag
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
    If Exists(Filename) = False Then MyMsg "Sorry, Platform file not found! " & Filename: Exit Sub
    If OverwriteProject = True Then LoadSymFile Filename, 3
    MLReView
    
End Sub

'---- Process selection of a new CPU from the list
Private Sub cboCPU_Click()
    Dim Filename As String
    If MLCFlag = False Then Exit Sub
    If ViewerReady = False Then Exit Sub
    
    Filename = ExeDir & cboCPUFile.List(cboCPU.ListIndex)
    If Exists(Filename) = False Then MyMsg "Sorry, CPU file not found! " & Filename: Exit Sub
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
        Filename = FileOpenSave("", 0, 2, "Load ASM Project File"): If Filename = "" Then Exit Sub
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
    Dim LA As String, LAFlag As Boolean, VName As String, VStr As String
        
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
                    Case 1 '[PROJECT] Section
                        VName = GetVNameU(Tmp): VStr = GetVstr(Tmp) 'Parse line
                        
                        Select Case VName
                            Case "LA": If Len(VStr) = 4 Then LAFlag = True: LA = VStr 'Load Address for project found
                            Case "DIVLEN": txtDivLen.Text = VStr
                            Case "INLCOL": txtInlineCol = VStr
                        End Select
                        
                    Case 2 '[SYMBOLS] Section
                        lstSYM.AddItem Tmp
                        
                    Case 3 '[TABLES] Section
                        lstDT.AddItem Tmp
                        lstDT.Selected(lstDT.NewIndex) = True
                        
                    Case 4 '[LABELS] Section
                        lstULabels.AddItem Tmp
                        
                    Case 5 '[COMMENTS] Section
                        lstCmnt.AddItem Tmp
                        
                    Case 6 '[ENTRYPT] Section
                        lstEntryPt.AddItem Tmp
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
    Print #FIO, "LA="; txtLA.Text               'Save the specified Load Address
    Print #FIO, "DIVLEN="; txtDivLen.Text       'Save the Divider Length
    Print #FIO, "INLCOL="; txtInlineCol.Text    'Save the Inline Comment Column
    
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
            MyMsg "Select the TAB for the type of entry you want first, or use the quick-add buttons at the top of the window!"
            
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
            Tmp2 = InputBox("Types: Byte Tables(A/T=Text,B/H=Hex,D=Decimal,Z=Binary),W=Word,R=RTS,V=Vect" & Cr & Cr & "HHHH,HHHH,TYPE{##},DESCRIPTION", "Add Table", Tmp)
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
    Dim p As Integer, Tmp As String, Tmp2 As String, l As Integer
    
    l = Len(LPrefix)
    p = 1
    If Left(Str, l) = LPrefix Then p = l + 1          'Skip over prefix
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
    
    If Flag = False Then MyMsg "All numbers must be >0!": Exit Sub
    
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
    
    MyMsg "File imported! " & Str(C) & " symbols loaded."
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
    Dim i As Integer, C As Integer
    
    C = 1
    
    For i = lstLabels.ListCount - 1 To 1 Step -1
       If lstLabels.List(i) = lstLabels.List(i - 1) Then
            C = C + 1: lstLabels.RemoveItem (i)
        Else
            If C > 1 Then lstLabels.List(i) = lstLabels.List(i) & " (" & Format(C) & ")"  'add count to remaining entry
            C = 1
        End If
    Next i
    
End Sub

'---- Remove Duplicate External JSR entries
Private Sub cmdRemDupJSR_Click()
    Dim i As Integer, C As Integer
    
    C = 1
    
    For i = lstJSR.ListCount - 1 To 1 Step -1
       If lstJSR.List(i) = lstJSR.List(i - 1) Then
            C = C + 1: lstJSR.RemoveItem (i)
        Else
            If C > 1 Then lstJSR.List(i) = lstJSR.List(i) & " (" & Format(C) & ")"  'add count to remaining entry
            C = 1
        End If
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
    Dim C1 As Integer, C2 As Integer
    
    Filename = ExeDir & "ml-config.txt"
    If Exists(Filename) = False Then MyMsg "ML Config file is missing!": Exit Sub
        
    FIO = FreeFile
    Open Filename For Input As FIO
    
    TMode = 0: C1 = 0: C2 = 0
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
                            cboPlatform.List(C1) = Tmp2
                            cboPlatFile.List(C1) = Mid(Tmp, p + 1)
                            C1 = C1 + 1
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

    Dim C As Single, W As Integer, H As Integer, H2 As Integer, DifCount As Integer
    Dim Tmp As String, Tmp2 As String
    Dim HXLine As String, TLine As String, ALine As String, BLine As String, CLine As String
    Dim Flag As Boolean, MaxW As Integer, LCount As Integer, VLen2 As Integer
    Dim Lo As Integer, Hi As Integer, Address As Long, BMASK As Integer
    Dim CBMFlag As Boolean, ASMFlag As Boolean, CmpFlag As Boolean, HLen As String

    BMASK = 255: If cb7bit.value = vbChecked Then BMASK = 127 'Enable 7-bit view
    lstBIN.Clear
    
    If cbWide.value = vbChecked Then MaxW = 15 Else MaxW = 7
    
    If cbHexFmt.value = vbChecked Then ASMFlag = True Else ASMFlag = False  'ASMbler format flag
    If cbShowP.value = vbChecked Then Flag = True Else Flag = False         'Show Printable
    If cbShowCBM.value = vbChecked Then CBMFlag = True Else CBMFlag = False 'Show CBM
    
    CmpFlag = False                                             'Compare Show Flag
    
    HLen = (MaxW + 1) * 3                                      'Length of Hex bytes
    If ASMFlag = True Then HLen = HLen + MaxW                   'Compensate for ASM $
    
    If cbCmpShow.value = vbChecked Then
        CmpFlag = True
        VLen2 = Len(VBuf2)
        lstBIN.AddItem "FILE COMPARE"
        lstBIN.AddItem "Left  File: " & FileNameOnly(VName) & "    Length=" & Str(VLen) & " bytes"
        lstBIN.AddItem "Right File: " & lblCFile.Caption & "    Length=" & Str(Len(VBuf2)) & " bytes"
        lstBIN.AddItem ""
        lstBIN.AddItem String(7 + 6 * (MaxW + 1), "*")
    End If
    
    C = 0: W = 0: Tmp = "": TLine = "": ALine = "": LCount = 0: DifCount = 0 'Initialize
    
    If cbHexSync.value = vbChecked Then
        Address = MyDec(txtLA.Text)                             'Use Address specified in ASM project
    Else
        Address = VLA                                           'Use Load Address from file
        If cbLA.value = vbUnchecked Then Address = MyDec(txtLA.Text)
    End If
    
    '-- Loop through buffer(s)
    Do
        '-- Reached Width setting... Add to output
        If W > MaxW Then
            If CmpFlag = True Then
                lstBIN.AddItem HXLine & TLine & CLine & ALine & BLine
            Else
                If Flag = True Then lstBIN.AddItem HXLine & TLine & ALine
                If Flag = False Then lstBIN.AddItem HXLine & TLine
            End If
            W = 0: LCount = LCount + 1
        End If
        
        W = W + 1                                               'Count bytes processed
        
        '-- Check for start of new line
        If W = 1 Then
            ALine = " ; ": BLine = "": CLine = ""               'Set initial strings
            If Flag = True Then BLine = " ; "                    'Compare printable string
            If CmpFlag = True Then CLine = "; "                 'Compare HEX differences
            TLine = ""
            HXLine = MyHex(Address, 4) & ": "                    'Start with HEX address
            If ASMFlag = True Then HXLine = HXLine & ".BYT "    'If ASM format add ".BYT"
        End If
        
        C = C + 1: Address = Address + 1                        'Move to Next byte
 
        '-- Build the HEX string
        Tmp = Mid(VBuf, C, 1): H = Asc(Tmp)                      'Get its value
        If ASMFlag = True Then
            TLine = TLine & "$" & MyHex(H, 2)                    'Add the hex string
            If W <= MaxW Then TLine = TLine & ","                'Add a comma if not last
        Else
            TLine = TLine & MyHex(H, 2) & " "
        End If
        
        '-- Build the Compare string
        If CmpFlag = True Then
            If C <= VLen2 Then
                Tmp = Mid(VBuf2, C, 1): H2 = Asc(Tmp)                 'Get its value
                If H = H2 Then
                    CLine = CLine & "== "
                Else
                    CLine = CLine & MyHex(H2, 2) & " "                    'Build hex string
                    DifCount = DifCount + 1
                End If
            Else
                CLine = CLine & "   "
            End If
        End If
        
        '-- Build the Printable bytes string
        If Flag = True Then
            '-- Original File
            Select Case (H And BMASK)
                Case 0 To 31
                    If CBMFlag = True Then
                        ALine = ALine & Chr((H And Mask) + 64)      'Converts CTRL chrs to Letter range
                    Else
                        ALine = ALine & "."                         'Un-Printable
                    End If
                Case 32 To 127: ALine = ALine & Chr(H And BMASK)    'Printable
                Case Else: ALine = ALine & "."                      'Un-Printable
            End Select
            
            '-- Compare File
            If CmpFlag = True Then
                If H = H2 Then
                    BLine = BLine & "="                                     'Values are the same so show "="
                Else
                    Select Case (H2 And BMASK)
                        Case 0 To 31
                            If CBMFlag = True Then
                                BLine = BLine & Chr((H2 And Mask) + 64)      'Converts CTRL chrs to Letter range
                            Else
                                BLine = BLine & "."                         'Un-Printable
                            End If
                        Case 32 To 127: BLine = BLine & Chr(H2 And BMASK)    'Printable
                        Case Else: BLine = BLine & "."                      'Un-Printable
                    End Select
                End If
            End If
        End If
        
    Loop While (C < VLen) 'And (LCount < 32766)
    
    '----- Handle the final line
    
    Tmp = String(HLen, " ")                                                  ' temp spacing string
    
    If TLine <> "" Then
        If CmpFlag = True Then
            lstBIN.AddItem HXLine & Left(TLine & Tmp, HLen) & Left(CLine & Tmp, HLen) & "  " & Left(ALine & Tmp, MaxW + 4) & BLine
            If DifCount = 0 Then
                Tmp2 = "Files are IDENTICAL!"
            Else
                Tmp2 = Str(DifCount) & " differences"
            End If
            lblDifTxt.Caption = Tmp2
            lstBIN.List(3) = "RESULTS: " & Tmp2
        Else
            If Flag = False Then lstBIN.AddItem HXLine & Left(TLine & Tmp, HLen) & Left(ALine & Tmp, HLen)
            If Flag = True Then lstBIN.AddItem HXLine & TLine
        End If
    End If
    
    If lstBIN.Visible = True Then lstBIN.SetFocus
    
End Sub

'---- Save HEX Listing
Private Sub cmdHSave_Click()
    Dim Filename As String, FIO As Integer, j As Integer, n As Integer
    
    n = lstBIN.ListCount: If n = 0 Then MyMsg "No lines to save!": Exit Sub
    
    Filename = FileOpenSave(FileBase(VFileName), 1, 5, "Save HEX Listing as TXT")
    If Filename = "" Then Exit Sub
    
    FIO = FreeFile
    Open Filename For Output As FIO
    
    For j = 0 To lstBIN.ListCount
    Print #FIO, lstBIN.List(j)
    Next j
    Close FIO

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

'---- Search for a String or Hex bytes
' Parses the search field and converts hex digits if required then searches from the top of the file

Private Sub cmdHexFind_Click()
    Dim SS As String, HH As String, SH As String
    Dim SL As Integer, j As Integer, p As Integer
    
    HH = txtHSS.Text: SL = Len(HH)
    
    If SL = 0 Then MyMsg "Enter String, or start with $ to search for hex byte(s).": Exit Sub
    If Left(HH, 1) = "$" Then
        HH = Mid(HH, 2): SL = SL - 1
        If (SL Mod 2) > 0 Then MyMsg "HEX digits must be in pairs.": Exit Sub
        
        SS = ""
        For j = 1 To SL Step 2
            SH = Mid(HH, j, 2)                                  'Get 2 hex digits
            SS = SS & Chr(MyDec(SH))                            'Add character to searchstring
        Next j
    Else
        SS = HH                                                 'Use original string as entered
    End If
    
    HexSearch SS                                                'Search from the TOP
    
End Sub

'---- Search for NEXT occurance
Private Sub cmdHNext_Click()
    HexSearch ""
End Sub

'---- Search for specified text string or hex bytes
' SS is the search string. If specified causes the search to start from the TOP
' If ommitted uses the last string and continues searching from last position

Private Sub HexSearch(ByVal SS As String)
    Static LastPos As Integer, LastSS As String                 'Remembers these between calls
    Dim MaxW As Integer, LL As Integer, L2 As Integer, p As Integer
    Dim HA As String
    
    If SS = "" Then SS = LastSS Else LastPos = 1                'If no searchstring then use previous, else start from top
    LastSS = SS                                                 'Remember the Searchstring
    
    If LastPos > Len(VBuf) Then LastPos = 1                     'Wrap back to top
   
    p = InStr(LastPos, VBuf, SS)                                'Do a binary search
    If p = 0 Then p = InStr(LastPos, VBuf, SS, vbTextCompare)   'If not found search textually
    
    If p = 0 Then
        MyMsg "No more occurances."                             'No results. Display message and exit
        Exit Sub
    End If
    
    LastPos = p + Len(SS)                                       'Set position for next search
    If cbWide.value = vbChecked Then MaxW = 16 Else MaxW = 8    'What is the view line lenght?
    LL = (p - 1) \ MaxW                                         'Which line is the found string on?
    L2 = p - LL * MaxW - 1                                      'offset on line
    
    lstBIN.ListIndex = LL                                       'Select the line containing the string
    HA = MyHex(MyDec(Left(lstBIN.List(LL), 4)) + L2, 4)         'Hex Address of found
    
    lblSResults.Caption = "Found at $" & HA & ", Offset:" & Str(p - 1) 'Results message

End Sub

'---- Compare to a second file
Private Sub cmdCompare_Click()
    Dim FIO As Integer, Tmp As String, P00Flag As Boolean, FLen As Long

    Filename = FileOpenSave("", 1, 0, "Load Compare file")
    If Filename = "" Then Exit Sub
    
    If Exists(Filename) = False Then MyMsg "Viewer: File '" & Filename & "' not found!": Exit Sub
        
    P00Flag = False                                        'Assume normal file
    If FileExtU(Filename) = "P00" Then P00Flag = True     'P00 file found!
  
    '-- Load the file to the buffer, update and display file details
    FIO = FreeFile
    Open Filename For Binary As FIO: FLen = intLOF(FIO)
        If P00Flag = True Then P00Buf = Input(26, FIO): FLen = FLen - 26      'Skip over header
        If cbLA.value = vbChecked Then
            VBuf2 = Input(2, FIO): FLen = FLen - 2                               'Read the Load address
        End If
        
        If FLen > 32760 Then FLen = 32760
        VBuf2 = Input(FLen, FIO)                                                 'Read contents to buffer
    Close FIO
    
    lblCFile.Caption = FileNameOnly(Filename)
    If VBuf = VBuf2 Then lblDifTxt.Caption = "The files are identical!"
    HEXView
End Sub


'============
'SEQ Viewer
'============
Sub SEQView()
    Dim FIO As Integer, C As Integer, Tmp As String, TLine As String, H As Integer

    lstSEQ.Clear
    
    C = 1: Tmp = "": TLine = ""
    Do
        If Len(TLine) > 80 Then lstSEQ.AddItem TLine: TLine = ""
        Tmp = Mid(VBuf, C, 1): H = Asc(Tmp)
        Select Case H
            Case 32 To 127: TLine = TLine & Tmp
            Case 192 To 218: TLine = TLine & Chr(H And 127)
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
    
    '-- Read shared buffer and determine what type of bitmap file it is
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
    Dim c0 As Long, C1 As Long 'Pixel on and off colours - new May 2017
    
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
    C1 = CBMColor(0)                                        'Black Foreground Colour - new 2017
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
                    For k = 1 To nxt
                      Pel = Asc(Mid(Dat, dpos, 1) & Nu): dpos = dpos + 1
                      GoSub PaintBit
                    Next k
                    
                  Case 65 To 127
                    For k = 0 To 7
                      pat(k) = Asc(Mid(Dat, dpos, 1) & Nu): dpos = dpos + 1
                    Next k
                    
                    For l = 1 To (nxt And 63)
                      For k = 0 To 7
                        Pel = pat(k): GoSub PaintBit
                      Next k
                    Next l
                    
                  Case 129 To 255
                    DT = Asc(Mid(Dat, dpos, 1) & Nu): dpos = dpos + 1
                    For k = 1 To (nxt - 128)
                      Pel = DT
                      GoSub PaintBit
                    Next k
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
        If (Pel And Pow(k2)) Then Picture1.PSet (XX - k2, YY), C1 'Set Black dot
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
                    K3 = 0
                    If (Pel And Pow(k2)) Then K3 = K3 + 1
                    If (Pel And Pow(k2 + 1)) Then K3 = K3 + 2
                    
                    Select Case K3
                        Case 0: colput& = CBMColor(BG)
                        Case 1: colput& = CBMColor((S And 240) / 16)
                        Case 2: colput& = CBMColor(S And 15)
                        Case 3: colput& = CBMColor(C And 15)
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
    If Exists(Filename) = False Then MyMsg "Picture formats file missing!!!": Exit Sub
    
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

Private Sub cmdLoadVPL_Click()
    Dim Filename As String
    
    Filename = FileOpenSave(FileBase(VFileName), 0, 7, "Load VICE Palette")
    If Filename <> "" Then
        LoadVPL Filename
        BMPView 're-draw the image with new palette
    End If

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

'--- Load VICE Palette File
' NOTE: VPL files have single LineFeed character between lines
Private Sub LoadVPL(ByVal Filename As String)
    Dim FIO As Integer, C As Integer, R As String, G As String, B As String
    Dim Ch As String, Tmp As String
    
    C = 0               'Colour Index
    
    FIO = FreeFile
    Open Filename For Input As FIO
    
    Do While Not EOF(FIO)
        Ch = Input(1, FIO)
        If Ch = LF Then
            If Left(Tmp, 1) <> "#" Then
                If Len(Tmp) > 9 Then
                    R = MyDec(Mid(Tmp, 1, 2))
                    G = MyDec(Mid(Tmp, 4, 2))
                    B = MyDec(Mid(Tmp, 7, 2))
                    CBMColor(C) = RGB(R, G, B)
                    C = C + 1: If C > 15 Then Exit Do
                End If
            End If
            Tmp = ""        'Clear out string
        Else
            Tmp = Tmp & Ch  'Add char to line string
        End If
    Loop
    
    Close FIO
End Sub
'--- Common File Open or Save Dialog
' You can specify a default filename, a File Filter list index (0-7), and Window Title
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
        Case 3: CommonDialog.Filter = "Bitmap Files (*.BMP)|*.BMP"
        Case 4: CommonDialog.Filter = "ASM Files (*.ASM,*.TXT)|*.ASM;*.TXT"
        Case 5: CommonDialog.Filter = "Text Files (*.TXT)|*.TXT"
        Case 6: CommonDialog.Filter = "Binary Files (*.BIN,*.ROM,*.FON)|*.BIN"
        Case 7: CommonDialog.Filter = "VICE Palette Files (*.VPL)|*.VPL"
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

Private Sub txtLA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ViewIt ViewMode, VFileName, VName, VExt 're-load the file
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
Private Sub cbHexFmt_Click()
    HEXView
End Sub
Private Sub cbCmpShow_Click()
    HEXView
End Sub

'---- Font Updates

'---- SEQ Updates
Private Sub cbIgnoreLF_Click()
    SEQView
End Sub

