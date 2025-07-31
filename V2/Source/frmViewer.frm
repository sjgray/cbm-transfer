VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmViewer 
   Appearance      =   0  'Flat
   BackColor       =   &H00808000&
   Caption         =   "Viewer:"
   ClientHeight    =   12405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20730
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmViewer.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   12405
   ScaleWidth      =   20730
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame frFont 
      BorderStyle     =   0  'None
      Caption         =   "Font Viewer"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8700
      Left            =   2400
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   17415
      Begin VB.Frame frControls 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   600
         TabIndex        =   130
         Top             =   0
         Width           =   16845
         Begin VB.CheckBox cbSetSize 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "256"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1200
            TabIndex        =   251
            Top             =   60
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.ComboBox cboTheme 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmViewer.frx":0442
            Left            =   3720
            List            =   "frmViewer.frx":0464
            Style           =   2  'Dropdown List
            TabIndex        =   138
            ToolTipText     =   "Pick colour Theme"
            Top             =   0
            Width           =   1245
         End
         Begin VB.TextBox txtCSkip 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9960
            TabIndex        =   137
            Text            =   "0"
            ToolTipText     =   "Set number of bytes to skip (decimal)"
            Top             =   30
            Width           =   645
         End
         Begin VB.CommandButton cmdSB 
            Appearance      =   0  'Flat
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
            Height          =   270
            Index           =   0
            Left            =   10680
            TabIndex        =   136
            Top             =   30
            Width           =   345
         End
         Begin VB.CommandButton cmdSB 
            Appearance      =   0  'Flat
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   11040
            TabIndex        =   135
            Top             =   30
            Width           =   285
         End
         Begin VB.CommandButton cmdSB 
            Appearance      =   0  'Flat
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   11340
            TabIndex        =   134
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmdSB 
            Appearance      =   0  'Flat
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   11610
            TabIndex        =   133
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmdSB 
            Appearance      =   0  'Flat
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   11880
            TabIndex        =   132
            Top             =   30
            Width           =   285
         End
         Begin VB.CommandButton cmdSB 
            Appearance      =   0  'Flat
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
            Height          =   270
            Index           =   5
            Left            =   12180
            TabIndex        =   131
            Top             =   30
            Width           =   345
         End
         Begin VB.Label lblTheme 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   6
            Left            =   3360
            TabIndex        =   332
            ToolTipText     =   "Current Character Outline Colour"
            Top             =   30
            Width           =   285
         End
         Begin VB.Label lblTheme 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   5
            Left            =   3060
            TabIndex        =   301
            ToolTipText     =   "Pixel/Character Divider Colour"
            Top             =   30
            Width           =   285
         End
         Begin VB.Label lblWidth 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fit"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   6
            Left            =   8970
            TabIndex        =   271
            Top             =   30
            Width           =   375
         End
         Begin VB.Label lblWidth 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   5
            Left            =   8490
            TabIndex        =   255
            Top             =   30
            Width           =   435
         End
         Begin VB.Label lblBorder 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   14850
            TabIndex        =   254
            ToolTipText     =   "Increase Border"
            Top             =   60
            Width           =   135
         End
         Begin VB.Label lblBorder 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   14310
            TabIndex        =   253
            ToolTipText     =   "Decrease Border"
            Top             =   60
            Width           =   135
         End
         Begin VB.Label lblBorderSize 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   14460
            TabIndex        =   171
            ToolTipText     =   "Border Size"
            Top             =   60
            Width           =   345
         End
         Begin VB.Label Label 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Border:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   23
            Left            =   13740
            TabIndex        =   170
            Top             =   60
            Width           =   555
         End
         Begin VB.Label lblWidth 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   6930
            TabIndex        =   162
            Top             =   30
            Width           =   255
         End
         Begin VB.Label lblTheme 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   1860
            TabIndex        =   157
            ToolTipText     =   "Foreground Colour"
            Top             =   30
            Width           =   285
         End
         Begin VB.Label lblTheme 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   2160
            TabIndex        =   156
            ToolTipText     =   "Background Colour"
            Top             =   30
            Width           =   285
         End
         Begin VB.Label lblTheme 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   2760
            TabIndex        =   155
            ToolTipText     =   "Screen Border Colour"
            Top             =   30
            Width           =   285
         End
         Begin VB.Label lblTheme 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   3
            Left            =   2460
            TabIndex        =   154
            ToolTipText     =   "Multicolour#1"
            Top             =   30
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblTheme 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   4
            Left            =   2460
            TabIndex        =   153
            ToolTipText     =   "Multicolour#2"
            Top             =   180
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Offset:"
            Height          =   195
            Index           =   22
            Left            =   9420
            TabIndex        =   152
            Top             =   60
            Width           =   525
         End
         Begin VB.Label lblEndRange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            Height          =   195
            Left            =   12570
            TabIndex        =   151
            Top             =   60
            Width           =   60
         End
         Begin VB.Label lblZoom 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1x"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   5040
            TabIndex        =   150
            Top             =   30
            Width           =   255
         End
         Begin VB.Label lblZoom 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2x"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   5340
            TabIndex        =   149
            Top             =   30
            Width           =   255
         End
         Begin VB.Label lblZoom 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "3x"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   5640
            TabIndex        =   148
            Top             =   30
            Width           =   255
         End
         Begin VB.Label lblZoom 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "4x"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   5940
            TabIndex        =   147
            Top             =   30
            Width           =   255
         End
         Begin VB.Label lblZoom 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "5x"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   4
            Left            =   6240
            TabIndex        =   146
            Top             =   30
            Width           =   255
         End
         Begin VB.Label lblWidth 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "16"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   7230
            TabIndex        =   145
            Top             =   30
            Width           =   255
         End
         Begin VB.Label lblWidth 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "32"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   7530
            TabIndex        =   144
            Top             =   30
            Width           =   255
         End
         Begin VB.Label lblWidth 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "64"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   7830
            TabIndex        =   143
            Top             =   30
            Width           =   255
         End
         Begin VB.Label lblWidth 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "128"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   4
            Left            =   8130
            TabIndex        =   142
            Top             =   30
            Width           =   315
         End
         Begin VB.Label lblZoom 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "6x"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   5
            Left            =   6540
            TabIndex        =   141
            Top             =   30
            Width           =   255
         End
         Begin VB.Label lblChrHeight 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "8x8"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   140
            Top             =   30
            Width           =   495
         End
         Begin VB.Label lblChrHeight 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "8x16"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   1
            Left            =   600
            TabIndex        =   139
            Top             =   30
            Width           =   495
         End
      End
      Begin VB.PictureBox cmdFontMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   90
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   316
         ToolTipText     =   "Font Editor Menu"
         Top             =   90
         Width           =   285
      End
      Begin VB.Frame frEditor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7395
         Left            =   6420
         TabIndex        =   275
         Top             =   420
         Visible         =   0   'False
         Width           =   10725
         Begin VB.CommandButton cmdClearMacro 
            Caption         =   "Clear"
            Height          =   285
            Left            =   6390
            TabIndex        =   327
            ToolTipText     =   "Clear the Macro"
            Top             =   90
            Width           =   645
         End
         Begin VB.CommandButton cmdPlay 
            Caption         =   "Play"
            Height          =   285
            Left            =   7230
            TabIndex        =   326
            ToolTipText     =   "Play Macro"
            Top             =   90
            Width           =   675
         End
         Begin VB.PictureBox cmdSEDMenu 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            ScaleHeight     =   19
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   19
            TabIndex        =   325
            ToolTipText     =   "Screen Designer Menu"
            Top             =   90
            Width           =   285
         End
         Begin VB.CheckBox cbCBM 
            Caption         =   "CBM"
            Height          =   225
            Left            =   4410
            TabIndex        =   323
            ToolTipText     =   "Enable for CBM. Disable for ASCII"
            Top             =   120
            Value           =   1  'Checked
            Width           =   795
         End
         Begin VB.CommandButton cmdScreenRefresh 
            Appearance      =   0  'Flat
            Caption         =   "Refresh"
            Height          =   315
            Left            =   9390
            TabIndex        =   281
            ToolTipText     =   "Refresh Screen"
            Top             =   60
            Width           =   825
         End
         Begin VB.CommandButton cmdInsert 
            Appearance      =   0  'Flat
            Caption         =   "INS"
            Height          =   315
            Left            =   3420
            TabIndex        =   279
            ToolTipText     =   "Insert selected Character (Shift-Insert)"
            Top             =   60
            Width           =   405
         End
         Begin VB.ComboBox cboScnFmt 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmViewer.frx":04C0
            Left            =   420
            List            =   "frmViewer.frx":04D0
            Style           =   2  'Dropdown List
            TabIndex        =   278
            ToolTipText     =   "Screen Format"
            Top             =   60
            Width           =   1125
         End
         Begin VB.PictureBox picScreen 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   6000
            Left            =   570
            ScaleHeight     =   400
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   640
            TabIndex        =   276
            Top             =   780
            Width           =   9600
         End
         Begin VB.Label lblREC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "REC"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   5880
            TabIndex        =   328
            ToolTipText     =   "RECord Macro"
            Top             =   120
            Width           =   435
         End
         Begin VB.Label lblRVS 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "RVS"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3870
            TabIndex        =   303
            ToolTipText     =   "RVS Mode (Click or CTRL-R)"
            Top             =   90
            Width           =   435
         End
         Begin VB.Label lblActive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "X"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   10290
            TabIndex        =   302
            ToolTipText     =   "Active Indicator / Close"
            Top             =   90
            Width           =   255
         End
         Begin VB.Label lblCursor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "@"
            Height          =   195
            Left            =   1590
            TabIndex        =   280
            ToolTipText     =   "Current Row,Col"
            Top             =   90
            Width           =   195
         End
         Begin VB.Label lblScnBorder 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H80000008&
            Height          =   6675
            Left            =   150
            TabIndex        =   277
            Top             =   450
            Width           =   10425
         End
      End
      Begin VB.Frame frTools 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   8025
         Left            =   60
         TabIndex        =   204
         Top             =   420
         Width           =   1515
         Begin VB.PictureBox cmdShift 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   435
            Index           =   3
            Left            =   960
            ScaleHeight     =   29
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   29
            TabIndex        =   320
            ToolTipText     =   "Shift Right"
            Top             =   300
            Width           =   435
         End
         Begin VB.PictureBox cmdShift 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   435
            Index           =   2
            Left            =   60
            ScaleHeight     =   29
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   29
            TabIndex        =   319
            ToolTipText     =   "Shift Left"
            Top             =   300
            Width           =   435
         End
         Begin VB.PictureBox cmdShift 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   435
            Index           =   1
            Left            =   510
            ScaleHeight     =   29
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   29
            TabIndex        =   318
            ToolTipText     =   "Shift Down"
            Top             =   540
            Width           =   435
         End
         Begin VB.PictureBox cmdShift 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   435
            Index           =   0
            Left            =   510
            ScaleHeight     =   29
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   29
            TabIndex        =   317
            ToolTipText     =   "Shift Up"
            Top             =   90
            Width           =   435
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Compare"
            Height          =   285
            Index           =   36
            Left            =   60
            TabIndex        =   273
            ToolTipText     =   "Compare"
            Top             =   7350
            Width           =   1365
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "New"
            Height          =   285
            Index           =   35
            Left            =   60
            TabIndex        =   272
            ToolTipText     =   "Start NEW set"
            Top             =   6960
            Width           =   1365
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Sel Set"
            Height          =   285
            Index           =   34
            Left            =   60
            TabIndex        =   252
            ToolTipText     =   "Select ALL"
            Top             =   4290
            Width           =   675
         End
         Begin VB.CheckBox cbShiftMode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1200
            TabIndex        =   235
            ToolTipText     =   "When checked pixels wrap to opposite side. When unset pixels are LOST!"
            Top             =   810
            Value           =   1  'Checked
            Width           =   195
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Clear"
            Height          =   255
            Index           =   4
            Left            =   60
            TabIndex        =   234
            ToolTipText     =   "Clear Character to Bg"
            Top             =   1500
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "RVS"
            Height          =   255
            HelpContextID   =   5
            Index           =   5
            Left            =   750
            TabIndex        =   233
            ToolTipText     =   "Invert pixels"
            Top             =   1500
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Bold"
            Height          =   255
            Index           =   6
            Left            =   60
            TabIndex        =   232
            ToolTipText     =   "Create Bold character"
            Top             =   1770
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Flip H"
            Height          =   255
            Index           =   10
            Left            =   60
            TabIndex        =   231
            ToolTipText     =   "Flip character top to bottom"
            Top             =   2310
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Flip V"
            Height          =   255
            Index           =   11
            Left            =   750
            TabIndex        =   230
            ToolTipText     =   "Flip character left to right"
            Top             =   2310
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Tall T"
            Height          =   255
            Index           =   12
            Left            =   60
            TabIndex        =   229
            ToolTipText     =   "Create Double-Tall TOP"
            Top             =   3120
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Tall B"
            Height          =   255
            Index           =   13
            Left            =   750
            TabIndex        =   228
            ToolTipText     =   "Create Double-Tall BOTTOM"
            Top             =   3120
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Wide L"
            Height          =   255
            Index           =   14
            Left            =   60
            TabIndex        =   227
            ToolTipText     =   "Create Double-Wide LEFT"
            Top             =   3390
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Wide R"
            Height          =   255
            Index           =   15
            Left            =   750
            TabIndex        =   226
            ToolTipText     =   "Create Double-Wide RIGHT "
            Top             =   3390
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "2x TL"
            Height          =   255
            Index           =   16
            Left            =   60
            TabIndex        =   225
            ToolTipText     =   "Create 2x TOP-LEFT"
            Top             =   3660
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "2x TR"
            Height          =   255
            Index           =   17
            Left            =   750
            TabIndex        =   224
            ToolTipText     =   "Create 2x TOP-RIGHT"
            Top             =   3660
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "2x BL"
            Height          =   255
            Index           =   18
            Left            =   60
            TabIndex        =   223
            ToolTipText     =   "Create 2x BOTTOM-LEFT"
            Top             =   3930
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "2x BR"
            Height          =   255
            Index           =   19
            Left            =   750
            TabIndex        =   222
            ToolTipText     =   "Create 2x BOTTOM-RIGHT"
            Top             =   3930
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Rot L"
            Height          =   255
            Index           =   8
            Left            =   60
            TabIndex        =   221
            ToolTipText     =   "Rotate character LEFT"
            Top             =   2040
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Rot R"
            Height          =   255
            Index           =   9
            Left            =   750
            TabIndex        =   220
            ToolTipText     =   "Rotate character RIGHT"
            Top             =   2040
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Und"
            Height          =   255
            Index           =   7
            Left            =   750
            TabIndex        =   219
            ToolTipText     =   "Create Underlined character (below crosshair)"
            Top             =   1770
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "SWAP"
            Height          =   285
            Index           =   20
            Left            =   60
            TabIndex        =   218
            ToolTipText     =   "Swap Sets 1 and 2"
            Top             =   5700
            Width           =   1365
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Sel ALL"
            Height          =   285
            Index           =   21
            Left            =   750
            TabIndex        =   217
            ToolTipText     =   "Select ALL"
            Top             =   4290
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Caption         =   "Copy"
            Height          =   285
            Index           =   22
            Left            =   60
            TabIndex        =   216
            ToolTipText     =   "Copy CHR or RANGE to clipboard"
            Top             =   4680
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Paste"
            Height          =   285
            Index           =   23
            Left            =   750
            TabIndex        =   215
            ToolTipText     =   "Paste clipboard to current selected position"
            Top             =   4680
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Restore Range "
            Height          =   285
            Index           =   24
            Left            =   60
            TabIndex        =   214
            ToolTipText     =   "Restore Character or Range"
            Top             =   6360
            Width           =   1365
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Restore"
            Height          =   285
            Index           =   25
            Left            =   60
            TabIndex        =   213
            ToolTipText     =   "Restore"
            Top             =   6660
            Width           =   1365
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Del C"
            Height          =   255
            Index           =   29
            Left            =   750
            TabIndex        =   212
            ToolTipText     =   "Delete COL to right of crosshair"
            Top             =   2850
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Ins C"
            Height          =   255
            Index           =   28
            Left            =   60
            TabIndex        =   211
            ToolTipText     =   "Insert COL to right of crosshair"
            Top             =   2850
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Del R"
            Height          =   255
            Index           =   27
            Left            =   750
            TabIndex        =   210
            ToolTipText     =   "Delete ROW below crosshair"
            Top             =   2580
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Ins R"
            Height          =   255
            Index           =   26
            Left            =   60
            TabIndex        =   209
            ToolTipText     =   "Insert ROW below crosshair"
            Top             =   2580
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Set Restore Pt"
            Height          =   285
            Index           =   30
            Left            =   60
            TabIndex        =   208
            ToolTipText     =   "Se a Restore Point"
            Top             =   6060
            Width           =   1365
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Cut"
            Height          =   315
            Index           =   31
            Left            =   60
            TabIndex        =   207
            ToolTipText     =   "Cut selection"
            Top             =   4980
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Insert"
            Height          =   315
            Index           =   32
            Left            =   750
            TabIndex        =   206
            ToolTipText     =   "Insert clipboard to current position"
            Top             =   4980
            Width           =   675
         End
         Begin VB.CommandButton cmdTool 
            Appearance      =   0  'Flat
            Caption         =   "Append"
            Height          =   315
            Index           =   33
            Left            =   330
            TabIndex        =   205
            ToolTipText     =   "Append clipboard to end of file"
            Top             =   5310
            Width           =   795
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shift:"
            Height          =   195
            Index           =   7
            Left            =   60
            TabIndex        =   239
            Top             =   60
            Width           =   405
         End
         Begin VB.Label lblPixelMode 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "BG"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   60
            TabIndex        =   238
            ToolTipText     =   "Draw using Background colour"
            Top             =   1140
            Width           =   435
         End
         Begin VB.Label lblPixelMode 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "FG"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   525
            TabIndex        =   237
            ToolTipText     =   "Draw using Foreground colour"
            Top             =   1140
            Width           =   405
         End
         Begin VB.Label lblPixelMode 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "XOR"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   960
            TabIndex        =   236
            ToolTipText     =   "Toggle pixel colour"
            Top             =   1140
            Width           =   435
         End
      End
      Begin VB.Frame frChr 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   7605
         Left            =   1650
         TabIndex        =   158
         Top             =   420
         Width           =   2055
         Begin VB.CommandButton cmdChrSel 
            Caption         =   ">"
            Height          =   255
            Index           =   5
            Left            =   1680
            TabIndex        =   245
            ToolTipText     =   "Next Char"
            Top             =   930
            Width           =   300
         End
         Begin VB.CommandButton cmdChrSel 
            Caption         =   "<"
            Height          =   255
            Index           =   4
            Left            =   1350
            TabIndex        =   244
            ToolTipText     =   "Previous Char"
            Top             =   930
            Width           =   300
         End
         Begin VB.CommandButton cmdChrSel 
            Caption         =   ">"
            Height          =   255
            Index           =   3
            Left            =   1020
            TabIndex        =   243
            ToolTipText     =   "Next Char in Set"
            Top             =   930
            Width           =   300
         End
         Begin VB.CommandButton cmdChrSel 
            Caption         =   "<"
            Height          =   255
            Index           =   2
            Left            =   690
            TabIndex        =   242
            ToolTipText     =   "Previous Char in Set"
            Top             =   930
            Width           =   300
         End
         Begin VB.CommandButton cmdChrSel 
            Caption         =   ">"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   241
            ToolTipText     =   "Next Set"
            Top             =   930
            Width           =   300
         End
         Begin VB.CommandButton cmdChrSel 
            Caption         =   "<"
            Height          =   255
            Index           =   0
            Left            =   30
            TabIndex        =   240
            ToolTipText     =   "Previous Set"
            Top             =   930
            Width           =   300
         End
         Begin VB.PictureBox picChr 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1830
            Left            =   30
            ScaleHeight     =   122
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   129
            TabIndex        =   159
            Top             =   1260
            Width           =   1935
         End
         Begin VB.Label lblOver 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   -60
            TabIndex        =   324
            Top             =   5430
            Width           =   1935
         End
         Begin VB.Label lblBufSel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "R.Pt"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   3
            Left            =   1470
            TabIndex        =   270
            ToolTipText     =   "Restore Point"
            Top             =   30
            Width           =   495
         End
         Begin VB.Label lblBufSel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Clip"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   960
            TabIndex        =   269
            ToolTipText     =   "Clipboard"
            Top             =   30
            Width           =   465
         End
         Begin VB.Label lblBufSel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   510
            TabIndex        =   268
            ToolTipText     =   "Buffer#2"
            Top             =   30
            Width           =   405
         End
         Begin VB.Label lblBufSel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   30
            TabIndex        =   267
            ToolTipText     =   "Buffer#1"
            Top             =   30
            Width           =   435
         End
         Begin VB.Label LabelC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "###"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   2
            Left            =   1350
            TabIndex        =   250
            Top             =   360
            Width           =   615
         End
         Begin VB.Label LabelC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "CHR"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   1
            Left            =   690
            TabIndex        =   249
            Top             =   360
            Width           =   615
         End
         Begin VB.Label LabelC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            Caption         =   "SET"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   0
            Left            =   30
            TabIndex        =   248
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblChrNum 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   690
            TabIndex        =   247
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblChrSet 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   30
            TabIndex        =   246
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblRange 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3855
            Left            =   30
            TabIndex        =   161
            Top             =   1260
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label lblChrSel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1350
            TabIndex        =   160
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblFStat 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1695
            Left            =   30
            TabIndex        =   163
            Top             =   5880
            Width           =   1935
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1290
         Left            =   3900
         ScaleHeight     =   86
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   119
         TabIndex        =   19
         Top             =   450
         Visible         =   0   'False
         Width           =   1785
      End
   End
   Begin VB.Frame frML 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Machine Language Disassembler"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11025
      Left            =   60
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   11100
      Begin VB.CommandButton cmdDTAdd 
         Appearance      =   0  'Flat
         Caption         =   "L"
         Height          =   315
         Index           =   9
         Left            =   8130
         TabIndex        =   203
         ToolTipText     =   "Make Litte-endian Word Block"
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdDTAdd 
         Appearance      =   0  'Flat
         Caption         =   "A"
         Height          =   315
         Index           =   8
         Left            =   9090
         TabIndex        =   202
         ToolTipText     =   "Make Assembly Block with Hex bytes"
         Top             =   120
         Width           =   285
      End
      Begin VB.CommandButton cmdMLSplit 
         Appearance      =   0  'Flat
         Caption         =   "/"
         Height          =   255
         Left            =   4680
         TabIndex        =   193
         ToolTipText     =   "Toggle Split View"
         Top             =   150
         Width           =   225
      End
      Begin VB.ListBox lstML2 
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   5310
         MultiSelect     =   2  'Extended
         TabIndex        =   192
         Top             =   1020
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdAddComment 
         Appearance      =   0  'Flat
         Caption         =   "[ ]"
         Height          =   315
         HelpContextID   =   7
         Index           =   8
         Left            =   9450
         TabIndex        =   183
         ToolTipText     =   "Add Block Comment"
         Top             =   120
         Width           =   315
      End
      Begin VB.Frame frBlock 
         Caption         =   "Add/Edit Block Comment"
         Height          =   2325
         Left            =   4050
         TabIndex        =   178
         Top             =   1770
         Visible         =   0   'False
         Width           =   7875
         Begin VB.CommandButton cmdBCancel 
            Appearance      =   0  'Flat
            Caption         =   "Cancel"
            Height          =   435
            Left            =   1050
            TabIndex        =   181
            Top             =   270
            Width           =   855
         End
         Begin VB.CommandButton cmdSaveBlock 
            Appearance      =   0  'Flat
            Caption         =   "Save"
            Height          =   435
            Left            =   150
            TabIndex        =   180
            Top             =   270
            Width           =   855
         End
         Begin VB.TextBox txtBlock 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1425
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   179
            Top             =   780
            Width           =   5955
         End
         Begin VB.Label Label 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmViewer.frx":04F8
            ForeColor       =   &H80000008&
            Height          =   585
            Index           =   10
            Left            =   3480
            TabIndex        =   186
            Top             =   180
            Width           =   4125
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblCPos 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1980
            TabIndex        =   185
            Top             =   570
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "At Location:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   9
            Left            =   1980
            TabIndex        =   184
            Top             =   360
            Width           =   915
         End
         Begin VB.Label lblBlockAddress 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "0000"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2940
            TabIndex        =   182
            Top             =   360
            Width           =   360
         End
      End
      Begin VB.CommandButton cmdDTAdd 
         Appearance      =   0  'Flat
         Caption         =   "Z"
         Height          =   315
         Index           =   7
         Left            =   8760
         TabIndex        =   129
         ToolTipText     =   "Make Binary Byte Block"
         Top             =   120
         Width           =   285
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         Caption         =   "< ]"
         Height          =   315
         Index           =   3
         Left            =   3930
         TabIndex        =   128
         ToolTipText     =   "Bottom Up "
         Top             =   120
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3630
         TabIndex        =   127
         ToolTipText     =   "Next Up"
         Top             =   120
         Width           =   285
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3330
         TabIndex        =   126
         ToolTipText     =   "Next Down"
         Top             =   120
         Width           =   285
      End
      Begin VB.Frame frInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3930
         TabIndex        =   124
         Top             =   480
         Width           =   8715
         Begin VB.Label Label 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Line#:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   8
            Left            =   60
            TabIndex        =   197
            Top             =   180
            Width           =   465
         End
         Begin VB.Label lblLineNum 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#####"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   600
            TabIndex        =   194
            ToolTipText     =   "Click to Enter line# to jump to"
            Top             =   180
            Width           =   585
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "Click table entry for info"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1740
            TabIndex        =   125
            Top             =   150
            Width           =   8430
         End
      End
      Begin VB.CommandButton cmdAddEP 
         Caption         =   "Ent.Pt"
         Height          =   315
         Left            =   5640
         TabIndex        =   123
         ToolTipText     =   "Add Label"
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdDTAdd 
         Appearance      =   0  'Flat
         Caption         =   "X"
         Height          =   315
         Index           =   6
         Left            =   8460
         TabIndex        =   120
         ToolTipText     =   "Make Hidden Block"
         Top             =   120
         Width           =   285
      End
      Begin VB.CommandButton cmdAddComment 
         Appearance      =   0  'Flat
         Caption         =   "*C*"
         Height          =   315
         Index           =   4
         Left            =   11370
         TabIndex        =   87
         ToolTipText     =   "Add Comment with * Separator"
         Top             =   120
         Width           =   435
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "**"
         Height          =   315
         HelpContextID   =   7
         Index           =   7
         Left            =   12660
         TabIndex        =   86
         ToolTipText     =   "Add * Separator"
         Top             =   120
         Width           =   405
      End
      Begin VB.CommandButton cmdDTAdd 
         Appearance      =   0  'Flat
         Caption         =   "W"
         Height          =   315
         Index           =   5
         Left            =   7830
         TabIndex        =   85
         ToolTipText     =   "Make Big Endian Word Block"
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdAddLabel 
         Caption         =   "Label"
         Height          =   315
         Left            =   5040
         TabIndex        =   84
         ToolTipText     =   "Add Label"
         Top             =   120
         Width           =   555
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "=="
         Height          =   315
         Index           =   6
         Left            =   12240
         TabIndex        =   83
         ToolTipText     =   "Add = Separator"
         Top             =   120
         Width           =   405
      End
      Begin VB.CommandButton cmdAddComment 
         Caption         =   "---"
         Height          =   315
         Index           =   5
         Left            =   11820
         TabIndex        =   82
         ToolTipText     =   "Add - Separator"
         Top             =   120
         Width           =   405
      End
      Begin VB.CommandButton cmdAddComment 
         Appearance      =   0  'Flat
         Caption         =   "=C="
         Height          =   315
         Index           =   3
         Left            =   10890
         TabIndex        =   81
         ToolTipText     =   "Add Comment with = Separator"
         Top             =   120
         Width           =   465
      End
      Begin VB.CommandButton cmdAddComment 
         Appearance      =   0  'Flat
         Caption         =   "-C-"
         Height          =   315
         Index           =   2
         Left            =   10380
         TabIndex        =   80
         ToolTipText     =   "Add Comment with - Separator"
         Top             =   120
         Width           =   465
      End
      Begin VB.CommandButton cmdAddComment 
         Appearance      =   0  'Flat
         Caption         =   "C"
         Height          =   315
         Index           =   1
         Left            =   10080
         TabIndex        =   79
         ToolTipText     =   "Add Standalone Comment"
         Top             =   120
         Width           =   285
      End
      Begin VB.CommandButton cmdAddComment 
         Appearance      =   0  'Flat
         Caption         =   ";C"
         Height          =   315
         Index           =   0
         Left            =   9780
         TabIndex        =   78
         ToolTipText     =   "Add Inline Comment"
         Top             =   120
         Width           =   285
      End
      Begin VB.CommandButton cmdDTAdd 
         Appearance      =   0  'Flat
         Caption         =   "V"
         Height          =   315
         Index           =   4
         Left            =   7530
         TabIndex        =   77
         ToolTipText     =   "Make Vector Block"
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdDTAdd 
         Appearance      =   0  'Flat
         Caption         =   "R"
         Height          =   315
         Index           =   3
         Left            =   7230
         TabIndex        =   76
         ToolTipText     =   "Make RTS vector block"
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdDTAdd 
         Appearance      =   0  'Flat
         Caption         =   "T"
         Height          =   315
         Index           =   2
         Left            =   6930
         TabIndex        =   75
         ToolTipText     =   "Make Text Block"
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdDTAdd 
         Appearance      =   0  'Flat
         Caption         =   "H"
         Height          =   315
         Index           =   1
         Left            =   6630
         TabIndex        =   74
         ToolTipText     =   "Make Hex Block"
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdDTAdd 
         Caption         =   "D"
         Height          =   315
         Index           =   0
         Left            =   6330
         TabIndex        =   73
         ToolTipText     =   "Make Dec Byte Block"
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdFindAll 
         Appearance      =   0  'Flat
         Caption         =   "All"
         Height          =   315
         Left            =   2580
         TabIndex        =   27
         ToolTipText     =   "Find all occurences"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         Caption         =   "[ >"
         Height          =   315
         Index           =   0
         Left            =   2970
         TabIndex        =   21
         ToolTipText     =   "Top Down"
         Top             =   120
         Width           =   345
      End
      Begin VB.CommandButton cmdFind 
         Appearance      =   0  'Flat
         Caption         =   "Find"
         Height          =   315
         Left            =   1980
         TabIndex        =   20
         ToolTipText     =   "Find Text"
         Top             =   120
         Width           =   555
      End
      Begin VB.CommandButton cmdRefresh 
         Appearance      =   0  'Flat
         Caption         =   "Refresh"
         Height          =   315
         Left            =   1110
         TabIndex        =   17
         Top             =   120
         Width           =   765
      End
      Begin VB.ListBox lstML 
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   3990
         MultiSelect     =   2  'Extended
         TabIndex        =   4
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Frame frTView 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9855
         Left            =   30
         TabIndex        =   28
         Top             =   450
         Width           =   3825
         Begin VB.Frame frTrace 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4425
            Left            =   0
            TabIndex        =   112
            Top             =   2085
            Width           =   3870
            Begin VB.CheckBox cbTraceLog 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Save Log"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   90
               TabIndex        =   330
               Top             =   1320
               Value           =   1  'Checked
               Width           =   1035
            End
            Begin VB.CheckBox cbVerb 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Verbose"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   90
               TabIndex        =   329
               Top             =   1050
               Value           =   1  'Checked
               Width           =   1035
            End
            Begin VB.CheckBox cbMLAddLabels 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   " Add Labels"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   90
               TabIndex        =   119
               Top             =   3090
               Value           =   1  'Checked
               Width           =   1155
            End
            Begin VB.CommandButton cmdAddTables 
               Appearance      =   0  'Flat
               Caption         =   "Add To Tables"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   90
               TabIndex        =   115
               Top             =   2340
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.CommandButton cmdTrace 
               Appearance      =   0  'Flat
               Caption         =   "START"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   675
               Left            =   90
               TabIndex        =   114
               Top             =   270
               Width           =   1005
            End
            Begin VB.ListBox lstEP 
               Appearance      =   0  'Flat
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
               Height          =   3930
               ItemData        =   "frmViewer.frx":05AC
               Left            =   1290
               List            =   "frmViewer.frx":05AE
               Sorted          =   -1  'True
               TabIndex        =   113
               Top             =   240
               Width           =   2340
            End
         End
         Begin VB.Frame frMLSettings 
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7035
            Left            =   -90
            TabIndex        =   42
            Top             =   2010
            Width           =   3870
            Begin VB.CheckBox cbOpUCase 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Uppercase Opcodes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1710
               TabIndex        =   265
               ToolTipText     =   "Include Equates in output"
               Top             =   3600
               Width           =   2055
            End
            Begin VB.CheckBox cbHHHH 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Add Hex Address to Standalone Comments"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   150
               TabIndex        =   201
               Top             =   4800
               Width           =   3585
            End
            Begin VB.CommandButton cmdProjSave 
               Caption         =   "Save"
               Height          =   315
               Left            =   750
               TabIndex        =   199
               ToolTipText     =   "Save Project"
               Top             =   60
               Width           =   615
            End
            Begin VB.CheckBox cbCompareOut 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Compare"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   2340
               TabIndex        =   196
               Top             =   5580
               Value           =   1  'Checked
               Width           =   1275
            End
            Begin VB.CommandButton cmdReassemble 
               Appearance      =   0  'Flat
               Caption         =   "Re-Assemble"
               Height          =   345
               Left            =   1110
               TabIndex        =   195
               ToolTipText     =   "Re-Assemble with ACME"
               Top             =   5520
               Width           =   1155
            End
            Begin VB.TextBox txtLineInc 
               Appearance      =   0  'Flat
               BackColor       =   &H00000080&
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   3210
               MaxLength       =   3
               TabIndex        =   190
               Text            =   "10"
               Top             =   3180
               Width           =   435
            End
            Begin VB.TextBox txtStartLine 
               Appearance      =   0  'Flat
               BackColor       =   &H00000080&
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1350
               MaxLength       =   4
               TabIndex        =   188
               Text            =   "100"
               Top             =   3180
               Width           =   675
            End
            Begin VB.CheckBox cbBlock 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Show Block Comments"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   150
               TabIndex        =   187
               ToolTipText     =   "Include Equates in output"
               Top             =   3840
               Value           =   1  'Checked
               Width           =   3615
            End
            Begin VB.TextBox txtInlineCol 
               Appearance      =   0  'Flat
               BackColor       =   &H00000080&
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2100
               MaxLength       =   2
               TabIndex        =   122
               Text            =   "50"
               Top             =   2850
               Width           =   345
            End
            Begin VB.CommandButton cmdImport 
               Caption         =   "Import"
               Height          =   315
               Left            =   2040
               TabIndex        =   93
               ToolTipText     =   "Import Symbols"
               Top             =   5940
               Width           =   1605
            End
            Begin VB.CheckBox cbIncSym 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Include Symbol comments"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   150
               TabIndex        =   72
               Top             =   4560
               Value           =   1  'Checked
               Width           =   3585
            End
            Begin VB.ComboBox cboCPUFile 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               CausesValidation=   0   'False
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "frmViewer.frx":05B0
               Left            =   2940
               List            =   "frmViewer.frx":05B2
               Style           =   2  'Dropdown List
               TabIndex        =   71
               Top             =   2700
               Visible         =   0   'False
               Width           =   1305
            End
            Begin VB.ComboBox cboCPU 
               Appearance      =   0  'Flat
               BackColor       =   &H00000080&
               CausesValidation=   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmViewer.frx":05B4
               Left            =   870
               List            =   "frmViewer.frx":05BB
               Style           =   2  'Dropdown List
               TabIndex        =   69
               Top             =   1200
               Width           =   2835
            End
            Begin VB.TextBox txtDivLen 
               Appearance      =   0  'Flat
               BackColor       =   &H00000080&
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2100
               MaxLength       =   2
               TabIndex        =   68
               Text            =   "80"
               Top             =   2520
               Width           =   345
            End
            Begin VB.ComboBox cboPlatFile 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               CausesValidation=   0   'False
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "frmViewer.frx":05CD
               Left            =   2820
               List            =   "frmViewer.frx":05CF
               Style           =   2  'Dropdown List
               TabIndex        =   66
               Top             =   2520
               Visible         =   0   'False
               Width           =   1305
            End
            Begin VB.ComboBox cboPlatform 
               Appearance      =   0  'Flat
               BackColor       =   &H00000080&
               CausesValidation=   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmViewer.frx":05D1
               Left            =   870
               List            =   "frmViewer.frx":05D8
               Style           =   2  'Dropdown List
               TabIndex        =   65
               Top             =   870
               Width           =   2835
            End
            Begin VB.CommandButton cmdMLHelp 
               Caption         =   "Help"
               Height          =   465
               Left            =   600
               TabIndex        =   63
               ToolTipText     =   "Display HELP file"
               Top             =   6420
               Width           =   2385
            End
            Begin VB.CheckBox cbLabelBlanks 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Add blank line before Labels"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   150
               TabIndex        =   62
               Top             =   4320
               Value           =   1  'Checked
               Width           =   3585
            End
            Begin VB.ComboBox cboPrefix 
               Appearance      =   0  'Flat
               BackColor       =   &H00000080&
               CausesValidation=   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmViewer.frx":05EA
               Left            =   1080
               List            =   "frmViewer.frx":05F1
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   2190
               Width           =   2625
            End
            Begin VB.ComboBox cboTarget 
               Appearance      =   0  'Flat
               BackColor       =   &H00000080&
               CausesValidation=   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmViewer.frx":0603
               Left            =   870
               List            =   "frmViewer.frx":0616
               Style           =   2  'Dropdown List
               TabIndex        =   58
               Top             =   1860
               Width           =   2835
            End
            Begin VB.CommandButton cmdSaveASM 
               Appearance      =   0  'Flat
               Caption         =   "Save..."
               Height          =   345
               Left            =   1110
               TabIndex        =   54
               ToolTipText     =   "Save disassembly to file"
               Top             =   5130
               Width           =   735
            End
            Begin VB.CheckBox cbSpaceRTS 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Add blank line after RTS/RTI instructions"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   150
               TabIndex        =   53
               Top             =   4080
               Value           =   1  'Checked
               Width           =   3615
            End
            Begin VB.CommandButton cmdPurge 
               Caption         =   "Purge"
               Height          =   315
               Left            =   1110
               TabIndex        =   52
               ToolTipText     =   "Purge unselected symbol entries"
               Top             =   5940
               Width           =   765
            End
            Begin VB.CommandButton cmdClrTables 
               Caption         =   "New"
               Height          =   315
               Left            =   2940
               TabIndex        =   51
               ToolTipText     =   "Clear Lists and start a new project"
               Top             =   60
               Width           =   735
            End
            Begin VB.CheckBox cbClearOnLoad 
               Caption         =   "Clear Lists on Load"
               Height          =   195
               Left            =   120
               TabIndex        =   50
               ToolTipText     =   "Uncheck if you want to keep existing entries when loading"
               Top             =   570
               Value           =   1  'Checked
               Width           =   1755
            End
            Begin VB.CommandButton cmdProjSaveAs 
               Caption         =   "Save As..."
               Height          =   315
               Left            =   1650
               TabIndex        =   49
               ToolTipText     =   "Prompt and Save Project "
               Top             =   60
               Width           =   885
            End
            Begin VB.CommandButton cmdProjLoad 
               Caption         =   "Load..."
               Height          =   315
               Left            =   30
               TabIndex        =   48
               ToolTipText     =   "Load Lists from a file"
               Top             =   60
               Width           =   705
            End
            Begin VB.CommandButton cmdCopyClip2 
               Appearance      =   0  'Flat
               Caption         =   "Copy To &Clipboard"
               Height          =   345
               Left            =   2010
               TabIndex        =   46
               ToolTipText     =   "Paste disassembly to clipboard"
               Top             =   5130
               Width           =   1635
            End
            Begin VB.CheckBox cbEquates 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Show Equates"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   150
               TabIndex        =   45
               ToolTipText     =   "Include Equates in output"
               Top             =   3600
               Width           =   1395
            End
            Begin VB.ComboBox cboMLFmt 
               Appearance      =   0  'Flat
               BackColor       =   &H00000080&
               CausesValidation=   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmViewer.frx":064C
               Left            =   870
               List            =   "frmViewer.frx":065F
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   1530
               Width           =   2835
            End
            Begin VB.Label lblUpdated 
               BackColor       =   &H80000018&
               Caption         =   "-"
               Height          =   225
               Left            =   1980
               TabIndex        =   200
               ToolTipText     =   "Last Update"
               Top             =   570
               Width           =   1575
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Increment:"
               Height          =   195
               Index           =   21
               Left            =   2340
               TabIndex        =   191
               Top             =   3210
               Width           =   810
            End
            Begin VB.Label Label 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Starting Line#:"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   18
               Left            =   150
               TabIndex        =   189
               Top             =   3210
               Width           =   1125
            End
            Begin VB.Label Label 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Inline Comment col:"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   17
               Left            =   150
               TabIndex        =   121
               Top             =   2880
               Width           =   1530
            End
            Begin VB.Label lblChanged 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   2610
               TabIndex        =   94
               ToolTipText     =   "Project Status (Green=OK, Red=Changed, Grey=No Project Loaded)"
               Top             =   120
               Width           =   195
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Symbols:"
               Height          =   195
               Index           =   20
               Left            =   360
               TabIndex        =   92
               Top             =   6000
               Width           =   675
            End
            Begin VB.Label Label 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "CPU:"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   12
               Left            =   390
               TabIndex        =   70
               Top             =   1260
               Width           =   360
            End
            Begin VB.Label Label 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Comment Divider length:"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   67
               Top             =   2550
               Width           =   1920
            End
            Begin VB.Label Label 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Platform:"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   11
               Left            =   150
               TabIndex        =   64
               Top             =   930
               Width           =   690
            End
            Begin VB.Label Label 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Disassembly:"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   19
               Left            =   90
               TabIndex        =   61
               Top             =   5190
               Width           =   975
            End
            Begin VB.Label Label 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Label Prefix:"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   15
               Left            =   120
               TabIndex        =   59
               Top             =   2250
               Width           =   915
            End
            Begin VB.Label Label 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Target:"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   14
               Left            =   270
               TabIndex        =   57
               Top             =   1920
               Width           =   540
            End
            Begin VB.Label Label 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "View Fmt:"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   13
               Left            =   90
               TabIndex        =   44
               Top             =   1590
               Width           =   750
            End
         End
         Begin VB.ListBox lstEntryPt 
            Appearance      =   0  'Flat
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
            Height          =   225
            ItemData        =   "frmViewer.frx":06B4
            Left            =   90
            List            =   "frmViewer.frx":06B6
            TabIndex        =   118
            Top             =   1590
            Width           =   705
         End
         Begin VB.ListBox lstJSR 
            Appearance      =   0  'Flat
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
            Height          =   225
            ItemData        =   "frmViewer.frx":06B8
            Left            =   2940
            List            =   "frmViewer.frx":06BA
            Sorted          =   -1  'True
            TabIndex        =   89
            Top             =   1320
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.ListBox lstLabels 
            Appearance      =   0  'Flat
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
            Height          =   225
            ItemData        =   "frmViewer.frx":06BC
            Left            =   2010
            List            =   "frmViewer.frx":06BE
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   55
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.ListBox lstCmnt 
            Appearance      =   0  'Flat
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
            Height          =   225
            ItemData        =   "frmViewer.frx":06C0
            Left            =   2940
            List            =   "frmViewer.frx":06C2
            Sorted          =   -1  'True
            TabIndex        =   41
            Top             =   1590
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.ListBox lstULabels 
            Appearance      =   0  'Flat
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
            Height          =   225
            ItemData        =   "frmViewer.frx":06C4
            Left            =   2190
            List            =   "frmViewer.frx":06C6
            Sorted          =   -1  'True
            TabIndex        =   39
            Top             =   1590
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.ListBox lstDT 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmViewer.frx":06C8
            Left            =   1530
            List            =   "frmViewer.frx":06CA
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   38
            Top             =   1590
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.ListBox lstSYM 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmViewer.frx":06CC
            Left            =   840
            List            =   "frmViewer.frx":06CE
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   37
            Top             =   1590
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.CommandButton cmdSymAdd 
            Appearance      =   0  'Flat
            Caption         =   "Add"
            Height          =   315
            Left            =   2130
            TabIndex        =   36
            ToolTipText     =   "Add an entry"
            Top             =   750
            Width           =   495
         End
         Begin VB.CommandButton cmdSymDel 
            Appearance      =   0  'Flat
            Caption         =   "Del"
            Height          =   315
            Left            =   2670
            TabIndex        =   35
            ToolTipText     =   "Delete current entry"
            Top             =   750
            Width           =   495
         End
         Begin VB.CommandButton cmdSYMGoto 
            Appearance      =   0  'Flat
            Caption         =   "Find"
            Height          =   315
            Left            =   3210
            TabIndex        =   34
            ToolTipText     =   "Find Selected"
            Top             =   750
            Width           =   555
         End
         Begin VB.CommandButton cmdSymSave 
            Appearance      =   0  'Flat
            Caption         =   "Save"
            Height          =   315
            Left            =   660
            TabIndex        =   33
            ToolTipText     =   "Save file"
            Top             =   750
            Width           =   555
         End
         Begin VB.CommandButton cmdSymLoad 
            Appearance      =   0  'Flat
            Caption         =   "Load"
            Height          =   315
            Left            =   60
            TabIndex        =   32
            ToolTipText     =   "Load a file"
            Top             =   750
            Width           =   555
         End
         Begin VB.CommandButton cmdRemDupLbls 
            Appearance      =   0  'Flat
            Caption         =   "Remove Duplicates"
            Height          =   315
            Left            =   60
            TabIndex        =   91
            ToolTipText     =   "Remove Duplicate Entries"
            Top             =   750
            Width           =   1845
         End
         Begin VB.CommandButton cmdRemDupJSR 
            Caption         =   "Remove Duplicates"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   90
            ToolTipText     =   "Remove Duplicate Entries"
            Top             =   750
            Width           =   1845
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "EntryPt"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   60
            TabIndex        =   117
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TRACER"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1050
            TabIndex        =   111
            Top             =   60
            Width           =   870
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ext JSR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   8
            Left            =   2910
            TabIndex        =   88
            Top             =   60
            Width           =   840
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Gen Labels"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   7
            Left            =   1950
            TabIndex        =   56
            Top             =   60
            Width           =   930
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PROJECT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   60
            TabIndex        =   47
            Top             =   60
            Width           =   960
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Comments"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   2910
            TabIndex        =   40
            Top             =   360
            Width           =   840
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tables"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   1560
            TabIndex        =   31
            Top             =   360
            Width           =   630
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Labels"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   5
            Left            =   2220
            TabIndex        =   30
            Top             =   360
            Width           =   660
         End
         Begin VB.Label lblTView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Symbols"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   810
            TabIndex        =   29
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.CheckBox cbAuto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   870
         TabIndex        =   26
         ToolTipText     =   "Automatically Refresh"
         Top             =   180
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.Image imgShowInfo 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4380
         Picture         =   "frmViewer.frx":06D0
         ToolTipText     =   "Toggle Info box"
         Top             =   150
         Width           =   255
      End
      Begin VB.Image imgBW 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   300
         Picture         =   "frmViewer.frx":0A86
         Top             =   150
         Width           =   255
      End
      Begin VB.Label lblShw 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   60
         TabIndex        =   25
         Top             =   30
         Width           =   255
      End
      Begin VB.Label lblGood 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   570
         TabIndex        =   16
         ToolTipText     =   "Disassembly Status (Green=OK, Red=Problems)"
         Top             =   150
         Width           =   225
      End
      Begin VB.Label lblEA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   195
         Left            =   13170
         TabIndex        =   15
         ToolTipText     =   "Address range"
         Top             =   150
         Width           =   90
      End
   End
   Begin VB.Frame frMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      TabIndex        =   282
      Top             =   0
      Width           =   16815
      Begin VB.TextBox txtLA 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7890
         TabIndex        =   285
         Text            =   "0000"
         ToolTipText     =   "Load Address from File, or Entered manually"
         Top             =   90
         Width           =   465
      End
      Begin VB.CheckBox cbLA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "LA:"
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   7350
         TabIndex        =   284
         ToolTipText     =   "Set if file includes Load Address  (first 2 bytes)"
         Top             =   120
         Value           =   1  'Checked
         Width           =   555
      End
      Begin VB.CheckBox cbLockView 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Lock"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6510
         TabIndex        =   283
         ToolTipText     =   "Lock to Current View"
         Top             =   120
         Width           =   675
      End
      Begin VB.Shape shOverflow 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   210
         Left            =   9390
         Shape           =   3  'Circle
         Top             =   120
         Width           =   210
      End
      Begin VB.Label lblVSize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "00000"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   8820
         LinkTimeout     =   0
         TabIndex        =   299
         Top             =   135
         Width           =   450
      End
      Begin VB.Label lblSz 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8460
         TabIndex        =   298
         Top             =   120
         Width           =   345
      End
      Begin VB.Label lblSSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "||"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   10770
         TabIndex        =   297
         ToolTipText     =   "Return split to CENTRE"
         Top             =   90
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lblSSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   11070
         TabIndex        =   296
         ToolTipText     =   "Move Split RIGHT"
         Top             =   90
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lblSSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   10470
         TabIndex        =   295
         ToolTipText     =   "Move Split LEFT"
         Top             =   90
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lblSelect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   10050
         TabIndex        =   294
         ToolTipText     =   "Select LEFT/RIGHT View"
         Top             =   90
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblSplit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   9690
         TabIndex        =   293
         ToolTipText     =   "Toggle Dual View Mode"
         Top             =   90
         Width           =   345
      End
      Begin VB.Label lblView 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         Caption         =   "PIC"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   5
         Left            =   5520
         TabIndex        =   292
         Top             =   60
         Width           =   915
      End
      Begin VB.Label lblView 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         Caption         =   "ASM"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   4
         Left            =   4560
         TabIndex        =   291
         Top             =   60
         Width           =   915
      End
      Begin VB.Label lblView 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "FONT"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   3
         Left            =   3600
         TabIndex        =   290
         Top             =   60
         Width           =   915
      End
      Begin VB.Label lblView 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         Caption         =   "HEX"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   2
         Left            =   2640
         TabIndex        =   289
         Top             =   60
         Width           =   915
      End
      Begin VB.Label lblView 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         Caption         =   "SEQ"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   1
         Left            =   1680
         TabIndex        =   288
         Top             =   60
         Width           =   915
      End
      Begin VB.Label lblView 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "BASIC"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   0
         Left            =   720
         TabIndex        =   287
         Top             =   60
         Width           =   915
      End
      Begin VB.Label lblViewAs 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "View As:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   286
         Top             =   105
         Width           =   645
      End
   End
   Begin VB.Timer SETimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   12540
      Top             =   570
   End
   Begin VB.Frame frBlank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   9870
      TabIndex        =   23
      Top             =   1230
      Visible         =   0   'False
      Width           =   2925
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Viewer with button above..."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   24
         Top             =   180
         Width           =   2640
      End
   End
   Begin VB.Frame frBasic 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "BASIC Lister"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   30
      TabIndex        =   1
      Top             =   450
      Visible         =   0   'False
      Width           =   9780
      Begin VB.PictureBox picView 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   0
         Left            =   1470
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   59
         TabIndex        =   258
         Top             =   810
         Width           =   885
      End
      Begin VB.VScrollBar vsView 
         Height          =   375
         Index           =   0
         Left            =   2370
         TabIndex        =   256
         Top             =   810
         Width           =   300
      End
      Begin VB.ListBox lstView 
         Appearance      =   0  'Flat
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
         Height          =   225
         Index           =   0
         ItemData        =   "frmViewer.frx":0E3C
         Left            =   105
         List            =   "frmViewer.frx":0E3E
         TabIndex        =   2
         Top             =   810
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Frame frBOpts 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   96
         Top             =   30
         Width           =   9465
         Begin VB.ComboBox cboColWidth2 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmViewer.frx":0E40
            Left            =   1050
            List            =   "frmViewer.frx":0E50
            Style           =   2  'Dropdown List
            TabIndex        =   322
            Top             =   360
            Width           =   1440
         End
         Begin VB.PictureBox cmdFSize 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   420
            ScaleHeight     =   19
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   19
            TabIndex        =   311
            Top             =   390
            Width           =   285
         End
         Begin VB.PictureBox cmdEncode 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   90
            ScaleHeight     =   19
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   19
            TabIndex        =   310
            Top             =   390
            Width           =   285
         End
         Begin VB.ComboBox cboMode 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmViewer.frx":0E76
            Left            =   570
            List            =   "frmViewer.frx":0E8C
            Style           =   2  'Dropdown List
            TabIndex        =   104
            ToolTipText     =   "Set BASIC Level"
            Top             =   0
            Width           =   1920
         End
         Begin VB.CommandButton cmdCpyClip 
            Appearance      =   0  'Flat
            Caption         =   "To &Clipboard"
            Height          =   315
            Left            =   8130
            TabIndex        =   103
            ToolTipText     =   "Copy View to clipboard"
            Top             =   30
            Width           =   1215
         End
         Begin VB.CommandButton cmdSave 
            Appearance      =   0  'Flat
            Caption         =   "E&xport"
            Height          =   315
            Index           =   0
            Left            =   8130
            TabIndex        =   101
            ToolTipText     =   "Export View to file"
            Top             =   390
            Width           =   1215
         End
         Begin VB.CheckBox cbUC 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "UCase)"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6750
            TabIndex        =   97
            ToolTipText     =   "Expand using UpperCase"
            Top             =   480
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox cbPad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Pad &Tokens"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5190
            TabIndex        =   98
            ToolTipText     =   "Append SPACE to tokens"
            Top             =   240
            Width           =   1515
         End
         Begin VB.CheckBox cbOneLine 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Break Multi"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6750
            TabIndex        =   99
            ToolTipText     =   "Break multi-statement lines (list one statement per line)"
            Top             =   30
            Width           =   1200
         End
         Begin VB.CheckBox cbExp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Expand &Special ("
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5190
            TabIndex        =   100
            ToolTipText     =   "Expand special characters (ie {RVS} )"
            Top             =   480
            Width           =   1560
         End
         Begin VB.CheckBox cbRev 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Reverse Text"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5190
            TabIndex        =   102
            ToolTipText     =   "Reverse display of Text"
            Top             =   30
            Width           =   1485
         End
         Begin VB.CheckBox cbMV 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "MV"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6750
            TabIndex        =   177
            ToolTipText     =   "Enable Magic Voice tokens"
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label lblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2640
            TabIndex        =   257
            Top             =   300
            Width           =   2460
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "BASIC:"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   108
            Top             =   60
            Width           =   480
         End
         Begin VB.Label lblGuess 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "-"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   3885
            TabIndex        =   107
            ToolTipText     =   "Computer model"
            Top             =   45
            Width           =   1245
         End
         Begin VB.Label Label 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "LOAD:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   2670
            TabIndex        =   106
            Top             =   45
            Width           =   480
         End
         Begin VB.Label lblLoadAdr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "-"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3210
            TabIndex        =   105
            ToolTipText     =   "Load Address"
            Top             =   45
            Width           =   600
         End
      End
      Begin VB.Label lblBView 
         BackStyle       =   0  'Transparent
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
         Left            =   0
         TabIndex        =   300
         ToolTipText     =   "Toggle Options pane"
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame frBIN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      Caption         =   "Binary Viewer"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   60
      TabIndex        =   5
      Top             =   2940
      Visible         =   0   'False
      Width           =   11610
      Begin VB.CheckBox cbByte 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Hex ("
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   750
         TabIndex        =   331
         ToolTipText     =   "Enable character view"
         Top             =   450
         Value           =   1  'Checked
         Width           =   675
      End
      Begin VB.CheckBox cbUpper 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Case)"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1440
         TabIndex        =   321
         ToolTipText     =   "Change case of Hex digits A-Z"
         Top             =   450
         Width           =   735
      End
      Begin VB.PictureBox cmdFSize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   390
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   315
         Top             =   390
         Width           =   285
      End
      Begin VB.PictureBox cmdEncode 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   60
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   313
         Top             =   390
         Width           =   285
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         Caption         =   "E&xport"
         Height          =   315
         Index           =   2
         Left            =   6210
         TabIndex        =   309
         ToolTipText     =   "Export View to file"
         Top             =   60
         Width           =   1230
      End
      Begin VB.CheckBox cbDifs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Difs only"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   9300
         TabIndex        =   274
         ToolTipText     =   "Show only lines with Differences"
         Top             =   450
         Width           =   1065
      End
      Begin VB.CheckBox cbGraphics 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Graphics ("
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3120
         TabIndex        =   266
         ToolTipText     =   "Enable CBM Graphics"
         Top             =   450
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.ListBox lstView 
         Appearance      =   0  'Flat
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
         Height          =   225
         Index           =   2
         ItemData        =   "frmViewer.frx":0EFD
         Left            =   90
         List            =   "frmViewer.frx":0EFF
         TabIndex        =   264
         Top             =   720
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.VScrollBar vsView 
         Height          =   405
         Index           =   2
         Left            =   2370
         TabIndex        =   262
         Top             =   720
         Width           =   300
      End
      Begin VB.PictureBox picView 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   2
         Left            =   1440
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   260
         Top             =   720
         Width           =   915
      End
      Begin VB.CommandButton cmdMD5 
         Caption         =   "MD5"
         Height          =   285
         Left            =   10440
         TabIndex        =   198
         ToolTipText     =   "Calculate MD5 ID"
         Top             =   420
         Width           =   525
      End
      Begin VB.CheckBox cbCmpShow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show:"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   9300
         TabIndex        =   174
         ToolTipText     =   "Show Compare file with differences"
         Top             =   90
         Width           =   795
      End
      Begin VB.CommandButton cmdCompare 
         Caption         =   "Load Compare..."
         Height          =   315
         Left            =   7470
         TabIndex        =   173
         ToolTipText     =   "Load a compare file"
         Top             =   60
         Width           =   1455
      End
      Begin VB.CheckBox cbHexFmt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ASM Fmt"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   7470
         TabIndex        =   172
         ToolTipText     =   "Use ASM .BYTE format"
         Top             =   450
         Width           =   975
      End
      Begin VB.CommandButton cmdHNext 
         Appearance      =   0  'Flat
         Caption         =   "Next"
         Height          =   285
         Left            =   3780
         TabIndex        =   168
         ToolTipText     =   "Find NEXT occurance"
         Top             =   60
         Width           =   585
      End
      Begin VB.CommandButton cmdHexFind 
         Appearance      =   0  'Flat
         Caption         =   "Find"
         Height          =   285
         Left            =   3090
         TabIndex        =   167
         Top             =   60
         Width           =   675
      End
      Begin VB.TextBox txtHSS 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   510
         TabIndex        =   165
         ToolTipText     =   "Enter text string to search for, or start with ""$""  to search for HEX value(s) - Do not use spaces"
         Top             =   60
         Width           =   2505
      End
      Begin VB.CheckBox cbHexSync 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ASM Sync"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6300
         TabIndex        =   110
         ToolTipText     =   "Sync to ASM list"
         Top             =   450
         Width           =   1065
      End
      Begin VB.CheckBox cbWide 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Wide"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5070
         TabIndex        =   95
         ToolTipText     =   "Wide (16 bytes)"
         Top             =   450
         Width           =   705
      End
      Begin VB.CheckBox cb7bit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "7-bit)"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4200
         TabIndex        =   22
         ToolTipText     =   "Strip off upper bit"
         Top             =   450
         Width           =   705
      End
      Begin VB.CheckBox cbShowP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Font"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2310
         TabIndex        =   7
         ToolTipText     =   "Enable character view"
         Top             =   450
         Value           =   1  'Checked
         Width           =   705
      End
      Begin VB.Label lblCFile 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "no file"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   10410
         TabIndex        =   176
         ToolTipText     =   "Compare Filename"
         Top             =   90
         Width           =   2895
      End
      Begin VB.Label lblDifTxt 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Height          =   195
         Left            =   11010
         TabIndex        =   175
         ToolTipText     =   "Compare Summary"
         Top             =   450
         Width           =   45
      End
      Begin VB.Label lblSResults 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "No search set"
         Height          =   195
         Left            =   4380
         TabIndex        =   169
         Top             =   90
         Width           =   1035
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FIND:"
         Height          =   195
         Index           =   24
         Left            =   60
         TabIndex        =   166
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.PictureBox Pix 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   18960
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Frame frBMP 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Pic Viewer"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   9870
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton cmdLoadVPL 
         Appearance      =   0  'Flat
         Caption         =   "Load VPL..."
         Height          =   285
         Left            =   1080
         TabIndex        =   164
         ToolTipText     =   "Load VICE Palette file"
         Top             =   90
         Width           =   975
      End
      Begin VB.CommandButton cmdBSave 
         Appearance      =   0  'Flat
         Caption         =   "Save..."
         Height          =   285
         Left            =   60
         TabIndex        =   12
         ToolTipText     =   "Save to BMP file"
         Top             =   90
         Width           =   975
      End
      Begin VB.PictureBox picBMP 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   60
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   9
         Top             =   510
         Width           =   975
      End
      Begin VB.Label lblMoment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "One moment... loading BMP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   116
         Top             =   780
         Width           =   1980
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comment:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   2610
         TabIndex        =   14
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Format:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   2610
         TabIndex        =   13
         Top             =   60
         Width           =   585
      End
      Begin VB.Label lblBType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3450
         TabIndex        =   11
         Top             =   60
         Width           =   45
      End
      Begin VB.Label lblBComment 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3450
         TabIndex        =   10
         Top             =   270
         Width           =   45
      End
   End
   Begin VB.Frame frSEQ 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Caption         =   "SEQ Viewer"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   30
      TabIndex        =   6
      Top             =   1770
      Visible         =   0   'False
      Width           =   9780
      Begin VB.PictureBox cmdFSize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   390
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   314
         Top             =   90
         Width           =   285
      End
      Begin VB.PictureBox cmdEncode 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   60
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   312
         Top             =   90
         Width           =   285
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         Caption         =   "E&xport"
         Height          =   315
         Index           =   1
         Left            =   8340
         TabIndex        =   308
         ToolTipText     =   "Export View to file"
         Top             =   60
         Width           =   1230
      End
      Begin VB.CheckBox cbIgnoreUnP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Un-Printable"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4110
         TabIndex        =   307
         ToolTipText     =   "Replace Un-printable with ""."""
         Top             =   120
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.ComboBox cboColWidth 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmViewer.frx":0F01
         Left            =   810
         List            =   "frmViewer.frx":0F17
         Style           =   2  'Dropdown List
         TabIndex        =   305
         Top             =   60
         Width           =   1440
      End
      Begin VB.CheckBox cbIgnoreCR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "CR"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2850
         TabIndex        =   304
         ToolTipText     =   "Strip Carriage Returns"
         Top             =   120
         Value           =   1  'Checked
         Width           =   555
      End
      Begin VB.ListBox lstView 
         Appearance      =   0  'Flat
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
         Height          =   225
         Index           =   1
         ItemData        =   "frmViewer.frx":0F4C
         Left            =   90
         List            =   "frmViewer.frx":0F4E
         TabIndex        =   263
         Top             =   480
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.VScrollBar vsView 
         Height          =   405
         Index           =   1
         Left            =   2370
         TabIndex        =   261
         Top             =   480
         Width           =   300
      End
      Begin VB.PictureBox picView 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   1
         Left            =   1440
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   259
         Top             =   480
         Width           =   915
      End
      Begin VB.CheckBox cbIgnoreLF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "LF"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3480
         TabIndex        =   109
         ToolTipText     =   "Strip Linefeeds"
         Top             =   120
         Value           =   1  'Checked
         Width           =   495
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Strip:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   2370
         TabIndex        =   306
         Top             =   120
         Width           =   405
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   13050
      Top             =   510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2021 Steve J. Gray
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

Public VBufNum          As Integer                                              'Current Buffer View (font editor)
Public VBufLast         As Integer                                              'Last BufferNum (font editor)

Public VBuf             As String                                               'ViewFile Visible/Edit Buffer - All viewers share this buffer
Public VBuf1            As String                                               'ViewFile Primary   Buffer (font editor)
Public VBuf2            As String                                               'ViewFile Secondary Buffer (font editor)
Public VClip            As String                                               'ViewFile Clipboard Buffer (font editor)
Public VRestore         As String                                               'ViewFile Restore   Buffer (font editor)

Public VFileName As String, VName As String, VExt As String                     'ViewFile Info
Public VLen As Long, VLA As Long                                                'ViewFile Length, Load Address
Public VP00Buf As String, VP00Flag As Boolean                                   'ViewFile P00 buffer, and flag
Public ViewReady As Boolean, ViewBusy As Boolean                                'Flag when processing
Public ViewMode As Integer, ViewMode2 As Integer                                'Which tabs are displayed
Public LockV1 As Integer, LockV2 As Integer                                     'Which tabs are locked
Public SplitMode As Boolean, SplitSize As Integer                               'Dual-view split

Public LastTheme As Integer

'==== Bitmap Viewer
Const NUMB = 20, GEO = -1, HRBW = 0, HR = 1, MC = 2

Dim PBuf As String                                              'Picture Buffer
Dim PicName As String
Dim CBMColor(15) As Long                                        'VIC-II colour values for bitmaps and character
Dim ImageType As Integer
Dim PFIO As Integer                                             'Picture file#, shared with multiple subs (needs re-writing)

Dim EncodeL(4)          As Integer                              'List Font Encoding
Dim LFontW(2)           As Integer                              'List Font Widths
Dim LFontH(2)           As Integer                              'List Font Heights

Dim xInit As String, xFile As String

Dim p_name(0 To NUMB)   As String
Dim p_sa(1 To NUMB)     As Long
Dim p_len(1 To NUMB)    As Long
Dim p_bitmap(1 To NUMB) As Long
Dim p_screen(1 To NUMB) As Long
Dim p_colour(1 To NUMB) As Long
Dim p_back(1 To NUMB)   As Long
Dim p_type(0 To NUMB)   As Integer


'==== BASIC Viewer

Dim Token(358) As String

'==== FONT Viewer
Dim FontOffset          As Integer                              'Font Offset
Dim ChrSetSize          As Integer                              'Number of characters in a 'set' (128 or 256)
Dim ChrSetNum           As Integer                              'Character Set Number
Dim ChrNum              As Integer                              'Character Number in Set
Dim SelChr              As Integer                              'Selected Character in File (all sets)
Dim SelChr2             As Integer                              'End of Range

Dim FontH               As Integer                              'Font Height (8 or 16)
Dim ChrZoom             As Integer                              'Zoom Factor
Dim SelZoom             As Integer
Dim ChrWIndex           As Integer                              'Number of characters per line
Dim ChrHIndex           As Integer
Dim ChrHeight           As Integer
Dim ChrLineMax          As Integer                              'Max Chrs per font line
Dim LastChrLineMax      As Integer                              'Max Chrs per font line
Dim ChrEditMode         As Boolean                              'Edit Mode Flag
Dim DesignerFlag        As Boolean                              'Screen Designer Mode Flag
Dim ChrPos              As Integer                              'Current Edit Chr Byte Offset Position
Dim ChrPosEnd           As Integer                              'Current Range End Pos
Dim ChrTop              As Integer                              'Pointer to current character
Dim CrosshairR          As Integer                              'Pixel Row Marker
Dim CrosshairC          As Integer                              'Pixel Col Marker
Dim PixelMode           As Integer                              'Pixel Draw Mode (set, reset, xor)
Dim BorderSize          As Integer                              'Border size

Dim BorderFlag          As Boolean                              'Display border between characters
Dim OutlineFlag         As Boolean                              'Outline each character (experimental)
Dim SelHiFlag           As Boolean                              'Flag to Highlight Current Character
Dim MCFlag              As Boolean                              'Multi-Colour Mode
Dim BitFlag             As Boolean                              'Update pixel bits?
Dim RangeFlag           As Boolean                              'True, when Range is valid
Dim RedrawFlag          As Boolean                              'Flag to force update of pixel bitmap

Dim CMat(15)            As String * 1                           'array for one character
Dim Tr(15)              As Integer                              'translation array

'==== Screen Designer (Part of Font Editor)

Dim SEBuf As String                                             'Screen Designer - Screen Memory Buffer
Dim SERow As Integer, SECol As Integer, BlinkFlag As Boolean    'Screen Designer - Cursor Position, Blink Flag
Dim SECursorPos As Integer                                      'Screen Designer - Offset of Cursor position into buffe
Dim SEMaxRow As Integer, SEMaxCol As Integer                    'Screen Designer - Screen Limits
Dim SEW As Integer, SEH As Integer                              'Screen Designer - Width and Height of each chr
Dim SERVSFlag As Boolean                                        'Screen Designer -
Dim SECBMFlag As Boolean                                        'Screen Designer - Typing Mode CBM or ASCII
Dim RECFlag As Boolean                                          'Screen Designer - RECORD Flag

Dim Macro(9) As String                                          'Macro Strings
Dim MacroNum As Integer                                         'Current Macro#

Dim SBit(7, 7) As Integer, DBit(7, 7) As Integer                'source/dest bit arrays for rotation

'==== ML Viewer
Dim OP(255) As String                                           '6502 Opcodes
Dim OpModeLen As String                                         'Opcode Addresing Mode Lengths (number of bytes for specified addressing mode)
Dim OpB As String, OpJ As String, OpZ As String                 'Tracer opcode groups: Branches, Jumps, Stops
Dim OpDesc As String                                            'Opcode Description from file
Dim LastFile As String, LastComment As String, LastSymPos As Integer

'==== Common
Public ProjFlag As Boolean, MLCFlag As Boolean, InfoFlag As Boolean, MLSplitFlag As Boolean
Public ChangeFlag As Boolean
Public MLTabNum As Integer
Public OpCodeFlag As Boolean, ShowTables As Boolean
Public DOTORG As String, DOTWORD As String, DOTBYTE As String, DOTTEXT As String, DOTHEX As String, DOTBIN As String

Dim TabColour(5, 7) 'cols,rows

Public LPrefix As String, ProjFilename As String

Option Explicit


'---- COMMON: Load the Form
Private Sub Form_Load()
    Dim i As Integer, Filename As String
    
    On Error Resume Next
    
    ViewerReady = False                                 'Make sure changing drop-down menus doesn't cause other code to run
    
    Me.Height = 8955                                    'Set the form height
    Me.Width = 12000                                    'Set the form width
   
    
    '--- Copy Button Icons
    
    For i = 1 To 2
        cmdEncode(i).Picture = cmdEncode(0).Picture     'Character Encoding Menu Icon
        cmdFSize(i).Picture = cmdFSize(0).Picture       'FontSize Menu Item
    Next i
    
    '--- Set Menu Defaults
    
    For MenuNum = 0 To 2
        EncodeL(MenuNum) = 0                            'Default Font Encoding for each List
        SetEncodeTip MenuNum                            'Set the Tooltip
        SetListFontWH 1                                 'Use first menu size as Default Font Width for each list
    Next
    
    SetChrSize                                          'Set Character Set
   
    '--- Set up the GUI
    
    SetVTheme                                            'Set the Theme Colours
    
    ViewMode = 0: ViewMode2 = -1                        'Default View Modes
    SplitMode = False: SplitSize = 50                   'Dual-view mode
        
    '--- Set Combo List Defaults
    
    cboMode.ListIndex = 0                               'MLView
    cboMLFmt.ListIndex = 0                              'MLView output format combo
    cboTarget.ListIndex = 0                             'MLView targe assembler combo
    cboPrefix.ListIndex = 0                             'MLView label prefix combo
    cboPlatform.ListIndex = 0                           'MLView platform combo
    cboCPU.ListIndex = 0                                'MLView CPU combo
        
    '--- Set Flags and Other Defaults
    
    ProjFlag = False                                    'ML Viewer
    MLCFlag = False                                     'ML Viewer
    InfoFlag = False
    MLSplitFlag = False
    ShowTables = False                                  'ML Viewer
    MLTabNum = 0                                        'ML Viewer
    SetTarget 0                                         'Target Assembler
    SetPrefix 0                                         'Label Prefix
            
    '--- Font Viewer/Editor
    
    ChrSetNum = 0                                       'Font Viewer/Editor - Character Set Number
    ChrNum = 0                                          'Font Viewer/Editor - Character Number is set
    SelChr = 0: SelChr2 = 0                             'Font Viewer/Editor - Selected character(s) in file (all sets)
    
    RangeFlag = False                                   'Font Viewer/Editor - Valid Range selected flag
    
    ChrZoom = 2                                         'Font Viewer/Editor - Character Set Zoom Index
    ChrWIndex = 5                                       'Font Viewer/Editor - Width Index
    
    ChrHeight = 8                                       'Font Viewer/Editor - Character Height (8 or 16)
    ChrHIndex = 0                                       'Font Viewer/Editor - Height selection index 0 or 1

    ChrPos = 1: ChrPosEnd = 8                           'Font Viewer/Editor - Start/end positions into buffer
        
    BitFlag = True                                      'Font Viewer - Flag to Update Pixel set
    RedrawFlag = True                                   'Font Viewer - Flag to Re-Draw entire Font View
        
    ChrEditMode = False                                 'Font Viewer/Editor - Edit Mode Flag
    DesignerFlag = False                                'Font Viewer/Editor - Screen Designer Mode Flag
    OutlineFlag = False                                 'Font Viewer/Editor - Outline each character
    BorderFlag = True                                   'Font Viewer/Editor - Borders on
    SECBMFlag = True                                    'Font Viewer/Editor - Type in CBM Mode
    BorderSize = 1                                      'Font Viewer/Editor - Border size
    CrosshairR = 0: CrosshairC = 0                      'Font Viewer/Editor - Character Edit Crosshairs

    SelZoom = 16                                        'Font Viewer/Editor - Selected Character Zoom Factor
    PixelMode = 2                                       'Font Viewer/Editor - Pixel Drawing Mode (0=BG,1=FG,2=XOR)
    
    VBufNum = -1                                        'Font Viewer/Editor - Current Buffer# -1=None
    'txtBorder.ListIndex = 0
    MacroNum = 1                                        'Screen Designer    - Current Macro#
    
    '---
    
    SetFontMenu
    
    '--- Theme
    
    Call SetColor                                       'Setup C64 colours
        
    Filename = ExeDir & "cbmxfer.vpl"                   'Check for default VPL file
    If Exists(Filename) = True Then
        LoadVPL Filename                                'Load the file
        cmdLoadVPL.ToolTipText = "Loaded cbmxfer.vpl"   'Set the Tooltip
    End If
    
    DoEvents                                            'Refresh everything
    
    ViewerReady = True                                  'Now we can allow the file to be viewed
    
End Sub

'---- COMMON: Process the Dropped File
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Data.GetFormat(vbCFFiles) Then
        Dim vFn As Variant
        
        For Each vFn In Data.Files
            ViewIt ViewMode, vFn, "", ""                            'vFn is name of file dropped
            Exit For                                                'only process the first dropped file!
        Next
    End If
End Sub

'-- COMMON: Unload the Form? - Check if ASM Project needs saving before Exiting
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If OverwriteProject = False Then Cancel = True                              'Set CANCEL=True to prevent form from Exiting

End Sub

'--- COMMON: Resize the Form
Private Sub Form_Resize()
    
    DoEvents
    DrawVLayout

End Sub

'---- COMMON: ViewIt
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

    If ViewerReady = False Then Exit Sub
    If SrcFile = "" Then Exit Sub
    If Exists(SrcFile) = False Then MyMsg "Viewer: File '" & SrcFile & "' not found!": Exit Sub
    
    ViewerReady = False
    
    VFileName = SrcFile: VName = SrcName: VExt = SrcExt                         'ViewFile Details
    VP00Flag = False                                                            'Assume normal file
    If FileExtU(VFileName) = "P00" Then VP00Flag = True                         'P00 file found!
    
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
        If VLen > 32760 Then                                                    'Lenght is too long!
            VLen = 32760                                                        'Set to Max
            ChrEditMode = False                                                 '
            shOverflow.Visible = True                                           'Show Overflow indicator
        End If
        
        VBuf = Input(VLen, FIO)                                                 'Read contents to buffer
        
    Close FIO
    
    VBuf1 = VBuf                                                                'Backup buffer for Font Editor
    txtCSkip.Text = "0"                                                         'Set File offset to zero
    
    lblVSize.Caption = Format(VLen)                                             'File Size
    cbLA.Enabled = True                                                         'Re-enable LA checkbox
    
    Tmp = "Viewer: " & FileNameOnly(VName)                                      'Titlebar string
    If VP00Flag = True Then Tmp = Tmp & " (Contained inside P00)"               'Add P00 note
    Me.Caption = Tmp                                                            'Set Window Titlebar
        
    ViewerReady = True                                                          'Allow interactive changes
    
    If cbLockView.value = vbUnchecked Then SelectNewTab Mode
    UpdateViews                                                                 'Update Tab Views
    
    For Lo = 0 To 2
        CalcScroll Lo                                                           'Recalulate Scrollbars
    Next
End Sub

'-- COMMON: SelectNewTab - Handle clicking view tab
Private Sub SelectNewTab(ByVal NewTabNum As Integer)
    
    If ViewerReady = False Then Exit Sub
    
    If (lblSelect.Caption = "<") Or (SplitMode = False) Then
        If NewTabNum <> ViewMode Then
            If NewTabNum = ViewMode2 Then ViewMode2 = ViewMode
            ViewMode = NewTabNum: LockV1 = NewTabNum
            RefreshContent ViewMode
        End If
    Else
        If NewTabNum <> ViewMode2 Then
            If NewTabNum = ViewMode Then ViewMode = ViewMode2
            ViewMode2 = NewTabNum: LockV2 = NewTabNum
            RefreshContent ViewMode2
        End If
    End If
    
    SetViewTabs
    DrawVLayout
    
End Sub

'---- COMMON: UpdateViews - Update contents of left and/or right
Private Sub UpdateViews()

    RefreshContent ViewMode                             'Update content in left view
    If SplitMode = True Then RefreshContent ViewMode2   'Update content in right view

End Sub

'---- COMMON: Refresh Content
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

    'ViewBusy = False
    
End Sub

'---- COMMON: Draw View Layout
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
    
    '-- Hide all the frames except Menu
    
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
    
    W = Me.Width - 120                                                          'Width of frame
    H = Me.Height - 800                                                         'Height of frame
    If W < 4400 Then W = 4400                                                   'Min frame size
    If H < 3700 Then H = 3700                                                   'Min frame size
    
    frMenu.Width = W
    
    L1 = 90                                                                     'Left edge of main frame
    T1 = 450                                                                    'Top  edge of main frame - Below Menu frame
    W1 = W - L1 - L1
    W2 = W
    H1 = H - frMenu.Height + 240                                                'Frame Height
    L2 = T1                                                                     'Set for single-view mode
    
    '-- Calculate Split mode sizes
    
    If SplitMode = True Then
        For i = 0 To 2: lblSSize(i).Visible = True: Next                        'Show the split re-sizers
        W1 = (W - L1 - L1) * (SplitSize / 100)                                  'Calc new width of LEFT frame
        W2 = (W - L1 - L1) * ((100 - SplitSize) / 100) - L1                     'Calc new width pf RIGHT frame
        L2 = L1 + W1 + L1                                                       'Calculate Left offset
    End If
    
    '-- Position the frames
    
    SetFrame ViewMode, L1, T1, W1, H1, True                                     'Position and Show Frame on LEFT
    SetFrame ViewMode2, L2, T1, W2, H1, SplitMode                               'Position and Show Frame on RIGHT (if SplitMode=TRUE)
    
    DoEvents
    
    If ViewMode < 3 Then CalcScroll ViewMode                                    'Re-calculate scrollbars
    
    DoEvents
End Sub

'---- COMMON: Update top line buttons
Private Sub SetViewTabs()
    Dim i As Integer

    For i = 0 To 5
        If (i = ViewMode) Or ((i = ViewMode2) And (SplitMode = True)) Then
            lblView(i).Font.Bold = True
            lblView(i).ForeColor = TabColour(i, 0)
            lblView(i).BackColor = TabColour(i, 1)
        Else
            lblView(i).Font.Bold = False
            lblView(i).ForeColor = TabColour(i, 2)
            lblView(i).BackColor = TabColour(i, 3)
        End If
    Next
    
    DoEvents
    
End Sub

'---- COMMON: SetFrame
' Arrange View Elements
' N=Frame#, Size: L=Left,T=Top,W=Width,H=Height, VisFlag=Frame Visible?
' In Dual-View Mode FLAG=TRUE
Sub SetFrame(ByVal N As Integer, ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal H As Single, ByVal VisFlag As Boolean)
    Dim L2 As Single, T2 As Single, W2 As Single, H2 As Single  'Second copy for modification
    Dim LL As Single, TT As Single, WW As Single, HH As Single
    Dim W3 As Single, H3 As Single
    Dim W4 As Single, H4 As Single
    
    L2 = L:   T2 = T: W2 = W: H2 = H                            'Copy of original size requested
    
    LL = 105: TT = 210                                          'Left/Top for Header or controls at top of frame
    HH = H - 440                                                'Adjust height for border area
    WW = W - 200                                                'Adjust width
    
    W3 = W - 200                                                'Width
    H3 = H - 600
    
    Select Case N
        Case -1                                                 '---- BLANK FRAME with message
        
            frBlank.Move L, T, W2, H2
            frBlank.Visible = VisFlag
            
        Case 0                                                  '---- BASIC
                    
            If lblBView.Caption = "<<" Then
                TT = 890: HH = H - 1000
                frBOpts.Visible = True                          'Show BASIC Options
            Else
                frBOpts.Visible = False                         'Hide Options
            End If
            
            frBasic.Move L, T, W2, H2                           'Set the Frame size
            picView(0).Move LL, TT, W3 - 300, HH                'Set BASIC pic area
            vsView(0).Move W3 - 200, TT, 300, HH                'Set Scrollbar size
            frBasic.Visible = VisFlag                           'Set frame visiblity
            RefreshVIEW 0                                       'Re-draw the listing
    
        Case 1                                                  '---- SEQ
        
            TT = 450: HH = H - 570                              'SAdjust for Options
            frSEQ.Move L, T, W2, H2                             'Set Frame Size
            vsView(1).Move W3 - 200, TT, 300, HH                'Set Scrollbar
            picView(1).Move LL, TT, W3 - 300, HH                'Set SEQ pic area
            frSEQ.Visible = VisFlag                             'Set frame visiblity
            RefreshVIEW 1                                       'Re-draw the listing
            
        Case 2                                                  '---- BIN
        
            TT = 810: HH = H - 900                              'Adjust for Options
            frBIN.Move L, T, W2, H2                             'Set Frame Size
            vsView(2).Move W3 - 200, TT, 300, HH                'Set Scrollbar Size
            picView(2).Move LL, TT, W3 - 300, HH                'Set SEQ pic area
            frBIN.Visible = VisFlag                             'Set frame visiblity
            RefreshVIEW 2                                       'Re-draw the listing
            
        Case 3                                                  '---- FONT
        
            frFont.Move L, T, W2, H2                            'Set frame size
            frFont.Visible = VisFlag                            'Set frame visiblity
            
            If ChrWIndex > 4 Then                               'If Auto-scale or Fit are selected
                SetChrLineMax                                   'Calculate new width (old width is saved)
                If ChrLineMax <> LastChrLineMax Then            'Try to avoid constant font re-draw
                    RedrawFlag = True                           'Force redraw all
                    DrawChrSet                                  'Draw the character set only
                End If
            End If
            
        Case 4                                                  '---- ASM
        
            frML.Move L, T, W2, H2                              'Move and size the MAIN frame
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
                LL = L + 3960: WW = W - 4130                    'Reduce Width when Tables are visible
                lblShw.Caption = "<<"                           'Set the show/hide label
            End If
            
            frInfo.Move LL, T + 145, WW                         'Position Info frame
            lblInfo.Width = WW - 240                            'Size the Info text area
            
            '---- Show the ML window(s)
            
            If MLSplitFlag = True Then
                HH = (HH - 300) / 2
                lstML.Move LL, TT, WW, HH                       'Position The output list
                lstML2.Move LL, TT + HH + 150, WW, HH           'Position The second output list
                lstML2.Visible = True                           'Show second ML View
            Else
                lstML2.Visible = False                          'Hide second ML View
                lstML.Move LL, TT, WW, HH                       'Position The output list
            End If
            
            frBlock.Move LL, TT, WW, HH                         'Make Block Comment frame equal to ML list
            txtBlock.Width = WW - 300
            txtBlock.Height = HH - 900
            
            If ShowTables = True Then
                LL = 15: TT = 1320: WW = 3870
                HH = H3 - TT - 60: W4 = WW - 120
                
                frTView.Move 120, 490, WW, H3
                frMLSettings.Move LL, 700, W4, HH + 480         'The Settings Frame
                frTrace.Move LL, 700, W4, HH + 480              'The Tracer frame
                
                lstEntryPt.Move LL, TT, W4, HH                  'The Entry Points list
                lstSYM.Move LL, TT, W4, HH                      'The Symbols list
                lstDT.Move LL, TT, W4, HH                       'The Data Tables list
                lstULabels.Move LL, TT, W4, HH                  'The Generated Labels list
                lstCmnt.Move LL, TT, W4, HH                     'The Comment list
                lstLabels.Move LL, TT, W4, HH                   'The Labels list
                lstJSR.Move LL, TT, W4, HH                      'The External JSR list
                                
                lstEP.Height = HH                               'The Tracer Entry Point List

                DrawMLTabs                                      'Draw Tabs
            End If
    
            frTView.Visible = ShowTables                        'Show or Hide

        Case 5                                                  '---- PIC

            frBMP.Move L, T, W2, H2                             'Set frame position and size
            frBMP.Visible = VisFlag                             'Set frame visiblity
            
    End Select
    
    DoEvents
    
End Sub

'---- COMMON: Adjust Dual-View Split Sizing
Private Sub lblSSize_Click(Index As Integer)

    SetSplit Index, False   'Normal step size

End Sub

'---- COMMON: Click to Set Split
Private Sub lblSSize_DblClick(Index As Integer)
    
    SetSplit Index, True                                                        'Doubles the step size when user clicks too fast and generates Double-click
    
End Sub

'---- COMMON: Adjust Dual-View Split proportions
Private Sub SetSplit(ByVal Index As Integer, ByVal Flag As Boolean)
    Dim N As Integer
    
    N = 5: If Flag = True Then N = 10 'Step Size
    
    Select Case Index
        Case 0: SplitSize = SplitSize - N: If SplitSize < 20 Then SplitSize = 20    'Move split LEFT
        Case 1: SplitSize = SplitSize + N: If SplitSize > 80 Then SplitSize = 80    'Move split RIGHT
        Case 2: SplitSize = 50                                                      'Return to MIDDLE
    End Select
    
    DrawVLayout
    
End Sub

'---- COMMON: View Tab was clicked
Private Sub lblView_Click(Index As Integer)
    
    If ViewerReady = True Then SelectNewTab Index
    
End Sub

'---- COMMON: Lock the current view
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

'---- COMMON: Toggle Single or Dual-View Mode
Private Sub lblSplit_Click()

    If lblSplit.Caption = "+" Then
        EnableSplit
    Else
        lblSplit.Caption = "+"
        lblSelect.Visible = False
        SplitMode = False
    End If
    
    DrawVLayout

End Sub

'---- COMMON: Enable Split View
Private Sub EnableSplit()

        lblSplit.Caption = "-"                                                          'Change + to -
        lblSelect.Visible = True                                                        'Show the Select button
        SplitMode = True                                                                'Enable Splir View
        
End Sub


'========================================
'BASIC Viewer
'========================================

'---- BAS: View a File - Main Routine
Private Sub BASView()
    Dim pLo As Integer, pHi As Integer                                                  'Program Line Links
    Dim lLo As Integer, lHi As Integer, LNum As Long                                    'Line numbers
    Dim i As Integer                                                                    'Bufer position counter
    Dim C As Integer, C2 As Integer                                                     'Character value
    Dim Mode As Integer, MaxWidth As Integer
    
    Dim Ch As String                                                                    'Character string
    Dim Tmp As String
    Dim BGuess As String, UnK As String, Pad As String, TLine As String
    
    Dim First   As Boolean, Quote   As Boolean
    Dim RevText As Boolean, OneLine As Boolean
    Dim EncFlag As Boolean, ExpFlag As Boolean, UCFlag As Boolean                       'Font Encoding
    
    Me.Show: DoEvents
    
    If Token(0) = "" Then LoadTokens                                                    'Load Tokens if first run
    
    EncFlag = False: If EncodeL(0) < 2 Then EncFlag = True                              'Set Encode Flag. TRUE = use CBM Font character (do not translate)
    
    UnK = "{unknown}"
    RevText = (cbRev.value = 1)                                                         'Reverse text case
    ExpFlag = (cbExp.value = 1)                                                         'Expand special characters
    UCFlag = (cbUC.value = 1)                                                           'Uppercase special characters
    OneLine = (cbOneLine.value = 1)                                                     'One statement per line mode
    
    Pad = "": If cbPad.value = 1 Then Pad = " "                                         'Padding of tokens
        
    Mode = cboMode.ListIndex                                                            'Basic Mode dropdown
    lblGuess.Caption = ""
        
    If VLen < 2 Then Exit Sub                                                           'Exit if Invalid BASIC program
    
    lblLoadAdr.Caption = MyHex(VLA, -4)                                                 'Show Load Address
    
    If Mode = 0 Then
        Select Case VLA                                                                 '-- Guess the target machine based on Load Address
            Case 3:             Mode = 2: BGuess = "CBM2"
            Case 1024, 1025:    Mode = 1: BGuess = "PET"
            Case 2049:          Mode = 1: BGuess = "C64"
            Case 4097, 4609:    Mode = 1: BGuess = "Vic20"
            Case 4096, 8192:    Mode = 3: BGuess = "C16/Plus4"
            Case 7169:          Mode = 4: BGuess = "C128 Basic 7"
            Case 12289:         Mode = 4: BGuess = "LCD Basic 3.6"
            Case Else:          Mode = 1: BGuess = "Unknown"
        End Select
        
        lblGuess.Caption = BGuess                                                       'Show the Guess
    End If
    
    C = cboColWidth2.ListIndex                                                          'Column Width Index
    Select Case C
        Case 1: MaxWidth = 80                                                           '80 COL - PET,CBM-II, C128 VDC Screen
        Case 2: MaxWidth = 40                                                           '40 COL - PET,C64,TED,P500
        Case 3: MaxWidth = 22                                                           '22 COL - VIC-20
        Case Else: MaxWidth = 255                                                       'Maximum possible BASIC line length
    End Select
    
    lstView(0).Clear                                                                    'Clear the Listing
    i = 1                                                                               'Pointer to BASIC DATA

    Do
        If i > VLen Then Exit Do                                                        'Exit if end of buffer
        
        pLo = Asc(Mid(VBuf, i, 1)):     If i + 1 > VLen Then Exit Do                    'Exit if past end of buffer
        pHi = Asc(Mid(VBuf, i + 1, 1)): If (pHi + pLo) = 0 Then Exit Do                 'program link=0 means END OF PROGRAM

        If (i + 3) > VLen Then Exit Do                                                  'Exit if past end of buffer
        
        '-- Get the Line Number
        lLo = Asc(Mid(VBuf, i + 2, 1))                                                  'Get LO byte of Line#
        lHi = Asc(Mid(VBuf, i + 3, 1))                                                  'Get HI byte of Line#
        LNum = lHi * 256! + lLo                                                         'Make Line number
        TLine = Format(LNum) & " "                                                      'Start a new line with the Line#
        i = i + 4                                                                       'Increment pointer

        Quote = False                                                                   'Turn off Quote Mode

        '-- Start Parsing the Line
        
        Do
            If (i > VLen) Then Exit Do                                                  'End of file
            
            C = Asc(Mid(VBuf, i, 1))                                                    'PETSCII Chacter Code
            Tmp = ""                                                                    'String to Add
            i = i + 1                                                                   'Index to Character
                        
            If Len(TLine) >= MaxWidth Then                                              'Check Maximum Line Width to display (0=no max)
                lstView(0).AddItem TLine: TLine = ""                                   'Break the line here
            End If
            
            If (C = 0) Then                                                             'NULL
                lstView(0).AddItem TLine                                                'NULL = End of line. Add it to the listbox ********************
                Exit Do                                                                 '
            End If
            
            
            '-- Do Inside Quotes or Non-Token characters
            
            If (Quote = True) Or (C < 128) Then                                         '==== Handle Non-Tokens or Characters inside Quotes
                Select Case C
                    Case 1 To 31                                                        'Special keys (cursor etc)
                        If ExpFlag = True Then
                            Tmp = Token(297 + C - 1)                                    'Get the TOKEN string
                            If UCFlag Then Tmp = UCase(Tmp)                             'Convert to Uppercase
                            
                            If EncodeL(0) < 6 Then
                                Tmp = "[" & Tmp & "]"                                   'Add brackets ie: [blue] - since CBM does not have "{}"
                            Else
                                Tmp = "{" & Tmp & "}"                                   'Add braces,  ie: {blue}
                            End If
                        Else
                            If EncFlag = True Then Tmp = Chr(C)                         'Keep Original Code character
                        End If
                        
                    Case 32, 160                                                        '-- Handle SPACE
                        Tmp = " "                                                       'SPACE
                    
                    Case 34                                                             '-- Handle QUOTE
                        Tmp = Qu
                        Quote = Not Quote                                               'Toggle Quote Mode
                    
                    Case 33 To 64                                                       '-- Handle Normal Character and ":" separator
                        Tmp = Chr(C)
                        If Tmp = ":" Then                                               'Statement Divider?
                            If (OneLine = True) And (Quote = False) Then
                                Tmp = "": lstView(0).AddItem TLine                      'Add the line **********************
                                TLine = Space$(Len(Format(LNum)) + 1)                   'Start next line with SPACE indenting
                            End If
                        End If
                        
                    Case 65 To 90                                                       '-- Handle A-Z
                        If RevText Then C = Reverse(C)                                  'Reverse the Case
                        Tmp = Chr(C)
                    
                    Case 97 To 122                                                      '-- Handle a-z
                        If RevText Then C = Reverse(C)                                  'Reverse the Case
                        Tmp = Chr(C)
                        
                    Case 129 To 159                                                     '-- Hsndle Special keys (Colours, Cursor etc)
                        If ExpFlag = True Then
                            Tmp = Token(328 + C - 129)                                  'Get the TOKEN string
                            If UCFlag Then Tmp = UCase(Tmp)
                            If EncodeL(0) < 6 Then
                                Tmp = "[" & Tmp & "]"                                   'Add brackets ie: [blue] - since CBM does not have "{}"
                            Else
                                Tmp = "{" & Tmp & "}"                                   'Add braces,  ie: {blue}
                            End If
                        Else
                            If EncFlag = True Then Tmp = Chr(C)                         'If PETSCII mode then just include it
                        End If
                        
                    Case 193 To 218                                                     '-- Handle a to z
                        C = C - 96: If RevText Then C = Reverse(C)
                        Tmp = Chr(C)
                        
                    Case Else                                                           '-- Handle Everything else
                        If EncFlag = False Then                                         'Everything else is an unknown graphic symbol
                            Tmp = "{" & Hex(C) & "}"                                    'Show Hex code for Graphic character, eg: {FF}
                        Else
                            Tmp = Chr(C)                                                'If PETSCII mode then just include it
                        End If
                        
                End Select
                
                TLine = TLine & Tmp
                
            Else
            
                '-----------------Convert to Tokens
                Select Case Mode
                    Case 1                                                              '-- BASIC 1/2
                        Select Case C
                            Case 128 To 203, 255: Tmp = Token(C - 128)                  'Common Tokens
                            Case 254                                                    'Expansion C64 Tokens
                                C2 = Asc(Mid(VBuf, i, 1)): i = i + 1                    'Get second Token byte
                                If cbMV.value = vbChecked Then
                                    Select Case C2                                      'Handle MagicVoice Tokens
                                        Case 128: Tmp = "SAY": lblGuess.Caption = "C64 MagicVoice"
                                        Case 129: Tmp = "RATE"
                                        Case 130: Tmp = "VOC"
                                        Case 132: Tmp = "RDY"
                                        Case Else: Tmp = "[" & Format(C2) & "]"
                                    End Select
                                Else
                                    If (C2 > 127) And (C2 < 159) Then Tmp = Token(266 + C2 - 128): lblGuess.Caption = "C64 Exp"
                                End If
                        End Select

                    Case 2                                                              '-- BASIC 4/4+
                        Select Case C
                            Case 128 To 203, 255: Tmp = Token(C - 128)                  'Common Tokens
                            Case 204 To 232: Tmp = Token(128 + C - 204)                 'Basic4/4+ Tokens
                        End Select
                        
                    Case 3                                                              '-- BASIC 3.5
                            
                        Select Case C
                            Case 254                                                    '-- Expansion TED Tokens
                                C2 = Asc(Mid(VBuf, i, 1)): i = i + 1                    'Get second Token byte
                                If cbMV.value = vbChecked Then
                                    Select Case C2                                      'Handle MagicVoice Tokens
                                        Case 1: Tmp = "RATE"
                                        Case 2: Tmp = "VOC"
                                        Case 4: Tmp = "RDY"
                                        Case 10: Tmp = "SAY": lblGuess.Caption = "V364 Speech"
                                        Case Else: Tmp = "[" & Format(C2) & "]"
                                    End Select
                                Else
                                    Tmp = "[" & Format(C2) & "]": lblGuess.Caption = "TED Exp"
                                End If
                            Case Else: Tmp = Token(C - 128) 'Common Tokens/Basic3.5
                        End Select


                    Case 4                                                              '-- BASIC 7
                        Select Case C
                            Case 128 To 205, 207 To 253, 255: Tmp = Token(C - 128)      'Common Tokens/Basic3.5
                            Case 206                                                    'CE Tokens; CE02 to CE0A
                                C2 = Asc(Mid(VBuf, i, 1)): i = i + 1
                                If C2 > 1 And C2 < 11 Then Tmp = Token(194 + C2 - 2)
                            Case 254                                                    'FE Tokens; FE02 to FE26
                                C2 = Asc(Mid(VBuf, i, 1)): i = i + 1
                                If C2 > 1 And C2 < 39 Then Tmp = Token(157 + C2 - 2)
                        End Select

                    
                    Case 5                                                              '-- BASIC 10
                       Select Case C
                            Case 128 To 205, 207 To 253, 255: Tmp = Token(C - 128)      'Common Tokens/Basic3.5
                            Case 206                                                    'CE Tokens; CE02 to CE0A
                                C2 = Asc(Mid(VBuf, i, 1)): i = i + 1
                                If C2 > 1 And C2 < 11 Then Tmp = Token(194 + C2 - 2)
                            Case 254                                                    'FE Tokens; FE02 to FE3D
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
        lblNote.Caption = "There are " & Format(VLen - i - 1) & " bytes following BASIC end!"
    Else
        lblNote.Caption = ""
    End If
  
    RefreshVIEW 0
    CalcScroll 0
    
End Sub

'---- BAS: Toggle Options pane
Private Sub lblBView_Click()
    
    If lblBView.Caption = ">>" Then
        lblBView.Caption = "<<"
    Else
        lblBView.Caption = ">>"
    End If
    
    DrawVLayout

End Sub

'---- BAS: Save Listing to File
Private Sub cmdSave_Click(Index As Integer)
    Dim FIO As Integer, Filename As String, J As Integer
    
    Filename = FileOpenSave(FileBase(LastFile), 1, 5, "Save Listing as Text")
    If Filename = "" Then Exit Sub
    
    FIO = FreeFile
    Open Filename For Output As FIO
    For J = 0 To lstView(Index).ListCount - 1
        Print #FIO, lstView(Index).List(J)
    Next
    Close FIO
    ChDir ExeDir

NoFile:

End Sub

'---- BAS: Copy BASIC listing to clipboard
Private Sub cmdCpyClip_Click()
    Dim J As Integer, Tmp As String
    
    For J = 0 To lstView(0).ListCount - 1
        Tmp = Tmp & lstView(0).List(J) & vbCrLf
    Next J
    
    Clipboard.Clear
    Clipboard.SetText Tmp

End Sub

'--- BAS: Load Token strings into array
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
    
    If cboTheme.ListIndex = -1 Then
        cboTheme.ListIndex = 0                                          'Set First Theme as default
        cboScnFmt.ListIndex = 1                                         'Set 40 col as default
    End If
    
    If BitFlag = True Then CreatePixels                                 'Create pixel multicolour or normal font pixels

    If MCFlag = True Then
        lblTheme(3).Visible = True: lblTheme(4).Visible = True          'Show Multi-colour colour boxes
    Else
        lblTheme(3).Visible = False: lblTheme(4).Visible = False        'Hide MC boxes
    End If

    If RangeFlag = True Then
        picChr.Visible = False: lblRange.Visible = True                 'Show the Range info
    Else
        picChr.Visible = True: lblRange.Visible = False                 'Hide the Range info
    End If
    
    If ChrHIndex = 0 Then
        picChr.Height = 1945                                            'Height for 8x8 characters
    Else
        picChr.Height = 3880                                            'Height for 8x16 characters
    End If
    lblOver.Top = lblRange.Top + lblRange.Height + 30                   'Move Over-character display
    lblFStat.Top = lblOver.Top + lblOver.Height + 30                    'Move status box
    lblFStat.Width = picChr.Width
    
    If VBufNum = -1 Then                                                'First-time Initialization
        SetChrLineMax
        UpdateSelectors                                                 'Update all Selectors
        VBufNum = 0
        UpdateBuf                                                       'Update Edit Buffer
        SetEditBuf 0                                                    'Default to First Buffer
    End If
    
    SetEditLayout
    DrawChrSet
    
    DoEvents
    
End Sub
 
'---- FONT: Draws the Complete Character Set
' Uses offset, BorderFlag, OutlineFlag, Zoom, and selected colours
Public Sub DrawChrSet()
    Dim J As Long, V As Integer, K As Integer
    Dim X As Integer, Y As Integer                                      'Co-ordinates
    Dim TopX As Integer, TopY As Integer                                'Top-left for character set
    Dim R As Integer, C As Integer                                      'Row and Col
    Dim RW As Integer, CW As Integer
    Dim MaxR As Integer, MaxH As Integer                                'Maximums
    Dim CZ As Integer, RZ As Integer, PZ As Integer                     'Zoomed sizes
    Dim ChrNum As Integer, OutFlag As Boolean
    Dim C1 As Long, C2 As Long, Thick As Integer, Thick2 As Integer     'Outline Colour and Thickness
    Dim CCZ As Integer, RRZ As Integer, YYZ As Integer                  'To help speed up drawing
    Dim FH As Integer
    
    'If VBuf = "" Then MsgBox "VBuf Empty!?": Exit Sub                   'Hide font and exit
    If VBuf = "" Then picV.Visible = False: Exit Sub                    'Hide font and exit
    
    FH = ChrHeight                                                      'Chr Height in pixels
    ChrNum = 0                                                          'Start at Zero
    C = 0: R = 0: X = 0: Y = 0
    TopX = 0: TopY = 0                                                  'Top-Left Offset
    MaxR = 64                                                           'Max Row (Max Col is global)
    
    picV.Cls                                                            'Clear to background
    
    CW = 8: RW = FH                                                     'Chr width, Row width
    PZ = CW * ChrZoom                                                   'Scale factor for drawing one line of pixels
    Thick = ChrZoom \ 2                                                 'Outline thickness
    Thick2 = Thick * 2                                                  'Speedup
    
    FontOffset = Val(txtCSkip.Text): If FontOffset < 0 Then FontOffset = 0
    If FontOffset > 32767 Then FontOffset = 32767                       'Enforce Maximum Offset
    
    If BorderFlag = True Then
        CW = CW + BorderSize                                            'Adjust for Border
        RW = RW + BorderSize
        TopX = BorderSize + ChrZoom: TopY = TopX                        'Add space at top
        If (OutlineFlag = True) And (BorderSize > 2) And (ChrZoom > 1) Then
            OutFlag = True                                              'If room for Outline then set Flag to allow it
            picV.DrawWidth = Thick                                      'Set thickness for drawing outlines
        End If
    End If
    
    C1 = vbWhite                                                        'Outline Colour
    C2 = lblTheme(6).BackColor                                          'Selected Character Colour
    
    CZ = CW * ChrZoom                                                   'Size of one character including borders
    RZ = RW * ChrZoom                                                   'Size of one character including borders
    FontH = FH                                                          'Set for calculating chr when clicked
            
    If RedrawFlag = True Then                                           '--- Redraw Flag=true forces update of entire view
        picV.Width = (CZ * ChrLineMax + TopY) * Screen.TwipsPerPixelX   'Set Width
        picV.Height = (RZ * MaxR + TopX) * Screen.TwipsPerPixelY        'Set Height
        picV.BackColor = lblTheme(5).BackColor                          'Set FG and BG colours
        picV.Cls                                                        'Clear the bitmap
        picV.Visible = False                                            'Hide the bitmap so drawing is faster
        DoEvents                                                        'Force update
    End If
    
    CCZ = TopX: RRZ = TopY                                              'Set drawing position to Top-Left
    YYZ = Y * ChrZoom                                                   'Speedup
    
    VLen = Len(VBuf)                                                    'Re-calculate font size buffer in case it has been modified by CUT or INSERT
    
    For J = FontOffset + 1 To VLen
        V = Asc(Mid(VBuf, J, 1))                                        'Get the pixels
        YYZ = Y * ChrZoom
        '----paintpicture {srcimg},destX,destY,destW,destH ,srcX,srcY,srcW,srcH,mode
        If (RangeFlag = True) Then
            If (ChrNum >= SelChr) And (ChrNum <= SelChr2) Then
                picV.PaintPicture Pix.Image, CCZ, RRZ + YYZ, PZ, ChrZoom, 0, V, 8, 1, vbNotSrcCopy              'blit the pixels - Selected character
            Else
                picV.PaintPicture Pix.Image, CCZ, RRZ + YYZ, PZ, ChrZoom, 0, V, 8, 1                            'blit the pixels - Un-selected character
            End If
        Else
            picV.PaintPicture Pix.Image, CCZ, RRZ + YYZ, PZ, ChrZoom, 0, V, 8, 1                                'blit the pixels - Un-selected character
        End If
        
        '-- Handle Outlines around each character
        If OutFlag = True Then
            picV.Line (CCZ - Thick, RRZ - Thick)-Step(8 * ChrZoom + Thick2, FontH * ChrZoom + Thick2), C1, B      'Draw the outline
        End If
        
        '-- Handle Outline around Selected Character
        If SelHiFlag = True Then
            If ChrNum = SelChr Then picV.Line (CCZ - Thick, RRZ - Thick)-Step(8 * ChrZoom + Thick2, FontH * ChrZoom + Thick2), C2, B  'Draw the outline
        End If

        '-- Move to next position
        Y = Y + 1                                                       'Next scanline
        
        If Y = FH Then                                                  '-- Reached character height
            Y = Y - FH: ChrNum = ChrNum + 1                             'Move back to top of line position
            C = C + 1: If C >= ChrLineMax Then C = 0: R = R + 1         'Next Character. Check end of line
            CCZ = TopX + C * CZ: RRZ = TopY + R * RZ                    'Pre-calc to speed up draw
        End If
                
        'If R > MaxR Then Exit For                                       'Exit if at bottom of visible area
    Next J
    
    If C > 0 Then R = R + 1
    
    If R < MaxR Then
        If R = 0 Then R = 1                                             'Fix if single row
        picV.Height = (RZ * R + TopX) * Screen.TwipsPerPixelY
    End If
        
    lblEndRange.Caption = "to" & Str(J)                                 'Show Range
    
    picV.Visible = True                                                 'Show the set
    DoEvents
    
    ShowSelChr                                                          'Display Selected/Edit character
    
    RedrawFlag = False
    
End Sub

'========================================
'Font View Subs
'========================================

'---- FONT: Set Dropdown Menu checks
' Sets the CHECKED property of ALL menu items
Private Sub SetFontMenu()
    
    frmMenu.mnuFont(1).Checked = ChrEditMode                            'Font Editor mode
    frmMenu.mnuFont(2).Checked = DesignerFlag                           'Screen Designer Mode
    frmMenu.mnuFont(3).Checked = MCFlag                                 'Multicolour
    frmMenu.mnuFont(4).Checked = BorderFlag                             'Border
    frmMenu.mnuFont(5).Checked = OutlineFlag                            'Outline
    frmMenu.mnuFont(6).Checked = SelHiFlag                              'Screen Designer Mode
    
    '--- Enable/Disable options depending on Screen Designer Active
    
    frmMenu.mnuFont(4).Enabled = Not DesignerFlag                      'Border
    frmMenu.mnuFont(5).Enabled = Not DesignerFlag                      'Outline
    frmMenu.mnuFont(6).Enabled = Not DesignerFlag                      'Screen Designer Mode
        
    
End Sub

'---- FONT: Update Selectors: Buffer, Height, Zoom, Line Width, Pixel Mode
Private Sub UpdateSelectors()

    SetBufSelector
    SetChrHeightSelector
    SetChrZoomSelector
    SetChrWidthSelector
    SetPixelModeSelector

End Sub

'---- FONT: Set Character Height Selector
Private Sub SetChrHeightSelector()
    Dim i As Integer
    
    For i = 0 To 1
        If i = ChrHIndex Then
            lblChrHeight(i).Font.Bold = True
            lblChrHeight(i).ForeColor = TabColour(0, 4)
            lblChrHeight(i).BackColor = TabColour(0, 5)
        Else
            lblChrHeight(i).Font.Bold = False
            lblChrHeight(i).ForeColor = TabColour(0, 6)
            lblChrHeight(i).BackColor = TabColour(0, 7)
        End If
    Next i
    
End Sub

'---- FONT: Set Character Zoom Selector
Private Sub SetChrZoomSelector()
    Dim i As Integer
    
    For i = 0 To 5
        If i = (ChrZoom - 1) Then
            lblZoom(i).Font.Bold = True
            lblZoom(i).ForeColor = TabColour(4, 4)
            lblZoom(i).BackColor = TabColour(4, 5)  'Selected
        Else
            lblZoom(i).Font.Bold = False
            lblZoom(i).ForeColor = TabColour(4, 6)
            lblZoom(i).BackColor = TabColour(4, 7)  'Un-selected
        End If
    Next i
    
End Sub

'---- FONT: Set Character Width Selector
Private Sub SetChrWidthSelector()
    Dim i As Integer
    
    For i = 0 To 6
        If i = ChrWIndex Then
            lblWidth(i).Font.Bold = True
            lblWidth(i).ForeColor = TabColour(5, 4)
            lblWidth(i).BackColor = TabColour(5, 5)                                     'Selected
        Else
            lblWidth(i).Font.Bold = False
            lblWidth(i).ForeColor = TabColour(5, 6)
            lblWidth(i).BackColor = TabColour(5, 7)                                     'Un-selected
        End If
    Next i
    
    
End Sub

'---- FONT: Set Pixel Mode Selector
Private Sub SetPixelModeSelector()
    Dim i As Integer
    
    For i = 0 To 2
        If i = PixelMode Then
            lblPixelMode(i).Font.Bold = True
            lblPixelMode(i).ForeColor = TabColour(4, 4)
            lblPixelMode(i).BackColor = TabColour(4, 5)
        Else
            lblPixelMode(i).Font.Bold = False
            lblPixelMode(i).ForeColor = TabColour(4, 6)
            lblPixelMode(i).BackColor = TabColour(4, 7)

            lblPixelMode(i).Font.Bold = False: lblPixelMode(i).ForeColor = vbBlack
        End If
    Next i
    
End Sub

'---- FONT: Change Theme
Private Sub lblTheme_Click(Index As Integer)
    
    frmColourPicker.Show vbModal                                                'Set form to show Modal
    
    If PickedColour >= 0 Then
        lblTheme(Index).BackColor = PickedColour                                'Set the new colour
        BitFlag = True                                                          'The pixel bitmap must be re-rendered in the new colour
        RedrawFlag = True                                                       'Flag to Re-draw everything
        FONTView
    End If
    SetScreenTheme
    
End Sub

'---- FONT: Adjust Border Size
Private Sub lblBorder_Click(Index As Integer)
    
    Select Case Index
        Case 0: BorderSize = BorderSize - 1: If BorderSize < 1 Then BorderSize = 1          'Decrease Sie
        Case 1: BorderSize = BorderSize + 1: If BorderSize > 16 Then BorderSize = 16        'Increase Size
    End Select
    lblBorderSize.Caption = Format(BorderSize)                                              'Set border size label
    
    UpdateChrSetView                                                                        'Update Character Set View
    
End Sub

'---- FONT: Click to select a Buffer
Private Sub lblBufSel_Click(Index As Integer)
    
    UpdateBuf                                                                   'Updates Last
    SetEditBuf Index

End Sub

'---- FONT: Update Buffer
' Copies the Visible buffer to the selected buffer
Private Sub UpdateBuf()
        
    Select Case VBufNum                                                         'Use buffer#
        Case 0: VBuf1 = VBuf                                                    'Make Buffer#1 Visible
        Case 1: VBuf2 = VBuf                                                    'Make Buffer#2 Visible
        Case 2: VClip = VBuf                                                    'Make Clipboard Visible
        Case 3: VRestore = VBuf                                                 'Make Restore Point Visible
    End Select

End Sub

'---- FONT: Set Visible Buffer
' Makes the selected buffer the current buffer to edit/view
Private Sub SetEditBuf(ByVal Index As Integer)

    ' Check if selected buffer is empty. If so then exit
    Select Case Index
        Case 0: If VBuf1 = "" Then Exit Sub
        Case 1: If VBuf2 = "" Then Exit Sub
        Case 2: If VClip = "" Then Exit Sub
        Case 3: If VRestore = "" Then Exit Sub
    End Select

    VBufNum = Index                                                             'Remember the Buffer number
    
    Select Case Index
        Case 0: VBuf = VBuf1                                                    'Make Buffer#1 Visible
        Case 1: VBuf = VBuf2                                                    'Make Buffer#2 Visible
        Case 2: VBuf = VClip                                                    'Make Clipboard Visible
        Case 3: VBuf = VRestore                                                 'Make Restore Point Visible
    End Select
    
    UpdateChrSetView
    SetBufSelector
    
End Sub
Private Sub SetBufSelector()
    Dim i As Integer
    
    For i = 0 To 3
        If i = VBufNum Then
            lblBufSel(i).Font.Bold = True                                       'Selected
            lblBufSel(i).ForeColor = TabColour(0, 4)
            lblBufSel(i).BackColor = TabColour(0, 5)
        Else
            lblBufSel(i).Font.Bold = False                                      'Unselected
            lblBufSel(i).ForeColor = TabColour(0, 6)
            lblBufSel(i).BackColor = TabColour(0, 7)
            
        End If
    Next i
End Sub

'---- FONT: Click to Set Pixel Drawing Mode
Private Sub lblPixelMode_Click(Index As Integer)
    
    PixelMode = Index
    SetPixelModeSelector

End Sub

'---- FONT: Popup the Font Editor Menu
Private Sub cmdFontMenu_Click()

    MenuForm = 2                                                        'The Viewer Form
    PopupMenu frmMenu.mnuF                                              'Display the Menu
    
End Sub

'---- FONT: Dispatch Font Menu Selection
Public Sub DoFMenu(ByVal Index As Integer)

    Select Case Index
        Case 1: ToggleEdit                                              'Toggle Edit mode
        Case 2: ToggleDesigner
        Case 3: ToggleMC                                                'Toggle Multi-colour
        Case 4: ToggleBorder                                            'Toggle Border
        Case 5: ToggleOutline
        Case 6: ToggleSelHi                                             'Toggle Selected Character Higlighting
        Case 7: SaveBMP                                                 'Save current view as Bitmap
        Case 8: SaveFont 0                                              'Save entire font
        Case 9: SaveFont 1                                              'Save Range
        
        Case 100 To 107                                                '--- Convert Font Menu
            
            ConvertFont Index - 100                                     'Convert the font
            SelChr = 0: SelChr2 = 0                                     'Reset Selected
            SetRange
            FONTView                                                    'Re-Draw character set
            
        Case 201: SECLS                                                 'Clear Screen
        Case 202: SEReset                                               'Reset Machine
        Case 203: SELoad                                                'Load Buffer
        Case 204: SESave                                                'Save Buffer
        Case 205: SEExport                                              'Export Buffer
        Case 206: SELoadMacro                                           'Load Macro
        Case 207: SESaveMacro                                           'Save Macro
        Case 208: SETogREC                                              'Toggle Record
        Case 209: SESaveBMP                                             'Save as Bitmap
        
        Case 300 To 306                                                 '--- Encoding Menu
            
            EncodeL(MenuNum) = Index - 301                              'Set Encoding for this list
            SetEncodeTip MenuNum                                        'Set the Tooltip
            
            Select Case MenuNum                                         'Re-read file since encoding will change how the characters are decoded
                Case 0: BASView
                Case 1: SEQView
                Case 2: HEXView
            End Select
        
        Case 500 To 599                                                 '--- Font Size Menu
            
            SetListFontWH Index                                         'Set the List Font Width and Height
            
    End Select
    
End Sub

'---- FONT: Save the Font Bitmap
Private Sub SaveBMP()
    Dim Filename As String
    
    Filename = FileOpenSave(FileBase(VFileName), 1, 3, "Save as BMP")
    picV.Picture = picV.Image 'crop to visible
    
    If Filename <> "" Then SavePicture picV.Image, Filename

End Sub

'-- FONT: Change Zoom Factor
Private Sub lblZoom_Click(Index As Integer)

    ChrZoom = Index + 1                                                         'Set Zoom factor
    SetChrZoomSelector
    UpdateChrSetView

End Sub

'---- FONT: Change Width
Private Sub lblWidth_Click(Index As Integer)

    ChrWIndex = Index                                                           'Set Font line Width
    SetChrWidthSelector
    UpdateChrSetView
    
End Sub

'---- FONT: Update Character Set View
' Updates ChrLineMax and redraws only the new character set
Private Sub UpdateChrSetView()

    SetChrLineMax                                                               'Calculate new Character Max width
    RedrawFlag = True                                                           'Force redraw all
    DrawChrSet                                                                  'Draw the character set only
    SetEditLayout                                                               'Set Edit Layout
    
End Sub

'---- Set Character Max Width (8-128)
'Index: 0 to 4 = Fixed 8/16/32/64/128
'       5      = Auto (contrained)
'       6      = Fit to width
Public Sub SetChrLineMax()
    Dim W As Integer, W2 As Integer, CW As Integer
    
    LastChrLineMax = ChrLineMax                                                 'Remember Previous for form resizing
    
    W = frFont.Width                                                            'Width of Frame
    W2 = frChr.Left + frChr.Width                                               'Position of Chr Edit frame + width
    CW = (W - W2 - 300) / (ChrZoom * 134)                                       'Space inbetween
    
    Select Case ChrWIndex                                                       'Width Index
        Case 0: ChrLineMax = 8
        Case 1: ChrLineMax = 16
        Case 2: ChrLineMax = 32
        Case 3: ChrLineMax = 64
        Case 4: ChrLineMax = 128
        Case 5                                                                  '--- Auto Constrained Width
            Select Case CW
                Case 0 To 15: ChrLineMax = 8
                Case 16 To 31: ChrLineMax = 16
                Case 32 To 63: ChrLineMax = 32
                Case 64 To 127: ChrLineMax = 64
                Case Else: ChrLineMax = 128                                     'Anything bigger is limited to 128
            End Select

        Case 6
            ChrLineMax = CW                                                     '--- Auto Max Width
            
    End Select
    
End Sub

'---- FONT: Change Character Height (8 or 16 pixel tall)
Private Sub lblChrHeight_Click(Index As Integer)
    Dim Tmp As String
    
    If Index = ChrHIndex Then Exit Sub                                          'Ignore click if already in same mode
    
    ChrHIndex = Index
    
    If Index = 0 Then
        ChrHeight = 8: Tmp = "16 to 8"
    Else
        ChrHeight = 16: Tmp = "8 to 16"
    End If
    
    If ChrEditMode = True Then
        If MsgBox("Do you want to convert this font from " & Tmp & " pixel format?", vbYesNo, "Convert Font") = vbYes Then
            ConvertFont Index                                   'Convert font from 8 to 16 (index=1) or 16 to 8 (index=0)
        End If
    End If
    
    SetChrHeightSelector
    RedrawFlag = True                                                   'Force redraw all
    SetRange                                                            'Display Selection info
    FONTView                                                            'Re-Draw character set
    
End Sub

'---- FONT: Click to Set Character Size
Private Sub cbSetSize_Click()
    
    SelChr = 0: SelChr2 = 0: RangeFlag = False                          'Reset selected character and range
    
    SetChrSize                                                          'Change the Character Set size
    CalcChrDisplay                                                      'Calculate selection
    ShowSelChr                                                          'Display the Selected Character

End Sub

'---- FONT: Set Character "Set" Size
' CBM Fonts have either 128 characters without RVS (PET), or 256 with RVS (most others)
Private Sub SetChrSize()
    
    ChrSetSize = 128                                                    'Assume 128 characters in set (no RVS)
    If cbSetSize.value = vbChecked Then ChrSetSize = 256                'If checked then 256 (includes RVS)
    
    ShowSelChrInfo                                                      'Show the new character info
    
End Sub

'---- FONT: Toggle Multicolour mode
' This requires a change to the Pixel Bitmap
Private Sub ToggleMC()
    
    MCFlag = Not MCFlag                                                 'Set Multicolour Flag
    BitFlag = True                                                      'Re-draw Pixel Bitmap
    RedrawFlag = True                                                   'Force Redraw all
    
    SetFontMenu                                                         'Update checkmarks
    FONTView                                                            'Draw character set

End Sub

'---- FONT: Toggle Border
Private Sub ToggleBorder()
    
    BorderFlag = Not BorderFlag
    SetFontMenu                                                         'Update checkmarks
    UpdateChrSetView

End Sub

'-- FONT: Toggle Outline
Private Sub ToggleOutline()
    
    OutlineFlag = Not OutlineFlag
    SetFontMenu                                                         'Update checkmarks
    UpdateChrSetView

End Sub

'-- FONT: Toggle Selected Character Higlighting
Private Sub ToggleSelHi()

    SelHiFlag = Not SelHiFlag                                           'Toggle the Mode
    SetFontMenu                                                         'Update checkmarks
    UpdateChrSetView

End Sub

'-- FONT: Set Viewer Colour Theme
Private Sub cboTheme_Click()
    Dim N As Integer
    
    N = cboTheme.ListIndex: If N < 0 Then N = 0                                 'Get choice and validate
    SetFontTheme N                                                              'Set the theme
    
End Sub

'---- Set the Font Theme
' Sets the theme to match specific CBM machine colour themes
Private Sub SetFontTheme(ByVal N As Integer)
    Dim FG As Long, BG As Long, BO As Long, DI As Long

    DI = vbBlack                                                                        'Default Divider Colour = Black
    If N > 4 Then DI = RGB(128, 128, 128)                                               'Medium Grey
    
    Select Case N
        Case 0: FG = CBMColor(14): BG = CBMColor(6): BO = CBMColor(14)                  'C64
        Case 1, 3, 9: FG = CBMColor(6): BG = CBMColor(1): BO = CBMColor(3)              'SX-64 / VIC-20 / P500
        Case 2: FG = CBMColor(13): BG = CBMColor(11): BO = CBMColor(13)                 'C128
        Case 4: FG = CBMColor(0): BG = CBMColor(1): BO = CBMColor(4)                    'TED
        Case 5: FG = CBMColor(1): BG = CBMColor(0): BO = CBMColor(0)                    'PET White
        Case 6: FG = CBMColor(5): BG = CBMColor(0): BO = CBMColor(0)                    'PET Green
        Case 7: FG = CBMColor(7): BG = CBMColor(0): BO = CBMColor(0)                    'PET Amber
        Case 8: FG = CBMColor(5): BG = CBMColor(0): BO = CBMColor(0)                    'CBM-II 256
        Case Else: FG = vbWhite: BG = vbBlack: BO = vbGreen
    End Select
    
    lblTheme(0).BackColor = FG: lblTheme(1).BackColor = BG              'Foreground and Background Colours
    lblTheme(2).BackColor = BO: lblTheme(5).BackColor = DI              'Border and Divider Colours
    
    DoEvents
    
    RedrawFlag = True
    BitFlag = True
        
    FONTView
    SetScreenTheme
    
End Sub

'---- FONT: Click ENTER to Change Font Start Offset
Private Sub txtCSkip_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then FONTView

End Sub

'---- FONT: Create Pixels
' Creates a bitmap containing the pixel representation for all values from 0 to 255, using the current colour palette
' If multicolour mode is disabled it uses the foreground (0) and backgound colour (1)
' If Multicolour mode is enabled  it uses 4 pairs of pixels to determine colour (0 to 3)
Public Sub CreatePixels()
    Dim J As Integer, K As Integer, CI As Integer
    Dim MC(3) As Long                                                   'Array to hold multicolour values
    
    MC(0) = lblTheme(1).BackColor                                       'Background colour
    MC(1) = lblTheme(3).BackColor                                       'Register colour #1
    MC(2) = lblTheme(4).BackColor                                       'Register colour #2
    MC(3) = lblTheme(0).BackColor                                       'Foreground Colour
    
    Pix.ForeColor = lblTheme(0).BackColor
    Pix.BackColor = lblTheme(1).BackColor
    Pix.Cls
        
    If MCFlag = True Then
        '-- Create a 4-colour bitmap with pixels to match binary representation of pixel pairs (row=value,cols 0 to 7=pixel)
        For J = 0 To 255
            For K = 0 To 7 Step 2
                CI = 0                                                  'Colour Index
                If (J And Pow(K)) Then CI = CI + 2                      'Check first bit
                If (J And Pow(K + 1)) Then CI = CI + 1                  'Check second bit
                Pix.ForeColor = MC(CI)                                  'Set the colour of the pixel to draw
                Pix.PSet (7 - K, J)                                     'Set the first pixel
                Pix.PSet (6 - K, J)                                     'Set the second pixel
            Next K
        Next J
    Else
        '-- Create a 2-colour bitmap with pixels to match binary representation of value (row=value,cols 0 to 7=pixel)
        For J = 0 To 255
            For K = 0 To 7
                If (J And Pow(K)) Then Pix.PSet (7 - K, J)
            Next K
        Next J
    End If
    
    BitFlag = False                                                     'Bitmaps are created

End Sub

'---- FONT: Handle Character Selection Buttons
' Set.... Either 128 or 256 depending on whether RVS characters are included (user must select)
' Chr.... The character number in the set
' Num.... The character number in the file (all characters)
Private Sub cmdChrSel_Click(Index As Integer)
    Dim MaxSet As Integer, V As Integer
    
    V = Len(VBuf)
    MaxSet = Int(V / ChrHeight / ChrSetSize) - 1
    
    Select Case Index
        Case 0: ChrSetNum = ChrSetNum - 1: If ChrSetNum < 0 Then ChrSetNum = 0              'Set -
        Case 1: ChrSetNum = ChrSetNum + 1: If ChrSetNum > MaxSet Then ChrSetNum = MaxSet + 1  'Set +
        Case 2: ChrNum = ChrNum - 1: If ChrNum < 255 Then ChrNum = 0                        'Chr -
        Case 3: ChrNum = ChrNum + 1: If ChrNum > 255 Then ChrNum = 255                      'Chr +
        Case 4:
            SelChr = SelChr - 1: If SelChr < 0 Then SelChr = 0                              'Num -
            CalcChrDisplay
        Case 5:
            SelChr = SelChr + 1                                                             'Num +
            CalcChrDisplay
    End Select
    
    SelChr = ChrSetNum * ChrSetSize + ChrNum                                                'Calculate new Selected Character
    SelChr2 = SelChr
    
    SetRange
    ShowSelChr
    
    If DesignerFlag = True Then DrawEditScreen 2                                            'Redraw the Edit screen
End Sub

'---- FONT: Display the Character Info
Public Sub ShowSelChrInfo()
    Dim Tmp As String
    
    lblChrSet.Caption = Format(ChrSetNum + 1, "000"): lblChrSet.ToolTipText = MyHex(ChrSetNum + 1, -2)
    lblChrNum.Caption = Format(ChrNum, "000"): lblChrNum.ToolTipText = MyHex(ChrNum, -2)
    lblChrSel.Caption = Format(SelChr, "000"): lblChrSel.ToolTipText = MyHex(SelChr, -4)
    
    lblFStat.Caption = "Crosshairs: Row=" & Format(CrosshairR) & ", Col=" & Format(CrosshairC) _
        & Cr & Cr & "Chr: RIGHT-CLICK to set crosshairs." & Cr & Cr & "Chr Set: CLICK on first chr, RIGHT-CLICK on last to set RANGE."
    
    Tmp = "Range:" & Cr & Cr & "From: " & Format(SelChr) & Cr & "To..: " & Format(SelChr2) & Cr & Cr & "(" & Format(SelChr2 - SelChr + 1) & " selected)"
    If Len(VClip) > 0 Then Tmp = Tmp & Cr & Cr & Format(Len(VClip)) & " bytes in clipboard"
    lblRange.Caption = Tmp
    
    DoEvents
    
End Sub

'---- FONT: Show the Selected Character
' Shows a single character enlarged for display or editing
' Updates character Information boxes
Public Sub ShowSelChr()
    Dim R As Integer, C As Integer, X As Integer, Y As Integer, i As Integer
    Dim XYOff As Integer
    Dim RW As Integer, CW As Integer
    Dim C1 As Long, C2 As Long, C3 As Long
    Dim Tmp As String, OutFlag As Boolean
    
    
    If ChrLineMax < 8 Then Exit Sub
    If SelChr < 0 Then SelChr = 0
    
    ShowSelChrInfo                                                                  'Display Character Info
    
    OutFlag = False                                                                 'Outline Flag
    
    RW = FontH: CW = 8: XYOff = 0                                                   'Pixels in one char
    
    If BorderFlag = True Then
        RW = RW + BorderSize: CW = CW + BorderSize
        XYOff = BorderSize + ChrZoom                                                'Adjust for border
    End If
    
    If CrosshairR >= ChrHeight Then CrosshairR = ChrHeight - 1                      'When switching from 16 to 8 pixel tall
    
    '-- Calc position
    R = Int(SelChr / ChrLineMax)                                                    'Calculate Row
    C = SelChr - R * ChrLineMax                                                     'Calculare Col
    X = C * CW * ChrZoom + XYOff: Y = R * RW * ChrZoom + XYOff                      'Calculate X/Y coordinates
    
    C1 = lblTheme(5).BackColor                                                      'Use theme colour 5 for border
    C2 = vbWhite                                                                    'Use White for crosshairs
    
    If picV.Height >= FontH * ChrZoom * 15 Then
        '-- Draw the Character using set on screen (not buffer)
        picChr.PaintPicture picV.Image, 0, 0, SelZoom * 8, SelZoom * FontH, X, Y, 8 * ChrZoom, FontH * ChrZoom
        
        '-- Draw Divider lines between pixels
        If BorderFlag = True Then
            For i = 0 To 16
                picChr.Line (0, i * SelZoom)-Step(160, 2), C1, BF                   'Draw Horizontal Lines
            Next i
            
            For i = 0 To 8
                picChr.Line (i * SelZoom, 0)-Step(2, 320), C1, BF                   'Draw Vertical Lines
            Next i
        End If
        
        '-- Draw Crosshairs
        If ChrEditMode = True Then
            picChr.Line (0, CrosshairR * SelZoom + 1)-Step(160, 0), C2               'Draw Horizontal Crosshair
            picChr.Line (CrosshairC * SelZoom + 1, 0)-Step(0, 320), C2               'Draw Vertical Crosshair
        End If
        
    End If
    
End Sub

'---- FONT: Handle clicking on Character Set image
' Select a character or Range for Editing
' LEFT BUTTON=Select character (Range Start), RIGHT BUTTON=Select Range End
Private Sub picV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim T As Integer, R As Integer, C As Integer, RW As Integer, CW As Integer
    Dim UpdateFlag As Boolean

    UpdateFlag = False
    
    RW = FontH: CW = 8
    If BorderFlag = True Then RW = RW + BorderSize: CW = CW + BorderSize
    
    R = Int(Y / (RW * ChrZoom))                                                 'Calculate Row
    C = Int(X / (CW * ChrZoom)): If C > ChrLineMax Then C = ChrLineMax          'Calculate Col
    T = R * ChrLineMax + C                                                      'Calculate Character
    
    If ((T + 1) * FontH) > VLen Then Exit Sub                                   'If past end of file then abort
    
    If (Shift > 0) Or (Button = 2) Then
       SelChr2 = T                                                              'Set Range End
       RangeFlag = True: UpdateFlag = True
    Else
       SelChr = T                                                               'Set the selected character
       SelChr2 = T                                                              'Set Range End to be the same
       If RangeFlag = True Then RangeFlag = False: UpdateFlag = True
       CalcChrDisplay                                                           'Calculate Character values display
    End If
        
    If RangeFlag = True Then
        If SelChr > SelChr2 Then
            T = SelChr: SelChr = SelChr2: SelChr2 = T                           'Swap Range Start and End points
        End If
    End If
        
    SetRange                                                                    'Set selection start and end points in string
    
    If (UpdateFlag = True) Then
        FONTView                                                                'Re-display the Font
    Else
        If SelHiFlag = True Then UpdateChrSetView
    End If
    
    ShowSelChr                                                                  'Display the Selected Character
    
End Sub

Private Sub picV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim R As Integer, C As Integer, T As Integer, SetN As Integer, ChrN As Integer
    Dim RW As Integer, CW As Integer
    
    If X < 2 Then lblOver.Caption = "": Exit Sub 'Clear when on edge
    
    RW = FontH: CW = 8
    If BorderFlag = True Then RW = RW + BorderSize: CW = CW + BorderSize
    
    R = Int(Y / (RW * ChrZoom))                                                 'Calculate Row
    C = Int(X / (CW * ChrZoom)): If C > ChrLineMax Then C = ChrLineMax          'Calculate Col
    T = R * ChrLineMax + C                                                      'Calculate Character
    SetN = (T \ ChrSetSize) + 1
    ChrN = T Mod ChrSetSize
    lblOver.Caption = "Set:" & Str(SetN) & ", Chr:" & Str(ChrN)
    
End Sub

'---- FONT: Clear "Over Character" display when mouse moves away from character set
Private Sub frFont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblOver.Caption = ""
End Sub

'---- FONT: Calculate Character Display Numbers
' Calculate Set# and Chr# based on "SelChr" and font size
Private Sub CalcChrDisplay()
    
    ChrSetNum = Int(SelChr / ChrSetSize)
    ChrNum = SelChr Mod ChrSetSize

End Sub

'---- FONT: Set Range
' Set Range Start and End byte positions based on character positions
Private Sub SetRange()

    ChrPos = SelChr * ChrHeight + 1                                  'Set the Start byte
    ChrPosEnd = SelChr2 * ChrHeight + ChrHeight                      'Set the End byte
    If (ChrPos = 0) Or (ChrPosEnd > VLen) Then ChrPos = 0: ChrPosEnd = 0  'If out of range then abort
    
    Debug.Print "SetRange: ChrPos=" & Str(ChrPos) & " ChrPosEnd=" & Str(ChrPosEnd)
    
End Sub

'---- FONT: Clear the Range
Private Sub ClearRange():

    SelChr = 0: SelChr2 = 0                                             'Reset Start and End Character Range
    RangeFlag = False                                                   'Clear the Range Flag
    
    'SetRange                                                            'Calc Range bytes
    
End Sub

'---- FONT: Change Skip-bytes
Private Sub cmdSB_Click(Index As Integer)
    Dim FontOffset As Integer
    
    FontOffset = Val(txtCSkip.Text)
    Select Case Index
        Case 0: FontOffset = FontOffset - 256
        Case 1: FontOffset = FontOffset - ChrHeight
        Case 2: FontOffset = FontOffset - 1
        Case 3: FontOffset = FontOffset + 1
        Case 4: FontOffset = FontOffset + ChrHeight
        Case 5: FontOffset = FontOffset + 256
    End Select
    
    If FontOffset < 0 Then FontOffset = 0
    
    txtCSkip.Text = Format(FontOffset)
    
    DrawChrSet
    
End Sub

'---- FONT: Toggle Edit Mode
Private Sub ToggleEdit()
    Dim N As Integer, X As Integer, Y As Integer
    
    If ChrEditMode = False Then
        If shOverflow.Visible = True Then MyMsg "Sorry, font is too big to edit!": Exit Sub
    End If
    
    ChrEditMode = Not ChrEditMode                                   'Toggle the Edit Mode
    frmMenu.mnuFont(1).Checked = ChrEditMode                                    'Set the Checkmark
   
'-- Check if font contains some multiple of 128 characters (128 x 8=1024 bytes).
'   If not ask to padd.

    If ChrEditMode = True Then
        N = Len(VBuf)
        If (N Mod 1024) <> 0 Then
            X = (Int(N / 1024) + 1) * 1024                                      'Calculate next biggest font size
            Y = X - N
            If MsgBox("This Font size is not a multiple of 128 characters." & Cr & "Do you want to fix this?" & Cr & "Bytes=" & Str(N) & " Expected=" & Str(X) & " Bytes to add=" & Str(Y), vbYesNo, "Pad Buffer") = vbYes Then
                VBuf = VBuf & String(Y, Nu)                                     'Pad it
                VLen = VLen + Y                                                 'Adjust size

            End If
        End If
    End If
    
'--
    SetEditLayout                                                               'Set Edit Mode Layout
    If ChrWIndex > 4 Then UpdateChrSetView

End Sub

'---- FONT: Set Layout according to Edit Mode
' Moves Edit Box and Character positions and sets Tool visibility
Private Sub SetEditLayout()

    If ChrEditMode = False Then
        frChr.Left = 30                                             'Tools not visible
        frTools.Visible = False                                     'Hide Tools
        lblFStat.Visible = False                                    'Hide Info Area
    Else
        frChr.Left = frTools.Left + frTools.Width                   'Tools visible
        frTools.Visible = True                                      'Show Tools
        lblFStat.Visible = True                                     'Show Info Area
    End If
    
    picV.Left = frChr.Left + frChr.Width                            '2390
    frEditor.Left = picV.Left + picV.Width + 90                   'Set Editor Position
    
    ShowSelChr
    
End Sub

'---- FONT: Save Font to File
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
        Print #FFIO, VBuf;                                                      'Write entire font
    Else
        Print #FFIO, Mid(VBuf, FontOffset + ChrPos, ChrPosEnd - ChrPos + 1);    'Write RANGE only
    End If
    Close FFIO
    
End Sub

'---- FONT: Handle clicking on the Selected Character Box
' This will exit if not in Edit mode
' Edit the character by clicking on a pixel - Uses currend Pixel Draw Mode
' Set Cross-hairs if SHIFT or RIGHT-BUTTON
Private Sub picChr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim R As Integer, C As Integer, bv As Integer, nv As Integer, TV As Integer, P As Integer
    Dim PP As Integer
    
    If ChrEditMode = False Then Exit Sub
    
    '-- convert x/y to row/col
    R = Y \ SelZoom: If R > ChrHeight - 1 Then R = ChrHeight - 1
    C = X \ SelZoom: If C > 7 Then C = 7
    
    '-- Set Cross-hair Markers (Shift-click or Right-click)
    If (Shift > 0) Or (Button = 2) Then
        CrosshairR = R: CrosshairC = C                                    'Set new crosshair position
        ShowSelChr                                                         'Display the character
        Exit Sub                                                        'Exit
    End If
    
    '-- Edit Pixel
    PP = FontOffset + ChrPos + R                                        'Position of byte to update
    bv = Asc(Mid(VBuf, PP, 1))                                          'Get byte for row
    P = Pow(7 - C)                                                      'Get pixel bit value
    
    Select Case PixelMode                                               'PixelMode determines how pixels are drawn
        Case 0: nv = bv And (255 - P)                                   'Set to Background
        Case 1: nv = bv Or P                                            'Set to Foreground
        Case 2: nv = bv Xor P                                           'XOR
    End Select
    
    Mid(VBuf, PP, 1) = Chr(nv)                                          'Update the byte
    
    DrawChrSet                                                          'Draw Character Set
    
End Sub

'---- FONT: Handle clicking on an Arrow Icon
Private Sub cmdShift_Click(Index As Integer)
    FontOp Index
End Sub

'---- FONT: Handle clicking on a Tool Button
Private Sub cmdTool_Click(Index As Integer)
    FontOp Index
End Sub

'---- FONT: Perform Font Operation
' This performs the selected operation on a single character or a range
' Handles Character Shifting or Rotating, Clearing, Reversing, Underlining, Bolding, Mirroring, Expanding, Doubling and Rotating
' Handles Character Insering and Deleting Rows or Columns
' Handles Clipboard Cut, Copy, Paste, Append
' Handles Set Selecting, Swapping, Copying, Restoring

Private Sub FontOp(ByVal Index As Integer)
    Dim a As Integer, B As Integer, C As Integer, K As Integer
    Dim cc As Integer, cStart As Integer
    Dim Row As Integer, Col As Integer
    Dim V As Integer, nv As Integer, nv2 As Integer, nv3 As Integer
    Dim Tmp As String, Tmp2 As String
    Dim Flag As Boolean, Bit As Integer
    Dim Max As Integer, J As Integer                                    'Max for compare
    
    Flag = False: If cbShiftMode.value = vbChecked Then Flag = True     'Flag set when Shift Mode is on
    
    VLen = Len(VBuf):  lblVSize.Caption = Format(VLen)                  'Set buffer length in case of changes
    
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
        Case 24: GoSub RestoreRange                                     'Restore Character(s)
        Case 25: GoSub RestoreAll                                       'Restore Selected Range from RestorPoint
        Case 26: GoSub InsRow                                           'Insert blank Row below crosshair
        Case 27: GoSub DelRow                                           'Delete Row below crosshair
        Case 28: GoSub InsCol                                           'Insert blank Col to right of crosshair
        Case 29: GoSub DelCol                                           'Delete column to right of crosshair
        Case 30: GoSub SetRestorePoint                                  'Set a restore point
        Case 31: GoSub CutClip                                          'Cut range from set
        Case 32: GoSub InsClip                                          'Insert clipboard to start position
        Case 33: GoSub AppendClip                                       'Append clipboard to end of font
        Case 34: GoSub SelectSet                                        'Select current set
        Case 35: GoSub NewFont                                          'Create a new font
        Case 36: GoSub Compare                                          'Compare sets 1 and 2 with results in clipboard
    End Select
    
    If Len(VBuf) <> VLen Then
        VLen = Len(VBuf): lblVSize.Caption = Format(VLen)               'Set buffer length in case of changes
        UpdateChrSetView                                                'View has changed. Display the Character Set
    Else
        DrawChrSet                                                      'View not changed, so just do a simple re-draw
    End If
    
    
    
    Exit Sub
    
'-------------------- Font Operations

ShiftUp:
    For J = ChrPos To ChrPosEnd Step ChrHeight
        DoEvents
        Tmp = Mid(VBuf, J, 1)
        For K = 1 To ChrHeight - 1
            Mid(VBuf, J + K - 1, 1) = Mid(VBuf, J + K, 1)               'Copy to byte above
        Next K
        If Flag = True Then
            Mid(VBuf, J + ChrHeight - 1, 1) = Tmp                       'Wrap to bottom line
        Else
            Mid(VBuf, J + ChrHeight - 1, 1) = Nu                        'Clear bottom line
        End If
    Next J
    Return

ShiftDown:
    For J = ChrPos To ChrPosEnd Step ChrHeight
        DoEvents
        Tmp = Mid(VBuf, J + ChrHeight - 1, 1)
        For K = ChrHeight - 2 To 0 Step -1
            Mid(VBuf, J + K + 1, 1) = Mid(VBuf, J + K, 1)               'Copy to byte above
        Next K
        If Flag = True Then
            Mid(VBuf, J, 1) = Tmp                                       'Wrap to top line
        Else
            Mid(VBuf, J, 1) = Nu                                        'Clear top line
        End If
    Next J
    Return

ShiftLeft:
    For J = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, J, 1))                                        'Read a byte
        Bit = 0: If (V And 128) > 0 Then Bit = 1
        nv = (V * 2) Mod 256                                            'Shift the pixels
        If Flag = True Then nv = nv + Bit
        Mid(VBuf, J, 1) = Chr(nv)                                       'Write it
    Next J
    Return

ShiftRight:
    For J = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, J, 1))                                        'Read a byte
        Bit = 0: If (V And 1) > 0 Then Bit = 128
        nv = V \ 2                                                      'Shift the pixels
        If Flag = True Then nv = nv + Bit
        Mid(VBuf, J, 1) = Chr(nv)                                       'Write it
    Next J
    Return
  
Clear:
    For J = ChrPos To ChrPosEnd
        Mid(VBuf, J, 1) = Nu                                            'Set to null
    Next J
    Return
    
RVS:
    For J = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, J, 1))                                        'Get byte value
        Mid(VBuf, J, 1) = Chr(255 - V)                                  'RVS it and write it
    Next J
    Return
    
BoldFont:
    For J = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, J, 1))
        nv2 = Int(V / 2)                                                'Shift the pixels
        nv = V Or nv2                                                   'Merge them
        Mid(VBuf, J, 1) = Chr(nv)                                       'Write it
    Next J
    Return
    
Underlined:
    For J = ChrPos To ChrPosEnd Step ChrHeight
        Mid(VBuf, J + CrosshairR, 1) = Chr(255)                          'Set byte to all 1's
    Next J
    Return
    
RotateRight:
    If ChrHeight = 16 Then MyMsg "Rotation only supported on 8x8 characters!": Return
    C = 0

    For J = ChrPos To ChrPosEnd Step ChrHeight
        DoEvents
        ChrTop = J                                                      'Set current character position
        GoSub ClearBitArrays                                            'Clear arrays for next character (all bits to zero)
        GoSub ReadChr                                                   'Get bytes and fill Source Bit array
        '---- Do Rotation 90
        For Row = 0 To 7
            For Col = 0 To 7
                DBit(7 - Col, Row) = SBit(Row, Col)
            Next Col
        Next Row
        GoSub WriteChr                                                  'Write the Dest Bit Array back as bytes
    Next J
    Return

RotateLeft:
    If ChrHeight = 16 Then MyMsg "Rotation only supported on 8x8 characters!": Return
    
    For J = ChrPos To ChrPosEnd Step ChrHeight
        DoEvents
        ChrTop = J                                                      'Set current character position
        GoSub ClearBitArrays                                            'Clear arrays for next character (all bits to zero)
        GoSub ReadChr                                                   'Read 8 bytes and fill Source Bit array
        '---- Do Rotation 270
        For Row = 0 To 7
            For Col = 0 To 7
                DBit(Col, 7 - Row) = SBit(Row, Col)
            Next Col
        Next Row
        GoSub WriteChr                                                  'Write the Dest Bit Array back as bytes
    Next J
    Return

MirrorH:
    For J = ChrPos To ChrPosEnd Step ChrHeight
        DoEvents
        For K = 0 To ChrHeight - 1
            CMat(K) = Mid(VBuf, J + K, 1)                               'Read to array in order
        Next K
       
        For K = 0 To ChrHeight - 1
            Mid(VBuf, J + K, 1) = CMat(ChrHeight - K - 1)               'Write to output in reverse order
        Next K
    Next J
    Return

MirrorV:
    GoSub SetupMirrorArray

    For J = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, J, 1))                                        'Read to array in order
        a = Int(V / 16): B = V Mod 16                                   'Calculate HI and LO nibbles
        nv = Tr(B) * 16 + Tr(a)                                         'Reverse the bits
        Mid(VBuf, J, 1) = Chr(nv)                                       'Write to output
    Next J
    Return

DoubleTall:
    For J = ChrPos To ChrPosEnd Step ChrHeight
        DoEvents
        For K = 0 To ChrHeight - 1
            CMat(K) = Mid(VBuf, J + K, 1)
        Next K
        C = cStart
        For K = 1 To ChrHeight - 1 Step 2
            Mid(VBuf, J + K - 1, 1) = CMat(C)
            Mid(VBuf, J + K, 1) = CMat(C)
            C = C + 1
        Next K
    Next J
    Return

DoubleWide:
    GoSub Setup2XArray
    For J = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, J, 1))                                        'Read byte, convert to ascii
        a = Int(V / 16): B = V Mod 16                                   'Calculate HI/LO nibbles
        If cc = 0 Then
            nv = Tr(a)                                                  'Translate HI
        Else
            nv = Tr(B)                                                  'Translate LO
        End If
        Mid(VBuf, J, 1) = Chr(nv)                                       'Write it
    Next J
    Return

DoubleSize:
    GoSub Setup2XArray
    For J = ChrPos To ChrPosEnd Step ChrHeight
        For K = 0 To ChrHeight - 1
            CMat(K) = Mid(VBuf, J + K, 1)                               'Read byte, convert to ascii
        Next K
        C = cStart
        For K = 1 To ChrHeight Step 2
            V = Asc(CMat(C))                                            'Get row byte
            a = Int(V / 16): B = V Mod 16                               'Calculate HI/LO nibbles
            If cc = 0 Then
                nv = Tr(a)                                              'Translate HI
            Else
                nv = Tr(B)                                              'Translate LO
            End If
            Mid(VBuf, J + K - 1, 1) = Chr(nv)                           'Write
            Mid(VBuf, J + K, 1) = Chr(nv)                               'Write
            C = C + 1
        Next K
    Next J
    Return

InsRow:
    For J = ChrPos To ChrPosEnd Step ChrHeight
        For K = ChrHeight - 2 To CrosshairR Step -1
            Mid(VBuf, J + K + 1, 1) = Mid(VBuf, J + K, 1)
        Next K
        Mid(VBuf, J + CrosshairR, 1) = Nu
    Next J
    Return
    
DelRow:
    For J = ChrPos To ChrPosEnd Step ChrHeight
        For K = CrosshairR To ChrHeight - 2
            Mid(VBuf, J + K, 1) = Mid(VBuf, J + K + 1, 1)
        Next K
        Mid(VBuf, J + ChrHeight - 1, 1) = Nu
    Next J
    Return
    
InsCol:
    '-- calculate pixel masks
    a = 0: For J = 7 To (8 - CrosshairC) Step -1: a = a + Pow(J): Next J     'LEFT side mask
    B = 0: For J = (7 - CrosshairC) To 0 Step -1: B = B + Pow(J): Next J     'RIGHT side mask
    
    '-- insert
    For J = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, J, 1))                                            'Get byte value
        nv2 = V And a                                                       'mask left side
        nv3 = (V And B) \ 2                                                 'mask right side and shift
        Mid(VBuf, J, 1) = Chr(nv2 + nv3)                                    'recombine and write
    Next J
    Return

DelCol:
    '-- calculate pixel masks
    a = 0: For J = 7 To (8 - CrosshairC) Step -1: a = a + Pow(J): Next J     'LEFT side mask
    B = 0: For J = (6 - CrosshairC) To 0 Step -1: B = B + Pow(J): Next J     'RIGHT side mask
    
    '-- delete
    For J = ChrPos To ChrPosEnd
        V = Asc(Mid(VBuf, J, 1))                                            'Get byte value
        nv2 = V And a                                                       'Mask left side
        nv3 = (V And B) * 2                                                 'Mask right side and shift
        Mid(VBuf, J, 1) = Chr(nv2 + nv3)                                    'Recombine and write
    Next J
    Return

    
'---------------------------- Clipboard subs
EmptyClip:
    MsgBox "Clipboard is empty!": Return                                        'Ignore if nothing to paste

CopyClip:
    If RangeFlag = True Then VClip = Mid(VBuf, ChrPos, ChrPosEnd - ChrPos + 1)
    Return

PasteClip:
    V = Len(VClip): If V = 0 Then GoTo EmptyClip
    For J = ChrPos To ChrPosEnd Step V                                          'Single/Multiple copy as needed
        Mid(VBuf, J, V) = VClip                                                 'Paste it once
    Next J
    Return
    
CutClip:
    If RangeFlag = False Then Return
    GoSub CopyClip                                                              'Copy it
    Tmp = "": If ChrPos > 1 Then Tmp = Left(VBuf, ChrPos - 1)                   'Data before cut range
    Tmp2 = "": If ChrPosEnd < Len(VBuf) Then Tmp2 = Mid(VBuf, ChrPosEnd + 1)    'Data after  cut range
    VBuf = Tmp & Tmp2                                                           'Make new buffer without cut range
    ClearRange                                                                  'Clear the Range
    Return

InsClip:
    V = Len(VClip): If V = 0 Then GoTo EmptyClip
    Tmp = "": If ChrPos > 1 Then Tmp = Left(VBuf, ChrPos - 1)                   'Data before insert point
    Tmp2 = "": If ChrPos < Len(VBuf) Then Tmp2 = Mid(VBuf, ChrPos)              'Data after  insert point
    VBuf = Tmp & VClip & Tmp2                                                   'Make new buffer with clip addeed
    ClearRange                                                                  'Clear the Range
    Return

AppendClip:
    V = Len(VClip): If V = 0 Then GoTo EmptyClip
    VBuf = VBuf & VClip                                                         'Append VClip
    ClearRange                                                                  'Clear the Range
    Return
    
RestoreRange:
    If VRestore = "" Then MsgBox "No Restore set!": Return                      'was: VRestore = VBuf2
    For J = ChrPos To ChrPosEnd
        Mid(VBuf, J, 1) = Mid(VRestore, J, 1)
    Next J
    Return
    
RestoreAll:
    If VRestore <> "" Then VBuf = VRestore                                      'If RestorePt is set then restore visible
    RedrawFlag = True                                                           'Force Re-draw when changing sets or if byte count has changed
    Return
    
SetRestorePoint:
    VRestore = VBuf
    Return
    
SwapSets:
    If VBuf2 = "" Then
        If MsgBox("Buffer 2 is empty. Would you like to copy Buffer 1?", vbYesNo) = vbNo Then Return
        VBuf2 = VBuf1
    End If
    
    UpdateBuf                                                                   'Update the Edited Buffer
    Tmp = VBuf1                                                                 'Remember set 1
    VBuf1 = VBuf2                                                               'Swap set 1 and 2
    VBuf2 = Tmp
    VBuf = VBuf1                                                                'Update Edit buffer
    RedrawFlag = True                                                           'Force Re-draw when changing sets or if byte count has changed
    ClearRange                                                                  'Clear the Range
    Return

SelectAll:
    SelChr = 0: SelChr2 = (VLen \ ChrHeight) - 1                                'Seleect Entire Range
    RangeFlag = True                                                            'Set Range true
    SetRange                                                                    'Set Range bytes positions
    Return
    
SelectSet:
    SelChr = ChrSetNum * ChrSetSize: SelChr2 = SelChr + ChrSetSize - 1          'Select Set
    RangeFlag = True                                                            'Set Range true
    SetRange                                                                    'Set Range bytes positions
    Return

NewFont:
    Tmp = InputBox("How many characters in the font?", "New Font", Format(ChrSetSize))
    If Tmp = "" Then Return
    nv = Val(Tmp)
    nv2 = nv * ChrHeight
    VBuf = String(nv2, Nu)                                                      'Make an empty font
    RedrawFlag = True
    Return
    
Compare:
    If MsgBox("This will compare Set#1 to Set#2 and place diff in clipboard", vbOKCancel) = vbCancel Then Return
    
    VClip = ""                                                                  'Clear the clipboard
    Max = Len(VBuf1): J = Len(VBuf2): If J < Max Then Max = J                   'Max compare size (minimum buffer sizes)
    
    For J = 1 To Max
        nv = Asc(Mid(VBuf1, J, 1))                                              'Get first byte value
        nv2 = Asc(Mid(VBuf2, J, 2))                                             'Get second byte value
        nv3 = nv Xor nv2                                                        'XOR them
        VClip = VClip & Chr(nv3)                                                'Add them to the string
    Next J
    Return

'=============================================================================== Manipulation Routines

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
    
'--- Read 8 bytes and fill the Source Bit array with 0's and 1's
ReadChr:
    For Row = 0 To 7
        V = Asc(Mid(VBuf, ChrTop + Row, 1))                                 'Get a byte/value
        If V > 0 Then                                                       'Only do bits if non-zero
            For Col = 0 To 7
                If (V And Pow(Col)) <> 0 Then SBit(Row, Col) = 1            'Set the bit array
            Next Col
        End If
    Next Row
    Return
    
'--- Write DBit Array out as 8 bytes
WriteChr:
    For Row = 0 To 7
            V = 0                                                           'Reset to zero
            For Col = 0 To 7
                If DBit(Row, Col) = 1 Then V = V + Pow(Col)                 'Add the value of the bit position
            Next Col
        Mid(VBuf, ChrTop + Row, 1) = Chr(V)                                 'Store bytes
    Next Row
    Return
    
End Sub

'---- Convert Font
Private Sub ConvertFont(ByVal N As Integer)
    Dim J As Integer, K As Integer, L As Integer
    Dim H As Integer, B As Integer, P As Integer
    Dim Tmp As String, Pad As String
    
    Select Case N
        Case 0: B = 8: H = 16: P = 0        'Read 16 bytes, write  8 no padding       - 8x8  font
        Case 1: B = 8: H = 8: P = 8         'Read  8 bytes, write  8 plus 8 padding   - 8x16 font
        Case 2: B = 5: H = 5: P = 3         'Read  5 bytes, write  5 plus 3 padding   - 5x7  sideways font
        Case 3: B = 7: H = 7: P = 1         'Read  7 bytes, write  7 plus 1 padding   - 5x7  upright font
        Case 4: B = 14: H = 14: P = 2       'Read 14 bytes, write 14 plus 2 padding   - 8x14 EGA font
        Case 5: B = 16: H = 32: P = 0       'Read 32 bytes, write 16 no padding       - 8x16 font
        Case 6: B = 6: H = 6: P = 2         'Read  6 bytes, write  6 plus 2 padding   - 6x8  font
    End Select
    
    Pad = String(P, Nu)                     'Make zero-padding string
    VBuf2 = ""                              'Buffer for converted font
    
    Select Case N
        ' Convert using parameters defined above
        Case 0 To 6
                For J = 1 To Len(VBuf) Step H
                    Tmp = Mid(VBuf, J, B)                                       'Get 8 bytes
                    VBuf2 = VBuf2 & Tmp & Pad                                   'Copy them plus padding if needed
                Next J
                      
        ' Convert 128 character set to 256 by adding 128 RVS characters
        Case 7
                H = 128 * ChrHeight
                
                For J = 1 To Len(VBuf) Step H
                    Tmp = Mid(VBuf, J, H)                                       'Get entire 128 characters
                    Pad = ""
                    For K = 1 To H
                        B = 255 - Asc(Mid(Tmp, K, 1))                           'RVS each byte
                        Pad = Pad & Chr(B)                                      'Add to string
                    Next K
                    VBuf2 = VBuf2 & Tmp & Pad                                   'Copy original plus padding
                Next J
        
        ' Convert Galaksija
        ' Data is ordered by Raster lines. There are 16 Rasters per character (ie: character height=16 pixels)
        ' IE: 128 RASTER#1 bytes, then 128 RASTER#2 bytes,... 128 RASTER#16 bytes.
        Case 8
                VBuf2 = String(2048, Nu)                                        'Zero out the string
                For J = 0 To 127                                                '128 characters total.
                    For K = 0 To 15                                             '16 bytes per character. Data is ordered by raster lines (K)
                        Tmp = Mid(VBuf, J + (128 * K) + 1, 1)                   'Get 1 byte (128 byte offsets)
                        Mid(VBuf2, J * 16 + K + 1) = Tmp                        'Relocate 1 byte (contiguous)
                    Next K
                Next J
        
    End Select
    
    VBuf = VBuf2                                                                'Update main buffer
    VClip = ""                                                                  'Clear clipboard
    VLen = Len(VBuf): lblVSize.Caption = Format(VLen)                           'Update buffer length
    
End Sub


'========================================
' Screen Designer
'========================================

'---- SCREEN DESIGNER: Draw Editor Screen
' Draws some or all of the screen
' MODE:         0=Chacter at Cursor, 1=Current Line, 2=Entire Screen
' SERow/SECol:  Cursor Row/Col
' SEMaxROW/COL: Screen Dimensions
' SECBMMode:    CBM/ASCII mode for handling RVS characters
' ChrSetSize:   128/256. 256 includes RVS characters in the set

' Uses SCREEN CODE/ASCII encoding (actual font rom). Uses the selected font "SET"
Private Sub DrawEditScreen(ByVal Mode As Integer)
    Dim R As Integer, C As Integer, V As Integer
    Dim PR As Integer, PC As Integer, TopY As Integer
    Dim SX As Integer, SY As Integer, SW As Integer, SH As Integer                  'Source X/Y/W/H
    Dim DX As Integer, DY As Integer, DW As Integer, DH As Integer                  'Destination X/Y/W/H
    Dim R1 As Integer, R2 As Integer, C1 As Integer, C2 As Integer                  'Rabge to Draw
    Dim OP As Long                                                                  'Blit Opcode
    Dim RvsFlag As Boolean, InvFlag As Boolean                                      'For Cursor and RVS characters
    
    If DesignerFlag = False Then Exit Sub                                           'Exit if not in Edit mode
    
    If SEBuf = "" Then Exit Sub                                                     'Exit if no buffer!
    
    '-- Set Screen Area to Update
    
    Select Case Mode
        Case 0                                                                      'Re-Draw at Cursor
            R1 = SERow: R2 = R1
            C1 = SECol: C2 = C1
        Case 1                                                                      'Re-Draw Current Line
            R1 = SERow: R2 = R1
            C1 = 0: C2 = SEMaxCol - 1
            BlinkFlag = False
        Case 2                                                                      'Re-Draw All
            R1 = 0: R2 = SEMaxRow - 1
            C1 = 0: C2 = SEMaxCol - 1
            SETimer.Enabled = False
            BlinkFlag = False                                                       'Force BLINK OFF
    End Select
        
    '-- Calculate Blit Sizing
    
    SW = 8 * ChrZoom
    SH = ChrHeight * ChrZoom
    DW = SEW * 8                                                                    'Scaled width of screen chr
    DH = SEH * ChrHeight                                                            'Scaled height of screen chr
    
    TopY = ChrSetNum * ((ChrSetSize \ ChrLineMax) * SH)                              'Pixel at top of set
    
    If Mode = 2 Then picScreen.Cls                                                  'Clear the bitmap
    
    '-- Set RVS Flag
 
    If (ChrSetSize = 256) And (SECBMFlag = True) Then
        RvsFlag = True                                                              'RVS characters are Included (CBM/256)
    Else
        RvsFlag = False                                                             'No included (CBM/128 or ASCII)
    End If
    
    '-- Loop to Update the Area
    
    For R = R1 To R2
        DY = R * DH                                                                 'Dest Y Co-ordinte (Pre-calc for speed)
        For C = C1 To C2
            V = GetChr(R, C)                                                        'Screen code (0-255)
            OP = vbSrcCopy                                                          'Default to NORMAL BLIT
            InvFlag = False                                                         'Default to NO INVERT
            
            If RvsFlag = False Then                                                 '-- No RVS characters in set
                If V > ChrSetSize Then                                              'Chr not in Set?
                    V = V Mod ChrSetSize: InvFlag = True                            'Strip HI BIT, Set Invert Flag
                End If
            End If
            
            If BlinkFlag = True Then                                                '-- CURSOR BLINK
                If RvsFlag = True Then
                    V = (V + 128) Mod 256                                           'Set includes RVS so toggle upper bit
                Else
                    InvFlag = Not InvFlag                                           'Flip the INVERT
                End If
            End If
            
            PR = Int(V / ChrLineMax)                                                'Picture Character ROW
            PC = V - (PR * ChrLineMax)                                              'Picture Character COL
            SX = PC * SW                                                            'Src Font X Co-ordinate
            SY = TopY + (PR * SH)                                                   'Src Font Y Co-ordinate
            DX = C * DW                                                             'Dest X Co-ordinate
            
            If InvFlag = True Then OP = vbNotSrcCopy                                 'If INVERT then set BLIT opcode to Invert
            
            picScreen.PaintPicture picV.Image, DX, DY, DW, DH, SX, SY, SW, SH, OP   'Blit it to screen

        Next C
    Next R
    
    picScreen.Visible = True
    SETimer.Enabled = True
    DoEvents
    
End Sub

'---- SCREEN DESIGNER: Popup the Screen Designer Menu
Private Sub cmdSEDMenu_Click()

    MenuForm = 2                                                        'The Viewer Form
    PopupMenu frmMenu.mnuScrEd                                          'Display the Screen Designer Menu
    
End Sub

'---- SCREEN DESIGNER: Cursor Blink
Private Sub SETimer_Timer()
    BlinkFlag = Not BlinkFlag                                           'Toggle the cursor
    DrawEditScreen 0                                                    'Draw character at cursor
End Sub

'---- SCREEN DESIGNER: Force Screen Refresh
Private Sub cmdScreenRefresh_Click()
    DrawEditScreen 2
    ScreenFocus
End Sub

'---- SCREEN DESIGNER: Set Focus on Screen for typing
Private Sub ScreenFocus()
    picScreen.SetFocus
    SETimer.Enabled = True
End Sub

'---- SCREEN DESIGNER: Change Typing Indicator when Got/Lost Focus
Private Sub picScreen_GotFocus()
    lblActive.BackColor = vbGreen
    SETimer.Enabled = True
End Sub
Private Sub picScreen_lostFocus()
    lblActive.BackColor = vbBlack
    SETimer.Enabled = False
End Sub

'---- SCREEN DESIGNER: Play Macro
' This plays back the macro. Assumes PETSCII.

Private Sub cmdPlay_Click()
    Dim i As Integer, V As Integer
    Dim Mac As String
    
    Mac = Macro(MacroNum)
    
    For i = 1 To Len(Mac)
        V = Asc(Mid(Mac, i, 1))
        Select Case V
            Case 0 To 31
                If V = 18 Then SERVSFlag = True                             'RVS ON
                If V = 13 Then WriteChr V                                   'RETURN
                
            Case 128 To 159
                If V = 146 Then SERVSFlag = False                             'RVS OFF
                
            Case Else
                WriteChr V
        End Select
    Next i
    
End Sub

'---- SCREEN DESIGNER: Type into Screen
Private Sub picScreen_KeyPress(KeyAscii As Integer)
    Dim K As Integer, V As Integer
    
    K = KeyAscii
    
    BlinkFlag = False                                                       'Make sure to draw it normal
    DrawEditScreen 0                                                        'Hide the cursor
    
    Select Case K
        Case 13: Cursor 5                                                   'Do RETURN
        Case 8: DoBackspace                                                 'Do BACKSPACE
        Case Else
            If SECBMFlag = True Then
                V = ASCIItoScreen(K)                                        'Convert to screen code
            Else
                V = K                                                       'Use the ASCII value directly
            End If
            
            If SERVSFlag = True Then V = (V + 128) Mod 256                  'Handle RVS
            
            If RECFlag = True Then
                If Len(Macro(MacroNum)) < 32766 Then Macro(MacroNum) = Macro(MacroNum) & Chr(V) 'Record the keystroke
            End If
            
            WriteChr V

    End Select

End Sub

'---- SCREEN DESIGNER: Handle Non-ASCII Keypresses
Private Sub picScreen_KeyDown(KeyCode As Integer, Shift As Integer)

    RestoreCursor
    
    Select Case KeyCode
        Case 36: If Shift = 1 Then Cursor 6 Else Cursor 0                   'Home or Clear Screen
        Case 37: Cursor 1                                                   'Cursor LEFT
        Case 39: Cursor 2                                                   'Cursor RIGHT
        Case 38: Cursor 3                                                   'Cursor UP
        Case 40: Cursor 4                                                   'Cursor DOWN
        Case 45:
            If Shift = 0 Then
                DoInsert                                                    'Insert Space
            Else
                WriteChr ChrNum                                             'Insert Selected Character
            End If
        Case 46: DoDelete                                                   'Delete
        Case 82: If Shift = 2 Then ToggleRVS: KeyCode = 0                   'CTRL-R = Toggle RVS mode
        
        Case Else:
            'Debug.Print "KeyCode="; KeyCode; " Shift="; Shift
    End Select
    
End Sub

'---- SCREEN DESIGNER: Write Character To Screen
' Adds to buffer, updates screen bitmap, advances cursor
Private Sub WriteChr(ByVal V As Integer)

    Mid(SEBuf, SECursorPos, 1) = Chr(V)
    BlinkFlag = False
    DrawEditScreen 0
    Cursor 2
    ScreenFocus
    
End Sub

'---- SCREEN DESIGNER: Insert a SPACE at Cursor
Private Sub DoInsert()
    Dim Tmp As String
    
    Tmp = " " & Mid(SEBuf, SECursorPos, SEMaxCol - SECol - 1)               'Get character from Cursor to END of line-1
    Mid(SEBuf, SECursorPos, Len(Tmp)) = Tmp                                 'Paste them at the Cursor
    DrawEditScreen 1                                                        'Redraw the Line
    
End Sub

'---- SCREEN DESIGNER: DELETE (Backspace) to left of Cursor
Private Sub DoBackspace()
    Dim Tmp As String
    
    If SECol = 0 Then Exit Sub                                              'Exit if at start of line
    Tmp = Mid(SEBuf, SECursorPos, SEMaxCol - SECol) & " "                   'Take string from COL to end of line
    Mid(SEBuf, SECursorPos - 1, Len(Tmp)) = Tmp                             'Paste it to COL - 1
    Cursor 1                                                                'Cursor LEFT
    DrawEditScreen 1                                                        'Redraw the Line
    
End Sub

'---- SCREEN DESIGNER: Delete at CURSOR
Private Sub DoDelete()
    Dim Tmp As String
    
    Tmp = Mid(SEBuf, SECursorPos + 1, SEMaxCol - SECol - 1) & " "           'Take string from COL+1 to end of line
    Mid(SEBuf, SECursorPos, Len(Tmp)) = Tmp                                 'Paste it to COL
    DrawEditScreen 1                                                        'Redraw the Line
    
End Sub

'---- SCREEN DESIGNER: Cursor Movement
' D: 0=Home, 1=left,2=right,3=up,4=down,5=Carriage Return
Private Sub Cursor(ByVal D As Integer)
    Dim N As Integer
    
    Select Case D
        Case 0: SECol = 0: SERow = 0                                            'Home the cursor
        Case 1: SECol = SECol - 1                                               'Cursor LEFT
        Case 2: SECol = SECol + 1                                               'Cursor RIGHT
        Case 3: SERow = SERow - 1                                               'Cursor UP
        Case 4: SERow = SERow + 1                                               'Cursor DOWN
        Case 5: SERow = SERow + 1: SECol = 0                                    'Carriage Return
        Case 6: ClearScreen True                                                'Clear Screen
    End Select
    
    If SECol >= SEMaxCol Then SECol = 0: SERow = SERow + 1                      'Check bounds and wrap
    If SECol < 0 Then SECol = SEMaxCol - 1: SERow = SERow - 1
    If SERow >= SEMaxRow Then SERow = 0
    If SERow < 0 Then SERow = SEMaxRow - 1
    
    SECursorPos = SERow * 80 + SECol + 1                                        'Remember the Cursor Position in the buffer
    
    N = Asc(Mid(SEBuf, SECursorPos, 1))
    
    lblCursor.Caption = "@ " & Str(SERow + 1) & "," & Str(SECol + 1) & "=" & Format(N, "###") & "/$" & MyHex(N, 2)         'Show the Cursor Position
End Sub

'---- SCREEN DESIGNER: Restore Cursor Character
Private Sub RestoreCursor()
    
    BlinkFlag = False                                                'Make sure to draw it normal
    DrawEditScreen 0                                            'Update character only

End Sub

'---- SCREEN DESIGNER: Set Screen Theme
Private Sub SetScreenTheme()

    lblScnBorder.BackColor = lblTheme(2).BackColor
    picScreen.ForeColor = lblTheme(0).BackColor
    
    If DesignerFlag = True Then DrawEditScreen 2
    
End Sub

'---- SCREEN DESIGNER: Get Character at specific Row and Col
' Returns screen code 0 to 255
Private Function GetChr(ByVal R As Integer, ByVal C As Integer) As Integer
    Dim BPos As Integer
    
    BPos = R * 80 + C + 1
    GetChr = Asc(Mid(SEBuf, BPos, 1))
    
End Function

'---- SCREEN DESIGNER: Reset "Machine"
' Clears the screen a shows power-on message
Private Sub ResetMachine()
        Dim Tmp As String, S1 As String, S2 As String, S3 As String
        Dim S4 As String, S5 As String, S6 As String, CBASIC As String
        
        CBASIC = " commodore basic "
        S2 = Cr & Cr & " 64k ram system  38911 basic"
        S3 = ""
        S4 = " bytes free" & Cr & Cr
        S5 = "ready." & Cr
        
        
        Select Case cboTheme.ListIndex
            Case 0 'C64
                S1 = "    **** " & CBASIC & "v2 ****"
                
            Case 1 'SX-64
                S1 = "     *****  sx-64 basic v2.0  *****"
                
            Case 2 'C128
                S1 = CBASIC & "v7.0 122365 bytes free" & Cr
                S2 = "   (c)1986 commodore electronics, ltd." & Cr
                S3 = "         (c)1977 microsoft corp." & Cr
                S4 = "           all rights reserved" & Cr & Cr
            
            Case 3 'VIC-20
                S1 = "**** cbm basic v2 ****"
                S2 = "": S3 = "2583"
                
            Case 4 'TED
                S1 = CBASIC & "3.5 60671 bytes free" & Cr
                S2 = " 3-plus-1 on key f1" & Cr & Cr: S4 = ""
                
            Case 5, 6, 7 'PET Green/White/Amber
                S1 = "***" & CBASIC & "4.0 ***" & Cr & Cr
                S2 = "": S3 = " 31743"
                
            Case 8 'CBM-II
                S1 = "***" & CBASIC & "256, v4.0 ***" & Cr & Cr
                S2 = "": S3 = "": S4 = ""
                
            Case 9 'CBM-II P500
                S1 = "***" & CBASIC & "128, v4.0 ***" & Cr & Cr
                S2 = "": S3 = "": S4 = ""
        End Select
                
        Tmp = S1 & S2 & S3 & S4 & S5
                
        ClearScreen False
        PutStr Tmp                                                          'Write the Banner text
        
        DrawEditScreen 2
End Sub

'---- SCREEN DESIGNER: Click to Reset Machine
Private Sub SEReset()
    ResetMachine
    ScreenFocus
End Sub

'---- SCREEN DESIGNER: Click to Clear the Screen
Private Sub SECLS()
    ClearScreen True
    ScreenFocus
    DoEvents
End Sub

'--- SCREEN DESIGNER: Insert Currently Selected Character
' Copies the selected "edit" character at the current cursor position
Private Sub cmdInsert_Click()
    WriteChr ChrNum
End Sub

'--- SCREEN DESIGNER: Toggle CBM Typing Mode
' When CBM Mode is Enabled, typed ASCII characters are converted to SCREEN.
' This changes the RVS indicator to be RVS or HI depending on the mode.
Private Sub cbCBM_Click()
    If cbCBM.value = vbChecked Then
        SECBMFlag = True                                                'Type in CBM Mode (screen codes)
        lblRVS.Caption = "RVS"                                          'CBM Mode = "RVS"
    Else
        SECBMFlag = False                                               'Type in ASCII Mode
        lblRVS.Caption = "HI"                                           'ASCII Mode = "HI"
    End If
    
    ScreenFocus
    
End Sub

'---- SCREEN DESIGNER: Click to Toggle RVS Mode
Private Sub lblRVS_Click()
    ToggleRVS
End Sub

'---- SCREEN DESIGNER: Click to Toggle RVS Mode
Private Sub lblREC_Click()
    SETogREC
End Sub


'---- SCREEN DESIGNER: Toggle RVS Mode
' When enabled the Upper BIT of the character typed is SET.
' In CBM Mode this selects the equivilent "RVS" Character.
' IN ASCII Mode this selects the extended character.
Private Sub ToggleRVS()
    
    SERVSFlag = Not SERVSFlag
    If SERVSFlag = True Then lblRVS.BackColor = vbRed Else lblRVS.BackColor = vbBlack
    ScreenFocus

End Sub

'---- SCREEN DESIGNER: Toggle Macro Record Mode
Private Sub SETogREC()
    
    RECFlag = Not RECFlag
    If RECFlag = True Then lblREC.BackColor = vbRed Else lblREC.BackColor = vbBlack
    ScreenFocus

End Sub

'---- SCREEN DESIGNER: Clear the Macro
Private Sub cmdClearMacro_Click()
    Macro(MacroNum) = ""
End Sub

'---- SCREEN DESIGNER: Click to Close Designer Window
Private Sub lblActive_Click()
    ToggleDesigner
End Sub

'---- SCREEN DESIGNER: Toggle Screen Designer
Private Sub ToggleDesigner()
    Dim FlagX As Boolean
        
    DesignerFlag = Not DesignerFlag                               'Toggle it
    
    FlagX = False
    
    If DesignerFlag = True Then
    
        frEditor.Visible = True: DoEvents
        
        BlinkFlag = False                                           'Start the cursor off
        BorderFlag = False                                          'Turn Border off
        If ChrZoom > 2 Then ChrZoom = 2: FlagX = True                           'Set Zoom Scale
        If ChrWIndex > 4 Then ChrWIndex = 2: ChrLineMax = 32: FlagX = True   'Set Max Characters per Line
        
        SetChrZoomSelector                                          'Set Zoom Indicator
        SetChrWidthSelector                                         'Set Width Indicator
        SetScreenTheme                                              'Set screen and border colours
        
        If FlagX = True Then RedrawFlag = True: UpdateChrSetView
        
        If SEBuf = "" Then
            InitEditor                                   '-- Initialize the Editor
            ResetMachine
        End If

        SetFocus
    Else
        SETimer.Enabled = False
        frEditor.Visible = False
    End If
    
    SetFontMenu
    
    
End Sub

'---- SCREEN DESIGNER: Init the Editor
Private Sub InitEditor()
        
        SEBuf = String(4000, " ")                                               'Initialize the buffer for max 80x50 characters
        
        picScreen.ForeColor = lblTheme(0).BackColor                             'Set colours
        picScreen.BackColor = lblTheme(1).BackColor
        lblScnBorder.BackColor = lblTheme(2).BackColor
        
        SetScnEdFmt 1

End Sub

'---- SCREEN DESIGNER: Clear the Screen
Private Sub ClearScreen(ByVal Flag As Boolean)
        SEBuf = String(4000, " ")                                               'Initialize the buffer for max 80x50 characters
        Cursor 0                                                                'Home Cursor
        
        If Flag = True Then DrawEditScreen 2                                 'Draw the screen
End Sub

'---- SCREEN DESIGNER: Put a string to the Screen - For pre-editing printing
Private Sub PutStr(ByVal Tmp As String)
    Dim i As Integer, K As Integer, V As Integer
    
    For i = 1 To Len(Tmp)
        K = Asc(Mid(Tmp, i, 1))                                                 'Get the character Code
        If K = 13 Then
            Cursor 5
        Else
            If SECBMFlag = True Then
                V = ASCIItoScreen(K)                                            'Convert to Screencode
            Else
                V = K                                                           'Use ASCII directly
            End If
            
            PutCh V                                                             'Print it
        End If
    Next i
        
End Sub

'---- SCREEN DESIGNER: Put character to Screen
' Places a single character (0-255) to the current cursor position and increments cursor and handles line wrap
Private Sub PutCh(ByVal N As Integer)
    
    Mid(SEBuf, SECursorPos) = Chr(N)                                            'Write character to buffer at cursor position
    Cursor 2                                                                    'Increment column and Wrap to next Row if needed
    
End Sub

'---- SCREEN DESIGNER: Click to change Screen Format
Private Sub cboScnFmt_Click()
    SetScnEdFmt cboScnFmt.ListIndex
End Sub

'--- SCREEN DESIGNER: Set Screen Format
' Sets Max Rows/Col and W/H scale for screen draw
Private Sub SetScnEdFmt(ByVal Index As Integer)

    Select Case Index
        Case 0: SEMaxRow = 23: SEMaxCol = 22: SEW = 3: SEH = 2                  '22 x 23
        Case 1: SEMaxRow = 25: SEMaxCol = 40: SEW = 2: SEH = 2                  '40 x 25
        Case 2: SEMaxRow = 25: SEMaxCol = 80: SEW = 1: SEH = 2                  '80 x 25
        Case 3: SEMaxRow = 50: SEMaxCol = 80: SEW = 1: SEH = 1                  '80 x 50
    End Select
    
    If ChrHeight = 16 Then SEH = 1                                              'Adjust for 8x16 fonts
    
    picScreen.Width = SEMaxCol * 8 * SEW * Screen.TwipsPerPixelX                '80x25=9600  (80 x 8 x 15twip)
    picScreen.Height = SEMaxRow * ChrHeight * SEH * Screen.TwipsPerPixelY       '80x25=6000
    
    If SERow >= SEMaxRow Then SERow = 0
    If SECol >= SEMaxCol Then SECol = 0                                         'Reset Cursor
    
    If DesignerFlag = True Then DrawEditScreen 2
    
End Sub

'---- SCREEN DESIGNER: Load Screen Buffer
Private Sub SELoad()
    LoadSEScreen True
End Sub

'---- SCREEN DESIGNER: Save Screen Buffer
Private Sub SESave()
    SESaveScreen True
End Sub

'---- SCREEN DESIGNER: Save Screen in Current Format
Private Sub SEExport()
    SESaveScreen False
End Sub

'---- SCREEN DESIGNER: Save Screen as Binary
' FLAG: True=Save Complete screen (4K buffer), False=Save as Current Screen Format (ie: Only visible character)
Private Sub SESaveScreen(ByVal Flag As Boolean)
    Dim FIO As Integer, Filename As String, J As Integer, S As Integer
    
    Filename = FileOpenSave(FileBase(LastFile), 1, 8, "Save Designer Screen")  '1=Save,0=CBM Files
    If Filename = "" Then Exit Sub
    
    If Overwrite(Filename) = False Then Exit Sub            'Exit if user does not want to overwrite the file
    
    FIO = FreeFile
    Open Filename For Output As FIO
    
    If Flag = False Then                                     '-- Save as Screen Format
        For J = 0 To SEMaxRow - 1
            S = J * 80 + 1                                  'Position of start of line (all lines are 80 bytes long
            Print #FIO, Mid(SEBuf, S, SEMaxCol);            'Write the correct number of bytes for this line
        Next
    Else
        Print #FIO, SEBuf;                                  '-- Save Complete Buffer
    End If
    
    Close FIO
    ChDir ExeDir
End Sub

'---- SCREEN DESIGNER: Load Screen Buffer
' FLAG: True=Load Buffer
Private Sub LoadSEScreen(ByVal Flag As Boolean)
    Dim FIO As Integer, Filename As String, FLen As Single, FLen2 As Single
    Dim Tmp As String
    
    Filename = FileOpenSave(FileBase(LastFile), 0, 8, "Load Designer Screen")  '0=Load,8=CBM Files
    If Filename = "" Then Exit Sub
    If Exists(Filename) = False Then MsgBox "File not found!": Exit Sub
    
    FIO = FreeFile
    Open Filename For Input As FIO
        FLen = LOF(FIO): FLen2 = FLen
        If FLen2 > 4000 Then FLen2 = 4000                               'Limit the size to load
        Tmp = Input(FLen2, FIO)                                         'Load the bytes
    Close FIO
    
    SEBuf = Tmp & String(4000 - FLen, " ")                              'Set the buffer. Pad with spaces if needed
    
    ChDir ExeDir

    If FLen < 4000 Then MsgBox "The file was smaller then 4000 bytes long and has been padded."
    If FLen > 4000 Then MsgBox "The file was larger then 4000 bytes long and has been truncated."
    
    Cursor 0
    DrawEditScreen 2
    ScreenFocus
    
End Sub

'---- SCREEN DESIGNER: Save the Screen as Bitmap
Private Sub SESaveBMP()
    Dim Filename As String
    
    Filename = FileOpenSave("screen.bmp", 1, 3, "Save Screen as BMP")
        
    If Filename <> "" Then SavePicture picScreen.Image, Filename

End Sub

'---- SCREEN DESIGNER: Load Macro
Private Sub SELoadMacro()
    Dim FIO As Integer, Filename As String, FLen As Single, FLen2 As Single
    Dim Tmp As String
    
    Filename = FileOpenSave(FileBase(LastFile), 0, 0, "Load Macro")  '0=Load,8=CBM Files
    If Filename = "" Then Exit Sub
    If Exists(Filename) = False Then MsgBox "File not found!": Exit Sub
    
    FIO = FreeFile
    Open Filename For Binary As FIO
        FLen = LOF(FIO): FLen2 = FLen
        If FLen2 > 32760 Then FLen2 = 32760                             'Limit the size to load
        Tmp = Input(FLen2, FIO)                                         'Load the bytes
    Close FIO
    
    Macro(MacroNum) = Tmp                                               'Set the Macro
    
    ChDir ExeDir
    
End Sub

'---- SCREEN DESIGNER: Save Macro
Private Sub SESaveMacro()

End Sub


'========================================
' ML VIEW - Machine Language Disassembler
'========================================
Sub MLView()
    Dim GoodFlag As Boolean
    Dim J As Integer
    Dim C As Integer, StartC As Integer                                 'Counter
    Dim Tmp As String, Tmp2 As String, Tmp3 As String, Tmp4 As String   'Temp strings
    Dim TmpB As String
    
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
    Dim CommentCol As Integer, DivStr As String                         'Divider String
    
    Dim LNum As Long, LInc As Integer                                   'Line Numbers
    Dim a As Integer, P As Integer, P2 As Integer
    
    Dim DTMode As Boolean, DTCount As Integer, DTType As String         'Data Table variables
    Dim DTCountMax As Integer, DTMax As Integer, DTPos As Integer       'Data Table variables
    Dim DTStart As Long, DTEnd As Long, DTAscMode As Integer            'Data Table variables
    Dim DTComment As String, DTAddress As String, DTOutStr As String    'Data Table variables
        
    Dim Pass As Integer
    Dim RTSOption As Boolean, SymComment As Boolean, DivLen As Integer  'options
    Dim HHHHFlag As Boolean                                             'Flag to addd Hex address to comments
    Dim ABlockFlag As Boolean, ABlockAddr As Long, Tblk As String       'Assembly Block: Flag, Start Address,Temp
    
    ViewerReady = False
        
    LInc = Val(txtLineInc.Text)                                         'Get Line# Increment
        
    '---- Options
    
    RTSOption = False: If cbSpaceRTS.value = vbChecked Then RTSOption = True
    SymComment = False: If cbIncSym.value = vbChecked Then SymComment = True
    HHHHFlag = False: If cbHHHH.value = vbChecked Then HHHHFlag = True
    DivLen = Val(txtDivLen.Text)
    CommentCol = Val(txtInlineCol.Text)
    
    '--------------------------------------------
    ' Load Support Files and Config settings etc
    '--------------------------------------------
    
    '---- Load ML Config File
    If MLCFlag = False Then LoadMLConfig
    If MLCFlag = False Then MyMsg "ML Config file is missing!": Exit Sub
        
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
    
    DTMax = lstDT.ListCount - 1                     'Last Data Table Entry
    OutFmt = cboMLFmt.ListIndex
    
    lstML.Visible = False                           'Hide the ML output
    lstML2.Visible = False                          'Hide the second ML output
    lstLabels.Clear                                 'Clear [GEN] labels list
    lstJSR.Clear                                    'Clear [JSR] list
    lblGood.BackColor = vbYellow: GoodFlag = True   'Set status box colour
       
    DoEvents
    
    '=========================================================================================
    ' This is the PASS loop. In PASS 1 labels are generated. In PASS 2 the output is generated
    '=========================================================================================
    
    For Pass = 1 To 2
        C = StartC                                                      'Start position 1 or 3 depending if load address is skipped
        
        lblEA.Caption = "Disassembling... PASS#" & Str(Pass)
        lblEA.BackColor = vbYellow
        DoEvents
        
        lstML.Clear                                                     'Clear the output
        lstML2.Clear                                                    'Clear the second list
        
        ABlockFlag = False                                              'Clear the ABlockFlag
        DTMode = False                                                  'Clear the DTMode flag
        DTCount = 0: DTPos = -1: DTStart = 0: DTEnd = 0                 'Reset Data Table pointer
        LNum = Val(txtStartLine.Text)                                   'Set the Starting Line#
        C = 1
        Address = VLA: If cbLA.value = vbUnchecked Then Address = MyDec(txtLA.Text)
        txtLA.Text = MyHex(Address, 4)
        StartAddress = MyHex(Address, 4)
        EndAddress = MyHex(Address + VLen - 1, 4)
        
                
        '====================================
        ' PASS 2 - Add Equates
        '====================================
        
        If (Pass = 2) And (cbEquates.value = vbChecked) Then
            If OutFmt = 2 Then
                lstML.AddItem Format(LNum) & " ; Disassembly of: " & FileNameOnly(VName) & "  DATE: " & Date: LNum = LNum + LInc
                lstML.AddItem Format(LNum) & " ;": LNum = LNum + LInc
                lstML.AddItem Format(LNum) & " ; ----- Equates": LNum = LNum + LInc
                lstML.AddItem Format(LNum) & " ;": LNum = LNum + LInc
            Else
                lstML.AddItem "; Disassembly of: " & FileNameOnly(VName) & "  DATE: " & Date
                lstML.AddItem ";"
                lstML.AddItem "; ---- Equates"
                lstML.AddItem ";"
            End If
            
            For J = 0 To lstSYM.ListCount - 1
                If lstSYM.Selected(J) = True Then
                    Tmp = lstSYM.List(J)
                    T1 = "": If OutFmt = 2 Then T1 = Format(LNum) & " ": LNum = LNum + LInc
                    Tmp2 = Pad(GetField(Tmp, 2), 20)
                    Tmp3 = GetField(Tmp, 1): If Left(Tmp3, 2) = "00" Then Tmp3 = Mid(Tmp3, 3)    'Remove zero page leading zeros
                    Tmp4 = Pad(Tmp2 & "=$" & Tmp3, CommentCol)
                    lstML.AddItem T1 & Tmp4 & ";" & GetField(Tmp, 3)
                End If
            Next J
            If OutFmt = 2 Then lstML.AddItem Format(LNum) & " ;" Else lstML.AddItem ";"
        End If
        
        '====================================
        ' PRE PASS - Add Code Origin
        '====================================
                
        If Pass = 2 Then
            If OutFmt = 2 Then
                lstML.AddItem Format(LNum) & " " & DOTORG & DOTHEX & StartAddress: LNum = LNum + LInc
                lstML.AddItem Format(LNum) & " ;": LNum = LNum + LInc
                lstML.AddItem Format(LNum) & " ; ---- Code": LNum = LNum + LInc
                lstML.AddItem Format(LNum) & " ;": LNum = LNum + LInc
            Else
                lstML.AddItem DOTORG & DOTHEX & StartAddress
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
            T2 = Pad(B0H, 16)                                       'Default to opcode byte and spacing
            T4 = ""                                                 'Formatted code
            T5 = ""                                                 'Comment area
            LastComment = ""                                        'Clear Last Comment
            SH = ""                                                 'HI
            SL = ""                                                 'LO
            SHL = ""                                                'Word
            DTMode = False                                          'Clear Data Table Mode
                        
            '===================================================
            ' PASS 2 only. Handle Symbols, Labels, and Comments
            '===================================================
            
            If Pass = 2 Then
                '---- Handle Comments
                UComment = FindComment(T0)                          'Check for a comment here. FindComment returns: type,text
                If UComment > "" Then
                    TmpB = UCase(Left(UComment, 1))                 'Check comment type (I,S or divider)
                    UComment = Mid(UComment, 3)                     'Strip away comment type
                    DivStr = ""
                    If TmpB <> "S" Then
                        DivStr = ";" & String(DivLen, TmpB)         'Generate comment with Divider using specified character
                        If OutFmt = 2 Then DivStr = " " & DivStr    'Format 2 need additional space infront
                    End If
                    
                    Select Case TmpB
                        Case "I"                                                        '"I"=Inline Comment (ignore)
                        Case "["                                                        '"["=Block Comment
                            If cbBlock.value = vbChecked Then
                                UComment = ";" & UComment                               'All block comments will have at least one ";" at the start.
                                P = 1
                                
                                Do
                                    P2 = InStr(P + 1, UComment, ";")
                                    If P2 = 0 Then Tmp = Mid(UComment, P)               'The rest of the line
                                    If P2 > 0 Then Tmp = Mid(UComment, P, P2 - P)       'The string up to the ";"
                                    
                                    If Len(Tmp) = 3 Then
                                        If Left(Tmp, 2) = ";/" Then
                                            Tmp = ";" & String(DivLen, Mid(Tmp, 3, 1))  'Convert to divider line"
                                        End If
                                    End If
                                    If OutFmt = 2 Then
                                        lstML.AddItem Format(LNum) & " " & Tmp          'ADD Line# and Comment <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                    Else
                                        lstML.AddItem Tmp                               'ADD Comment only <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                    End If
                                    LNum = LNum + LInc                                  'Increase Line #
                                    P = P2                                              'Position to the ";"
                                Loop While P2 > 0
                            End If
                            
                            UComment = ""                                               'Clear it. If type is "i" (inline) then we'll add it later
                            
                        Case Else                                                       'All other comment types (S=Standalone)
                        
                            If OutFmt = 2 Then
                                    ' FORMAT=2 has line# on every line
                                    If TmpB <> "S" Then lstML.AddItem Format(LNum) & DivStr: LNum = LNum + LInc     'Add a divider line
                                    If UComment > "" Then
                                        If HHHHFlag = True Then UComment = "[" & T0 & "] " & UComment               'Add Hex Address
                                        lstML.AddItem Format(LNum) & " ; " & UComment                               'Add comment string <<<<<<<<<<<<<<<<<
                                        LNum = LNum + LInc
                                        If TmpB <> "S" Then lstML.AddItem Format(LNum) & DivStr: LNum = LNum + LInc 'Add a divider line
                                    End If
                            Else
                                    'Other formats have no line numbers
                                    If TmpB <> "S" Then lstML.AddItem DivStr: LNum = LNum + LInc
                                    If UComment > "" Then
                                        If HHHHFlag = True Then UComment = "[" & T0 & "] " & UComment               'Add HEX Address
                                        lstML.AddItem "; " & UComment                                               'ADD comment <<<<<<<<<<<<<<<<<<<<<<<<<
                                        If TmpB <> "S" Then lstML.AddItem DivStr: LNum = LNum + LInc
                                    End If
                            End If
                            UComment = ""       'clear it since it's been used. if type is "i" (inline) then we'll add it later
                            
                    End Select
                End If
                
                '===================================================
                ' PASS 2 - Handle Labels
                '===================================================
                
                Tmp = FindUL(T0)    'Find User Label or Generated Label
                If Tmp > "" Then
                    ALabel = Tmp & ":"                                                                                      'Address Label
                    
                    Select Case OutFmt
                        Case 0, 1, 3                                                                                        'label on it's own
                            If cbLabelBlanks.value = vbChecked Then lstML.AddItem ";"                                       'ADD blank line <<<<<<<<<<<<<<<<
                            lstML.AddItem ALabel
                        Case 2                                                                                              'line number and label
                            If cbLabelBlanks.value = vbChecked Then lstML.AddItem Format(LNum) & " ;": LNum = LNum + LInc
                            lstML.AddItem Format(LNum) & " " & ALabel: LNum = LNum + LInc                                   'ADD blank line <<<<<<<<<<<<<<<<<
                        Case 4                                                                                              'label cmd param
                    End Select
                End If
            End If
            
            '===================================================
            ' PASS 1 and 2 - Handle Stepping through Data Tables
            '===================================================
        
            '-- IF we are NOT inside a Table range then check if New Data Table Range
            If DTStart = 0 Then
                '---- Search for the Next SELECTED Data Table Entry
                Do
                    If DTPos >= DTMax Then Exit Do
                    DTPos = DTPos + 1                                   'Go to next position
                    If lstDT.Selected(DTPos) = True Then Exit Do        'If it is selected then use it
                Loop
                
                
                '---- Check if there are more TABLES
                If DTPos <= DTMax Then
                    '---- Yes, so look at the current range entry.       Format: HHHH,HHHH,T{num},Comment
                    Tmp = lstDT.List(DTPos)                             'Get the line from the list
                    
                    DTStart = MyDec(Mid(Tmp, 1, 4))                     'Get Range Start
                    DTEnd = MyDec(Mid(Tmp, 6, 4))                       'Get Range End
                    Tmp = Mid(Tmp, 11)                                  'Get just the Type and Comment
                    DTType = UCase(Left(Tmp, 1))                        'Get Type (Asc,Byte,Word,Vector,RVector etc)
                    
                    P = InStr(Tmp, ","): If P = 0 Then P = 1            'Check for comma
                    DTCountMax = -1                                     'Default Items per line
                    If P > 2 Then DTCountMax = Val(Mid(Tmp, 2, P - 1))  'If specified, use {num} entries. Num must be single digit
                    If Pass = 2 Then DTComment = Mid(Tmp, P + 1)        'Get Comment
                Else
                    '---- No more TABLES, B0H set to highest byte $FFFF
                    DTStart = CLng(65536): DTEnd = CLng(65536): DTComment = "end"
                End If
            End If
            
            ABlockFlag = False
                            
            '===================================================
            ' PASS 1 and 2 - Handle Stepping through Data Tables
            '===================================================
            
            If Address >= DTStart Then
                MD = 0
                '---- Check if Table also has a symbol or label. If not. add one.
                If Address = DTStart Then                               '---- This is the first byte of the range
                    DTAscMode = 0                                       'Reset Asc mode
                    '---- It should have a label
                    Tmp = FindSym(DTStart)                              'Is there a SYMBOL?

                    If Tmp = "" Then                                    'No
                       Tmp = FindLabel(DTStart)                         'Is there a LABEL?
                       If Tmp = "" Then                                 'No
                            lstLabels.AddItem MyHex(DTStart, 4)         'Automatically ADD a label
                       End If
                    End If
                    
                    If DTType = "A" Then                                'Is it Assembly Block?
                        ABlockAddr = DTStart                            'Default to same as normal address
                        If Len(DTComment) = 4 Then                      'Must be exactly 4 bytes
                            If MyDec(DTComment) > 0 Then                'Is it >0?
                                ABlockAddr = MyDec(DTComment)           'If a real hex number then use it as the start address for the block
                            End If
                        End If
                    End If

                End If
                
                '===================================================
                ' PASS 1 and 2 - Data Tables
                '===================================================
                
                If Address <= DTEnd Then
                    '---- We are inside a data range!
                    ' In PASS 1 we can generally skip over everything except for "V" and "R" modes,
                    ' which need to add labels for the target addresses.

                    DTMode = True                                                                               'Set the Flag to indicate we are INSIDE a Table!
                    
                    If DTCount = 0 Then DTAddress = T1: DTOutStr = ""                                           'Initialize line string
                    If (DTCount > 0) And ((DTType <> "S") And (DTType <> "T")) Then DTOutStr = DTOutStr & ","   'Add a comma between entries unless String mode
                    
                    Select Case DTType  'Valid: S/T,B/H/$,D,Z/%,W,V,R,X,Y,A,L
                        Case "S", "T"                                                           '---- String/Text Directive
                            If Pass = 2 Then                                                    '---- We now need to build the output string, handling printable and non-printable bytes
                                
                                DTCountMax = 20
                                T3 = DOTTEXT                                                    'Set the "!TEXT" string
                                
                                '---- DTAscMode: 0=initial state, 1=non-printable, 2=printable (inside quotes)
                                Select Case B0C
                                    Case Qu                                                     '-- Quote
                                        If DTAscMode = 2 Then DTOutStr = DTOutStr & Qu & ","    'Finish off quote mode then comma
                                        DTOutStr = DTOutStr & "$22"
                                        DTAscMode = 1                                           'Set Non-Printable mode (hex values)
                                        
                                    Case " " To "z"                                             '-- Space or Letter
                                        If DTAscMode = 0 Then DTOutStr = DTOutStr & Qu          'Add Quote
                                        If DTAscMode = 1 Then DTOutStr = DTOutStr & "," & Qu    'Add Comma + Quote
                                        DTOutStr = DTOutStr & B0C
                                        DTAscMode = 2                                           'Set Printable mode (inside quotes)
                                        
                                    Case Else                                                   '-- Non-printable character
                                        If DTAscMode = 2 Then DTOutStr = DTOutStr & Qu & ","    'Add End Quote + Comma
                                        If DTAscMode = 1 Then DTOutStr = DTOutStr & ","         'Add Comma
                                        DTOutStr = DTOutStr & MyHex(B0A, -2)
                                        DTAscMode = 1                                           'Set Non-Printable mode (hex values)
                                End Select
                            End If
                            
                        Case "B", "H", "$"                                                      '---- Byte Directive (Hex)
                            If Pass = 2 Then
                                T3 = DOTBYTE                                                    'Set the "!BYTE" string
                                If DTCountMax < 1 Then DTCountMax = 8                           'Set Maximum entries per line
                                DTOutStr = DTOutStr & DOTHEX & B0H                                 'Add HEX byte
                            End If
                            
                        Case "D"                                                                '---- Byte Directive (Dec)
                            If Pass = 2 Then
                                T3 = DOTBYTE                                                    'Set the "!BYTE" string
                                If DTCountMax < 1 Then DTCountMax = 8                           'Set Maximum entries per line
                                DTOutStr = DTOutStr & B0A                                       'Add Decumal byte
                            End If
                            
                        Case "Z", "%"                                                           '---- Byte Directive (Binary)
                            If Pass = 2 Then
                                T3 = DOTBYTE                                                    'Set the "!BYTE" string
                                If DTCountMax < 1 Then DTCountMax = 4                           'Set Maximum entries per line
                                DTOutStr = DTOutStr & "%" & MyBin(B0A)                          'Add Binary byte
                            End If
                            
                        Case "L"                                                                '---- Little-Endian Word Directive (Hex)
                            If Pass = 2 Then
                                T3 = DOTWORD                                                    'Set the "!WORD" string
                                If DTCountMax < 1 Then DTCountMax = 8                           'Set Maximum entries per line
                                Address = Address + 1: C = C + 1                                'Increment address
                                B1A = Asc(Mid(VBuf, C, 1))                                      'Get next byte
                                SL = B0H                                                        'Lo Byte
                                SH = MyHex(B1A, 2)                                              'HI Byte
                                DTOutStr = DTOutStr & DOTHEX & SL & SH                             'Add HEX HEX to output list
                            End If
                            
                        Case "W"                                                                '---- Big Endian Word Directive (Hex)
                            If Pass = 2 Then
                                T3 = DOTWORD                                                    'Set the "!WORD" string
                                If DTCountMax < 1 Then DTCountMax = 8                           'Set Maximum entries per line
                                Address = Address + 1: C = C + 1                                'Increment address
                                B1A = Asc(Mid(VBuf, C, 1))                                      'Get next byte
                                SL = B0H                                                        'Lo Byte
                                SH = MyHex(B1A, 2)                                              'HI Byte
                                DTOutStr = DTOutStr & DOTHEX & SH & SL                             'Add HEX HEX to output list
                            End If
                            
                        Case "V"                                                                '---- "V" Word, Vector address
                            '---- Take the next byte and generate an address.
                            If DTCountMax < 1 Then DTCountMax = 6
                            Address = Address + 1: C = C + 1                                    'Increment address
                            B1A = Asc(Mid(VBuf, C, 1))                                          'Value of byte
                            TAddress = B1A * 256 + B0A                                          'Calculate Target Address (decimal)
                            JAddress = MyHex(TAddress, 4)                                       'Make it a string
                            SHL = DOTHEX & JAddress                                                'Make string for output

                            If Pass = 1 Then
                                If (JAddress >= StartAddress) And (JAddress <= EndAddress) Then
                                    lstLabels.AddItem JAddress                                  'Target is inside code range, so ADD a label for it
                                End If
                            Else
                                '---- PASS 2
                                T3 = DOTWORD                                                    'Set the "!WORD" string
                                Tmp = FindSL(JAddress)                                          'Look for Target address
                                If Tmp = "" Then Tmp = SHL
                                DTOutStr = DTOutStr & Tmp                                       'Add to output string
                            End If
                            
                        Case "R"                                                                '---- "R" Word, RTS Vector address
                            '---- Take the next byte and generate an address
                            If DTCountMax < 1 Then DTCountMax = 6
                            Address = Address + 1: C = C + 1
                            B1A = Asc(Mid(VBuf, C, 1))                                          'Value of byte
                            TAddress = B1A * 256 + B0A + 1                                      'Calculate Target Address (decimal) with Offset
                            JAddress = MyHex(TAddress, 4)                                       'Make it a string
                            SHL = DOTHEX & JAddress                                                'Make string for output
                            
                            If Pass = 1 Then
                                '---- PASS 1
                                If (JAddress >= StartAddress) And (JAddress <= EndAddress) Then
                                    lstLabels.AddItem JAddress                                  'Target is inside code range, so ADD a label for it
                                End If
                            Else
                                '---- PASS 2
                                T3 = DOTWORD                                                    'Set the "!WORD" string
                                Tmp = FindSL(JAddress)
                                If Tmp = "" Then Tmp = SHL
                                DTOutStr = DTOutStr & Tmp & "-1"                                'Add to output string with "-1" offset
                            End If
                        
                        Case "X"                                                                '---- "X" Hide the entire range
                            T3 = ""
                    
                        Case "A" '---- "A" Assembly Block
                            DTCountMax = 3
                            ABlockFlag = True                                                   'Set FLAG so that code is disassembled
                            
                    End Select
                    
                    '---- If "A" Block then do not update address as it will be done in the disassembly section below
                    If ABlockFlag = False Then
                        C = C + 1
                        Address = Address + 1                                                   'Increment address for each byte
                        DTCount = DTCount + 1                                                   'Store up x bytes
                    End If
                    
                    '===================================================
                    ' PASS 2 - Add TABLE data
                    '===================================================
                    If Pass = 2 Then
                        If (DTCount >= DTCountMax) Or (Address > DTEnd) Or (C > VLen) Then
                            '---- We've done 'DTCountMax' entries, or we've reached the end of the table or file
                            If (DTType = "S") Or (DTType = "T") Then
                                '-- we need to finish off the string properly
                                If DTAscMode = 2 Then DTOutStr = DTOutStr & Qu                  'Add ending quote
                                DTAscMode = 0
                            End If

                            If T3 > "" Then
                                '---- Add a line according to selected format
                                Tblk = T3 & DTOutStr                                            'common ending
                                
                                Select Case OutFmt
                                    Case 0, 1: Tmp = DTAddress & Tblk                           'addr cmd param (bytes not listed for 0)
                                    Case 2: Tmp = Format(LNum) & " " & Tblk                     'nnnn cmd param
                                    Case 3: Tmp = Pad("", 16) & Tblk                            'cmd param
                                    Case 4:
                                        Tmp = Pad(ALabel, 16) & Tblk                            'label cmd param
                                        ALabel = ""                                             'Blank it for multi-line tables
                                End Select
                                
                                If ABlockFlag = False Then
                                    J = Len(Tmp) + 1: If CommentCol > J Then J = CommentCol     'Calculate comment position. Try to line up except for really long line
                                    lstML.AddItem Pad(Tmp, J) & ";" & DTComment                 'ADD IT to output <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                    LNum = LNum + LInc                                          'Increment the Line Number
                                End If

                            End If
                            T4 = ""
                            DTCount = 0
                            ABlockFlag = False
                            
                        End If
                    End If
                
                    '---- Check if we are finished with the current table
                    If Address > DTEnd Then
                        If (RTSOption = True) And (ABlockFlag = False) Then                     'Handle RTS spacing option
                            Select Case OutFmt
                                Case 2: lstML.AddItem Format(LNum) & " ;": LNum = LNum + LInc   'Add a Line Number
                                Case Else: lstML.AddItem " "                                    'Add a Blank line
                            End Select
                        End If
                        
                        DTStart = 0                                                             'Clear TABLE Start value

                    End If
                End If
            End If
                        
            '=========================================================
            ' If NOT inside a Table range then process as valid opcode          EXCEPTION: Also if inside "A" Block Table!!!
            '=========================================================
            
            If (DTMode = False) Or (ABlockFlag = True) Then
                NM = Left(OP(B0A), Len(OP(B0A)) - 1)                            'Mneumonic string (eg: JSR or BBR0)
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
                        SHL = DOTHEX & JAddress                                    'Add the $ to HI string
                        SL = DOTHEX & SL                                           'Add the $ to LO string
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
                
                '===================================================
                ' PASS 2 - Build output string
                '===================================================
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
                            If B1A > 127 Then B1A = B1A - 256                                   'Calculate backwards branch
                            RAddress = MyHex(Address + B1A + 2, 4)                              'Make HHHH string
                            Tmp = FindSL(RAddress): If Tmp > "" Then T4 = " " & Tmp             'Lookup Relative Address in Symbols and Labels lists
                            T5 = ""
                        Case 11: T4 = " (" & SL & ",X)"     'k-Indexed Indirect Addressing with X
                        Case 12: T4 = " (" & SL & "),Y"     'l-Indexed Indirect Addressing with Y
                        Case 13: T4 = " (" & SHL & ")"      'm-Absolute Indirect
                        Case 14: T4 = " (" & SHL & ",X)"    'n-iax (65c02)
                        Case 15: T4 = " " & SL & "," & SH   'o-zpr (65c02) ***** need to convert SH to HHHH relative address
                        Case 16: T4 = " (" & SL & ")"       'p-izp (65c02)
                    End Select
                                    
                    '---- Handle inline comments
                    
                    If UComment > "" Then T5 = "; " & UComment                  'Use user comment string
                    
                    '=========================================================
                    ' PASS 2 - Output line in specified format
                    '=========================================================
                    
                    If ABlockFlag = False Then                                              '===================== Normal Format Output
                        Tblk = T3 & T4
                        Select Case OutFmt
                            Case 0: Tmp = T1 & T2 & Tblk                                    'addr: bytes cmd param
                            Case 1: Tmp = T1 & Tblk                                         'addr: cmd param
                            Case 2: Tmp = Format(LNum) & " " & Tblk                         'nnnn cmd param
                            Case 3: Tmp = Pad("", 16) & Tblk                                'cmd param
                            Case 4: Tmp = Pad(ALabel, 16) & Tblk                            'label cmd param
                        End Select
                                        
                        '------------------------------------------------------------------ Space after RTS/RTI option
                        If MD = 9 Then
                            If (T3 = "RTS") Or (T3 = "RTI") Then
                                If RTSOption = True Then
                                    If OutFmt = 2 Then
                                        LNum = LNum + LInc                                  'Next line number
                                        lstML.AddItem Format(LNum) & " ;"                   'ADD Line Number and semicolon <<<<<<<<<<<<<<<<<<<
                                    Else
                                        lstML.AddItem ";"                                   'ADD a blank line after RTI or RTS <<<<<<<<<<<<<<<
                                    End If
                                End If
                            End If
                        End If
                        
                    Else
                        '--- Handle "A" Block Output                                        '========================== "A" Block Output
                        Tmp2 = DOTBYTE & MyHex(B0A, -2)                                     'Add !BYTE" and HEX byte
                        If NB > 1 Then Tmp2 = Tmp2 & "," & MyHex(B1A, -2)                   'Add HEX byte
                        If NB > 2 Then Tmp2 = Tmp2 & "," & MyHex(B2A, -2)                   'Add HEX byte
                        Tmp2 = Pad(Tmp2, 20)                                                'Padd it out
                        
                        Tblk = ";" & MyHex(ABlockAddr, 4) & " " & T3 & T4                   'The Assembler Block Address
                        ABlockAddr = ABlockAddr + NB                                        'Increment it
                        
                        Select Case OutFmt
                            Case 0: Tmp = T1 & T2 & Tmp2 & Tblk                             'addr: HH HH HH !BYTE $HH,$HH,$HH ;addr: cmd param
                            Case 1: Tmp = T1 & Tmp2 & Tblk                                  'addr: !BYTE $HH,$HH,$HH          ;addr: cmd param
                            Case 2: Tmp = Format(LNum) & " " & Tmp2 & Tblk                  'nnnn !BYTE $HH,$HH,$HH ;cmd param
                            Case 3: Tmp = Pad("", 16) & Tmp2 & Tblk                         'bbbb !BYTE $HH,$HH,$HH ;cmd param
                            Case 4: Tmp = Pad(ALabel, 16) & Tmp2 & Tblk                     'label cmd param
                        End Select
                        
                    End If
                    
                    J = Len(Tmp) + 1: If CommentCol > J Then J = CommentCol                 'Position for comment
                    lstML.AddItem RTrim(Pad(Tmp, J) & T5)                                   'ADD IT to output <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                    
                End If
                
                '=========================================================
                ' PASS 1 and 2 - Advance Address and Line Number
                '=========================================================
                C = C + NB
                Address = Address + NB                                                      'increment address according to bytes used by opcode
                LNum = LNum + LInc                                                          'next line number
                ALabel = ""                                                                 'clear out label
                DoEvents                                                                    'added for long files
                
            End If
            
            If ABlockFlag = True Then
                If Address > DTEnd Then DTStart = 0                                         'Clear TABLE Start value
            End If
            
        Loop While C <= VLen
    Next Pass
    
    '=========================================================
    ' PASS 2 - Disassembly is complete
    '=========================================================

    lblEA.Caption = "Code from $" & StartAddress & " to $" & EndAddress & " (" & Format(C - 1) & " bytes)"
    lblEA.BackColor = &H8000000F
    
    If MLSplitFlag = True Then
        For J = 0 To lstML.ListCount
            lstML2.List(J) = lstML.List(J)                                  'Copy list item to second view
        Next J
    End If
    
    lstML.Visible = True                                                    'Show the first view
    If MLSplitFlag = True Then lstML2.Visible = True                           'Show the second view
    DoEvents
    
    If lstML.Visible = True Then lstML.SetFocus
    If GoodFlag = True Then lblGood.BackColor = vbGreen
    ViewerReady = True
    
End Sub

'---- ASM: Draw ML Viewer Side-panel elements
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

'---- ASM: Change Opcode Case
Private Sub cbOpUCase_Click()
    Dim i As Integer
    
    For i = 0 To 255
        If cbOpUCase.value = vbChecked Then
            OP(i) = UCase(OP(i))                    'Make uppercase
        Else
            OP(i) = LCase(OP(i))                    'Make lowercase
        End If
    Next i
    
End Sub

'---- ASM: Cancel Block Comment Edit
Private Sub cmdBCancel_Click()
    
    frBlock.Visible = False

End Sub

'---- ASM: Save Comment Block
Private Sub cmdSaveBlock_Click()
    Dim Tmp As String, V As Integer
   
    
    If lblCPos.Caption <> "" Then
        V = Val(lblCPos.Caption)
        lstCmnt.RemoveItem V                                                            'Remove the previous comment first
    End If
    
    Tmp = Replace(txtBlock.Text, Cr & LF, ";")                                          'Replace CR+LF with ";"
    lstCmnt.AddItem lblBlockAddress.Caption & ",[," & Tmp   'Add it
    frBlock.Visible = False
    
    MLReViewC
End Sub

'---- ASM: Re-Assemble using ACME
Private Sub cmdReassemble_Click()
    Dim Source As String, Dest As String, Tmp As String
    
    Source = FileBase(VFileName) & ".REASM"                                             'Use current file but with .reasm extension
    Dest = FileBase(VFileName) & ".out"                                                 'Use current file but with .out extension
    SaveASM Source                                                                      'Save the ASM listing
    KillFile Dest                                                                       'Delete the old output binary
    KillTemp                                                                            'Delete temp files
    
    DoEvents
    
    Tmp = " -o " & Quoted(Dest) & " " & Quoted(Source)
    'MsgBox Tmp
    
    If Exists(Source) Then
        frmMain.PubDoCommand CBMAcme, Tmp, "Assembling...", False                       'Use ACME to assemble the file
        If LastCMDError <> "" Then MsgBox LastCMDError                                  'Show errors
        
        If cbCompareOut.value = vbChecked Then                                          'Put results into HEX viewer
            If Exists(Dest) = True Then
                LoadCompare Dest                                                        'Load the binary file for comparison
                If SplitMode = False Then
                    EnableSplit                                                         'Enable SPLIT VIEW
                    ViewMode2 = 2                                                       'Enable HEX View
                    DrawVLayout                                                         'ReDraw Interface
                    cbCmpShow.value = vbChecked                                         'Enable Compare

                End If
                HEXView                                                                 'Refresh HEX view
            End If
        End If
        
        If Exists(TEMPFILE1) Then lblInfo.Caption = LastCMDError  'Reassembly did not complete
        
    Else
        MsgBox "The .REASM file was not created."
    End If
    
End Sub

'---- ASM: Viewer Project/Table Buttons
Private Sub lblTView_Click(Index As Integer)
    MLTabNum = Index
    DrawMLTabs
End Sub

'---- ASM: Jump to Specific Line# in listing
Private Sub lblLineNum_Click()
    Dim Tmp As String, N As Integer
        
    Tmp = InputBox("Jump to line#:", "Jump To", "")                                     'Get Line# from user
    
    N = Val(Tmp)                                                                        'Convert to numeric
    If N > 0 And N < lstML.ListCount Then
        lstML.ListIndex = N - 1                                                         'Move to the line
        lstML.Selected(N - 1) = True                                                    'Select the line
    End If
    
End Sub

'---- ASM: Single clicking on one of the Lists
Private Sub lstCmnt_Click()

    lblInfo.Caption = lstCmnt.List(lstCmnt.ListIndex)

End Sub

'---- ASM: Edit User Comment Table Entry
Private Sub lstCmnt_dblClick()
    Dim i As Integer, Tmp As String, Tmp2 As String, Tmp3 As String
    
    i = lstCmnt.ListIndex

    If i >= 0 Then
        Tmp = lstCmnt.List(i)
        
        If Mid(Tmp, 6, 1) = "[" Then                                                    '-- Block Comment
            Tmp2 = Mid(Tmp, 8)                                                          'Get the comment text
            Tmp3 = Replace(Tmp2, ";", Cr & LF)                                          'Replace ";" with CR+LF
            lblCPos.Caption = Format(i)                                                 'The Comment Position
            lblBlockAddress.Caption = Left(Tmp, 4)                                      'The Comment Address
            txtBlock.Text = Tmp3                                                        'The Comment Text
            frBlock.Visible = True                                                      'Make the Frame visible
            txtBlock.SetFocus
            DoEvents
        Else
            'Normal Comment
            Tmp2 = InputBox("Edit Comment:", "Edit Comment", Tmp)
            If Tmp2 > "" Then
                lstCmnt.RemoveItem i
                lstCmnt.AddItem Tmp2
                MLReViewC
            End If
        End If
    End If
    
End Sub

'---- ASM: Click on Entry Point List
Private Sub lstEntryPt_Click()
    lblInfo.Caption = lstEntryPt.List(lstEntryPt.ListIndex)
End Sub

'---- ASM: Click on External JSR List
Private Sub lstJSR_Click()
    lblInfo.Caption = lstJSR.List(lstJSR.ListIndex)
End Sub

'---- ASM: Click on Labels List
Private Sub lstLabels_Click()
    lblInfo.Caption = lstLabels.List(lstLabels.ListIndex)
End Sub

'---- ASM: Double-Click Labels
Private Sub lstLabels_DblClick()
    Dim Tmp As String, Tmp2 As String
    
    Tmp = lstLabels.List(lstLabels.ListIndex) & ",name,-"              'Make default text entry string
    Tmp2 = InputBox("HHHH,LABELNAME,DESCRIPTION", "Add Label from [GEN] label", Tmp)
    If Len(Tmp2) > 12 Then lstULabels.AddItem Tmp2: MLReViewA

End Sub

'---- ASM: Click to Start a Trace
Private Sub cmdTrace_Click()
    If lstEntryPt.ListCount = 0 Then MyMsg "You must add Entrypoints first!": Exit Sub
    TraceIt
End Sub

'---- ASM: Flow Tracing
Private Sub TraceIt()

    Dim i As Long, J As Integer                                 'Loop counters
    Dim Tmp As String                                           'Temp strings
    Dim PC  As Long, EA As Long                                 'Program Counter, Effective Address
    Dim CodeOffset As Long                                      'For calculating position in buffer
    Dim StartAddr As Long, EndAddr As Long                      'Start and end addresses of code range
    Dim TargetAddr As Long, TargetH As String                   'Target address dec and hex
    
    Dim Address As Long
    Dim OpByte As Integer                                       'Opcode byte            (0 to 255)
    Dim OpDef  As String                                        'Opcode definition      (ie: BRKi = BRK immediate)
    Dim OpStr  As String                                        'Opcode Mnuemonic       (ie: BRK)
    Dim OpMode As Integer                                       'Opcode Addressing Mode (ie: immediate)
    Dim OpLen As Integer                                        'Opcode Length
    Dim StopFlag As Boolean                                     'Flag to indicate flow is stopped (end of branch)
    Dim RangeS As Integer, RangeE As Integer                    'Range boundaries
    Dim VFlag As Boolean                                        'Verbose Tracer Output Flag
    
    ReDim Addr(32767) As Boolean                                'This is the address space (FALSE=data, TRUE=code)
    
    '-- Initialize
    
    lstML.Clear
    lstEP.Clear
    
    If cbVerb.value = vbChecked Then VFlag = True               'Verbose flag
    
    For i = 0 To 32767: Addr(i) = False: Next                   'Mark entire address space as data

    Address = VLA                                               'Assume Load Addres included
    If cbLA.value = Checked Then Address = MyDec(txtLA.Text)    'User Specified Address
    
    CodeOffset = Address - 1                                    'Offset between First byte in buffer and it's address
    'If cbLA.value = Checked Then CodeOffset = CodeOffset + 2    'Adjust for Load address bytes
    
    StartAddr = Address
    EndAddr = Address + VLen
    'If cbLA.value = Checked Then EndAddr = EndAddr - 2          'Adjust for Load address bytes
    
    PC = StartAddr:  EA = 1                                     'Program Counter
    
    lstML.AddItem "Target CPU: " & OpDesc
    lstML.AddItem "Jumps.....: " & OpJ
    lstML.AddItem "Branches..: " & OpB
    lstML.AddItem "Stops.....: " & OpZ

    lstML.AddItem "Start Address: " & MyHex(StartAddr, -4)
    lstML.AddItem "End Address..: " & MyHex(EndAddr, -4)
    lstML.AddItem "Code Bytes...: " & Str(EndAddr - StartAddr + 1)

    lstML.AddItem "Loading entry points..."
    
    For i = 0 To lstEntryPt.ListCount - 1
        lstEP.AddItem Left(lstEntryPt.List(i), 4)               'Add Hex adress to tracer list
    Next i
    
    cmdAddTables.Visible = False                                'Reset visibility
       
    lstML.AddItem "Starting Trace..."
        
    StopFlag = True                                             'Set to True so first EP is removed from list
    
    '===============
    ' Start of Trace
    '===============
    Do
        If StopFlag = True Then                                 '-- Are we STOPPED?
            i = lstEP.ListCount - 1                             'How many entry points?
            If i < 0 Then lstML.AddItem "Finished trace!": Exit Do
            
            Tmp = Left(lstEP.List(i), 4)                        'Read new EntryPoint (hex format HHHH)
            PC = MyDec(Tmp)                                     'Set Program Counter address (convert to decimal)
            If PC > EndAddr Then
                lstML.AddItem "Entrypoint out of range!"        'Oops, outside range! Abort
                Exit Do
            End If
            
            lstML.AddItem "Tracing from $" & Tmp & " (" & Format(PC) & ")"
            lstEP.RemoveItem i                                  'Remove EntryPoint from bottom of list
            
            StopFlag = False
            DoEvents
        End If
                
        '-- Get instruction
        EA = PC - CodeOffset                                    'Calculate buffer pos
        'lstML.AddItem "EA=" & Str(EA) & " PC=" & Str(PC)
        
        If Addr(EA) = False Then
            '-------------------------------------------------- These bytes have not been processed yet
            OpByte = Asc(Mid(VBuf, EA, 1))                      'Get the opcode
            OpDef = OP(OpByte)                                  'Opcode definition (OP array is global and should be loaded at ASM init)
            OpStr = UCase(Left(OpDef, Len(OpDef) - 1))          'Opcode string
            OpMode = Asc(Right(UCase(OpDef), 1)) - 64           'Opcode addressing mode (a-z)
            OpLen = Val(Mid(OpModeLen, OpMode, 1))              'Get instruction length
            Addr(EA) = True                                     'Mark byte as code (opcode)
            
            If VFlag = True Then lstML.AddItem Str(EA) & " " & OpStr
            
            If OpLen > 1 Then
                If EA + 1 > VLen Then
                    TargetAddr = PC
                Else
                    OpByte = Asc(Mid(VBuf, EA + 1, 1))              'Get the first parameter byte
                    Addr(EA + 1) = True                             'Mark it as code
                    If OpMode = 10 Then
                        If OpByte > 127 Then OpByte = OpByte - 256  'Calculate backwards branch
                        TargetAddr = PC + OpByte + 2                'Relative offset for branch
                    Else
                        TargetAddr = OpByte                         'LO byte of target
                    End If
                End If
            End If
            
            If OpLen = 3 Then
                If EA + 2 > VLen Then
                    TargetAddr = PC
                Else
                    OpByte = Asc(Mid(VBuf, EA + 2, 1))              'Get the second parameter byte 'zzzzzzzzzzzzzzzzzzzzzzz
                    Addr(EA + 2) = True                             'Mark it as code
                    TargetAddr = TargetAddr + 256& * OpByte         'HI byte of target gets combined with LO
                End If
            End If
            
            TargetH = MyHex(TargetAddr, 4)
            
            '-- Is opcode a flow change?
            If InStr(1, OpJ, OpStr) > 0 Then
                lstML.AddItem "Flow change at $" & TargetH
                If (TargetAddr >= StartAddr) And (TargetAddr <= EndAddr) Then
                    lstEP.AddItem TargetH                       'Add target to EP then STOP if inside code range
                    If VFlag = True Then lstML.AddItem "Adding " & TargetH & " to EP list."
                End If
                StopFlag = True                                 'STOP
            End If
            
            '-- Is opcode a flow split?
            If InStr(1, OpB, OpStr) > 0 Then
                lstML.AddItem "found flow split at $" & TargetH
                If (TargetAddr >= StartAddr) And (TargetAddr <= EndAddr) Then
                    lstEP.AddItem TargetH                       'Add target to EP if inside code range
                    If VFlag = True Then lstML.AddItem "Adding " & TargetH & " to EP list."
                End If
                
                Tmp = MyHex(PC + OpLen, 4)
                lstEP.AddItem Tmp                               'Add next opcode to EP
                If VFlag = True Then lstML.AddItem "Adding " & Tmp & " to EP list."
                
                StopFlag = True                                 'STOP
            End If
            
            '-- Is opcode a flow stop?
            If InStr(1, OpZ, OpStr) > 0 Then
                lstML.AddItem "Flow stop at" & Str(PC) & " - " & OpZ
                StopFlag = True                                 'STOP
            End If
        Else
            If VFlag = True Then lstML.AddItem "Found marked code - skipping."
            StopFlag = True                                     'Instruction already marked as code, treat as STOP
        End If
                
        PC = PC + OpLen                                         'Increment Pc
        EA = EA + OpLen                                         'Increment buffer pointer
        
        If PC > EndAddr Then
            StopFlag = True                                     'We hit the end of the file. Don't exit - there might be more Entry Points!
            lstML.AddItem "Hit end of code. Abnormal trace end."
        End If
        
    Loop
    
    '-- Completed Trace
    StopFlag = False                                            'Flag to indicate we are in a data area
    lstEP.Clear                                                 'List should be empty but clear just in case
    
    lstML.AddItem "Building data table list..."
    
    'PC = 1: If cbLA.value = Checked Then PC = 3                 'Skip load address bytes
    
    For i = 1 To VLen
        If Addr(i) = False Then
            If StopFlag = False Then RangeS = i: StopFlag = True
        Else
            If StopFlag = True Then
                RangeE = i - 1: StopFlag = False                'If we are in a data range then this is the end
                lstEP.AddItem MyHex(RangeS + CodeOffset, 4) & "," & MyHex(RangeE + CodeOffset, 4) & ",b,trace data" 'Add it to the list
                lstML.AddItem MyHex(RangeS + CodeOffset, 4) & "," & MyHex(RangeE + CodeOffset, 4) & ",b,trace data" 'Add it to the list
            End If
        End If
    Next i
    
    '-- Handle data block at the end of the code range
    If StopFlag = True Then
        RangeE = i - 1: StopFlag = False                        'If we are in a data range then this is the end
        lstEP.AddItem MyHex(RangeS + CodeOffset, 4) & "," & MyHex(RangeE + CodeOffset, 4) & ",b,trace data" 'Add it to the list
    End If
    
    '-- Save Log
    If cbTraceLog.value = vbChecked Then SaveTraceLog
    
    lstML.AddItem "Done! Click Add Tables if results look correct. Refresh to see results."
    
    If lstEP.ListCount > 0 Then cmdAddTables.Visible = True     'Make ADD button visible
End Sub

'---- Save Trace Log
Private Sub SaveTraceLog()
    Dim FIO As Integer, Filename As String, Tmp As String
    Dim J As Integer
   
    Filename = FileBase(VFileName) & ".tracelog"                        'Use name of view file
    
    FIO = FreeFile
    Open Filename For Output As FIO
    For J = 0 To lstML.ListCount - 1
        Print #FIO, lstML.List(J)
    Next J
    
    Close FIO
 
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

'---- ASM: Prompt to Save ASM Ouput to file
Private Sub cmdSaveASM_Click()
    Dim J As Integer, Filename As String, FIO As Integer
    
    Filename = FileOpenSave(FileBase(VFileName), 1, 4, "Save ASM code")
    If Filename = "" Then Exit Sub
    If Overwrite(Filename) = False Then Exit Sub
    SaveASM Filename
    
End Sub

'---- ASM: Save Output to specified File
Private Sub SaveASM(ByVal Filename As String)
    Dim FIO As Integer, J As Integer
    
    FIO = FreeFile
    Open Filename For Output As FIO
    
    For J = 0 To lstML.ListCount - 1
        Print #FIO, lstML.List(J)
    Next J
    Close FIO
    
End Sub

'---- ASM: Copy ouput to Clipboard
Private Sub cmdCopyClip2_Click()
    Dim J As Integer, Tmp As String
    
    For J = 0 To lstML.ListCount - 1
        Tmp = Tmp & lstML.List(J) & vbCrLf
    Next J
    
    Clipboard.Clear
    Clipboard.SetText Tmp

End Sub

'---- Set Change Flag to true and update Status
Sub SetChangeFlag()
    
    ChangeFlag = True
    ShowMLChange

End Sub

'---- Set Change Falg to true and update Status
Sub ClearChangeFlag()
    
    ChangeFlag = False
    ShowMLChange

End Sub

'---- Show Project Changed Status
Sub ShowMLChange()
    
    If frML.Visible = True Then
        If ProjFilename = "" Then
            lblChanged.BackColor = RGB(63, 63, 64)
        Else
            If ChangeFlag = True Then lblChanged.BackColor = vbRed Else lblChanged.BackColor = vbGreen
        End If
    End If
    
    DoEvents

End Sub

'---- Re-Views the file when options have changed but only if AutoRefresh is true
Sub MLReViewA()
    
    If cbAuto.value = vbChecked Then MLReView
    
    ShowMLChange

End Sub

'---- ASM: Views the file as above AND also sets the Changes Flag=True
Sub MLReViewC()

    SetChangeFlag
    
    If cbAuto.value = vbChecked Then MLReView

End Sub

'---- ASM: Re-Views the file when options have changed
Sub MLReView()
    Dim TopPos As Integer
    
    TopPos = lstML.TopIndex                                 'Remember the position
    
    If ViewerReady = True Then
        MLView
        ShowMLChange                                        'ML Project status
        If TopPos > lstML.ListCount Then TopPos = 0         'FIX: Large data block additions can make TopPos be past end
        lstML.TopIndex = TopPos                             'Restore the position
    End If

End Sub

'---- ASM: Jump to selected entry in currently visible table
Private Sub cmdSYMGoto_Click()
    Dim i As Integer, Tmp As String
        
    Select Case MLTabNum
        Case 2
            If lstEntryPt.ListCount = -1 Then Exit Sub      'No Entries, so exit
            i = lstEntryPt.ListIndex: If i < 0 Then i = 0   'Get index
            Tmp = Left(lstEntryPt.List(i), 4)               'Extract HEX Addres
    
        Case 3
            If lstSYM.ListCount = -1 Then Exit Sub          'No Entries, so exit
            i = lstSYM.ListIndex: If i < 0 Then i = 0       'Get index
            Tmp = GetField(lstSYM.List(i), 2)               'Get 2nd field = Symbol
        
        Case 4
            If lstDT.ListCount = -1 Then Exit Sub           'No Entries, so exit
            i = lstDT.ListIndex: If i < 0 Then i = 0        'Get Index
            Tmp = Left(lstDT.List(i), 4)                    'Extract HEX Address
            
        Case 5
            If lstULabels.ListCount = -1 Then Exit Sub      'No Entries, so exit
            i = lstULabels.ListIndex: If i < 0 Then i = 0   'Get Index
            Tmp = Left(lstULabels.List(i), 4)               'Extract HEX address
            
        Case 6
            If lstCmnt.ListCount = -1 Then Exit Sub         'No Entries, so exit
            i = lstCmnt.ListIndex: If i < 0 Then i = 0      'Get Index
            Tmp = Left(lstCmnt.List(i), 4)                  'Extract HEX Address
            
        Case 7
            If lstLabels.ListCount = -1 Then Exit Sub       'No Entries, so exit
            i = lstLabels.ListIndex: If i < 0 Then i = 0    'Get Index
            Tmp = Left(lstLabels.List(i), 4)                'Extract HEX Address
            
        Case 8
            If lstJSR.ListCount = -1 Then Exit Sub          'No Entries, so exit
            i = lstJSR.ListIndex: If i < 0 Then i = 0       'Get Index
            Tmp = Left(lstJSR.List(i), 4)                   'Extract HEX Address
    
    End Select
    
    JumpList Tmp, 1, False                                  'Find next match from current down, only hightlight 1

End Sub

'---- ASM: Find and jump to the next undefined opcode
Private Sub lblGood_Click()

    JumpList "???", 0, False

End Sub

'---- ASM: Find specified string
Private Sub cmdFind_Click()
    Dim Tmp As String
    
    Tmp = InputBox("Enter String to find:", "Find")
    If Tmp <> "" Then
        cmdFindAll.ToolTipText = ""
        JumpList Tmp, 0, False
    End If
    
End Sub

'---- ASM: Find ALL occurances of last search string
Private Sub cmdFindAll_Click()

    JumpList "", 0, True                                    'Find from TOP, highlight all lines

End Sub

'---- ASM: Jump to next occurance of search string
Private Sub cmdNext_Click(Index As Integer)
    
    JumpList "", Index, False

End Sub

'---- ASM: Search for string
' Blank string searches with same string
' MODE - Search method: 0=Top Down, 1=Current Down, 2=Current UP, 3=Bottom UP
' FLAG - TRUE = ALL matches
Sub JumpList(ByVal Txt As String, Mode As Integer, ByVal Flag As Boolean)
    Static LastTxt As String, Count As Integer 'These values are retained between calls
    
    Dim i As Integer, J As Integer, Max As Integer, Direction As Integer
    Dim Tmp As String
    
    If Txt = "" Then Txt = LastTxt
    If Txt = "" Then Exit Sub
    
    Max = lstML.ListCount - 1                           'Max entries
    Count = 0
    i = lstML.ListIndex                                 'Assume current position
    
    Select Case Mode
        Case 1: Direction = 1:      Tmp = "Down"
        Case 2: Direction = -1:     Tmp = "Up"
        Case 3: Direction = -1:     Tmp = "Bottom Up": i = lstML.ListCount - 1 'Start at END
        Case Else: Direction = 1:   Tmp = "Top Down": i = 0           'Start at TOP
    End Select
    
    lblInfo.Caption = "Search (" & Tmp & "): " & Txt
   
    Do
        i = i + Direction: If (i < 0) Or (i > Max) Then Exit Do
        If InStr(1, lstML.List(i), Txt, vbTextCompare) > 0 Then
            lstML.Selected(i) = True                                'Hilight it
            Count = Count + 1                                       'Count it
            
            If Flag = False Then
                J = i - 5: If J < 0 Or J > Max Then J = i
                lstML.TopIndex = J                                  'Move top of list to near found line
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

'---- ASM: Find an Address in the following order: SYMBOL, ULABEL, LABEL.
' Return SYMBOL name, ULABEL name, or "L_xxxx" LABEL string
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

'---- ASM: Find a User Label or Generated Label in the following order: ULABEL, LABEL.
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

'---- ASM: Lookup SYMBOL and return string. Also Set LastComment
' FORMAT of SYMBOL list entry: HHHH,symbolstring,comment
Private Function FindSym(ByVal Addr As String) As String
    Dim Tmp As String, Tmp2 As String, Tmp3 As String
    Dim R1 As Integer, R2 As Integer, R3 As Integer                         'binary search range
        
    R3 = lstSYM.ListCount - 1                                               'Range End position
    If R3 < 0 Then Exit Function                                            'Exit if no entries
    R1 = 0                                                                  'Range Start position
    LastComment = ""                                                        'Clear Last Comment string
    LastSymPos = 0                                                          'Clear Last SYM position
    
    Do
        R2 = (R1 + R3) \ 2                                                  'Calculate middle of range
        Tmp = lstSYM.List(R2)                                               'Check array at middle position
        Tmp2 = Left(Tmp, 4)                                                 'Extract address part
        
        If Tmp2 = Addr Then
            Tmp3 = GetField(Tmp, 2)
            If Tmp3 = "" Then Tmp3 = DOTHEX & Addr                          'If not symbol then just use address
            FindSym = MyTrim(Tmp3)                                          'Return string
            LastComment = GetField(Tmp, 3)                                  'Get the comment
            LastSymPos = R2                                                 'Remember it's position
            lstSYM.Selected(R2) = True
            Exit Do
        Else
            If Tmp2 > Addr Then R3 = R2 - 1 Else R1 = R2 + 1                'Adjust range end points depending on comparison
        End If
        If R1 > R3 Then FindSym = "": Exit Do                               'No more in range, so exit with NULL string
    Loop

End Function

'---- ASM: Find ULABEL Address and return Symbol string
' FORMAT of ULABEL list entry: HHHH,symbolstring
Private Function FindULabel(ByVal Addr As String) As String
    Dim Tmp As String, Tmp2 As String
    Dim R1 As Integer, R2 As Integer, R3 As Integer                         'binary search range
        
    R3 = lstULabels.ListCount - 1                                           'Range End position
    If R3 < 0 Then Exit Function                                            'Exit if no entries
    R1 = 0                                                                  'Range Start position
        
    Do
        R2 = (R1 + R3) \ 2                                                  'Calculate middle of range
        Tmp = lstULabels.List(R2)                                           'Check array at middle position
        Tmp2 = Left(Tmp, 4)                                                 'Extract Address
        
        If Tmp2 = Addr Then
            FindULabel = GetField(Tmp, 2)                                   'Substitute label
            Exit Do
        Else
            If Tmp2 > Addr Then R3 = R2 - 1 Else R1 = R2 + 1                'Adjust range end points depending on comparison
        End If
        If R1 > R3 Then FindULabel = "": Exit Do                            'No more in range, so exit with NULL string
        DoEvents
    Loop

End Function

'---- ASM: Find LABEL Address and return Address string
' FORMAT of LABEL list entry: HHHH
Private Function FindLabel(ByVal Addr As String) As String
    Dim Tmp As String, Tmp2 As String
    Dim R1 As Integer, R2 As Integer, R3 As Integer                         'Binary search range
        
    R3 = lstLabels.ListCount - 1                                            'Range End position
    If R3 < 0 Then Exit Function                                            'Exit if no entries
    R1 = 0                                                                  'Range Start position

    Do
        R2 = (R1 + R3) \ 2                                                  'Calculate middle of range
        Tmp = lstLabels.List(R2)                                            'Check array at middle position
        Tmp2 = Left(Tmp, 4)                                                 'Extract Address
                
        If Tmp2 = Addr Then
            FindLabel = Tmp2                                                'Return the label
            Exit Do
        Else
            If Tmp2 > Addr Then R3 = R2 - 1 Else R1 = R2 + 1                'Adjust range end points depending on comparison
        End If
        If R1 > R3 Then FindLabel = "": Exit Do                             'No more in range, so exit with NULL string
        DoEvents
    Loop

End Function

'---- ASM: Lookup comment for specified address and return "type,commentstring"
' FORMAT of COMMENT list entry: HHHH,type,commentstring
Private Function FindComment(ByVal Addr As String) As String
    Dim Tmp As String, Tmp2 As String
    Dim R1 As Integer, R2 As Integer, R3 As Integer 'binary search range
        
    R3 = lstCmnt.ListCount - 1                                              'Range End position
    If R3 < 0 Then Exit Function                                            'Exit if no entries
    R1 = 0                                                                  'Range Start position
    
    Do
        R2 = (R1 + R3) \ 2                                                  'Calculate middle of range
        Tmp = lstCmnt.List(R2)                                              'Check array at middle position
        Tmp2 = Left(Tmp, 4)                                                 'Extract Address
        
        If Tmp2 = Addr Then
            FindComment = Mid(Tmp, 6)                                       'Return the type and commentstring
            Exit Do
        Else
            If Tmp2 > Addr Then R3 = R2 - 1 Else R1 = R2 + 1                'Adjust range end points depending on comparison
        End If
        If R1 > R3 Then FindComment = "": Exit Do                           'No more in range, so exit with NULL string
    Loop

End Function

'---- ASM: Quick Add Label
Private Sub cmdAddLabel_Click()
    Dim RS As String, Tmp As String, Tmp2 As String, i As Integer
    
    Tmp = "Please select a line with an address first!"
    
    i = lstML.ListIndex: If i < 0 Then MyMsg Tmp: Exit Sub                 'Ooops, no line selected!
    RS = ExtractAddr(lstML.List(i)): If RS = "" Then MyMsg Tmp: Exit Sub   'Ooops, line didn't have an address!
 
    Tmp2 = InputBox("Add LABEL at " & RS & Cr & Cr & "Enter LABEL Name:", "Add Label", "")
    If Tmp2 > "" Then lstULabels.AddItem RS & "," & Tmp2: MLReViewC
    
End Sub

'---- ASM: Quick Add Entry Point
Private Sub cmdAddEP_Click()
    Dim RS As String, Tmp As String, Tmp2 As String, i As Integer
    
    Tmp = "Please select a line with an address first!"
    
    i = lstML.ListIndex: If i < 0 Then MyMsg Tmp: Exit Sub                 'Ooops, no line selected!
    RS = ExtractAddr(lstML.List(i)): If RS = "" Then MyMsg Tmp: Exit Sub   'Ooops, line didn't have an address!
 
    Tmp2 = InputBox("Add ENTRY POINT at " & RS & Cr & Cr & "Enter ENTRY POINT Name:", "Add Entry Point", "")
    If Tmp2 > "" Then lstEntryPt.AddItem RS & "," & Tmp2: MLReViewC
    
End Sub

'---- ASM: Quick Add Comment / Separator ( ;C / C / -C- / =C= / - / = )
Private Sub cmdAddComment_Click(Index As Integer)
    Dim RS As String, Tmp As String, Tmp2 As String, i As Integer
    
    Tmp = "Please select a line with an address first!"
    
    i = lstML.ListIndex: If i < 0 Then MyMsg Tmp: Exit Sub     'Oops, no line selected!
    RS = ExtractAddr(lstML.List(i)): If RS = "" Then MyMsg Tmp: Exit Sub        'Oops, line didn't have an address!
        
    Tmp = Mid("is-=*-=*[", Index + 1, 1): Tmp2 = ""
        
    Select Case Index '---- 0 to 4 need a Comment, 5 to 7 are dividers, 8 is BLOCK
        Case 0 To 4                                                             '-- Regular Comment
            Tmp2 = InputBox("Enter a comment at position " & RS & ":", "Enter Comment", "")
            If Tmp2 = "" Then Exit Sub
            
        Case 8                                                                  '-- Block Comment
            lblBlockAddress.Caption = RS                                        'Remember/Display Address
            lblCPos.Caption = ""                                                'Clear POSition index for add
            txtBlock.Text = ""                                                  'Clear the Comment
            frBlock.Visible = True: Exit Sub                                    'Make it visible
    End Select
    
    lstCmnt.AddItem RS & "," & Tmp & "," & Tmp2                                 'Add it
    
    MLReViewC
    
End Sub

'---- ASM: Quick Add Data Table (DHSRVW)
Private Sub cmdDTAdd_Click(Index As Integer)
    Dim Tmp As String, Tmp2 As String
    Dim i As Integer, P As Integer, RS As String, RE As String
    Dim Flag As Boolean
    
    Flag = False
    
    'Check if there is a range selected
    For i = 0 To lstML.ListCount - 1
        If lstML.Selected(i) = True Then
            If Flag = False Then RS = ExtractAddr(lstML.List(i)): Flag = True   'Found first selected line
            P = i                                                               'remember it
        Else
            If Flag = True Then RE = ExtractAddr(lstML.List(P)): Exit For       'Not selected so use last remembered line for end
        End If
    Next i
         
    If Flag = True Then
        If RE = "" Then RE = RS
        Select Case Index 'DHSRVWXZAL
            Case 0: Tmp = "D": Tmp2 = "Decimal Byte Table"
            Case 1: Tmp = "H": Tmp2 = "Hex Byte Table"
            Case 2: Tmp = "S": Tmp2 = "Text/String Table"
            Case 3: Tmp = "R": Tmp2 = "RTS Address Table (Generates Labels)"
            Case 4: Tmp = "V": Tmp2 = "Address Table (Generates Labels)"
            Case 5: Tmp = "W": Tmp2 = "Big Endian Word Table"
            Case 6: Tmp = "X": Tmp2 = "Hidden Table"
            Case 7: Tmp = "Z": Tmp2 = "Binary Byte Table"
            Case 8: Tmp = "A": Tmp2 = "Assembly Block"
            Case 9: Tmp = "L": Tmp2 = "Little-endian Word Table"
        End Select
                   
        Tmp2 = InputBox("Type : " & Tmp2 & Cr & "Range: " & RS & " to " & RE & Cr & Cr & "Enter a description:", "Add Table", "")
        
        If Tmp2 <> "" Then
            lstDT.AddItem RS & "," & RE & "," & Tmp & "," & Tmp2                'Add it
            lstDT.Selected(lstDT.NewIndex) = True                               'Make it selected
            MLReViewC
        End If
    Else
        MyMsg "Please select a range first!"
    End If
    
End Sub

'---- ASM: Display Data Table Info when clicked
Private Sub lstDT_Click()

    lblInfo.Caption = lstDT.List(lstDT.ListIndex)

End Sub

'---- ASM: Edit a Data Table Entry
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

'---- ASM: Handle clicking on an entry in the listing.
' If BINARY viewer is visible then try to synchronize address
Private Sub lstML_Click()
    Dim Tmp As String, Tmp2 As String, Tmp3 As String, Addr As String
    Dim R1 As Integer, R2 As Integer, R3 As Integer 'binary search range
    
    lblLineNum.Caption = lstML.ListIndex + 1                                'Show the line# of selected line
    
    If frBIN.Visible = True Then                                            'DualView with Hex visible - Try to find matching hex listing line
        Addr = Left(lstML.List(lstML.ListIndex), 4)                         'Address in ASM listing
        If Len(Addr) <> 4 Then Exit Sub
        Tmp = Right(Addr, 1): Tmp2 = "0"                                    'Last digit and replacement default
        If cbWide.value = vbUnchecked Then                                  'Handle wide listing
            If Tmp < "8" Then Tmp2 = "0" Else Tmp2 = "8"
        End If
        
        Mid(Addr, 4, 1) = Tmp2                                              'Replace the last digit
        
        R3 = lstView(2).ListCount - 1                                       'Range End position
        If R3 < 0 Then Exit Sub                                             'Exit if no entries
        R1 = 0                                                              'Range Start position
        
        Do
            R2 = (R1 + R3) \ 2                                              'Calculate middle of range
            Tmp = lstView(2).List(R2)                                       'Check array at middle position
            Tmp2 = Left(Tmp, 4)                                             'Extract address part
            
            If Tmp2 = Addr Then
                lstView(2).ListIndex = R2: Exit Do                          'Highlight the BIN line
            Else
                If Tmp2 > Addr Then R3 = R2 - 1 Else R1 = R2 + 1            'Adjust range end points depending on comparison
            End If
            If R1 > R3 Then Exit Do                                         'No more in range, so exit with NULL string
        Loop
    End If
        
End Sub

'---- ASM: Add a Symbol by Doubleclick of ML line
Private Sub lstML_DblClick()

    Call cmdSymAdd_Click

End Sub

'---- ASM: Display Symbol Info when clicked on
Private Sub lstSYM_Click()
    
    lblInfo.Caption = lstSYM.List(lstSYM.ListIndex)

End Sub

'---- ASM: Edit Symbol Table Entry
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

'---- ASM: Disply Label Info when clicked
Private Sub lstULabels_Click()

    lblInfo.Caption = lstULabels.List(lstULabels.ListIndex)

End Sub

'---- ASM: Edit User Label Table Entry
Private Sub lstULabels_dblClick()
    Dim i As Integer, Tmp As String, Tmp2 As String
    
    i = lstULabels.ListIndex
    If i >= 0 Then
        Tmp = lstULabels.List(i)
        If Mid(Tmp, 6, 1) = "[" Then
            lblBlockAddress.Caption = Left(Tmp, 4)          'Remember the address
            txtBlock.Text = Mid(Tmp, 8)
            frBlock.Visible = True
        Else
            Tmp2 = InputBox("Edit Label:", "Edit Label", Tmp)
            If Tmp2 > "" Then
                lstULabels.RemoveItem i
                lstULabels.AddItem Tmp2
                MLReViewC
            End If
        End If
    End If
    
End Sub

'---- ASM: Toggle displaying of Data and Symbol Table frames
Private Sub lblShw_Click()
    
    ShowTables = Not ShowTables
    DrawVLayout

End Sub

'---- ASM: Toggle Info frame
Private Sub imgShowInfo_Click()
    
    InfoFlag = Not InfoFlag
    DrawVLayout                                                                 'Draw the new View layout

End Sub

'---- ASM: Toggle Split ML View
Private Sub cmdMLSplit_Click()
    
    MLSplitFlag = Not MLSplitFlag                                               'Toggle it
    
    If MLSplitFlag = True Then
        If lstML2.ListCount = 0 Then MLReView                                   'Only refesh if second view is empty
    End If
    
    DrawVLayout                                                                 'Draw the new View layout

End Sub

'---- ASM: Prompt to Save Symbol Table to file
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

'---- ASM: Prompt for Loading a new Symbol Table File
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
    MLReViewA                                                               'Re-Display Listing
    
End Sub

'---- ASM: Process selection of a new Platform from the list
Private Sub cboPlatform_Click()
    Dim Filename As String, i As Integer
    
    If MLCFlag = False Then Exit Sub                                        'Exit if ???? Flag is false
    If ViewerReady = False Then Exit Sub                                    'Exit if BUSY
    
    i = cboPlatform.ListIndex: If i = 0 Then Exit Sub                       'Exit if no Platforms
    
    Filename = ExeDir & cboPlatFile.List(i)                                 'Get name of Platform file
    
    If Exists(Filename) = False Then MyMsg "Sorry, Platform file not found! " & Filename: Exit Sub
    If OverwriteProject = True Then LoadSymFile Filename, 3                 'Load Platform file (symbols)
    
    MLReView                                                                'Re-Display Listing
    
End Sub

'---- ASM: Process selection of a new CPU from the list
Private Sub cboCPU_Click()
    Dim Filename As String
    
    If MLCFlag = False Then Exit Sub                                        'Exit if ???? Flag is false
    If ViewerReady = False Then Exit Sub                                    'Exit if BUSY
    
    Filename = ExeDir & cboCPUFile.List(cboCPU.ListIndex)                   'Get filename from dropdown list
    If Exists(Filename) = False Then MyMsg "Sorry, CPU file not found! " & Filename: Exit Sub
    
    LoadOpcodes Filename                                                    'Load OpCodes file
    MLReView                                                                'Re-display listing

End Sub

'---- ASM: Check Project Changed status and Prompt for Saving Project if there is a change
' Returns TRUE if:
'   - project has not changed, or there is no project file
'   - project has changed and YES or NO is selected. If YES is selected then project will be saved first
' Returns FALSE if CANCEL is selected
Private Function OverwriteProject() As Boolean
    Dim Result As VbMsgBoxResult
    
    OverwriteProject = False                                                'Assume NOT ok to continue
    
    If (ProjFilename <> "") And (ChangeFlag = True) Then
        Result = MsgBox("Project has changed. Save Changes first?", vbYesNoCancel)
        If Result = vbCancel Then Exit Function                             'Exit if CANCEL clicked
        If Result = vbYes Then SaveProjFile True                            'ProjFilename 'YES=save project, NO=loose changes
    End If
    
    OverwriteProject = True                                                 'Return FLAG
    
End Function

'---- ASM: Prompt for project filename to load
Private Sub cmdProjLoad_Click()
    Dim Filename As String
    
    If OverwriteProject = True Then
        Filename = FileOpenSave("", 0, 2, "Load ASM Project File"): If Filename = "" Then Exit Sub
        LoadProjFile Filename
        MLReView                                                           'Re-Display Listing
    End If
End Sub

'---- ASM: Load specified Project File
' A Project file contains lines to be loaded into the tabels
' Each table group must be proceeded by a selection marker:
' [SYMBOLS] [TABLES] [LABELS] [COMMENTS]

Public Sub LoadProjFile(ByVal Filename As String)
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
                            Case "UPDATED": lblUpdated.Caption = VStr
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
        ViewerReady = False
        cbLA.value = vbUnchecked
        txtLA.Text = LA                                                     'Use Load address from Project file
        DoEvents
        ViewerReady = True                                                  'Set READY
    End If
    
    ShowTables = True                                                       'Enable Tables Panel
    ClearChangeFlag                                                         'Clear Change Flag
    
End Sub

'---- ASM: Prompt for Filename then save Project
Private Sub cmdProjSaveAs_Click()
    SaveProjFile False
End Sub

'---- ASM: Save to Existing Project
Private Sub cmdProjSave_Click()
    SaveProjFile True
End Sub

'---- ASM: Save specified Project File
' Set Flag to TRUE  to Save with Current project filename
' Set Flag to FALSE to Prompt for filename first
'
' A Project file contains lines to be loaded into the tabels
' Each table group must be proceeded by a selection marker:
' [SYMBOLS] [TABLES] [LABELS] [COMMENTS]

Private Sub SaveProjFile(ByVal Flag As Boolean)  'Filename As String)
    Dim FIO As Integer, Filename As String, Tmp As String
    Dim J As Integer, TMode As Integer
   
    Tmp = ProjFilename
    If Tmp = "" Then Tmp = FileBase(VFileName)                              'Use Project Filename as default, otherwise use name of view file
    
    If Flag = True Then
        Filename = Tmp
    Else
        Filename = FileOpenSave(Tmp, 1, 2, "Save ASM Project File")         'Prompt for Project filename
        If Filename = "" Then Exit Sub                                      'Exit if not supplied
        If Overwrite(Filename) = False Then Exit Sub                        'Exit if not overwritting existing project file
    End If
        
    FIO = FreeFile
    Open Filename For Output As FIO
    TMode = 0
    
    '-- [PROJECT]
    Print #FIO, "[PROJECT]"
    Print #FIO, "UPDATED="; Date & ", " & Time                              'Write last update
    Print #FIO, "LA="; txtLA.Text                                           'Save the specified Load Address
    Print #FIO, "DIVLEN="; txtDivLen.Text                                   'Save the Divider Length
    Print #FIO, "INLCOL="; txtInlineCol.Text                                'Save the Inline Comment Column
    
    '-- [ENTRY POINTS]
    If lstSYM.ListCount > 0 Then
        Print #FIO, "[ENTRYPT]"
        For J = 0 To lstEntryPt.ListCount - 1
            Print #FIO, lstEntryPt.List(J)                                  'Write the Entrypoint Entry
        Next J
    End If
    
    '-- [SYMBOLS]
    If lstSYM.ListCount > 0 Then
        Print #FIO, "[SYMBOLS]"
        For J = 0 To lstSYM.ListCount - 1
            Print #FIO, lstSYM.List(J)                                      'Write the Symbols Entry
        Next J
    End If
      
    '-- [TABLES]
    If lstDT.ListCount > 0 Then
        Print #FIO, "[TABLES]"
        For J = 0 To lstDT.ListCount - 1
            Print #FIO, lstDT.List(J)                                       'Write the Tables Entry
        Next J
    End If
    
    '-- [LABELS]
    If lstULabels.ListCount > 0 Then
        Print #FIO, "[LABELS]"
        For J = 0 To lstULabels.ListCount - 1
            Print #FIO, lstULabels.List(J)                                  'Write the Labels Entry
        Next J
    End If
    
    '-- [COMMENTS]
    If lstCmnt.ListCount > 0 Then
        Print #FIO, "[COMMENTS]"
        For J = 0 To lstCmnt.ListCount - 1
            Print #FIO, lstCmnt.List(J)                                     'Write the Comment Entry
        Next J
    End If

    Close FIO
    
    lblUpdated.Caption = Date & ", " & Time                                 'Update Date and Time of last change
    ProjFilename = Filename                                                 'Remember the project file
    ClearChangeFlag                                                         'Clear Change Flag
    
End Sub

'---- ASM: Click on the Clear Tables button
Private Sub cmdClrTables_Click()

    If OverwriteProject = True Then
        ClearTables                                                         'Clear All Tables
        ProjFilename = ""                                                   'Clear the Project Filename
        MLReViewA                                                           'Re-Display Listing
    End If
    
End Sub

'---- ASM: Clear All Tables
Private Sub ClearTables()

    lstEntryPt.Clear
    lstSYM.Clear
    lstDT.Clear
    lstULabels.Clear
    lstCmnt.Clear

End Sub

'---- ASM: Load specified List from File
Private Sub LoadSymFile(ByVal Filename As String, ByVal TabNum As Integer)
    Dim FIO As Integer, Tmp As String, Tmp2 As String, Mode As Integer
    Dim Addr As String, Sym As String, Cmnt As String, Flag As Boolean
    
    If Exists(Filename) = False Then Exit Sub
    
    Mode = 0: Tmp = FileExtU(Filename): If Tmp = "SY4" Then Mode = 1        'Check for 'SYM4' file
    
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
                        Case 0                                                      '--Standard format input
                            If Left(Tmp, 1) <> ";" Then lstSYM.AddItem Tmp
                        Case 1 '--SYM4 format
                            If Mid(Tmp, 2, 1) <> " " Then
                                Tmp2 = Mid(Tmp, 13, 4) & "," & Mid(Tmp, 2, 6) & "," & Mid(Tmp, 37)
                                lstSYM.AddItem Tmp2
                            End If
                        Case 2                                                      '--Regenerator Label format: HHHH SYMBOL
                            Addr = Left(Tmp, 4)                                     'Save Address
                            Sym = MyTrim(Mid(Tmp, 6))                               'Save Symbol
                            Tmp = FindSym(Addr)                                     'Check if there is an existing symbol
                            If Tmp = "" Then
                                lstSYM.AddItem Addr & "," & Sym & ","
                            End If
    
                        Case 3                                                      '--ReGenerator Comment format: HHHH Comment
                            Addr = Left(Tmp, 4)                                     'Save Address
                            Cmnt = MyTrim(Mid(Tmp, 6))                              'Save Symbol
                            Tmp = FindSym(Addr)                                     'Check if there is an existing symbol (LastSymPos will point to it)
                            If Tmp = "" Then
                                lstSYM.AddItem Addr & ",," & Cmnt
                            Else                                                    'Symbol was found, so update data
                                    If LastComment = "" Then                        'Only update if the symbol has no existing comment
                                    lstSYM.RemoveItem LastSymPos                    'Remove it
                                    lstSYM.AddItem Addr & "," & Tmp & "," & Cmnt    'Add replacement
                                End If
                            End If
                    End Select
                End If
            Wend

        Case 4
            If Flag = True Then lstDT.Clear                                         'Clear Data Tables
            While Not EOF(FIO)
                Line Input #FIO, Tmp                                                'Read a line
                If Left(Tmp, 1) <> ";" Then                                         'Check for Comment
                    lstDT.AddItem Tmp                                               'Add the line
                    lstDT.Selected(lstDT.NewIndex) = True                           'Select it
                End If
            Wend
                        
        Case 5
            If Flag = True Then lstULabels.Clear                                    'Clear User Labels
            While Not EOF(FIO)
                Line Input #FIO, Tmp: If Left(Tmp, 1) <> ";" Then lstULabels.AddItem Tmp
            Wend
        Case 6
            If Flag = True Then lstCmnt.Clear                                       'Clear Comment
            While Not EOF(FIO)
                Line Input #FIO, Tmp                                                'Get the line
                If Left(Tmp, 1) <> ";" Then lstCmnt.AddItem Tmp                     'Add it, if not a comment
            Wend

    End Select
    Close FIO

End Sub

'---- ASM: Add a new List Entry
Private Sub cmdSymAdd_Click()
    Dim i As Integer, P As Integer, Flag As Boolean
    Dim RS As String, RE As String, Tmp As String, Tmp2 As String
    
    i = lstML.ListIndex
    Tmp2 = ""
    If i > 0 Then Tmp2 = ExtractAddr(lstML.List(lstML.ListIndex))                       'Find the address on selected line
    
    Select Case MLTabNum
        Case 0, 1
            MyMsg "Select the TAB for the type of entry you want first, or use the quick-add buttons at the top of the window!"
            
        Case 2                                                                          '-- Entry Points
            Tmp = Tmp2 & ",-"                                                           'Make default text entry string
            Tmp2 = InputBox("HHHH,DESCRIPTION", "Add Entry Pointl", Tmp)
            If Len(Tmp2) > 3 Then
                Tmp = UCase(Mid(Tmp2, 1, 4)): Mid(Tmp2, 1, 4) = Tmp                     'Force address to uppercase
                lstEntryPt.AddItem Tmp2: MLReViewC                                      'Review plus set changeflag=true
            End If
            
        Case 3                                                                          '-- Symbols
            Tmp = Tmp2 & ",symbol,-"                                                    'Make default text entry string
            Tmp2 = InputBox("HHHH,SYMBOL,DESCRIPTION", "Add Symbol", Tmp)
            If Len(Tmp2) > 5 Then
                Tmp = UCase(Mid(Tmp2, 1, 4)): Mid(Tmp2, 1, 4) = Tmp                     'Force address to uppercase
                lstSYM.AddItem Tmp2: MLReViewC                                          'Review plus set changeflag=true
            End If
            
        Case 4                                                                          '-- Data Tables
            'Check if there is a range selected
            For i = 0 To lstML.ListCount - 1
                If lstML.Selected(i) = True Then
                    If Flag = False Then RS = ExtractAddr(lstML.List(i)): Flag = True   'Found first selected line
                    P = i                                                               'Remember it
                Else
                    If Flag = True Then RE = ExtractAddr(lstML.List(P)): Exit For       'Not selected so use last remembered line for end
                End If
            Next i
            
            If Flag = True Then Tmp = RS & "," & RE & ",b,-"
            Tmp2 = InputBox("Types: Byte Tables:" & Cr & "A/T=Text,B/H=Hex,D=Decimal,Z=Binary," & Cr & "W=Word,R=RTS,V=Vector" & Cr & Cr & "HHHH,HHHH,TYPE{##},DESCRIPTION", "Add Table", Tmp)
            If Len(Tmp2) > 12 Then
                Tmp = UCase(Mid(Tmp2, 1, 9)): Mid(Tmp2, 1, 9) = Tmp                     'Force addresses to uppercase
                lstDT.AddItem Tmp2
                lstDT.Selected(lstDT.NewIndex) = True
                MLReViewC                                                               'Review plus set changeflag=true
            End If
            
        Case 5                                                                          '-- User Labels
            Tmp = Tmp2 & ",name,-"                                                      'Make default text entry string
            Tmp2 = InputBox("HHHH,LABELNAME,DESCRIPTION", "Add Label", Tmp)
            
            If Len(Tmp2) > 7 Then
                Tmp = UCase(Mid(Tmp2, 1, 4)): Mid(Tmp2, 1, 4) = Tmp                     'Force addresses to uppercase
                lstULabels.AddItem Tmp2: MLReViewC
            End If
            
        Case 6                                                                          '-- Comments
            Tmp = Tmp2 & ",s,-"                                                         'Make default text entry string
            Tmp2 = InputBox("Types: I=In-line,S=Single,OTHER=Double Divider Chr" & Cr & "(For Single Divider leave comment empty)" & Cr & Cr & "HHHH,TYPE,COMMENT", "Add Comment", Tmp)
            If Len(Tmp2) > 6 Then
                Tmp = UCase(Mid(Tmp2, 1, 4)): Mid(Tmp2, 1, 4) = Tmp                     'Force address to uppercase
                lstCmnt.AddItem Tmp2: MLReViewC                                         'Review plus set changeflag=true
            End If
    End Select
    
End Sub

'---- ASM: Extracts the HEX Address from the string using current PREFIX
' If PREFIX is not found then look at start of line
Private Function ExtractAddr(ByVal Str As String) As String
    Dim P As Integer, Tmp As String, Tmp2 As String, L As Integer
    
    L = Len(LPrefix)                                                        'Length of Prefix
    P = 1                                                                   'Start at position 1
    If Left(Str, L) = LPrefix Then P = L + 1                                'Skip over prefix
    Tmp = UCase(Mid(Str, P, 4))                                             'Extract the hex address
    Tmp2 = Left(Tmp, 1)                                                     'Get first character
    If (Tmp2 < "0") Or (Tmp2 > "F") Then Exit Function                      'Exit if not 0-F
    If (Tmp2 <= "9") Or (Tmp2 >= "A") Then ExtractAddr = Tmp                'Check for valid 0-9 or A-F

End Function

'---- ASM: Remove the Current List Entry
' Uses global variable MLTabNum to determine list. If item is removed ChangeFlag is set true
Private Sub cmdSymDel_Click()
    Dim i As Integer
    
    Select Case MLTabNum
        Case 2
            i = lstEntryPt.ListIndex
            If i >= 0 Then lstEntryPt.RemoveItem (i): SetChangeFlag
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

'---- ASM: Select Prefix Option from Dropdown
Private Sub cboPrefix_Click()
    SetPrefix cboPrefix.ListIndex
    MLReViewA
End Sub

'---- ASM: Set Label Prefix Format
Private Sub SetPrefix(ByVal N As Integer)
    LPrefix = cboPrefix.List(N)
End Sub

'---- ASM: Set Target Assembler Listing Format
Private Sub cboTarget_Click()
    SetTarget cboTarget.ListIndex
    MLReViewA
End Sub

'---- ASM: Sets Target Assembler Directives
Private Sub SetTarget(ByVal N As Integer)
    Select Case N
        Case 0: DOTORG = "*=":    DOTWORD = "!WORD ": DOTBYTE = "!BYTE ": DOTTEXT = "!TEXT ":  DOTHEX = "$"     'acme
        Case 1: DOTORG = "*=":    DOTWORD = ".WORD ": DOTBYTE = ".BYTE ": DOTTEXT = ".TEXT ":  DOTHEX = "$"     'long
        Case 2: DOTORG = ".ORG ": DOTWORD = ".WOR ":  DOTBYTE = ".BYT ":  DOTTEXT = ".TXT ":   DOTHEX = "$"     'short
        Case 3: DOTORG = ".org ": DOTWORD = ".word ": DOTBYTE = ".byte ": DOTTEXT = ".ascii ": DOTHEX = "0x"    'AS6500 long
        Case 4: DOTORG = ".org ": DOTWORD = ".dw ":   DOTBYTE = ".db ":   DOTTEXT = ".ascii ": DOTHEX = "0x"    'AS6500 short
    End Select
End Sub

'---- ASM: Load opcodes from specified file
Private Sub LoadOpcodes(ByVal Filename As String)
    Dim Tmp As String, J As Integer, FIO As Integer
    
    If Exists(Filename) = False Then Exit Sub
    FIO = FreeFile
    Open Filename For Input As FIO
    
    Line Input #FIO, Tmp                                                        'CBM-Transfer header line
    Line Input #FIO, OpDesc                                                     'CPU description string
    Line Input #FIO, Tmp                                                        'Divider line
    
    For J = 0 To 255
        Input #FIO, Tmp: OP(J) = Tmp
    Next J
    
    Line Input #FIO, Tmp                                                        'Divider line
    Line Input #FIO, OpModeLen                                                  'Opcode lengths
    Line Input #FIO, OpJ                                                        'Tracer Jumps    (unconditional - single flow)
    Line Input #FIO, OpB                                                        'Tracer Branches (conditional - two flows)
    Line Input #FIO, OpZ                                                        'Tracer Stops    (end flow)
    'The rest of the file is ignored
    
    Close FIO
    OpCodeFlag = True
End Sub

'---- ASM: Import Symbol Entries
' Supports Fixed, Comma and Tab-delimited files using parameters entered by user
Private Sub cmdImport_Click()
    Dim Tmp As String, Tmp2 As String, Meth As String
    Dim Filename As String, FIO As Integer
    Dim Par(6) As Integer, Out As String, Flag As Boolean
    Dim C As Integer, i As Integer, p1 As Integer, P2 As Integer
    
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
    
    C = 0                                                                       'Count of symbols imported
    
    FIO = FreeFile
    Open Filename For Input As FIO
    
    While Not EOF(FIO)
        Line Input #FIO, Tmp
        If Left(Tmp, 1) <> ";" Then
            Out = ""
            For i = 1 To 3
                Tmp2 = ""                                                       'Clear Tmp2
                Select Case Meth
                    Case "C": Tmp2 = GetField(Tmp, Par(i))                      'Extract field from Comma-delimited line
                    Case "T": Tmp2 = GetDField(Tmp, "", Par(i))                 'Extract field from delimited line (Null Delimiter defaults to TAB)
                    Case "F"
                        p1 = Par(i * 2 - 1)                                     'Start Position
                        P2 = Par(i * 2)                                         'Length
                        If P2 > 0 Then Tmp2 = MyTrim(Mid(Tmp, p1, P2))          'Extract the field at position p1 with length p2 and trim it
                 End Select
                 If (i = 1) And (Left(Tmp2, 1) = "$") Then Tmp2 = Mid(Tmp2, 2)  'If Addr begins with "$" remove it!
                 Out = Out & Tmp2                                               'Build the string
                 If i < 3 Then Out = Out & ","                                  'Add a comma
            Next
            
            Tmp = Left(Out, 4)                                                  'Check Hex
            If Tmp >= "0000" And Tmp <= "FFFF" Then
                C = C + 1: lstSYM.AddItem Out                                   'Add it to the symbol list
            End If
        End If
    Wend
    Close FIO
    
    MyMsg "File imported! " & Str(C) & " symbols loaded."                       'Show results message
    MLReViewC
    
End Sub

'---- ASM: Purge Un-selected entries from SYMBOL table
Private Sub cmdPurge_Click()
    Dim i As Integer
    
    For i = lstSYM.ListCount - 1 To 0 Step -1
        If lstSYM.Selected(i) = False Then lstSYM.RemoveItem (i)
    Next i
    MLReViewC
End Sub

'---- ASM: Remove Duplicate Generated Label entries
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

'---- ASM: Remove Duplicate External JSR entries
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

'---- ASM: Toggle listing colours
Private Sub imgBW_Click()
    
    If lstML.BackColor = vbWhite Then
        lstML.BackColor = vbBlack: lstML.ForeColor = vbWhite
        lstML2.BackColor = vbBlack: lstML2.ForeColor = vbWhite
    Else
        lstML.BackColor = vbWhite: lstML.ForeColor = vbBlack
        lstML2.BackColor = vbWhite: lstML2.ForeColor = vbBlack
    End If

End Sub

'---- ASM: Display HELP file
Private Sub cmdMLHelp_Click()

    ViewFile ExeDir & "\ml-help.txt"

End Sub

'---- ASM: Load Config File
' The ML Config file contains lines to be loaded into the drop-down menus along with the specified file resource
' Each table group must be proceeded by a selection marker:
' [PLATFORM] [CPU] [PREFIX]

Private Sub LoadMLConfig()
    Dim FIO As Integer, Tmp As String, Tmp2 As String, TMode As Integer, Filename As String
    Dim C1 As Integer, C2 As Integer, P As Integer
    
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
                P = InStr(1, Tmp, ",") 'look for comma separator
                '---- Process according to current section marker
                Select Case TMode
                    Case 1                                                  '-- PLATFORM
                        If P > 0 Then
                            Tmp2 = Left(Tmp, P - 1)
                            cboPlatform.List(C1) = Tmp2
                            cboPlatFile.List(C1) = Mid(Tmp, P + 1)
                            C1 = C1 + 1
                        End If

                    Case 2                                                  '-- CPU
                        If P > 0 Then
                            Tmp2 = Left(Tmp, P - 1)
                            cboCPU.List(C2) = Tmp2
                            cboCPUFile.List(C2) = Mid(Tmp, P + 1)
                            C2 = C2 + 1
                        End If
                        
                    Case 3                                                  '-- Prefix
                        cboPrefix.AddItem Tmp
                End Select
            End If
        End If
    Wend
    
    Close FIO
    
    cboPlatform.ListIndex = 0
    cboCPU.ListIndex = 0
    cboPrefix.ListIndex = 0
    
    MLCFlag = True
    ViewerReady = True
    
End Sub

'========================================
'HEX/Binary Viewer
'========================================
Private Sub HEXView()

    Dim C As Single, W As Integer, H As Integer, H2 As Integer, HTmp As Integer
    Dim Tmp As String, Tmp2 As String
    
    Dim HexAddr As String                                                       'Hex Address of Line
    Dim HxBytes As String, HxBytes2 As String                                   'Hex Values Lines
    Dim PrLine  As String, PrLine2  As String                                   'Printable Characters Lines
    
    Dim Flag As Boolean, MaxW As Integer, LCount As Integer, VLen2 As Integer
    Dim Lo As Integer, Hi As Integer, Address As Long, BMASK As Integer
    Dim HLen As String, LDifs As Integer, DifCount As Integer
    
    Dim ByteFlag As Boolean, CBMFlag As Boolean, ASMFlag As Boolean, CmpFlag As Boolean
    Dim DifsFlag As Boolean, PrtFlag As Boolean, UCFlag As Boolean
    
    
    lstView(2).Clear                                                            'Clear the List
        
    BMASK = 255: If cb7bit.value = vbChecked Then BMASK = 127                   'Enable 7-bit view
     
    MaxW = 7
    
    If cbByte.value = vbChecked Then MaxW = 15 Else MaxW = 31                   'Wider when characters only
    If cbWide.value = vbChecked Then MaxW = MaxW * 2 + 1                        'Wide option: Off=8, On=16
    
    ByteFlag = cbByte.value                                                     'Show Bytes Flag
    ASMFlag = cbHexFmt.value                                                    'ASMbler format Flag
    PrtFlag = cbShowP.value                                                     'Show Printable
    CmpFlag = False                                                             'Assume Compare is false
    CBMFlag = cbGraphics.value                                                  'Set Graphics Flag
    DifsFlag = cbDifs.value                                                     'Set Differences Only Output Flag
    UCFlag = cbUpper.value
    
    If VBuf2 <> "" Then CmpFlag = cbCmpShow.value                               'Set Compare Show Flag only if VBuf2 is defined
    
    HLen = (MaxW + 1) * 3                                                       'Length of Hex bytes
    If ASMFlag = True Then HLen = HLen + MaxW                                   'Compensate for ASM $
    
    '-- Header for Compare Report
    
    If CmpFlag = True Then
        VLen2 = Len(VBuf2)
        lstView(2).AddItem "file compare"
        lstView(2).AddItem "left  file: " & LCase(FileNameOnly(VName)) & "    length=" & Str(VLen) & " bytes"
        lstView(2).AddItem "right file: " & LCase(lblCFile.Caption) & "    length=" & Str(Len(VBuf2)) & " bytes"
        lstView(2).AddItem ""
        lstView(2).AddItem String(7 + 6 * (MaxW + 1), "*")
    End If
    
    '-- Init Variables
    C = 0: W = 0: Tmp = "": HxBytes = "": HxBytes2 = "": LCount = 0: DifCount = 0    'Initialize
    LDifs = 0
    
    '-- Use Address from File or ASM
    
    If cbHexSync.value = vbChecked Then
        Address = MyDec(txtLA.Text)                                             'Use Address specified in ASM project
    Else
        Address = VLA                                                           'Use Load Address from file
        If cbLA.value = vbUnchecked Then Address = MyDec(txtLA.Text)
    End If
    
    '-- Loop through buffer(s)                                                  '== Start of Loop
    Do
        If W > MaxW Then GoSub AddToOutput                                      'Reached Width setting... Add to output
        W = W + 1                                                               'Count bytes processed
        
        '---- Check for start of new line
        
        If W = 1 Then                                                           '-- Initialize Strings for new Line
            HxBytes = "": HxBytes2 = " ; "                                      'Hex Bytes
            PrLine = "": PrLine2 = ""                                           'Printable characters
            If PrtFlag = True Then PrLine2 = " ; "                              'Compare printable string
            If CmpFlag = True Then PrLine = "; "                                'Compare HEX differences
            
            HexAddr = MyHex(Address, 4) & ": "                                  'Start with HEX address
            If UCFlag = False Then HexAddr = UCase(HexAddr)                     'Convert to lowercase
            If ASMFlag = True Then HexAddr = HexAddr & ".BYT "                  'Add ".BYTE" if ASM format
        End If
        
        C = C + 1: Address = Address + 1                                        'Move to Next byte
 
        '---- FILE1:                                                            '==== FILE1: Build the HEX string
        Tmp = Mid(VBuf, C, 1): H = Asc(Tmp)                                     'Get Byte from buffer and its Character code
        
        If ByteFlag = True Then
            If ASMFlag = True Then
                HxBytes = HxBytes & MyHex(H, -2)                                    'Add ASM format HEX string (with $)
                If W <= MaxW Then HxBytes = HxBytes & ","                           'Add ASM format COMMA if not last byte of line
            Else
                HxBytes = HxBytes & MyHex(H, 2) & " "                               'Add HEX string in normal format
            End If
        End If
              
        If PrtFlag = True Then                                                  '-- FILE1: Build Printables String
            HTmp = H And BMASK                                                  'Apply Mask
            If CBMFlag = True Then                                              '-- If CBMFlag = True then we allow everything but NULL
                If HTmp = 0 Then HTmp = 32                                      'If NULL then convert to SPACE (NULL will terminate string in listbox entry)
                PrLine = PrLine & Chr(HTmp)                                     'Add it
            Else
                Select Case HTmp                                                '-- When CBMFlag = False we only allow ASCII 32 to 127
                    Case 32 To 127: PrLine = PrLine & Chr(HTmp)                 'Printable
                    Case Else:      PrLine = PrLine & "."                       'Everything else is Un-Printable and replaced by "."
                End Select
            End If
        End If
        
        '======================================================================= Process FILE2
        If CmpFlag = True Then

            If C <= VLen2 Then                                                  'Check if compare file is shorter than FILE1
            
                Tmp = Mid(VBuf2, C, 1): H2 = Asc(Tmp)                           'Get Byte from Buffer and its Character Code
                                                                                '-- Build the HEX string
                If H = H2 Then
                    HxBytes2 = HxBytes2 & "== "                                 'Show SAME value as "=="
                Else
                    HxBytes2 = HxBytes2 & MyHex(H2, 2) & " "                    'Show DIFFERENT value as Hext
                    DifCount = DifCount + 1                                     'Add to Difference Count
                    LDifs = LDifs + 1                                       'Increment Line Difference count
                End If
                
                If PrtFlag = True Then                                          '-- Build the Printable bytes string
                    HTmp = H2 And BMASK                                         'Apply Mask
                    If HTmp = 0 Then HTmp = 32                                  'If NULL then convert to SPACE (NULL will terminate string in listbox entry)
                    If H = H2 Then
                        PrLine2 = PrLine2 & Chr(HTmp)                           'Add it
                    Else
                        If CBMFlag = True Then
                            PrLine2 = PrLine2 & Chr(HTmp)                       'Use as-is
                        Else
                            Select Case HTmp
                                Case 32 To 127: PrLine2 = PrLine2 & Chr(HTmp)   'Printable
                                Case Else:      PrLine2 = PrLine2 & "."         'Un-Printable
                            End Select
                        End If
                    End If
                End If
                
            Else                                                                '-- File2 is shorter so Pad the strings
                HxBytes2 = HxBytes2 & "   "                                     'Add SPACES to HEX bytes
                If CmpFlag = True Then PrLine2 = PrLine2 & " "                  'Add SPACE to Printable String
            End If
            
        End If
        
    Loop While (C < VLen) 'And (LCount < 32766)
    
    '---- Handle the final line
    
    Tmp = String(HLen, " ")                                                     'Temp spacing string
    
    If HxBytes <> "" Then GoSub AddToOutput

    '---- Report Differences
    
    If CmpFlag = True Then
        If DifCount = 0 Then
            Tmp2 = "Files are identical!"
        Else
            Tmp2 = Str(DifCount) & " differences"
        End If
        
        lblDifTxt.Caption = Tmp2
        lstView(2).List(3) = "results: " & Tmp2
    End If
        
    RefreshVIEW 2
    CalcScroll 2
    
    Exit Sub
    
'==== Add to output list

AddToOutput:
    If UCFlag = False Then                                                          'Change HEX bytes to Lowercase
        HxBytes = LCase(HxBytes)
        HxBytes2 = LCase(HxBytes2)
    End If
    
    If CmpFlag = True Then                                                          '-- Compare Option is Enabled - Show Both Files
        
        If (DifsFlag = False) Or ((DifsFlag = True) And (LDifs > 0)) Then           'Check DifsOnly Option
            If PrtFlag = True Then                                                  'Show Printables Enabled
                lstView(2).AddItem HexAddr & HxBytes & PrLine & HxBytes2 & PrLine2  'Add line with both files and printables
            Else
                lstView(2).AddItem HexAddr & HxBytes & HxBytes2                     'Add line with both files
            End If
        End If
    
    Else                                                                            '-- Compare option NOT Enabled - Show One File
    
        If PrtFlag = True Then                                                      'Show Printable Enabled
            lstView(2).AddItem HexAddr & HxBytes & PrLine                           'Add the line with Printable
        Else
            lstView(2).AddItem HexAddr & HxBytes                                    'Add the line without Printable
        End If
        
    End If
    
    W = 0: LCount = LCount + 1: LDifs = 0
    Return
    
End Sub

'================
' HEX Viewer Subs
'================

'---- HEX: Calculate MD5 Value
Private Sub cmdMD5_Click()
    Dim FIO As Integer, Tmp As String, Dest As String
        
    KillTemp
    Dest = FileBase(VFileName) & ".md5"                                         'Output to FILE with .MD5 extension
    Tmp = "-o" & Quoted(Dest) & " " & Quoted(VFileName)                         'Make the parameter string
    
    frmMain.PubDoCommand CBMMD5, Tmp, "Calculating MD5...", False               'Run the MD5.EXE file
            
    If Exists(Dest) = False Then
        lblDifTxt.Caption = "MD5 had no output"                                 'MD5 failed
        Exit Sub
    End If
    
    '-- Read in the complete output file
    FIO = FreeFile
    Open Dest For Input As FIO
        Tmp = Input(LOF(FIO), FIO)                                              'Read in the output
    Close FIO
   
    lblDifTxt.Caption = Left(Tmp, 32)                                           'Set the compare results label to the MD5 checksum
    
End Sub

'---- HEX: Sync HEX view with FONT Offset
Private Sub lstView_Click(Index As Integer)
    Dim Tmp As String
    
    If frFont.Visible = True Then
        Tmp = Left(lstView(2).List(lstView(2).ListIndex), 4)                    'Get the HEX address
        txtCSkip.Text = Format(MyDec(Tmp))                                      'Convert it to decimal and store it in the offset field
        FONTView                                                                'Re-display font
    End If
End Sub

'---- HEX: Initiate search when ENTER is pressed
Private Sub txtHSS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0                                                            'Clear to prevent beep
        HexFind
    End If
End Sub

'---- HEX: Click to Find
Private Sub cmdHexFind_Click()
    HexFind
End Sub

'---- HEX: Search for NEXT occurance
Private Sub cmdHNext_Click()
    HexSearch ""
End Sub

'---- HEX: Enter Search String or Hex bytes
' Parses the search field and converts hex digits if required then searches from the top of the file
Private Sub HexFind()
    Dim SS As String, HH As String, SH As String
    Dim SL As Integer, J As Integer, P As Integer
    
    HH = txtHSS.Text: SL = Len(HH)
    
    If SL = 0 Then MyMsg "Enter String, or start with $ to search for hex byte(s).": Exit Sub
    If Left(HH, 1) = "$" Then
        HH = Mid(HH, 2): SL = SL - 1                                                'Remove the '$'
        If (SL Mod 2) > 0 Then MyMsg "HEX digits must be in pairs.": Exit Sub
        
        SS = ""
        For J = 1 To SL Step 2
            SH = Mid(HH, J, 2)                                                      'Get 2 hex digits
            SS = SS & Chr(MyDec(SH))                                                'Add character to searchstring
        Next J
    Else
        SS = HH                                                                     'Use original string as entered
    End If
    
    HexSearch SS                                                                    'Search from the TOP
    
End Sub

'---- HEX: Search for specified text string or hex bytes
' SS is the search string. If specified causes the search to start from the TOP
' If ommitted uses the last string and continues searching from last position
Private Sub HexSearch(ByVal SS As String)
    Static LastPos As Integer, LastSS As String                                     'Remembers these between calls
    Dim MaxW As Integer, LL As Integer, L2 As Integer, P As Integer
    Dim HA As String
    
    If SS = "" Then SS = LastSS Else LastPos = 1                                    'If no searchstring then use previous, else start from top
    LastSS = SS                                                                     'Remember the Searchstring
    
    If LastPos > Len(VBuf) Then LastPos = 1                                         'Wrap back to top
   
    P = InStr(LastPos, VBuf, SS)                                                    'Do a binary search
    If P = 0 Then P = InStr(LastPos, VBuf, SS, vbTextCompare)                       'If not found search textually
    
    If P = 0 Then
        MyMsg "No more occurances."                                                 'No results. Display message and exit
        Exit Sub
    End If
    
    LastPos = P + Len(SS)                                                           'Set position for next search
    If cbWide.value = vbChecked Then MaxW = 16 Else MaxW = 8                        'What is the view line lenght?
    LL = (P - 1) \ MaxW                                                             'Which line is the found string on?
    L2 = P - LL * MaxW - 1                                                          'Offset on line
    
    lstView(2).ListIndex = LL                                                       'Select the line containing the string
    HA = MyHex(MyDec(Left(lstView(2).List(LL), 4)) + L2, 4)                         'Hex Address of found
    
    lblSResults.Caption = "Found at $" & HA & ", Offset:" & Str(P - 1)              'Results message

End Sub

'---- HEX: Compare to a second file
Private Sub cmdCompare_Click()
    Dim Filename As String

    Filename = FileOpenSave("", 0, 0, "Load Compare file")
    If Filename = "" Then Exit Sub
    
    If Exists(Filename) = False Then MyMsg "Viewer: File '" & Filename & "' not found!": Exit Sub
    LoadCompare Filename

End Sub

'---- HEX: Load specified File to Compare
Private Sub LoadCompare(ByVal Filename As String)
    Dim FIO As Integer
    Dim Tmp As String, P00Buf As String
    Dim P00Flag As Boolean
    Dim FLen As Long
    
    P00Flag = False                                                                 'Assume normal file
    If FileExtU(Filename) = "P00" Then P00Flag = True                               'P00 file found!
    
    '-- Load the file to the buffer, update and display file details
    FIO = FreeFile
    Open Filename For Binary As FIO: FLen = intLOF(FIO)
        If P00Flag = True Then P00Buf = Input(26, FIO): FLen = FLen - 26            'Skip over header
        If cbLA.value = vbChecked Then
            VBuf2 = Input(2, FIO): FLen = FLen - 2                                  'Read the Load address
        End If
        
        If FLen > 32760 Then FLen = 32760
        VBuf2 = Input(FLen, FIO)                                                    'Read contents to buffer
    Close FIO
    
    lblCFile.Caption = FileNameOnly(Filename)
    If VBuf = VBuf2 Then lblDifTxt.Caption = "The files are identical!"
    HEXView
End Sub


'============
'SEQ Viewer
'============
Sub SEQView()
    Dim C As Integer, H As Integer, MaxCol As Integer                               'Counter, Character Code, Maximum Column
    Dim Tmp As String, TLine As String                                              'Character and output strings
    Dim EncFlag As Boolean
    Dim StripCR As Boolean, StripLF As Boolean, StripUP As Boolean                  'Show PETSCII flag
    Dim UCFlag As Boolean
    
    lstView(1).Clear
    
    If cboColWidth.ListIndex < 0 Then cboColWidth.ListIndex = 0                     'Set Default COL Width
    
    MaxCol = Val(cboColWidth.List(cboColWidth.ListIndex))                           'Max Width
    
    EncFlag = False: If EncodeL(1) < 2 Then EncFlag = True                          'Flag to let show PETSCII
    StripCR = cbIgnoreCR.value                                                      'Flag to Ignore Carriage Returns
    StripLF = cbIgnoreLF.value                                                      'Flag to Ignore Line Feeds
    StripUP = cbIgnoreUnP.value                                                     'Flag to Ignore Un-Printable
    
    C = 1                                                                           'Line Counter
    Tmp = "": TLine = ""                                                            'Character and line strings
    
    Do
        Tmp = Mid(VBuf, C, 1): H = Asc(Tmp)                                         'Get the character and value
        
        Select Case H
            Case 0: Tmp = ""
            Case 10: Tmp = "": If StripLF = True Then H = 0                         'Remove LF
            Case 13: Tmp = "": If StripCR = True Then H = 0                         'Remove CR
            Case 32 To 127:
            Case 128 To 255: If EncFlag = False Then Tmp = Chr(H And 127)           'Strip high bit
            Case Else
                If EncFlag = False Then                                             'If Encode FLAG=FALSE then process
                    If StripUP = True Then Tmp = "" Else Tmp = "."                  'Remove if Ignore Unprintable, else convert to "."
                End If
        End Select

        TLine = TLine & Tmp                                                         'Add Character to line

        If (Len(TLine) >= MaxCol) Or (H = 13) Then
            lstView(1).AddItem TLine: TLine = ""                                    'Add the line
        End If
        
        C = C + 1                                                                   'Next line
    
    Loop While (C < VLen) And (C < 32767)                                           'Exit if End of file or Max# of lines
    
    If TLine <> "" Then lstView(1).AddItem TLine                                    'Add any remaining characters
    
    RefreshVIEW 1                                                                   'Display it!
    CalcScroll 1                                                                    'Calculate Scrollbar
    
End Sub


'===============
' Bitmap Viewer
'===============
Private Sub BMPView()
    Dim Comment As String, i As Integer, X As Integer, FLen As Long
    Dim TwipX As Integer, TwipY As Integer, LAOffset As Integer
        
    TwipX = Screen.TwipsPerPixelX
    TwipY = Screen.TwipsPerPixelY
    
    If p_name(0) = "" Then Call LoadPicFormats                                      'Load picture formats if needed
        
    lblBComment.Caption = "None"                                                    'Clear comment
    LAOffset = 0: If cbLA.value = vbChecked Then LAOffset = 2                       'Compensate for LA if selected
    ImageType = HRBW                                                                'Default to hi-res mono
    
    picBMP.Visible = False                                                          'Hide the picture
    lblMoment.Visible = True                                                        'Display loading message
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
            picBMP.Width = 640 * TwipX
            picBMP.Height = 720 * TwipY
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
        
        picBMP.Width = 320 * TwipX                                                      'Standard 320x200 bitmap
        picBMP.Height = 200 * TwipY
        Read_Bitmap VFileName
    End If
    
    picBMP.Visible = True                                                               'Hide the picture
    lblMoment.Visible = False                                                           'Display loading message
    DoEvents
    
End Sub

'---- Read GeoPaint Image
Private Sub Read_GeoPaint(ByVal Filename As String)
    Dim Dat As String, PBuf As String, M As String
    Dim i As Integer, J As Integer, L As Integer, FIO As Integer
    Dim BitPosH As Integer, BitPosV As Integer, DPos As Integer
    Dim LDat As Integer, Pel As Integer, XX As Integer, YY As Integer
    Dim DT As Integer, K As Integer, k2 As Integer, nxt As Integer
    Dim ValidSectors As Integer, Sector As Integer
    
    Dim C0 As Long, C1 As Long                                              'Pixel on and off colours - new May 2017
    
    ReDim blocks(0 To 44, 1 To 2)
    ReDim pat(0 To 7)
    
    Close PFIO
        
    PFIO = FreeFile
    Open Filename For Binary As PFIO
    
    PBuf = ReadBlock()                                                      'First sector
    PBuf = ReadBlock()                                                      'Second sector
    PBuf = ReadBlock()                                                      'Third sector
    
    ValidSectors = 0: Sector = 0

    picBMP.Cls
    
    C0 = CBMColor(1)                                                        'White Background Colour - new 2017
    C1 = CBMColor(0)                                                        'Black Foreground Colour - new 2017
    picBMP.BackColor = C0                                                   'Default to white background
    picBMP.Cls                                                              'Clear to background colour
    
    For i = 0 To 44
      M = Left(PBuf, 2)
      blocks(i, 1) = Asc(M)
      If blocks(i, 1) <> 0 Then
        blocks(i, 2) = Asc(Right(M, 1))
        ValidSectors = ValidSectors + 1
      End If
      PBuf = Mid(PBuf, 3)
    Next i
    
    '-- Display loop
    For i = 0 To 44
        If blocks(i, 1) > 0 Then
            Dat = ""
            For J = 1 To blocks(i, 1)
                PBuf = ReadBlock()
                If J = blocks(i, 1) Then PBuf = Left(PBuf, blocks(i, 2))
                Dat = Dat & PBuf
            Next J
            
            BitPosH = 0:  BitPosV = 0
            
            DPos = 1
            LDat = Len(Dat)
            
            DoEvents
            
            Do While (BitPosV < 16) And (LDat >= DPos)
                nxt = Asc(Mid(Dat, DPos, 1) & Nu): DPos = DPos + 1
                
                Select Case nxt
                  Case 1 To 63
                    For K = 1 To nxt
                      Pel = Asc(Mid(Dat, DPos, 1) & Nu): DPos = DPos + 1
                      GoSub PaintBit
                    Next K
                    
                  Case 65 To 127
                    For K = 0 To 7
                      pat(K) = Asc(Mid(Dat, DPos, 1) & Nu): DPos = DPos + 1
                    Next K
                    
                    For L = 1 To (nxt And 63)
                      For K = 0 To 7
                        Pel = pat(K): GoSub PaintBit
                      Next K
                    Next L
                    
                  Case 129 To 255
                    DT = Asc(Mid(Dat, DPos, 1) & Nu): DPos = DPos + 1
                    For K = 1 To (nxt - 128)
                      Pel = DT
                      GoSub PaintBit
                    Next K
                End Select
            Loop
            
            Sector = Sector + 1
        End If
    Next i
    
    Close PFIO
Exit Sub

'---- Paint Bits
PaintBit:
    XX = BitPosH * 8 + 7
    YY = i * 16 + BitPosV
    
    For k2 = 0 To 7
        If (Pel And Pow(k2)) Then picBMP.PSet (XX - k2, YY), C1 'Set Black dot
    Next k2
    
    BitPosV = BitPosV + 1
    
    If BitPosV = 8 Or BitPosV = 16 Then
        BitPosH = BitPosH + 1: BitPosV = BitPosV - 8
        If BitPosH > 79 Then BitPosH = BitPosH - 80: BitPosV = BitPosV + 8
    End If
    
    Return

End Sub

Private Sub Read_Bitmap(ByVal Filename As String)
    Dim Bitmap As String, Scrn As String, Col As String, Bk As String
    Dim Pel As Integer, BG As Integer, XX As Integer, YY As Integer
    Dim BitPosH As Integer, BitPosB As Integer, BitPosV As Integer, DPos As Integer, CPos As Integer
    Dim C As Integer, S As Integer
    Dim k2 As Integer, k3 As Integer
    Dim ColPut As Long
    
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
    
    BitPosH = 0: BitPosV = 0: DPos = 1: CPos = 1
    BG = Asc(Bk)
    
    picBMP.Cls
    DoEvents
        
    Do While (BitPosV < 200)
    
        Pel = Asc(Mid(Bitmap, DPos, 1))
        DPos = DPos + 1
        XX = BitPosH * 8 + 7
        YY = 0
        
        Select Case p_type(ImageType)
            Case HRBW                                                                                   'High-res Mono Mode
                For k2 = 0 To 7
                    picBMP.PSet (XX - k2, BitPosV), IIf(Pel And Pow(k2), CBMColor(0), CBMColor(1))
                Next k2
                
            Case HR 'High-res Colour Mode
                S = Asc(Mid(Scrn, CPos, 1))
                For k2 = 0 To 7
                    picBMP.PSet (XX - k2, BitPosV), IIf(Pel And Pow(k2), CBMColor((S And 240) / 16), CBMColor(S And 15))
                Next k2
                
            Case MC 'Multi-colour Mode
                S = Asc(Mid(Scrn, CPos, 1))
                C = Asc(Mid(Col, CPos, 1))
                For k2 = 0 To 6 Step 2
                    k3 = 0
                    If (Pel And Pow(k2)) Then k3 = k3 + 1
                    If (Pel And Pow(k2 + 1)) Then k3 = k3 + 2
                    
                    Select Case k3
                        Case 0: ColPut = CBMColor(BG)
                        Case 1: ColPut = CBMColor((S And 240) / 16)
                        Case 2: ColPut = CBMColor(S And 15)
                        Case 3: ColPut = CBMColor(C And 15)
                    End Select
                    
                    picBMP.PSet (XX - k2, BitPosV), ColPut&
                    picBMP.PSet (XX - k2 - 1, BitPosV), ColPut&
                Next k2
        End Select
    
        BitPosV = BitPosV + 1
        If BitPosV / 8 = BitPosV \ 8 Then
            BitPosH = BitPosH + 1: BitPosV = BitPosV - 8
            CPos = CPos + 1
            If BitPosH > 39 Then BitPosH = BitPosH - 40: BitPosV = BitPosV + 8
        End If
    Loop

End Sub

Private Sub LoadPicFormats()
    Dim Filename As String, Tmp As String
    Dim FIO As Integer, Num As Integer
       
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
    If Filename <> "" Then SavePicture picBMP.Image, Filename

End Sub

Private Sub cmdLoadVPL_Click()
    Dim Filename As String
    
    Filename = FileOpenSave(FileBase(VFileName), 0, 7, "Load VICE Palette")
    If Filename <> "" Then
        LoadVPL Filename
        BMPView 're-draw the image with new palette
        cmdLoadVPL.ToolTipText = "Load VICE Palette file [Current: " & Filename & "]"
    End If

End Sub

'---- Reads a chunk of 256 bytes from the bitmap file
Private Function ReadBlock() As String
    Dim Buf As String
    
    Buf = Space(254)
    Get #PFIO, , Buf
    ReadBlock = Buf

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

'=============================================================
' COMMON Routines
'=============================================================

'---- COMMON: File Open or Save Dialog
' You can specify a default filename, a File Filter list index (0-8), and Window Title
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
        Case 8: CommonDialog.Filter = "CBM Files (*.PRG,*.SEQ,*.USR)|*.PRG;*.SEQ;*.USR"
    End Select
    
    If Mode = 0 Then CommonDialog.ShowOpen Else CommonDialog.ShowSave   'MODE: 0=Open, 1=Save
        
    If CommonDialog.Filename = "" Then Exit Function
    
    FileOpenSave = CommonDialog.Filename
    Exit Function
NoFile:

End Function

'==========================================================
' Controls causing Re-load and refresh of output (Any View)
'==========================================================

'---- COMMON: Change the Load Address Mode
' When the Load Address Mode is changed the file has to be reloaded
' with or without the first two bytes included.
' Some files, like BASIC, require a Load Address.
' Others, like SEQ files, pictures, general data files etc, do not.
Private Sub cbLA_Click()

    DoEvents
    ViewIt ViewMode, VFileName, VName, VExt 're-load the file
    
End Sub

'---- COMMON: Edit Load Address
' When Load Address is not included we must manually supply a load address.
' For example, BIN and ROM files.
Private Sub txtLA_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0: ViewIt ViewMode, VFileName, VName, VExt 're-load the file
    End If
    
End Sub

'==========================================================
' Controls that cause a refresh of output (Any View)
'==========================================================
'These controls control the view, but do not require re-loading the file

'---- BASIC Updates

Private Sub cboMode_Click()
    If ViewerReady Then BASView
End Sub
Private Sub cboColWidth2_Click()
    BASView
End Sub
Private Sub cbRev_Click()
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
Private Sub cbMV_Click()
    BASView
End Sub

'---- SEQ Updates

Private Sub cboColWidth_Click()
    SEQView
End Sub

Private Sub cbIgnoreCR_Click()
    SEQView
End Sub

Private Sub cbIgnoreLF_Click()
    SEQView
End Sub

Private Sub cbIgnoreUnP_Click()
    SEQView
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
Private Sub cbByte_Click()
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

Private Sub cbDifs_Click()
    HEXView
End Sub

Private Sub cbGraphics_Click()
    HEXView
End Sub
Private Sub cbUpper_Click()
    HEXView
End Sub

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
Private Sub cbBlock_Click()
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

'========================================
' Common List Routines
'========================================

'---- COMMON: Calculate Scrollbar Properties
Private Sub CalcScroll(ByVal Index As Integer)
    Dim L As Integer, M As Integer, N As Integer
    Dim vH As Integer
    
    If (Index < 0) Or (Index > 2) Then Exit Sub                                  'Exit if Index out of range
    If LFontH(Index) < 1 Then Exit Sub                                          'Exit if Font Height not set
    If vsView(Index).Height < 500 Then Exit Sub                                 'Exit if scrollbar not set or too small
    
    vH = vsView(Index).Height / (15 * 8 * LFontH(Index))                        'Number of visible entries in list based on height of list and scalefactor
    
    N = lstView(Index).ListCount                                                'Number of Entries in list
    M = N - vH                                                                  'Subtract how many fit in listing
    
    If M < 0 Then
        M = 0: L = vH                                                           'Doesn't fill the window?
    Else
        L = Int(N / vH) * 100: If L > vH Then L = vH                            'Scrollbar size
    End If
    
    '-- Set the Scrollbar
    
    vsView(Index).Min = 0
    vsView(Index).Max = M                                                       'Set Maximum value to number of entries in list
    vsView(Index).SmallChange = 1
    vsView(Index).LargeChange = L                                               'Set Scrollbar size
    If vsView(Index).value > M Then vsView(Index).value = 0
    vsView(Index).Refresh
    
End Sub

'---- COMMON: Scroll the list
Private Sub vsView_Change(Index As Integer)
    
    RefreshVIEW Index

End Sub

'---- COMMON: Scroll the list
Private Sub vsView_Scroll(Index As Integer)
    
    RefreshVIEW Index

End Sub

'---- COMMON: Refresh List View using Current Parameters
Private Sub RefreshVIEW(ByVal Index As Integer)
    
    frmMain.DrawCBM lstView(Index), picView(Index), vsView(Index), EncodeL(Index), LFontW(Index), LFontH(Index)

End Sub

'---- COMMON: Display the Popup Font Size Menu
Private Sub cmdFSize_Click(Index As Integer)
    
    MenuForm = 2                                                        'Which Form to notify
    MenuNum = Index                                                     'Menu Number
    
    PopupMenu frmMenu.mnuFontSize                                       'Pop up the Font Size Menu
    
End Sub

'---- COMMON: Set the Tooltip for the Encode Drop-down
Private Sub SetEncodeTip(ByVal Index As Integer)
    Dim Tmp As String
    
    Select Case EncodeL(Index)
        Case 0: Tmp = "PETSCII Upper"
        Case 1: Tmp = "PETSCII Lower"
        Case 2: Tmp = "Screen Upper"
        Case 3: Tmp = "Screen Lower"
        Case 4: Tmp = "ASCII SuperPET"
        Case 5: Tmp = "ASCII"
    End Select
    
    cmdEncode(Index).ToolTipText = "Encoding: " & Tmp                                  'Set the Tooltip

End Sub

'---- COMMON Change the List Font Width and Height
' Look in frmMenu for menu item list for SELECT CASE settings
Private Sub SetListFontWH(ByVal Index As Integer)

    Dim W As Integer, H As Integer
    
    Select Case Index - 500                                             'Set Width and Height according to Menu Index
        Case 2: W = 2: H = 2
        Case 3: W = 4: H = 2
        Case 4: W = 1: H = 1
        Case 5: W = 2: H = 1
        Case 6: W = 2: H = 3
        Case 7: W = 4: H = 3
        Case Else: W = 1: H = 2
    End Select
    
    LFontW(MenuNum) = W                                                 'Set the Width
    LFontH(MenuNum) = H                                                 'Set the Height
    cmdFSize(MenuNum).ToolTipText = "Size: " & Format(W) & " x " & Format(H)       'Set the tooltip to show width and height
    
    RefreshVIEW MenuNum                                                 'Redraw the list
    CalcScroll MenuNum                                                  'Re-calculate scrollbar
    
End Sub

'---- COMMON: Popup the List Font Encoding Menu
Private Sub cmdEncode_Click(Index As Integer)
    
    MenuForm = 2                                                        'Which Form to notify
    MenuNum = Index                                                     'Menu Number
    
    PopupMenu frmMenu.mnuE                                              'Pop up the menu
    
End Sub

'========================================
' THEME
'========================================

'---- Set Theme
' ThemeBG=Title/Background
' ThemeFrBG=Frames Background
' ThemeListBG=Listbox Background
' ThemeListFG=Listbox Foreground
' ThemeFG=Text Labels

Public Sub SetVTheme()
    Dim i As Integer, J As Integer, Y As Integer
    
    On Local Error GoTo 0
    
    Me.BackColor = ThemeBG                                                          'The Form
    Me.ForeColor = ThemeFG                                                          'The Form

    '--- COMMON: Top elements
    
    frMenu.BackColor = ThemeBG
     
    lblViewAs.BackColor = ThemeBG:      lblViewAs.ForeColor = ThemeFG               'Common
    lblSplit.BackColor = ThemeFrBG:     lblSplit.ForeColor = ThemeFrFG              'Common
    lblSelect.BackColor = ThemeFrBG:    lblSelect.ForeColor = ThemeFrFG             'Common
    lblSz.BackColor = ThemeBG:          lblSz.ForeColor = ThemeFG                   'Common
    lblVSize.BackColor = ThemeBG:       lblVSize.ForeColor = ThemeFG                'Common
    cbLA.BackColor = ThemeBG:           cbLA.ForeColor = ThemeFG                    'Common
    cbLockView.BackColor = ThemeBG:     cbLockView.ForeColor = ThemeFG              'Common
    
    txtLA.BackColor = ThemeListBG:      txtLA.ForeColor = ThemeListFG               'Common
    
    
    For i = 0 To 2
        lblSSize(i).BackColor = ThemeFrBG: lblSSize(i).ForeColor = ThemeFrFG        'Common
    Next i
    
    For i = 2 To 24
        Label(i).BackColor = ThemeFrBG: Label(i).ForeColor = ThemeFrFG              'Common
    Next i
    
    '-- COMMON: Get Tab Colours Row 0-3=Tabs, 4-7=Font Editor Colours
    
    For i = 0 To 3                                                                  '[ ROW/Y
        For J = 0 To 5                                                              ' [ COL/X
            TabColour(J, i) = frmMain.GetTheme(135 + J * 9, 13 + (i * 5))           'Get Tab Colours from bitmap (x,y)
            TabColour(J, i + 4) = frmMain.GetTheme(23 + J * 9, 34 + (i * 5))        'Get Font Editor Colours from bitmap (x,y)
        Next J
    Next i

    '--- COMMON: Frames
    
    frBasic.BackColor = ThemeFrBG: frBasic.ForeColor = ThemeFrFG                    'Common
    frSEQ.BackColor = ThemeFrBG:   frSEQ.ForeColor = ThemeFrFG                      'Common
    frBIN.BackColor = ThemeFrBG:   frBIN.ForeColor = ThemeFrFG                      'Common
    frFont.BackColor = ThemeFrBG:  frFont.ForeColor = ThemeFrFG                     'Common
    frBMP.BackColor = ThemeFrBG:   frBMP.ForeColor = ThemeFrFG                      'Common
    frML.BackColor = ThemeFrBG:    frML.ForeColor = ThemeFrFG                       'Common
    frBlank.BackColor = ThemeFrBG: frBlank.ForeColor = ThemeFrFG                    'Common
    
    '--- BASIC
    
    frBOpts.BackColor = ThemeFrBG:      frBOpts.ForeColor = ThemeFrFG               'Basic
    cbRev.BackColor = ThemeFrBG:        cbRev.ForeColor = ThemeFrFG                 'Basic
    cbPad.BackColor = ThemeFrBG:        cbPad.ForeColor = ThemeFrFG                 'Basic
    cbExp.BackColor = ThemeFrBG:        cbExp.ForeColor = ThemeFrFG                 'Basic
    cbOneLine.BackColor = ThemeFrBG:    cbOneLine.ForeColor = ThemeFrFG             'Basic
    cbMV.BackColor = ThemeFrBG:         cbMV.ForeColor = ThemeFrFG                  'Basic
    cbUC.BackColor = ThemeFrBG:         cbUC.ForeColor = ThemeFrFG                  'Basic
        
    cboMode.BackColor = ThemeListBG:    cboMode.ForeColor = ThemeListFG             'Basic
    cboColWidth2.BackColor = ThemeListBG: cboColWidth2.ForeColor = ThemeListFG      'Basic
    lblLoadAdr.BackColor = ThemeListBG: lblLoadAdr.ForeColor = ThemeListFG          'Basic
    lblGuess.BackColor = ThemeListBG:   lblGuess.ForeColor = ThemeListFG            'Basic
    lblNote.ForeColor = ThemeFG
    lblBView.ForeColor = ThemeFrFG
    
    '--- SEQ
    
    cboColWidth.BackColor = ThemeListBG: cboColWidth.ForeColor = ThemeListFG        'Seq
        
    cbIgnoreLF.BackColor = ThemeFrBG:   cbIgnoreLF.ForeColor = ThemeFrFG            'Seq
    cbIgnoreCR.BackColor = ThemeFrBG:   cbIgnoreCR.ForeColor = ThemeFrFG            'Seq
    cbIgnoreUnP.BackColor = ThemeFrBG:  cbIgnoreUnP.ForeColor = ThemeFrFG           'Seq
    
    '--- HEX/BIN
    cbByte.BackColor = ThemeFrBG:       cbByte.ForeColor = ThemeFrFG                'Bin
    cbGraphics.BackColor = ThemeFrBG:   cbGraphics.ForeColor = ThemeFrFG            'Bin
    cbWide.BackColor = ThemeFrBG:       cbWide.ForeColor = ThemeFrFG                'Bin
    cbShowP.BackColor = ThemeFrBG:      cbShowP.ForeColor = ThemeFrFG               'Bin
    cb7bit.BackColor = ThemeFrBG:       cb7bit.ForeColor = ThemeFrFG                'Bin
    cbHexSync.BackColor = ThemeFrBG:    cbHexSync.ForeColor = ThemeFrFG             'Bin
    cbHexFmt.BackColor = ThemeFrBG:     cbHexFmt.ForeColor = ThemeFrFG              'Bin
    cbUpper.BackColor = ThemeFrBG:      cbUpper.ForeColor = ThemeFrFG
    cbCmpShow.BackColor = ThemeFrBG:    cbCmpShow.ForeColor = ThemeFrFG             'Bin
    cbDifs.BackColor = ThemeFrBG:       cbDifs.ForeColor = ThemeFrFG                'Bin
    
    lblCFile.ForeColor = ThemeFG                                                    'Bin
    lblSResults.BackColor = ThemeFrBG:  lblSResults.ForeColor = ThemeFrFG           'Bin
    lblDifTxt.BackColor = ThemeBG:      lblDifTxt.ForeColor = ThemeFG               'Bin
    
    txtHSS.BackColor = ThemeListBG: txtHSS.ForeColor = ThemeListFG                  'Bin
    
    '--- FONT
    
    frTools.BackColor = ThemeFrBG:      frTools.ForeColor = ThemeFrFG               'Font
    frControls.BackColor = ThemeFrBG:   frControls.ForeColor = ThemeFrFG            'Font
    frChr.BackColor = ThemeFrBG:        frChr.ForeColor = ThemeFrFG                 'Font

    cbSetSize.BackColor = ThemeFrBG:    cbSetSize.ForeColor = ThemeFrFG             'Font

    cbShiftMode.BackColor = ThemeFrBG:  cbShiftMode.ForeColor = ThemeFrFG           'Font
        
    lblEndRange.ForeColor = ThemeFG
    lblFStat.BackColor = ThemeBG: lblFStat.ForeColor = ThemeFG
    
    For i = 0 To 2
        LabelC(i).ForeColor = TabColour(2, 4)
        LabelC(i).BackColor = TabColour(2, 5) 'Set/Num/## Box Colours (top set)
    Next
    
    '--- Set Editor Info boxes
    lblChrSet.ForeColor = TabColour(3, 4): lblChrSet.BackColor = TabColour(4, 5)    'Character Set#
    lblChrNum.ForeColor = TabColour(3, 4): lblChrNum.BackColor = TabColour(4, 5)    'Character # in Set
    lblChrSel.ForeColor = TabColour(3, 4): lblChrSel.BackColor = TabColour(4, 5)    'Character #
    
    
    For i = 0 To 1
        lblBorder(i).ForeColor = ThemeFG
    Next i

    '--- FONT: SCREEN DESIGNER
    
    frEditor.BackColor = ThemeBG                                                    'Screen Designer
    lblCursor.BackColor = ThemeBG:          lblCursor.ForeColor = ThemeFG           'Screen Designer
    cboTheme.BackColor = ThemeListBG:       cboTheme.ForeColor = ThemeListFG        'Screen Designer
    cbCBM.BackColor = ThemeBG:              cbCBM.ForeColor = ThemeFG               'Screen Designer
    
    '--- ASM

    frInfo.BackColor = ThemeFrBG:           frInfo.ForeColor = ThemeFrFG            'Asm
    frTView.BackColor = ThemeFrBG:          frTView.ForeColor = ThemeFrFG           'Asm
    frMLSettings.BackColor = ThemeFrBG:     frMLSettings.ForeColor = ThemeFrFG      'Asm
    frTrace.BackColor = ThemeFrBG:          frTrace.ForeColor = ThemeFrFG           'Asm
    frBlock.BackColor = ThemeFrBG:          frBlock.ForeColor = ThemeFrFG           'Asm
    
    lblShw.ForeColor = ThemeFG
    lblEA.ForeColor = ThemeFG
    txtBlock.BackColor = ThemeListBG:       txtBlock.ForeColor = ThemeFG            'Asm
    lblLineNum.ForeColor = ThemeFG
    lblInfo.ForeColor = ThemeFG
    
    cbClearOnLoad.BackColor = ThemeFrBG:    cbClearOnLoad.ForeColor = ThemeFrFG     'Asm
    cbEquates.BackColor = ThemeFrBG:        cbEquates.ForeColor = ThemeFrFG         'Asm
    cbOpUCase.BackColor = ThemeFrBG:        cbOpUCase.ForeColor = ThemeFrFG         'Asm
    cbBlock.BackColor = ThemeFrBG:          cbBlock.ForeColor = ThemeFrFG           'Asm
    cbSpaceRTS.BackColor = ThemeFrBG:       cbSpaceRTS.ForeColor = ThemeFrFG        'Asm
    cbLabelBlanks.BackColor = ThemeFrBG:    cbLabelBlanks.ForeColor = ThemeFrFG     'Asm
    cbIncSym.BackColor = ThemeFrBG:         cbIncSym.ForeColor = ThemeFrFG          'Asm
    cbHHHH.BackColor = ThemeFrBG:           cbHHHH.ForeColor = ThemeFrFG            'Asm
    cbCompareOut.BackColor = ThemeFrBG:     cbCompareOut.ForeColor = ThemeFrFG      'Asm
    cbMLAddLabels.BackColor = ThemeFrBG:    cbMLAddLabels.ForeColor = ThemeFrFG     'Asm
    cbVerb.BackColor = ThemeFrBG:           cbVerb.ForeColor = ThemeFrFG            'Asm
    cboPlatform.BackColor = ThemeListBG:    cboPlatform.ForeColor = ThemeListFG     'Asm
    cboCPU.BackColor = ThemeListBG:         cboCPU.ForeColor = ThemeListFG          'Asm
    cboMLFmt.BackColor = ThemeListBG:       cboMLFmt.ForeColor = ThemeListFG        'Asm
    cboTarget.BackColor = ThemeListBG:      cboTarget.ForeColor = ThemeListFG       'Asm
    cboPrefix.BackColor = ThemeListBG:      cboPrefix.ForeColor = ThemeListFG       'Asm
    
    txtDivLen.BackColor = ThemeListBG:      txtDivLen.ForeColor = ThemeListFG      'Asm
    txtInlineCol.BackColor = ThemeListBG:   txtInlineCol.ForeColor = ThemeListFG    'Asm
    txtStartLine.BackColor = ThemeListBG:   txtStartLine.ForeColor = ThemeListFG    'Asm
    txtLineInc.BackColor = ThemeListBG:     txtLineInc.ForeColor = ThemeListFG      'Asm
    
    '--- BITMAP
    
    lblBType.ForeColor = ThemeFG                                                    'Bitmap
    lblBComment.ForeColor = ThemeFG                                                 'Bitmap
    lblMoment.ForeColor = ThemeFG                                                   'Bitmap
    
    Y = 67
    For i = 0 To 2
        frmMain.GetIcon cmdEncode(i), 246, Y                                        'Get Encode Icon
        frmMain.GetIcon cmdFSize(i), 267, Y                                         'Get Font Size Icon
    Next i

    frmMain.GetIcon cmdFontMenu, 288, Y                                             'Get Dropdown Menu Icon
    frmMain.GetIcon cmdSEDMenu, 288, Y                                              'Get Screen Editor Menu Icon
    
    Y = 93
    frmMain.GetIcon cmdShift(0), 225, Y                                             'Get Shift UP    Icon
    frmMain.GetIcon cmdShift(1), 257, Y                                             'Get Shift DOWN  Icon
    frmMain.GetIcon cmdShift(2), 289, Y                                             'Get Shift LEFT  Icon
    frmMain.GetIcon cmdShift(3), 321, Y                                             'Get Shift RIGHT Icon
    
    '--- Update GUI Elements
    
    For J = 0 To 2: RefreshVIEW J: Next                                             'Re-draw the list with the new theme bitmaps
    
    SetViewTabs                                                                     'Update View Tabs
    UpdateSelectors                                                                 'Update all character selector boxes

End Sub

