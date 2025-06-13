VERSION 5.00
Begin VB.Form frmOptions 
   Appearance      =   0  'Flat
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CBM-Transfer Options"
   ClientHeight    =   12600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16710
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12600
   ScaleWidth      =   16710
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame optFrame 
      BackColor       =   &H00404040&
      Caption         =   "Utility Paths"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   3255
      Index           =   1
      Left            =   30
      TabIndex        =   104
      Top             =   4170
      Width           =   5445
      Begin VB.PictureBox cmdUBrowse 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5010
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   117
         ToolTipText     =   "Browse"
         Top             =   2820
         Width           =   285
      End
      Begin VB.OptionButton optUPath 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "OpenCBM:"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   109
         Top             =   1080
         Width           =   1245
      End
      Begin VB.OptionButton optUPath 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "VICE,1541:"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   108
         Top             =   1395
         Width           =   1245
      End
      Begin VB.OptionButton optUPath 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NIBTools:"
         ForeColor       =   &H80000008&
         Height          =   405
         Index           =   2
         Left            =   90
         TabIndex        =   107
         Top             =   1665
         Width           =   1245
      End
      Begin VB.OptionButton optUPath 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "CBM-Link:"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   90
         TabIndex        =   106
         Top             =   2040
         Width           =   1245
      End
      Begin VB.OptionButton optUPath 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Other:"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   90
         TabIndex        =   105
         ToolTipText     =   "ACME, MD5"
         Top             =   2340
         Width           =   1245
      End
      Begin VB.TextBox txtUpath 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   120
         TabIndex        =   115
         ToolTipText     =   "Edit PATH here then press ENTER"
         Top             =   2820
         Width           =   4845
      End
      Begin VB.Label lblUPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   110
         Top             =   2400
         Width           =   4005
      End
      Begin VB.Label lblUPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   111
         Top             =   2070
         Width           =   4005
      End
      Begin VB.Label lblUPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   112
         Top             =   1755
         Width           =   4005
      End
      Begin VB.Label lblUPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   113
         Top             =   1425
         Width           =   4005
      End
      Begin VB.Label lblUPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   114
         Top             =   1110
         Width           =   4005
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmOptions.frx":0442
         ForeColor       =   &H00E0E0E0&
         Height          =   795
         Index           =   7
         Left            =   90
         TabIndex        =   116
         Top             =   240
         Width           =   5235
      End
   End
   Begin VB.Frame optFrame 
      BackColor       =   &H00404040&
      Caption         =   "General Options"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   3555
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   90
      Width           =   5445
      Begin VB.CommandButton cmdClearLog 
         Appearance      =   0  'Flat
         Caption         =   "Clear Log"
         Height          =   375
         Left            =   3720
         TabIndex        =   102
         Top             =   1020
         Width           =   1455
      End
      Begin VB.CheckBox cbIgnoreBadID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ignore BAD disk ID's when imaging D64"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   150
         TabIndex        =   100
         Top             =   2520
         Width           =   5115
      End
      Begin VB.CheckBox cbErr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "For results button, also show ERROR file"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   150
         TabIndex        =   99
         Top             =   2790
         Width           =   5115
      End
      Begin VB.CheckBox cbP00 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Write P00 files when copying from Disk Images"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   91
         Top             =   1500
         Width           =   5115
      End
      Begin VB.CheckBox cbDAD 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Start with DAD window open, main window Minimized"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   150
         TabIndex        =   89
         Top             =   3090
         Width           =   5115
      End
      Begin VB.CheckBox cbIgnoreD 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Treat Dxx files as regular files (to allow copying to large media)"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   150
         TabIndex        =   65
         Top             =   2250
         Width           =   5115
      End
      Begin VB.CommandButton cmdShowLog 
         Appearance      =   0  'Flat
         Caption         =   "Show Log"
         Height          =   375
         Left            =   3720
         TabIndex        =   63
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox cbLog 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Log all commands"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   150
         TabIndex        =   29
         Top             =   600
         Value           =   1  'Checked
         Width           =   3405
      End
      Begin VB.ComboBox cboDefDst 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "frmOptions.frx":0540
         Left            =   2235
         List            =   "frmOptions.frx":0550
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   210
         Width           =   1485
      End
      Begin VB.CheckBox cbConfirmCreate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Confirm D64 Creation when no files selected"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   150
         TabIndex        =   4
         Top             =   1980
         Value           =   1  'Checked
         Width           =   5115
      End
      Begin VB.CheckBox cbAutoRefreshDir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Automatically refresh directory after write to floppy"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   150
         TabIndex        =   3
         Top             =   1710
         Value           =   1  'Checked
         Width           =   5115
      End
      Begin VB.CheckBox cbPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Preview shell commands"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   150
         TabIndex        =   2
         Top             =   915
         Width           =   3405
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Default Destination Mode:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame optFrame 
      BackColor       =   &H00404040&
      Caption         =   "NibTools Options"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   3555
      Index           =   6
      Left            =   5520
      TabIndex        =   31
      Top             =   8700
      Visible         =   0   'False
      Width           =   5445
      Begin VB.CheckBox cbNibPrompt 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Prompt to Confirm "
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   300
         TabIndex        =   101
         Top             =   510
         Width           =   2475
      End
      Begin VB.CheckBox cbWriteD64 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Write D64"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   300
         TabIndex        =   98
         ToolTipText     =   "Also write D64 files using NibWrite"
         Top             =   2010
         Width           =   1065
      End
      Begin VB.TextBox txtNibConv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1830
         TabIndex        =   95
         ToolTipText     =   "Custom Switches for NibConv"
         Top             =   3180
         Width           =   3495
      End
      Begin VB.TextBox txtNibWrite 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1830
         TabIndex        =   94
         ToolTipText     =   "Custom Switches for NibWrite"
         Top             =   2850
         Width           =   3495
      End
      Begin VB.TextBox txtNibRead 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1830
         TabIndex        =   93
         ToolTipText     =   "Custom Switches for NibRead"
         Top             =   2520
         Width           =   3495
      End
      Begin VB.CheckBox cbNibCustom 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Custom: NIBREAD:"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   120
         TabIndex        =   92
         ToolTipText     =   "Use Custom switches (overrides switches above)"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CheckBox cbNibArg 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "-s = Use 1571 Fast Serial"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   86
         Tag             =   "-s"
         Top             =   1650
         Width           =   2265
      End
      Begin VB.TextBox txtRetries 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   4650
         TabIndex        =   71
         Text            =   "40"
         Top             =   1950
         Width           =   315
      End
      Begin VB.CheckBox cbRetries 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "-e = Set Retries to:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   3000
         TabIndex        =   70
         Top             =   1890
         Width           =   1755
      End
      Begin VB.CheckBox cbNBZ 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Compress to NBZ"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   510
         TabIndex        =   69
         ToolTipText     =   "Use NBZ format"
         Top             =   990
         Width           =   1695
      End
      Begin VB.CheckBox cbCreateD64 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Create D64 files"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   300
         TabIndex        =   68
         ToolTipText     =   "Create D64 files (converted from NIB)"
         Top             =   1470
         Width           =   1545
      End
      Begin VB.CheckBox cbCreateG64 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Create G64 files"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   300
         TabIndex        =   67
         ToolTipText     =   "Create G64 files (converted from NIB)"
         Top             =   1230
         Width           =   1545
      End
      Begin VB.CheckBox cbCreateNIB 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Create NIB files"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   300
         TabIndex        =   66
         ToolTipText     =   "Create NIB files (direct)"
         Top             =   750
         Width           =   1575
      End
      Begin VB.TextBox txtNibETrk 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2490
         TabIndex        =   61
         Text            =   "40"
         Top             =   1770
         Width           =   345
      End
      Begin VB.TextBox txtNibSTrk 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1830
         TabIndex        =   60
         Text            =   "1"
         Top             =   1770
         Width           =   345
      End
      Begin VB.CheckBox cbNibSE 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Track Range:"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   510
         TabIndex        =   58
         Top             =   1770
         Width           =   1335
      End
      Begin VB.CheckBox cbNibArg 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "-d = Default densities"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   57
         Tag             =   "-d"
         Top             =   1440
         Width           =   2265
      End
      Begin VB.CheckBox cbNibArg 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "-c = Disable capacity adjust"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   56
         Tag             =   "-c"
         Top             =   1230
         Width           =   2415
      End
      Begin VB.CheckBox cbNibArg 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "-g = Reduce gaps"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   55
         Tag             =   "-g"
         Top             =   1020
         Width           =   2265
      End
      Begin VB.CheckBox cbNibArg 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "-F = Fix short tracks"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   54
         Tag             =   "-F"
         Top             =   810
         Width           =   2235
      End
      Begin VB.CheckBox cbNibArg 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "-k = Disable killer tracks"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   53
         Tag             =   "-k"
         Top             =   600
         Width           =   2235
      End
      Begin VB.CheckBox cbNibArg 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "-h = Half tracks"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   52
         Tag             =   "-h"
         Top             =   390
         Width           =   2235
      End
      Begin VB.CheckBox cbNibArg 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "-l = Limit to 40 tracks "
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   51
         Tag             =   "-l"
         Top             =   180
         Width           =   2235
      End
      Begin VB.TextBox txtNibOpt 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2880
         TabIndex        =   34
         ToolTipText     =   "Refer to NibTools readme for valid options!"
         Top             =   2220
         Width           =   2445
      End
      Begin VB.CheckBox cbUseNib 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "ENABLE!"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "This option requires 1541/71 with parallel cable"
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "NIBCONV:"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   12
         Left            =   330
         TabIndex        =   97
         Top             =   3210
         Width           =   1425
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "NIBWRITE:"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   11
         Left            =   330
         TabIndex        =   96
         Top             =   2880
         Width           =   1425
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Switches:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   19
         Left            =   2220
         TabIndex        =   62
         Top             =   210
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   18
         Left            =   2250
         TabIndex        =   59
         Top             =   1800
         Width           =   165
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Additional Options:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   13
         Left            =   1320
         TabIndex        =   33
         Top             =   2250
         Width           =   1530
      End
   End
   Begin VB.Frame optFrame 
      BackColor       =   &H00404040&
      Caption         =   "Local Paths"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2640
      Index           =   2
      Left            =   30
      TabIndex        =   17
      Top             =   7470
      Visible         =   0   'False
      Width           =   5445
      Begin VB.PictureBox cmdDestBrowse 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5040
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   119
         ToolTipText     =   "Browse"
         Top             =   1920
         Width           =   285
      End
      Begin VB.PictureBox cmdSrcBrowse 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5040
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   118
         ToolTipText     =   "Browse"
         Top             =   840
         Width           =   285
      End
      Begin VB.CheckBox cbLastPaths 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Remember PATHs (otherwise default to those specified below)"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Tag             =   "Remember Paths"
         Top             =   270
         Width           =   5055
      End
      Begin VB.CheckBox cbPathHistory 
         BackColor       =   &H00404040&
         Caption         =   "Auto add to History"
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   3360
         TabIndex        =   88
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1845
      End
      Begin VB.CommandButton cmdClearHistory 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Clear History"
         Height          =   315
         Left            =   2040
         TabIndex        =   87
         ToolTipText     =   "Clear drop-down history"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdDstCurrent 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Use Current"
         Height          =   315
         Left            =   720
         TabIndex        =   85
         Top             =   2250
         Width           =   1215
      End
      Begin VB.CommandButton cmdSrcCurrent 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Use Current"
         Height          =   315
         Left            =   720
         TabIndex        =   84
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox DefaultSrcPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   270
         TabIndex        =   19
         Text            =   "c:\"
         Top             =   840
         Width           =   4725
      End
      Begin VB.TextBox DefaultDstPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   270
         TabIndex        =   18
         Text            =   "c:\"
         Top             =   1920
         Width           =   4725
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default LEFT Directory:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   21
         Top             =   600
         Width           =   1740
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default RIGHT Directory:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1875
      End
   End
   Begin VB.Frame optFrame 
      BackColor       =   &H00404040&
      Caption         =   "VICE Emulator Options"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   3195
      Index           =   5
      Left            =   5520
      TabIndex        =   27
      Top             =   5460
      Visible         =   0   'False
      Width           =   5445
      Begin VB.ComboBox cboPRG 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "frmOptions.frx":0588
         Left            =   1950
         List            =   "frmOptions.frx":05AA
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   1890
         Width           =   2175
      End
      Begin VB.OptionButton OptPRGMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Use Load address to select proper emulation "
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   1
         Left            =   1230
         TabIndex        =   43
         Top             =   2250
         Width           =   3855
      End
      Begin VB.OptionButton OptPRGMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Run:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   0
         Left            =   1230
         TabIndex        =   41
         Top             =   1890
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.ComboBox cbo80 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "frmOptions.frx":05F1
         Left            =   1950
         List            =   "frmOptions.frx":0616
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1350
         Width           =   2175
      End
      Begin VB.ComboBox cbo71 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "frmOptions.frx":0665
         Left            =   1950
         List            =   "frmOptions.frx":068D
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   990
         Width           =   2175
      End
      Begin VB.ComboBox cbo64 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "frmOptions.frx":06E4
         Left            =   1950
         List            =   "frmOptions.frx":0706
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   630
         Width           =   2175
      End
      Begin VB.CheckBox cbUseVice 
         BackColor       =   &H00404040&
         Caption         =   "Enable for 'Dxx' and 'PRG' files"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   300
         Value           =   1  'Checked
         Width           =   2835
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "For PRG Files:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   17
         Left            =   150
         TabIndex        =   42
         Top             =   1920
         Width           =   1050
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "For D80/82 files run:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   16
         Left            =   150
         TabIndex        =   40
         Top             =   1410
         Width           =   1560
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "For D71/81 files run:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   15
         Left            =   150
         TabIndex        =   38
         Top             =   1050
         Width           =   1560
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "For D64 files run:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   14
         Left            =   150
         TabIndex        =   35
         Top             =   690
         Width           =   1320
      End
   End
   Begin VB.Frame optFrame 
      BackColor       =   &H00404040&
      Caption         =   "Batch Imaging"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2655
      Index           =   8
      Left            =   11040
      TabIndex        =   72
      Top             =   6120
      Visible         =   0   'False
      Width           =   5445
      Begin VB.CheckBox cbDouble 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   " Double-sided"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   3450
         TabIndex        =   82
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optBatchMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Numbered, starting at:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   81
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton optBatchMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Use Disk Label"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   80
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optBatchMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Manual Filename Entry"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   79
         Top             =   600
         Value           =   -1  'True
         Width           =   2085
      End
      Begin VB.CheckBox cbLogContents 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Log Disk contents"
         Enabled         =   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   360
         TabIndex        =   78
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CheckBox cbLogLabels 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Log Disk Labels"
         Enabled         =   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   360
         TabIndex        =   77
         Top             =   1950
         Width           =   1935
      End
      Begin VB.TextBox txtBatchFN 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2430
         TabIndex        =   76
         Text            =   "disk-###.d64"
         Top             =   1350
         Width           =   1935
      End
      Begin VB.TextBox txtStartNum 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2430
         TabIndex        =   74
         Text            =   "1"
         Top             =   1050
         Width           =   885
      End
      Begin VB.CheckBox cbUseBatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Enable Batch Disk Imaging"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   180
         TabIndex        =   73
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "( #=digit, %=Side a/b, ^=Side A/B,  *=Side 1/2 )"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   9
         Left            =   1650
         TabIndex        =   83
         Top             =   1650
         Width           =   3645
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Filename Format:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   10
         Left            =   1050
         TabIndex        =   75
         Top             =   1380
         Width           =   1320
      End
   End
   Begin VB.Frame optFrame 
      BackColor       =   &H00404040&
      Caption         =   "X-Cable"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2355
      Index           =   3
      Left            =   30
      TabIndex        =   7
      Top             =   10140
      Visible         =   0   'False
      Width           =   5445
      Begin VB.CheckBox cbUseFirstDrive 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Use First Detected Drive"
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   3030
         TabIndex        =   103
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox CheckNoWarpMode 
         BackColor       =   &H00404040&
         Caption         =   "Disable &Warp Mode for D64 Transfer"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   165
         TabIndex        =   16
         Top             =   2025
         Width           =   3150
      End
      Begin VB.ComboBox cboDriveNum 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "frmOptions.frx":074D
         Left            =   2040
         List            =   "frmOptions.frx":075D
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   270
         Width           =   735
      End
      Begin VB.OptionButton optXMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Serial 1 (Slow!)"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   660
         TabIndex        =   12
         Top             =   1110
         Width           =   1935
      End
      Begin VB.OptionButton optXMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Serial 2"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   660
         TabIndex        =   11
         ToolTipText     =   "This only works with one serial device connected."
         Top             =   1380
         Width           =   1935
      End
      Begin VB.OptionButton optXMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Parallel"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   3
         Left            =   660
         TabIndex        =   10
         ToolTipText     =   "Requires a XP1541/XP1571 cable"
         Top             =   1635
         Width           =   1935
      End
      Begin VB.OptionButton optXMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Original (Very Slow!)"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   660
         TabIndex        =   9
         Top             =   855
         Width           =   1935
      End
      Begin VB.OptionButton optXMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Auto (Recommended)"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   4
         Left            =   660
         TabIndex        =   8
         ToolTipText     =   "Let OpenCBM select the most efficient transfer mode"
         Top             =   600
         Value           =   -1  'True
         Width           =   1965
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Device Number:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   315
         Width           =   1815
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mode:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   14
         Top             =   600
         Width           =   450
      End
   End
   Begin VB.Frame optFrame 
      BackColor       =   &H00404040&
      Caption         =   "Filenames"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1905
      Index           =   7
      Left            =   11010
      TabIndex        =   45
      Top             =   4170
      Visible         =   0   'False
      Width           =   5445
      Begin VB.CheckBox cbFNEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Confirm/Edit new filename"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   420
         TabIndex        =   64
         Top             =   1500
         Width           =   3135
      End
      Begin VB.TextBox txtFNChr 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3180
         MaxLength       =   5
         TabIndex        =   50
         Text            =   "-"
         Top             =   1050
         Width           =   285
      End
      Begin VB.OptionButton optFNMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Leave as-is then prompt for new filename"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   48
         Top             =   600
         Value           =   -1  'True
         Width           =   4275
      End
      Begin VB.OptionButton optFNMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Remove invalid characters"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   47
         Top             =   840
         Width           =   2775
      End
      Begin VB.OptionButton optFNMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Replace invalid characters with:"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   420
         TabIndex        =   46
         Top             =   1080
         Width           =   2715
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "When CBM filename contains invalid DOS characters:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   49
         Top             =   300
         Width           =   4080
      End
   End
   Begin VB.ListBox lstOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   3540
      IntegralHeight  =   0   'False
      ItemData        =   "frmOptions.frx":076F
      Left            =   60
      List            =   "frmOptions.frx":078E
      TabIndex        =   30
      Top             =   90
      Width           =   1185
   End
   Begin VB.Frame optFrame 
      BackColor       =   &H00404040&
      Caption         =   "CBMLink"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1260
      Index           =   4
      Left            =   5520
      TabIndex        =   22
      Top             =   4170
      Visible         =   0   'False
      Width           =   5445
      Begin VB.ComboBox cboLinkDev 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "frmOptions.frx":07F3
         Left            =   2910
         List            =   "frmOptions.frx":080F
         Style           =   2  'Dropdown List
         TabIndex        =   25
         ToolTipText     =   "Select X Device Unit Number"
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txtConStr 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1650
         TabIndex        =   24
         Text            =   "serial 19200,com1"
         Top             =   720
         Width           =   3675
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Unit (device)/Drive Number:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   26
         Top             =   300
         Width           =   2715
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Connection String:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   23
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5220
      TabIndex        =   1
      Top             =   3690
      Width           =   1545
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2021 Steve J. Gray
' ====================================================
'
' frmOptions - Program Options Window
'
' Based on GUI4CBM4WIN. The following (between "/" lines) is the notice
' included with the GUI4CBM4WIN source code:
'
' ////////////////////////////////////////////////////////////////////
' Copyright (C) 2004-2005 Leif Bloomquist
' Copyright (C) 2006      Wolfgang Moser
'
' This software Is provided 'as-is', without any express or implied
' warranty. In no event will the authors be held liable for any damages
' arising from the use of this software.
'
' Permission is granted to anyone to use this software for any purpose,
' including commercial applications, and to alter it and redistribute it
' freely, subject to the following restrictions:
'
'     1. The origin of this software must not be misrepresented; you must
'        not claim that you wrote the original software. If you use this
'        software in a product, an acknowledgment in the product
'        documentation would be appreciated but is not required.
'
'     2. Altered source versions must be plainly marked as such, and must
'        not be misrepresented as being the original software.
'
'     3. This notice may not be removed or altered from any source
'        distribution.
'
'/////////////////////////////////////////////////////////////////////////

Option Explicit

Public USelect As Integer

Private IsOnTop As Boolean


'---- Form Load - Initialize some things
Private Sub Form_Load()
    Dim J As Integer, Tmp As String
    
    On Error Resume Next
    
    Me.Height = 4590                            'Set the form width and height
    Me.Width = 6915
    
    Me.AlwaysOnTop = True                       'Make the Options Window be ON TOP
    
    cboDefDst.ListIndex = DstMode               'Set the default destination tab
    cboLinkDev.ListIndex = 0                    'Set the CBMLink device#
    cboDriveNum.ListIndex = 0                   'Set the OpenCBM drive#
    lstOpt.ListIndex = 0
    
    LinkUnit = 8: LinkDrive = 0                 'For CBM-Link
    USelect = 0                                 'Currently Selected Utility Path
    
    SetTheme                                    'Set the Theme colours
    
End Sub

Private Sub cmdClearHistory_Click()
    
    KillFile HistoryFile
    
    frmMain.txtLocalDir(0).Clear
    frmMain.txtLocalDir(1).Clear
    
End Sub

Private Sub cmdOK_Click()
    
    SetConfigOptions
    SaveINI
    
    Me.Hide

End Sub

Private Sub cmdDstCurrent_Click()
    
    DefaultDstPath.Text = LocalDir(1)

End Sub

Private Sub cmdSrcCurrent_Click()
    
    DefaultSrcPath.Text = LocalDir(0)

End Sub

Public Sub SetConfigOptions()
    Dim J As Integer, Tmp As String, Tmp2 As String
    
    Tmp = ""
    
    '--- Set up Paths
    SetAllPaths
    
    '--- NIB Options
    'Start and End tracks
    If cbNibSE.value = 1 Then Tmp = "-S" & MyTrim(txtNibSTrk.Text) & " " & "-E" & MyTrim(txtNibETrk.Text) & " "
    
    '--- Retries
    If cbRetries.value = 1 Then Tmp = Tmp & "-e" & MyTrim(txtRetries.Text) & " "
    
    '--- General Switches
    For J = 0 To 7
        If cbNibArg(J).value = 1 Then Tmp = Tmp & cbNibArg(J).Tag & " "
    Next J
    
    '--- Additional Switches
    Tmp = MyTrim(Tmp & txtNibOpt.Text) & " "
    NIBstr = Tmp
  
    '-- FN stuff
    FNChr = txtFNChr.Text
    
    '-- Path History
    AddPathFlag = cbPathHistory.value
    
    Batch2Sided = (cbDouble.value = 1)
    UseLP = (cbLastPaths.value = 1)
    UseBatch = (cbUseBatch.value = 1)
    P00Flag = cbP00.value
    StartDAD = cbDAD.value
    LogAll = (frmOptions.cbLog.value = vbChecked)
    UseNIB = (cbUseNib.value = vbChecked)
    CreateNIB = (cbCreateNIB.value = vbChecked)
    UseNBZ = (cbNBZ.value = vbChecked)
    CreateG64 = (cbCreateG64.value = vbChecked)
    CreateD64 = (cbCreateD64.value = vbChecked)
    WriteD64 = (cbWriteD64.value = vbChecked)
    UseNibCustom = (cbNibCustom.value = vbChecked)
    NIBPrompt = (cbNibPrompt.value = vbChecked)
    
    DstMode = cboDefDst.ListIndex
       
    LocalDir(1) = DefaultDstPath.Text
    LocalDir(0) = DefaultSrcPath.Text

    DefaultSrcPath.Text = LocalDir(0)
    DefaultDstPath.Text = LocalDir(1)
    AutoRefreshDir = (cbAutoRefreshDir.value = vbChecked)
    ConfirmD64 = (cbConfirmCreate.value = vbChecked)
    PreviewCheck = (cbPreview.value = vbChecked)
    IgnoreD = (cbIgnoreD.value = vbChecked)
    FNEdit = (cbFNEdit.value = vbChecked)
    
    UseVice = cbUseVice.value
    
    IgnoreBadID = (cbIgnoreBadID.value = vbChecked)
End Sub

Private Sub cmdShowLog_Click()
    
    If Exists(LogFile) = True Then
        ViewFile LogFile
    Else
        MyMsg "There is no log file yet."
    End If

End Sub

Private Sub cmdClearLog_Click()
    
    KillFile LogFile

End Sub

Private Sub cmdSrcBrowse_Click()
    
    Dim Tmp As String
    Tmp = GetBrowseDir(Me, "Select Default Source Path:")
    If Tmp <> "" Then DefaultSrcPath.Text = AddSlash(Tmp)

End Sub

Private Sub cmdDestBrowse_Click()
    
    Dim Tmp As String
    Tmp = GetBrowseDir(Me, "Select Default Destination Path:")
    If Tmp <> "" Then DefaultDstPath.Text = AddSlash(Tmp)

End Sub

'Private Sub cmdViceBrowse_Click()
'    Dim Tmp As String
'    Tmp = GetBrowseDir(Me, "Select Path containing VICE executables:")
'    If Tmp <> "" Then txtVicePath.Text = AddSlash(Tmp)
'End Sub
Private Sub CheckNoWarpMode_Click()
    
    If CheckNoWarpMode.value = vbChecked Then
        NoWarpString = "--no-warp"
    Else
        NoWarpString = ""
    End If

End Sub

Public Property Let AlwaysOnTop(ByVal bState As Boolean)
  
  Dim lFlag As Long
  If bState Then lFlag = HWND_TOPMOST Else lFlag = HWND_NOTOPMOST
  IsOnTop = bState
  SetWindowPos Me.hWnd, lFlag, 0&, 0&, 0&, 0&, (SWP_NOSIZE Or SWP_NOMOVE)

End Property

Private Sub lstOpt_Click()
    Dim J As Integer
    
    For J = 0 To 8: optFrame(J).Visible = False: Next
    J = lstOpt.ListIndex
    If J > 0 Then
        optFrame(J).Left = optFrame(0).Left
        optFrame(J).Top = optFrame(0).Top
        optFrame(J).Width = optFrame(0).Width
        optFrame(J).Height = optFrame(0).Height
    End If
    optFrame(J).Visible = True
    DoEvents
    
End Sub

Private Sub optBatchMode_Click(Index As Integer)
    
    BatchMode = Index

End Sub

Private Sub optFNMode_Click(Index As Integer)
    
    FNMode = Index

End Sub

'---- Set X-Cable Transfer Mode String
Private Sub optXMode_Click(Index As Integer)
    
    Select Case Index
        Case 0: TransferString = "original"
        Case 1: TransferString = "serial1"
        Case 2: TransferString = "serial2"
        Case 3: TransferString = "parallel"
        Case 4: TransferString = "auto"
    End Select

End Sub

Private Sub txtConStr_Change()

    frmMain.SetLinkString
    
End Sub

Private Sub txtStartNum_Change()

    DiskNum = Val(txtStartNum.Text)

End Sub

'---- Set Path when user presses ENTER
Private Sub txtUpath_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SetUPath txtUpath.Text                          'Use the textbox string
        KeyAscii = 0                                    'Clear key to avoid system sound
    End If

End Sub

'---- Click on Utility Path string to select it for editing
Private Sub lblUPath_Click(Index As Integer)
    
    optUPath(Index).value = True

End Sub

'---- Set the Current Path when radio button is clicked
Private Sub optUPath_Click(Index As Integer)
    
    txtUpath.Text = lblUPath(Index).Caption
    USelect = Index                                     'Remember which path we are working with

End Sub

'---- Allow user to select a Folder/Path
Private Sub cmdUBrowse_Click()
    
    Dim Tmp As String
    
    Tmp = GetBrowseDir(Me, "Select Utility Path:")
    If Tmp <> "" Then SetUPath Tmp

End Sub

'---- Set the Utility Path for the Selected Path
' A "\" is appended if needed
Private Sub SetUPath(ByVal Tmp As String)

    If Tmp <> "" Then Tmp = AddSlash(Tmp)               'Make sure there is a \ in the path name (unless blank)
    txtUpath.Text = Tmp                                 'Set the Editing path
    lblUPath(USelect).Caption = Tmp                     'Set the actual path
    UPath(USelect) = Tmp                                'Set Global
    
End Sub

'---- Set Theme
Public Sub SetTheme()
    Dim i As Integer, Y As Integer
    
    frmOptions.BackColor = ThemeBG
    lstOpt.BackColor = ThemeListBG: lstOpt.ForeColor = ThemeListFG
    
    For i = 0 To 8
        optFrame(i).BackColor = ThemeFrBG: optFrame(i).ForeColor = ThemeFrFG
    Next
          
    For i = 0 To 18
        Label(i).ForeColor = ThemeFrFG
    Next
    
    '--- General Options
    
    cboDefDst.BackColor = ThemeListBG:      cboDefDst.ForeColor = ThemeListFG
    
    cbLog.BackColor = ThemeFrBG:            cbLog.ForeColor = ThemeFrFG
    cbPreview.BackColor = ThemeFrBG:        cbPreview.ForeColor = ThemeFrFG
    cbP00.BackColor = ThemeFrBG:            cbP00.ForeColor = ThemeFrFG
    cbAutoRefreshDir.BackColor = ThemeFrBG: cbAutoRefreshDir.ForeColor = ThemeFrFG
    cbConfirmCreate.BackColor = ThemeFrBG:  cbConfirmCreate.ForeColor = ThemeFrFG
    cbIgnoreD.BackColor = ThemeFrBG:        cbIgnoreD.ForeColor = ThemeFrFG
    cbIgnoreBadID.BackColor = ThemeFrBG:    cbIgnoreBadID.ForeColor = ThemeFrFG
    cbErr.BackColor = ThemeFrBG:            cbErr.ForeColor = ThemeFrFG
    cbDAD.BackColor = ThemeFrBG:            cbDAD.ForeColor = ThemeFrFG
    
    '--- Utility Paths
    
    For i = 0 To 4
        optUPath(i).BackColor = ThemeFrBG:  optUPath(i).ForeColor = ThemeFrFG
        lblUPath(i).BackColor = ThemeFrBG:  lblUPath(i).ForeColor = ThemeFrFG
    Next i
    
    txtUpath.BackColor = ThemeListBG:       txtUpath.ForeColor = ThemeListFG
 
    '--- Local Paths
    
    cbLastPaths.BackColor = ThemeFrBG:      cbLastPaths.ForeColor = ThemeFrFG
    cbPathHistory.BackColor = ThemeFrBG:    cbPathHistory.ForeColor = ThemeFrFG
    DefaultSrcPath.BackColor = ThemeListBG: DefaultSrcPath.ForeColor = ThemeListFG
    DefaultDstPath.BackColor = ThemeListBG: DefaultDstPath.ForeColor = ThemeListFG
    
    '--- X-Cable
    
    For i = 0 To 4
        optXMode(i).BackColor = ThemeFrBG:  optXMode(i).ForeColor = ThemeFrFG
    Next i
    
    cboDriveNum.BackColor = ThemeListBG:    cboDriveNum.ForeColor = ThemeListFG
    cbUseFirstDrive.BackColor = ThemeFrBG:  cbUseFirstDrive.ForeColor = ThemeFrFG
    CheckNoWarpMode.BackColor = ThemeFrBG:  CheckNoWarpMode.ForeColor = ThemeFrFG
    
    '--- CBMLink
    
    cboLinkDev.BackColor = ThemeListBG:     cboLinkDev.ForeColor = ThemeListFG
    txtConStr.BackColor = ThemeListBG:      txtConStr.ForeColor = ThemeListFG
    
    '--- VICE
    
    cbo64.BackColor = ThemeListBG:          cbo64.ForeColor = ThemeListFG
    cbo71.BackColor = ThemeListBG:          cbo71.ForeColor = ThemeListFG
    cbo80.BackColor = ThemeListBG:          cbo80.ForeColor = ThemeListFG
    cboPRG.BackColor = ThemeListBG:         cboPRG.ForeColor = ThemeListFG
    
    For i = 0 To 1
        OptPRGMode(i).BackColor = ThemeFrBG: OptPRGMode(i).ForeColor = ThemeFrFG
    Next i
    
    cbUseVice.BackColor = ThemeFrBG:        cbUseVice.ForeColor = ThemeFrFG
    
    '--- NIBTOOLS
    
    For i = 0 To 7
        cbNibArg(i).BackColor = ThemeFrBG:  cbNibArg(i).ForeColor = ThemeFrFG
    Next i
    
    cbUseNib.BackColor = ThemeFrBG:         cbUseNib.ForeColor = ThemeFrFG
    cbNibPrompt.BackColor = ThemeFrBG:      cbNibPrompt.ForeColor = ThemeFrFG
    cbCreateNIB.BackColor = ThemeFrBG:      cbCreateNIB.ForeColor = ThemeFrFG
    cbNBZ.BackColor = ThemeFrBG:            cbNBZ.ForeColor = ThemeFrFG
    cbCreateG64.BackColor = ThemeFrBG:      cbCreateG64.ForeColor = ThemeFrFG
    cbCreateD64.BackColor = ThemeFrBG:      cbCreateD64.ForeColor = ThemeFrFG
    cbNibSE.BackColor = ThemeFrBG:          cbNibSE.ForeColor = ThemeFrFG
    cbWriteD64.BackColor = ThemeFrBG:       cbWriteD64.ForeColor = ThemeFrFG
    cbRetries.BackColor = ThemeFrBG:        cbRetries.ForeColor = ThemeFrFG
    cbNibCustom.BackColor = ThemeFrBG:      cbNibCustom.ForeColor = ThemeFrFG
    
    txtNibSTrk.BackColor = ThemeListBG:     txtNibSTrk.ForeColor = ThemeListFG
    txtNibETrk.BackColor = ThemeListBG:     txtNibETrk.ForeColor = ThemeListFG
    
    txtNibOpt.BackColor = ThemeListBG:      txtNibOpt.ForeColor = ThemeListFG
    txtNibRead.BackColor = ThemeListBG:     txtNibRead.ForeColor = ThemeListFG
    txtNibWrite.BackColor = ThemeListBG:    txtNibWrite.ForeColor = ThemeListFG
    txtNibConv.BackColor = ThemeListBG:     txtNibConv.ForeColor = ThemeListFG
    txtRetries.BackColor = ThemeListBG:     txtRetries.ForeColor = ThemeListFG
    
    '--- Filenames
    
    For i = 0 To 2
        optFNMode(i).BackColor = ThemeFrBG: optFNMode(i).ForeColor = ThemeFrFG
    Next i
    
    cbFNEdit.BackColor = ThemeFrBG:         cbFNEdit.ForeColor = ThemeFrFG
    txtFNChr.BackColor = ThemeListBG:       txtFNChr.ForeColor = ThemeListFG
    
    '--- Batch Imaging
    
    For i = 0 To 2
        optBatchMode(i).BackColor = ThemeFrBG:  optBatchMode(i).ForeColor = ThemeFrFG
    Next i
    
    cbUseBatch.BackColor = ThemeFrBG:       cbUseBatch.ForeColor = ThemeFrFG
    cbLogLabels.BackColor = ThemeFrBG:      cbLogLabels.ForeColor = ThemeFrFG
    cbLogContents.BackColor = ThemeFrBG:    cbLogContents.ForeColor = ThemeFrFG
    cbDouble.BackColor = ThemeFrBG:         cbDouble.ForeColor = ThemeFrFG
    
    txtStartNum.BackColor = ThemeListBG:    txtStartNum.ForeColor = ThemeListFG
    txtBatchFN.BackColor = ThemeListBG:     txtBatchFN.ForeColor = ThemeListFG
       
    '--- Icons
    
    Y = 67
    
    frmMain.GetIcon cmdUBrowse, 225, Y
    frmMain.GetIcon cmdSrcBrowse, 225, Y
    frmMain.GetIcon cmdDestBrowse, 225, Y
    DoEvents
    
End Sub











