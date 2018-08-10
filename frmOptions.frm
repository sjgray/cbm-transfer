VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CBM-Transfer Options"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6825
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame optFrame 
      Caption         =   "NibTools Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Index           =   5
      Left            =   6990
      TabIndex        =   36
      Top             =   60
      Visible         =   0   'False
      Width           =   5445
      Begin VB.CheckBox cbWriteD64 
         Caption         =   "Write D64"
         Height          =   315
         Left            =   450
         TabIndex        =   108
         ToolTipText     =   "Also write D64 files using NibWrite"
         Top             =   2010
         Width           =   1065
      End
      Begin VB.TextBox txtNibConv 
         Height          =   285
         Left            =   1830
         TabIndex        =   105
         ToolTipText     =   "Custom Switches for NibConv"
         Top             =   3180
         Width           =   3495
      End
      Begin VB.TextBox txtNibWrite 
         Height          =   285
         Left            =   1830
         TabIndex        =   104
         ToolTipText     =   "Custom Switches for NibWrite"
         Top             =   2850
         Width           =   3495
      End
      Begin VB.TextBox txtNibRead 
         Height          =   285
         Left            =   1830
         TabIndex        =   103
         ToolTipText     =   "Custom Switches for NibRead"
         Top             =   2520
         Width           =   3495
      End
      Begin VB.CheckBox cbNibCustom 
         Caption         =   "Custom: NIBREAD:"
         Height          =   315
         Left            =   120
         TabIndex        =   102
         ToolTipText     =   "Use Custom switches"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CheckBox cbNibArg 
         Caption         =   "-s = Use 1571 Fast Serial"
         Height          =   255
         Index           =   7
         Left            =   2880
         TabIndex        =   94
         Tag             =   "-s"
         Top             =   1650
         Width           =   2415
      End
      Begin VB.TextBox txtRetries 
         Height          =   315
         Left            =   4560
         TabIndex        =   77
         Text            =   "40"
         Top             =   1890
         Width           =   315
      End
      Begin VB.CheckBox cbRetries 
         Caption         =   "-e = Set Retries to:"
         Height          =   315
         Left            =   2880
         TabIndex        =   76
         Top             =   1890
         Width           =   1755
      End
      Begin VB.CheckBox cbNBZ 
         Caption         =   "Compress to NBZ"
         Height          =   315
         Left            =   810
         TabIndex        =   75
         ToolTipText     =   "Use NBZ format"
         Top             =   690
         Width           =   1575
      End
      Begin VB.CheckBox cbCreateD64 
         Caption         =   "Create D64 files"
         Height          =   315
         Left            =   420
         TabIndex        =   74
         ToolTipText     =   "Create D64 files (converted from NIB)"
         Top             =   1170
         Width           =   1575
      End
      Begin VB.CheckBox cbCreateG64 
         Caption         =   "Create G64 files"
         Height          =   315
         Left            =   420
         TabIndex        =   73
         ToolTipText     =   "Create G64 files (converted from NIB)"
         Top             =   930
         Width           =   1575
      End
      Begin VB.CheckBox cbCreateNIB 
         Caption         =   "Create NIB files"
         Height          =   315
         Left            =   420
         TabIndex        =   72
         ToolTipText     =   "Create NIB files (direct)"
         Top             =   450
         Width           =   1575
      End
      Begin VB.TextBox txtNibETrk 
         Height          =   315
         Left            =   2310
         TabIndex        =   66
         Text            =   "40"
         Top             =   1770
         Width           =   345
      End
      Begin VB.TextBox txtNibSTrk 
         Height          =   285
         Left            =   2310
         TabIndex        =   65
         Text            =   "1"
         Top             =   1440
         Width           =   345
      End
      Begin VB.CheckBox cbNibSE 
         Caption         =   "Start Track:"
         Height          =   255
         Left            =   1140
         TabIndex        =   63
         Top             =   1470
         Width           =   1275
      End
      Begin VB.CheckBox cbNibArg 
         Caption         =   "-d = Default densities"
         Height          =   255
         Index           =   6
         Left            =   2880
         TabIndex        =   62
         Tag             =   "-d"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CheckBox cbNibArg 
         Caption         =   "-c = Disable capacity adjust"
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   61
         Tag             =   "-c"
         Top             =   1230
         Width           =   2415
      End
      Begin VB.CheckBox cbNibArg 
         Caption         =   "-g = Reduce gaps"
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   60
         Tag             =   "-g"
         Top             =   1020
         Width           =   2415
      End
      Begin VB.CheckBox cbNibArg 
         Caption         =   "-F = Fix short tracks"
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   59
         Tag             =   "-F"
         Top             =   810
         Width           =   2415
      End
      Begin VB.CheckBox cbNibArg 
         Caption         =   "-k = Disable killer tracks"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   58
         Tag             =   "-k"
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox cbNibArg 
         Caption         =   "-h = Half tracks"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   57
         Tag             =   "-h"
         Top             =   390
         Width           =   2415
      End
      Begin VB.CheckBox cbNibArg 
         Caption         =   "-l = Limit to 40 tracks "
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   56
         Tag             =   "-l"
         Top             =   180
         Width           =   2415
      End
      Begin VB.TextBox txtNibOpt 
         Height          =   285
         Left            =   2880
         TabIndex        =   39
         ToolTipText     =   "Refer to NibTools readme for valid options!"
         Top             =   2220
         Width           =   2445
      End
      Begin VB.CheckBox cbUseNib 
         Caption         =   "ENABLE!"
         Height          =   315
         Left            =   120
         TabIndex        =   37
         ToolTipText     =   "This option requires 1541/71 with parallel cable"
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "NIBCONV:"
         Height          =   255
         Left            =   330
         TabIndex        =   107
         Top             =   3210
         Width           =   1425
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "NIBWRITE:"
         Height          =   255
         Left            =   330
         TabIndex        =   106
         Top             =   2880
         Width           =   1425
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Switches:"
         Height          =   195
         Left            =   2130
         TabIndex        =   67
         Top             =   210
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "End Track:"
         Height          =   195
         Left            =   1470
         TabIndex        =   64
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Additional Options:"
         Height          =   195
         Left            =   1530
         TabIndex        =   38
         Top             =   2250
         Width           =   1320
      End
   End
   Begin VB.Frame optFrame 
      Caption         =   "Paths"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   4260
      Visible         =   0   'False
      Width           =   5445
      Begin VB.CheckBox cbLastPaths 
         Caption         =   "Remember PATHs (otherwise default to those specified below)"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Tag             =   "Remember Paths"
         Top             =   240
         Width           =   5055
      End
      Begin VB.CheckBox cbPathHistory 
         Caption         =   "Auto add to History"
         Height          =   375
         Left            =   3360
         TabIndex        =   96
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CommandButton cmdClearHistory 
         Caption         =   "Clear History"
         Height          =   315
         Left            =   2040
         TabIndex        =   95
         ToolTipText     =   "Clear drop-down history"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdDstCurrent 
         Caption         =   "Use Current"
         Height          =   315
         Left            =   720
         TabIndex        =   93
         Top             =   2250
         Width           =   1215
      End
      Begin VB.CommandButton cmdSrcCurrent 
         Caption         =   "Use Current"
         Height          =   315
         Left            =   720
         TabIndex        =   92
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdDestBrowse 
         Caption         =   "..."
         Height          =   285
         Left            =   4830
         TabIndex        =   79
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton cmdSrcBrowse 
         Caption         =   "..."
         Height          =   285
         Left            =   4830
         TabIndex        =   78
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox DefaultSrcPath 
         Height          =   285
         Left            =   270
         TabIndex        =   19
         Text            =   "c:\"
         Top             =   840
         Width           =   4515
      End
      Begin VB.TextBox DefaultDstPath 
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Text            =   "c:\"
         Top             =   1920
         Width           =   4545
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default LEFT Directory:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   21
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default RIGHT Directory:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1785
      End
   End
   Begin VB.Frame optFrame 
      Caption         =   "Fonts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   8
      Left            =   11340
      TabIndex        =   98
      Top             =   6990
      Width           =   5445
      Begin VB.CheckBox cbUseCBMFont 
         Caption         =   "Use 'C64 User Mono' font for CBM file lists"
         Height          =   435
         Left            =   180
         TabIndex        =   99
         Top             =   300
         Width           =   3375
      End
   End
   Begin VB.Frame optFrame 
      Caption         =   "VICE Emulator Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Index           =   4
      Left            =   5700
      TabIndex        =   29
      Top             =   6570
      Visible         =   0   'False
      Width           =   5445
      Begin VB.CommandButton cmdViceBrowse 
         Caption         =   "..."
         Height          =   285
         Left            =   4830
         TabIndex        =   68
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cboPRG 
         Height          =   315
         ItemData        =   "frmOptions.frx":0442
         Left            =   1920
         List            =   "frmOptions.frx":0464
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   2400
         Width           =   2175
      End
      Begin VB.OptionButton OptPRGMode 
         Caption         =   "Use Load address to select proper emulation "
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   48
         Top             =   2700
         Width           =   3495
      End
      Begin VB.OptionButton OptPRGMode 
         Caption         =   "Run:"
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   46
         Top             =   2400
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.ComboBox cbo80 
         Height          =   315
         ItemData        =   "frmOptions.frx":04AB
         Left            =   1920
         List            =   "frmOptions.frx":04D0
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   1860
         Width           =   2175
      End
      Begin VB.ComboBox cbo71 
         Height          =   315
         ItemData        =   "frmOptions.frx":051F
         Left            =   1920
         List            =   "frmOptions.frx":0547
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1500
         Width           =   2175
      End
      Begin VB.ComboBox cbo64 
         Height          =   315
         ItemData        =   "frmOptions.frx":059E
         Left            =   1920
         List            =   "frmOptions.frx":05C0
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1140
         Width           =   2175
      End
      Begin VB.TextBox txtVicePath 
         Height          =   315
         Left            =   1620
         TabIndex        =   32
         Text            =   "C:\Program Files\WinVICE-2.1\"
         ToolTipText     =   "(do not include a filename!)"
         Top             =   600
         Width           =   3195
      End
      Begin VB.CheckBox cbUseVice 
         Caption         =   "Enable for 'Dxx' and 'PRG' files"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   300
         Value           =   1  'Checked
         Width           =   2835
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "For PRG Files:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   2430
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "For D80/82 files run:"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "For D71/81 files run:"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "For D64 files run:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Path to VICE folder:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   660
         Width           =   1395
      End
   End
   Begin VB.Frame optFrame 
      Caption         =   "Batch Imaging"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Index           =   7
      Left            =   11310
      TabIndex        =   80
      Top             =   4260
      Visible         =   0   'False
      Width           =   5445
      Begin VB.CheckBox cbDouble 
         Caption         =   " Double-sided"
         Height          =   255
         Left            =   3360
         TabIndex        =   90
         Top             =   1020
         Width           =   1755
      End
      Begin VB.OptionButton optBatchMode 
         Caption         =   "Numbered, starting at:"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   89
         Top             =   1080
         Width           =   1875
      End
      Begin VB.OptionButton optBatchMode 
         Caption         =   "Use Disk Label"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   88
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optBatchMode 
         Caption         =   "Manual Filename Entry"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   87
         Top             =   600
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.CheckBox cbLogContents 
         Caption         =   "Log Disk contents"
         Enabled         =   0   'False
         Height          =   195
         Left            =   360
         TabIndex        =   86
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CheckBox cbLogLabels 
         Caption         =   "Log Disk Labels"
         Enabled         =   0   'False
         Height          =   315
         Left            =   360
         TabIndex        =   85
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtBatchFN 
         Height          =   285
         Left            =   2280
         TabIndex        =   84
         Text            =   "disk-###.d64"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtStartNum 
         Height          =   285
         Left            =   2280
         TabIndex        =   82
         Text            =   "1"
         Top             =   1020
         Width           =   795
      End
      Begin VB.CheckBox cbUseBatch 
         Caption         =   "Enable Batch Disk Imaging"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "( #=digit, %=Side a/b, ^=Side A/B,  *=Side 1/2 )"
         Height          =   195
         Left            =   1800
         TabIndex        =   91
         Top             =   1620
         Width           =   3390
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Filename Format:"
         Height          =   195
         Left            =   960
         TabIndex        =   83
         Top             =   1380
         Width           =   1200
      End
   End
   Begin VB.Frame optFrame 
      Caption         =   "X-Cable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2590
      Index           =   2
      Left            =   150
      TabIndex        =   7
      Top             =   7170
      Visible         =   0   'False
      Width           =   5445
      Begin VB.CommandButton cmdResetBus 
         Caption         =   "&Reset Bus"
         Height          =   375
         Left            =   3000
         TabIndex        =   24
         ToolTipText     =   "Reset all device on the IEC bus"
         Top             =   1320
         Width           =   1170
      End
      Begin VB.CommandButton cmdDetect 
         Caption         =   "&Detect Drive"
         Height          =   375
         Left            =   3000
         TabIndex        =   23
         ToolTipText     =   "Detect all currently active device on the IEC bus"
         Top             =   840
         Width           =   1170
      End
      Begin VB.CheckBox CheckNoWarpMode 
         Caption         =   "Disable &Warp Mode for D64 Transfer"
         Height          =   255
         Left            =   165
         TabIndex        =   16
         Top             =   1935
         Width           =   3150
      End
      Begin VB.ComboBox cboDriveNum 
         Height          =   315
         ItemData        =   "frmOptions.frx":0607
         Left            =   1920
         List            =   "frmOptions.frx":0617
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   255
         Width           =   855
      End
      Begin VB.OptionButton optXMode 
         Caption         =   "Serial 1 (Slow!)"
         Height          =   255
         Index           =   1
         Left            =   660
         TabIndex        =   12
         Top             =   1110
         Width           =   1935
      End
      Begin VB.OptionButton optXMode 
         Caption         =   "Serial 2"
         Height          =   255
         Index           =   2
         Left            =   660
         TabIndex        =   11
         ToolTipText     =   "This only works with one serial device connected."
         Top             =   1380
         Width           =   1935
      End
      Begin VB.OptionButton optXMode 
         Caption         =   "Parallel"
         Height          =   255
         Index           =   3
         Left            =   660
         TabIndex        =   10
         ToolTipText     =   "Requires a XP1541/XP1571 cable"
         Top             =   1635
         Width           =   1935
      End
      Begin VB.OptionButton optXMode 
         Caption         =   "Original (Very Slow!)"
         Height          =   255
         Index           =   0
         Left            =   660
         TabIndex        =   9
         Top             =   855
         Width           =   1935
      End
      Begin VB.OptionButton optXMode 
         Caption         =   "Auto (Recommended)"
         Height          =   255
         Index           =   4
         Left            =   660
         TabIndex        =   8
         ToolTipText     =   "Let OpenCBM select the most efficient transfer mode"
         Top             =   600
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Device Number:"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   315
         Width           =   1710
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mode:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   14
         Top             =   600
         Width           =   450
      End
   End
   Begin VB.Frame optFrame 
      Caption         =   "Filenames"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   6
      Left            =   11370
      TabIndex        =   50
      Top             =   8070
      Visible         =   0   'False
      Width           =   5445
      Begin VB.CheckBox cbFNEdit 
         Caption         =   "Confirm/Edit new filename"
         Height          =   255
         Left            =   420
         TabIndex        =   70
         Top             =   1500
         Width           =   3135
      End
      Begin VB.TextBox txtFNChr 
         Height          =   315
         Left            =   3060
         MaxLength       =   5
         TabIndex        =   55
         Text            =   "-"
         Top             =   1020
         Width           =   375
      End
      Begin VB.OptionButton optFNMode 
         Caption         =   "Leave as-is then prompt for new filename"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   53
         Top             =   600
         Value           =   -1  'True
         Width           =   4275
      End
      Begin VB.OptionButton optFNMode 
         Caption         =   "Remove invalid characters"
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   52
         Top             =   840
         Width           =   2775
      End
      Begin VB.OptionButton optFNMode 
         Caption         =   "Replace invalid characters with:"
         Height          =   255
         Index           =   2
         Left            =   420
         TabIndex        =   51
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "When CBM filename contains invalid DOS characters:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   54
         Top             =   300
         Width           =   3825
      End
   End
   Begin VB.ListBox lstOpt 
      BackColor       =   &H00FFFF00&
      Height          =   3960
      ItemData        =   "frmOptions.frx":0629
      Left            =   60
      List            =   "frmOptions.frx":0648
      TabIndex        =   35
      Top             =   60
      Width           =   1155
   End
   Begin VB.Frame optFrame 
      Caption         =   "CBMLink"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Index           =   3
      Left            =   5700
      TabIndex        =   22
      Top             =   4260
      Visible         =   0   'False
      Width           =   5445
      Begin VB.ComboBox cboLinkDev 
         Height          =   315
         ItemData        =   "frmOptions.frx":069F
         Left            =   2760
         List            =   "frmOptions.frx":06BB
         Style           =   2  'Dropdown List
         TabIndex        =   27
         ToolTipText     =   "Select X Device Unit Number"
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txtConStr 
         Height          =   285
         Left            =   1530
         TabIndex        =   26
         Text            =   "serial 19200,com1"
         Top             =   690
         Width           =   3405
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Unit (device)/Drive Number:"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   28
         Top             =   300
         Width           =   2550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Connection String:"
         Height          =   195
         Left            =   135
         TabIndex        =   25
         Top             =   720
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5250
      TabIndex        =   1
      Top             =   3690
      Width           =   1545
   End
   Begin VB.Frame optFrame 
      Caption         =   "General Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3555
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   60
      Width           =   5445
      Begin VB.CheckBox cbIgnoreBadID 
         Caption         =   "Ignore BAD disk ID's when imaging D64"
         Height          =   375
         Left            =   150
         TabIndex        =   110
         Top             =   2520
         Width           =   4935
      End
      Begin VB.CheckBox cbErr 
         Caption         =   "For results button, also show ERROR file"
         Height          =   375
         Left            =   150
         TabIndex        =   109
         Top             =   2790
         Width           =   4935
      End
      Begin VB.CheckBox cbP00 
         Caption         =   "&Write P00 files when copying from Disk Images"
         Height          =   255
         Left            =   150
         TabIndex        =   101
         Top             =   1500
         Width           =   4935
      End
      Begin VB.CheckBox cbDAD 
         Caption         =   "Start with DAD window open, main window Minimized"
         Height          =   375
         Left            =   150
         TabIndex        =   97
         Top             =   3090
         Width           =   4935
      End
      Begin VB.CheckBox cbIgnoreD 
         Caption         =   "Treat Dxx files as regular files (to allow copying to large media)"
         Height          =   375
         Left            =   150
         TabIndex        =   71
         Top             =   2250
         Width           =   4935
      End
      Begin VB.CommandButton cmdShowLog 
         Caption         =   "Show Log"
         Height          =   375
         Left            =   3720
         TabIndex        =   69
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox cbLog 
         Caption         =   "&Log all commands"
         Height          =   375
         Left            =   150
         TabIndex        =   34
         Top             =   600
         Value           =   1  'Checked
         Width           =   3945
      End
      Begin VB.CheckBox cbCheckEXE 
         Caption         =   "&Check for EXE's in CBMXfer folder"
         Height          =   375
         Left            =   150
         TabIndex        =   33
         Top             =   870
         Value           =   1  'Checked
         Width           =   3945
      End
      Begin VB.ComboBox cboDefDst 
         Height          =   315
         ItemData        =   "frmOptions.frx":0702
         Left            =   2025
         List            =   "frmOptions.frx":0712
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   210
         Width           =   1485
      End
      Begin VB.CheckBox cbConfirmCreate 
         Caption         =   "&Confirm D64 Creation when no files selected"
         Height          =   375
         Left            =   150
         TabIndex        =   4
         Top             =   1980
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin VB.CheckBox cbAutoRefreshDir 
         Caption         =   "&Automatically refresh directory after write to floppy"
         Height          =   375
         Left            =   150
         TabIndex        =   3
         Top             =   1710
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin VB.CheckBox cbPreview 
         Caption         =   "&Preview shell commands"
         Height          =   345
         Left            =   150
         TabIndex        =   2
         Top             =   1185
         Width           =   3945
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Destination Mode:"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   5
         Top             =   255
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2017 Steve J. Gray
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

Private IsOnTop As Boolean

Private Sub Form_Load()
    Dim j As Integer, Tmp As String
    
    On Error Resume Next
    
    Me.AlwaysOnTop = True
    cboDefDst.ListIndex = DstMode
    cboLinkDev.ListIndex = 0
    cboDriveNum.ListIndex = 0
    lstOpt.ListIndex = 0
    
    CBMUnit = 8: CBMDrive = 0
    
End Sub

Private Sub cmdClearHistory_Click()
    KillFile PathHistory
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
    Dim j As Integer, Tmp As String, Tmp2 As String
    
    Tmp = ""
    '--- NIB Options
    'Start and End tracks
    If cbNibSE.value = 1 Then Tmp = "-S" & MyTrim(txtNibSTrk.Text) & " " & "-E" & MyTrim(txtNibETrk.Text) & " "
    
    '--- Retries
    If cbRetries.value = 1 Then Tmp = Tmp & "-e" & MyTrim(txtRetries.Text) & " "
    
    '--- General Switches
    For j = 0 To 7
        If cbNibArg(j).value = 1 Then Tmp = Tmp & cbNibArg(j).Tag & " "
    Next j
    
    '--- Additional Switches
    Tmp = MyTrim(Tmp & txtNibOpt.Text) & " "
    NIBstr = Tmp
  
    '-- FN stuff
    FNChr = txtFNChr.Text
    
    '-- Path History
    PathHistory = cbPathHistory.value
    
    CheckEXE = (frmOptions.cbCheckEXE.value = vbChecked)
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
    DstMode = cboDefDst.ListIndex
    UseVice = cbUseVice.value
    LocalDir(1) = DefaultDstPath.Text
    LocalDir(0) = DefaultSrcPath.Text
    DriveNum = Val(cboDriveNum.List(cboDriveNum.ListIndex))
    DefaultSrcPath.Text = LocalDir(0)
    DefaultDstPath.Text = LocalDir(1)
    AutoRefreshDir = (cbAutoRefreshDir.value = vbChecked)
    ConfirmD64 = (cbConfirmCreate.value = vbChecked)
    PreviewCheck = (cbPreview.value = vbChecked)
    IgnoreD = (cbIgnoreD.value = vbChecked)
    FNEdit = (cbFNEdit.value = vbChecked)
    VicePath = txtVicePath.Text
    IgnoreBadID = (cbIgnoreBadID.value = vbChecked)
End Sub

Private Sub cmdShowLog_Click()
    If Exists(LogFile) = True Then
        ViewFile LogFile
    Else
        MyMsg "There is no log file yet."
    End If
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

Private Sub cmdViceBrowse_Click()
    Dim Tmp As String
    Tmp = GetBrowseDir(Me, "Select Path containing VICE executables:")
    If Tmp <> "" Then txtVicePath.Text = AddSlash(Tmp)
End Sub

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

Private Sub cmdDetect_Click()
    frmMain.DetectDrives True
End Sub

Private Sub lstOpt_Click()
    Dim j As Integer
    
    For j = 0 To 8: optFrame(j).Visible = False: Next
    j = lstOpt.ListIndex
    If j > 0 Then
        optFrame(j).Left = optFrame(0).Left
        optFrame(j).Top = optFrame(0).Top
        optFrame(j).Width = optFrame(0).Width
        optFrame(j).Height = optFrame(0).Height
    End If
    optFrame(j).Visible = True
    DoEvents
    
End Sub

Private Sub optBatchMode_Click(Index As Integer)
    BatchMode = Index
End Sub

Private Sub optFNMode_Click(Index As Integer)
    FNMode = Index
End Sub

Private Sub cmdResetBus_Click()
    frmMain.PubDoCommand CBMCtrl, "reset", "Resetting drives, please wait..."
End Sub

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
