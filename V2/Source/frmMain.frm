VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   Caption         =   "CBM Transfer"
   ClientHeight    =   14505
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   18630
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
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   14505
   ScaleWidth      =   18630
   StartUpPosition =   3  'Windows Default
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
      Left            =   690
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   172
      Top             =   7890
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picTheme 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
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
      HasDC           =   0   'False
      Height          =   1935
      Left            =   1470
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   171
      Top             =   9330
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.FileListBox flbThemes 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00E0E0E0&
      Height          =   4905
      Left            =   16320
      Pattern         =   "theme-*.bmp"
      TabIndex        =   170
      Top             =   3450
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.ListBox lstImageFiles 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Index           =   3
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":BBE8
      Left            =   16290
      List            =   "frmMain.frx":BBEA
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   129
      Top             =   2610
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.ListBox lstImageFiles 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Index           =   2
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":BBEC
      Left            =   16260
      List            =   "frmMain.frx":BBEE
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   128
      Top             =   2010
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.ListBox lstImageFiles 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Index           =   1
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":BBF0
      Left            =   16260
      List            =   "frmMain.frx":BBF2
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   127
      Top             =   1410
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.ListBox lstImageFiles 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Index           =   0
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":BBF4
      Left            =   16260
      List            =   "frmMain.frx":BBF6
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   126
      Top             =   870
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox picCBM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
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
      Height          =   1440
      Left            =   990
      Picture         =   "frmMain.frx":BBF8
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   121
      Top             =   7680
      Width           =   15360
   End
   Begin VB.Frame frDDF 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Disk Image File"
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
      Height          =   6735
      Index           =   1
      Left            =   5370
      TabIndex        =   86
      Top             =   7560
      Visible         =   0   'False
      Width           =   5205
      Begin VB.PictureBox cmdImageMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   4260
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   161
         Top             =   510
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
         Left            =   3930
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   160
         ToolTipText     =   "Encoding"
         Top             =   510
         Width           =   285
      End
      Begin VB.PictureBox cmdScale 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   3600
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   159
         ToolTipText     =   "Scale Height"
         Top             =   510
         Width           =   285
      End
      Begin VB.PictureBox picDiskID 
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
         Height          =   240
         Index           =   1
         Left            =   2220
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   43
         TabIndex        =   137
         Top             =   540
         Width           =   645
      End
      Begin VB.PictureBox picDiskName 
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
         Height          =   240
         Index           =   1
         Left            =   120
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   135
         TabIndex        =   133
         Top             =   540
         Width           =   2025
      End
      Begin VB.VScrollBar vsImgDir 
         Height          =   5010
         Index           =   1
         Left            =   4350
         TabIndex        =   120
         Top             =   870
         Width           =   225
      End
      Begin VB.PictureBox picDir 
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
         ForeColor       =   &H00E0E0E0&
         Height          =   5010
         Index           =   1
         Left            =   120
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   334
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   281
         TabIndex        =   119
         Top             =   870
         Width           =   4215
      End
      Begin VB.CommandButton cmdDAll 
         Appearance      =   0  'Flat
         Caption         =   "++"
         Height          =   345
         Index           =   1
         Left            =   900
         TabIndex        =   93
         ToolTipText     =   "Select ALL files"
         Top             =   6300
         Width           =   405
      End
      Begin VB.CommandButton cmdDNone 
         Appearance      =   0  'Flat
         Caption         =   "--"
         Height          =   345
         Index           =   1
         Left            =   1320
         TabIndex        =   92
         ToolTipText     =   "Select None"
         Top             =   6300
         Width           =   375
      End
      Begin VB.CommandButton cmdDDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   345
         Index           =   1
         Left            =   1740
         TabIndex        =   91
         ToolTipText     =   "Delete the selected file(s)"
         Top             =   6300
         Width           =   765
      End
      Begin VB.CommandButton cmdDView 
         Appearance      =   0  'Flat
         Caption         =   "&View"
         Height          =   345
         Index           =   1
         Left            =   3360
         TabIndex        =   90
         ToolTipText     =   "View selected file"
         Top             =   6300
         Width           =   555
      End
      Begin VB.CommandButton cmdDRun 
         Appearance      =   0  'Flat
         Caption         =   "&Run"
         Height          =   345
         Index           =   1
         Left            =   3930
         TabIndex        =   89
         ToolTipText     =   "Run selected file in Vice"
         Top             =   6300
         Width           =   645
      End
      Begin VB.CommandButton cmdImageRefresh 
         Appearance      =   0  'Flat
         Caption         =   "Refresh"
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   88
         ToolTipText     =   "Delete the selected file(s)"
         Top             =   6300
         Width           =   735
      End
      Begin VB.CommandButton cmdDRename 
         Appearance      =   0  'Flat
         Caption         =   "Rename"
         Height          =   345
         Index           =   1
         Left            =   2520
         TabIndex        =   87
         ToolTipText     =   "Rename file(s)"
         Top             =   6300
         Width           =   795
      End
      Begin VB.Label lblExt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2940
         TabIndex        =   101
         ToolTipText     =   "Image Type"
         Top             =   510
         Width           =   645
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "File:"
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
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   96
         Top             =   120
         Width           =   285
      End
      Begin VB.Label lblDDFile 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   1
         Left            =   420
         TabIndex        =   95
         Top             =   90
         Width           =   4140
      End
      Begin VB.Label DFBlocksFree 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   94
         ToolTipText     =   "Blocks Free"
         Top             =   5940
         Width           =   4455
      End
   End
   Begin VB.Frame frSrc 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Directory on Local PC"
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
      Height          =   6735
      Index           =   1
      Left            =   30
      TabIndex        =   67
      Top             =   7560
      Width           =   5205
      Begin VB.PictureBox cmdLocalMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   4770
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   154
         Top             =   480
         Width           =   285
      End
      Begin VB.PictureBox cmdBrowse 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   4770
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   153
         Top             =   90
         Width           =   285
      End
      Begin VB.PictureBox cmdPathUp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   90
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   151
         Top             =   90
         Width           =   285
      End
      Begin VB.ComboBox cboFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   1
         ItemData        =   "frmMain.frx":53C3A
         Left            =   960
         List            =   "frmMain.frx":53C74
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   480
         Width           =   3735
      End
      Begin VB.ComboBox txtLocalDir 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   1
         ItemData        =   "frmMain.frx":53DC1
         Left            =   420
         List            =   "frmMain.frx":53DC3
         OLEDropMode     =   1  'Manual
         Sorted          =   -1  'True
         TabIndex        =   82
         Top             =   90
         Width           =   4275
      End
      Begin VB.CommandButton cmdSrcDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   77
         ToolTipText     =   "Delete selected file(s)"
         Top             =   6300
         Width           =   1065
      End
      Begin VB.CommandButton cmdSrcRename 
         Appearance      =   0  'Flat
         Caption         =   "R&ename"
         Height          =   315
         Index           =   1
         Left            =   1260
         TabIndex        =   76
         ToolTipText     =   "Rename Selected File(s)"
         Top             =   6300
         Width           =   1065
      End
      Begin VB.CommandButton cmdSrcRun 
         Appearance      =   0  'Flat
         Caption         =   "R&un"
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   75
         ToolTipText     =   "Run File or Image using Vice "
         Top             =   5940
         Width           =   1065
      End
      Begin VB.CommandButton cmdNewImage 
         Appearance      =   0  'Flat
         Caption         =   "&New Dnn"
         Height          =   315
         Index           =   1
         Left            =   1260
         TabIndex        =   73
         ToolTipText     =   "Create a new/blank CBM Image File"
         Top             =   5940
         Width           =   1065
      End
      Begin VB.CommandButton cmdSrcView 
         Appearance      =   0  'Flat
         Caption         =   "&View"
         Height          =   315
         Index           =   1
         Left            =   2370
         TabIndex        =   72
         ToolTipText     =   "View File or Disk Image file Contents"
         Top             =   6300
         Width           =   675
      End
      Begin VB.CommandButton cmdSrcRefresh 
         Appearance      =   0  'Flat
         Caption         =   "Re&fresh"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   71
         ToolTipText     =   "Refresh Directory"
         Top             =   5940
         Width           =   1065
      End
      Begin VB.DriveListBox drvLocal 
         Appearance      =   0  'Flat
         BackColor       =   &H00212226&
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   70
         Top             =   870
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.DirListBox dirLocal 
         Appearance      =   0  'Flat
         BackColor       =   &H00212226&
         ForeColor       =   &H00E0E0E0&
         Height          =   4590
         Index           =   1
         Left            =   120
         TabIndex        =   69
         Top             =   1170
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdSrcView2 
         Appearance      =   0  'Flat
         Caption         =   "&2"
         Height          =   345
         Index           =   1
         Left            =   3090
         TabIndex        =   68
         ToolTipText     =   "View File or Disk Image file Contents"
         Top             =   6270
         Width           =   375
      End
      Begin VB.FileListBox lstLocal 
         Appearance      =   0  'Flat
         BackColor       =   &H00212226&
         ForeColor       =   &H00E0E0E0&
         Height          =   4905
         Index           =   1
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   74
         Top             =   870
         Width           =   4995
      End
      Begin VB.Label BlockText 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "0"
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
         Index           =   1
         Left            =   3990
         TabIndex        =   118
         Top             =   6390
         Width           =   735
      End
      Begin VB.Label KBText 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "0"
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
         Index           =   1
         Left            =   3990
         TabIndex        =   117
         Top             =   6120
         Width           =   735
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show:"
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
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   84
         Top             =   540
         Width           =   450
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected:"
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
         Height          =   195
         Index           =   22
         Left            =   4020
         TabIndex        =   81
         Top             =   5880
         Width           =   675
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KB"
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
         Height          =   195
         Index           =   7
         Left            =   4815
         TabIndex        =   80
         Top             =   6120
         Width           =   210
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blks"
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
         Height          =   195
         Index           =   9
         Left            =   4800
         TabIndex        =   79
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label lblPathView 
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
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   78
         ToolTipText     =   "Drive and Folder View"
         Top             =   660
         Width           =   225
      End
   End
   Begin VB.Frame frMiddle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
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
      Height          =   6735
      Left            =   9660
      TabIndex        =   61
      Top             =   750
      Width           =   1035
      Begin VB.PictureBox cmdConfig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   240
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   169
         ToolTipText     =   "Settings"
         Top             =   1350
         Width           =   540
      End
      Begin VB.PictureBox cmdDAD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   75
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   60
         TabIndex        =   149
         ToolTipText     =   "Toggle DAD Window"
         Top             =   4260
         Width           =   900
      End
      Begin VB.PictureBox cmdHelp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   330
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   148
         ToolTipText     =   "Show Help File"
         Top             =   5100
         Width           =   330
      End
      Begin VB.PictureBox cmdAbout 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   330
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   147
         ToolTipText     =   "About CBM Transfer"
         Top             =   390
         Width           =   330
      End
      Begin VB.PictureBox cmdCopyLeft 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   90
         ScaleHeight     =   56
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   56
         TabIndex        =   146
         ToolTipText     =   "Copy Right to Left"
         Top             =   3180
         Width           =   840
      End
      Begin VB.PictureBox cmdCopyRight 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   90
         ScaleHeight     =   56
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   56
         TabIndex        =   145
         ToolTipText     =   "Copy Left to Right"
         Top             =   2250
         Width           =   840
      End
      Begin VB.CheckBox cbTest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "debug"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   102
         ToolTipText     =   "For Internal Testing"
         Top             =   4800
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblFGBG 
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
         Height          =   195
         Index           =   1
         Left            =   510
         TabIndex        =   174
         ToolTipText     =   "List BG Colour"
         Top             =   5700
         Width           =   285
      End
      Begin VB.Label lblFGBG 
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
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   173
         ToolTipText     =   "List FG Colour"
         Top             =   5700
         Width           =   285
      End
      Begin VB.Label lblSizer2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   30
         TabIndex        =   114
         ToolTipText     =   "Dual-view"
         Top             =   60
         Width           =   225
      End
      Begin VB.Label lblTheme 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dark"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   30
         TabIndex        =   113
         ToolTipText     =   "Theme Name. Click for Menu"
         Top             =   5910
         Width           =   975
      End
      Begin VB.Label cmdThemeSel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ">"
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
         Height          =   240
         Index           =   1
         Left            =   840
         TabIndex        =   112
         ToolTipText     =   "Change Theme"
         Top             =   5670
         Width           =   135
      End
      Begin VB.Label cmdThemeSel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<"
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
         Height          =   240
         Index           =   0
         Left            =   30
         TabIndex        =   111
         ToolTipText     =   "Change Theme"
         Top             =   5670
         Width           =   135
      End
      Begin VB.Label cmdResults 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "*"
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
         Height          =   165
         Left            =   390
         TabIndex        =   99
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblSizer 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   62
         ToolTipText     =   "Show/Hide Pane"
         Top             =   6420
         Width           =   225
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8790
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frDDF 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Disk Image File"
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
      Height          =   6735
      Index           =   0
      Left            =   4860
      TabIndex        =   19
      Top             =   750
      Visible         =   0   'False
      Width           =   4725
      Begin VB.PictureBox cmdImageMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   4260
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   158
         ToolTipText     =   "Directory Actions"
         Top             =   510
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
         Left            =   3930
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   157
         ToolTipText     =   "Encoding"
         Top             =   510
         Width           =   285
      End
      Begin VB.PictureBox cmdScale 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   3600
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   156
         ToolTipText     =   "Toggle Height"
         Top             =   510
         Width           =   285
      End
      Begin VB.PictureBox picDiskID 
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
         Height          =   240
         Index           =   0
         Left            =   2220
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   136
         Top             =   540
         Width           =   615
      End
      Begin VB.PictureBox picDiskName 
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
         Height          =   240
         Index           =   0
         Left            =   120
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   135
         TabIndex        =   132
         Top             =   540
         Width           =   2025
      End
      Begin VB.VScrollBar vsImgDir 
         Height          =   5010
         Index           =   0
         Left            =   4350
         TabIndex        =   123
         Top             =   900
         Width           =   225
      End
      Begin VB.PictureBox picDir 
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
         Height          =   5010
         Index           =   0
         Left            =   120
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   334
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   281
         TabIndex        =   122
         Top             =   900
         Width           =   4215
      End
      Begin VB.CommandButton cmdDRename 
         Appearance      =   0  'Flat
         Caption         =   "Rename"
         Height          =   345
         Index           =   0
         Left            =   2460
         TabIndex        =   85
         ToolTipText     =   "Rename file(s)"
         Top             =   6300
         Width           =   795
      End
      Begin VB.CommandButton cmdImageRefresh 
         Appearance      =   0  'Flat
         Caption         =   "Refresh"
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   59
         ToolTipText     =   "Delete the selected file(s)"
         Top             =   6300
         Width           =   765
      End
      Begin VB.CommandButton cmdDRun 
         Appearance      =   0  'Flat
         Caption         =   "&Run"
         Height          =   345
         Index           =   0
         Left            =   3930
         TabIndex        =   52
         ToolTipText     =   "Run selected file in Vice"
         Top             =   6300
         Width           =   675
      End
      Begin VB.CommandButton cmdDView 
         Appearance      =   0  'Flat
         Caption         =   "&View"
         Height          =   345
         Index           =   0
         Left            =   3300
         TabIndex        =   32
         ToolTipText     =   "View selected file"
         Top             =   6300
         Width           =   615
      End
      Begin VB.CommandButton cmdDDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   345
         Index           =   0
         Left            =   1800
         TabIndex        =   31
         ToolTipText     =   "Delete the selected file(s)"
         Top             =   6300
         Width           =   645
      End
      Begin VB.CommandButton cmdDNone 
         Appearance      =   0  'Flat
         Caption         =   "--"
         Height          =   345
         Index           =   0
         Left            =   1380
         TabIndex        =   29
         ToolTipText     =   "Select None"
         Top             =   6300
         Width           =   375
      End
      Begin VB.CommandButton cmdDAll 
         Appearance      =   0  'Flat
         Caption         =   "++"
         Height          =   345
         Index           =   0
         Left            =   960
         TabIndex        =   28
         ToolTipText     =   "Select ALL files"
         Top             =   6300
         Width           =   405
      End
      Begin VB.Label lblExt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   2910
         TabIndex        =   100
         ToolTipText     =   "Image Type"
         Top             =   510
         Width           =   645
      End
      Begin VB.Label DFBlocksFree 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   5940
         Width           =   4455
      End
      Begin VB.Label lblDDFile 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   0
         Left            =   510
         TabIndex        =   21
         Top             =   90
         Width           =   4050
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   90
         Width           =   315
      End
   End
   Begin VB.Frame frLink 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "CBM Drive via CBMLink"
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
      Height          =   6735
      Left            =   10740
      TabIndex        =   33
      Top             =   7560
      Visible         =   0   'False
      Width           =   5205
      Begin VB.PictureBox cmdImageMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   3600
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   167
         Top             =   480
         Width           =   285
      End
      Begin VB.PictureBox cmdEncode 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   3240
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   166
         Top             =   480
         Width           =   285
      End
      Begin VB.PictureBox cmdScale 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   2910
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   165
         Top             =   480
         Width           =   285
      End
      Begin VB.CommandButton cmdDNone 
         Appearance      =   0  'Flat
         Caption         =   "--"
         Height          =   345
         Index           =   3
         Left            =   4710
         TabIndex        =   143
         ToolTipText     =   "Select None"
         Top             =   5130
         Width           =   435
      End
      Begin VB.CommandButton cmdDAll 
         Appearance      =   0  'Flat
         Caption         =   "++"
         Height          =   345
         Index           =   3
         Left            =   4200
         TabIndex        =   142
         ToolTipText     =   "Select ALL files"
         Top             =   5130
         Width           =   465
      End
      Begin VB.PictureBox picDiskID 
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
         Height          =   240
         Index           =   3
         Left            =   2220
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   43
         TabIndex        =   139
         Top             =   540
         Width           =   645
      End
      Begin VB.PictureBox picDiskName 
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
         Height          =   240
         Index           =   3
         Left            =   120
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   135
         TabIndex        =   135
         Top             =   540
         Width           =   2025
      End
      Begin VB.VScrollBar vsImgDir 
         Height          =   5010
         Index           =   3
         Left            =   3900
         TabIndex        =   131
         Top             =   870
         Width           =   225
      End
      Begin VB.PictureBox picDir 
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
         Height          =   5010
         Index           =   3
         Left            =   120
         ScaleHeight     =   334
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   251
         TabIndex        =   130
         Top             =   870
         Width           =   3765
      End
      Begin VB.CommandButton cmdLinkScratch 
         Appearance      =   0  'Flat
         Caption         =   "Delete"
         Height          =   360
         Left            =   4185
         TabIndex        =   44
         ToolTipText     =   "Scratch (delete) selected file(s)"
         Top             =   3630
         Width           =   945
      End
      Begin VB.CommandButton cmdLinkRename 
         Appearance      =   0  'Flat
         Caption         =   "Rename"
         Height          =   360
         Left            =   4185
         TabIndex        =   43
         ToolTipText     =   "Rename selected file(s)"
         Top             =   4020
         Width           =   945
      End
      Begin VB.CommandButton cmdLinkStatus 
         Appearance      =   0  'Flat
         Caption         =   "Status"
         Height          =   360
         Left            =   4215
         TabIndex        =   42
         Top             =   6270
         Width           =   945
      End
      Begin VB.CommandButton cmdLinkReset 
         Appearance      =   0  'Flat
         Caption         =   "Reset"
         Height          =   360
         Left            =   4215
         TabIndex        =   41
         Top             =   5880
         Width           =   945
      End
      Begin VB.CommandButton cmdLinkFormat 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         Height          =   360
         Left            =   4185
         TabIndex        =   40
         ToolTipText     =   "Format disk in Floppy drive"
         Top             =   1995
         Width           =   945
      End
      Begin VB.CommandButton cmdLinkInit 
         Appearance      =   0  'Flat
         Caption         =   "Initialize"
         Height          =   330
         Left            =   4185
         TabIndex        =   39
         ToolTipText     =   "Reset the Drive"
         Top             =   1605
         Width           =   945
      End
      Begin VB.CommandButton cmdLinkValidate 
         Appearance      =   0  'Flat
         Caption         =   "Validate"
         Height          =   360
         Left            =   4185
         TabIndex        =   38
         ToolTipText     =   "Perform Disk Validation"
         Top             =   2685
         Width           =   945
      End
      Begin VB.CommandButton cmdLinkDir 
         Appearance      =   0  'Flat
         Caption         =   "Directory"
         Height          =   375
         Left            =   4185
         TabIndex        =   36
         Top             =   1185
         Width           =   945
      End
      Begin VB.ComboBox cboLinkDev 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "frmMain.frx":53DC5
         Left            =   780
         List            =   "frmMain.frx":53DE1
         Style           =   2  'Dropdown List
         TabIndex        =   35
         ToolTipText     =   "Select X Device Unit Number"
         Top             =   60
         Width           =   1230
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Device:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   98
         Top             =   120
         Width           =   540
      End
      Begin VB.Label lblLinkLastStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1020
         TabIndex        =   49
         ToolTipText     =   "Drive Status"
         Top             =   6345
         Width           =   3105
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drv Status:"
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
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   48
         Top             =   6375
         Width           =   840
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Files:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   15
         Left            =   4305
         TabIndex        =   47
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drive:"
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
         Height          =   195
         Index           =   13
         Left            =   4275
         TabIndex        =   46
         Top             =   5640
         Width           =   450
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select:"
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
         Height          =   195
         Index           =   11
         Left            =   4305
         TabIndex        =   45
         Top             =   4845
         Width           =   495
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disk:"
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
         Height          =   195
         Index           =   14
         Left            =   4230
         TabIndex        =   37
         Top             =   945
         Width           =   375
      End
      Begin VB.Label DFBlocksFree 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Blocks Free"
         Top             =   5940
         Width           =   4005
      End
   End
   Begin VB.Frame frSrc 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Directory on Local PC"
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
      Height          =   6735
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   750
      Width           =   4725
      Begin VB.PictureBox cmdLocalMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   4260
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   155
         ToolTipText     =   "Directory Actions"
         Top             =   480
         Width           =   285
      End
      Begin VB.PictureBox cmdBrowse 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   4260
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   152
         ToolTipText     =   "Browse for Folder"
         Top             =   90
         Width           =   285
      End
      Begin VB.PictureBox cmdPathUp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   60
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   150
         ToolTipText     =   "Move Up to Parent"
         Top             =   120
         Width           =   285
      End
      Begin VB.CommandButton cmdSrcView2 
         Appearance      =   0  'Flat
         Caption         =   "View &2"
         Height          =   315
         Index           =   0
         Left            =   2790
         TabIndex        =   66
         ToolTipText     =   "View File or Disk Image file Contents in Right Window"
         Top             =   6300
         Width           =   675
      End
      Begin VB.DirListBox dirLocal 
         Appearance      =   0  'Flat
         BackColor       =   &H00212226&
         ForeColor       =   &H00E0E0E0&
         Height          =   4590
         Index           =   0
         Left            =   120
         TabIndex        =   64
         Top             =   1170
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.DriveListBox drvLocal 
         Appearance      =   0  'Flat
         BackColor       =   &H00212226&
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   63
         Top             =   870
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox txtLocalDir 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":53E28
         Left            =   420
         List            =   "frmMain.frx":53E2A
         OLEDropMode     =   1  'Manual
         Sorted          =   -1  'True
         TabIndex        =   60
         Top             =   90
         Width           =   3795
      End
      Begin VB.CommandButton cmdSrcRefresh 
         Appearance      =   0  'Flat
         Caption         =   "Re&fresh"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   53
         ToolTipText     =   "Refresh Directory"
         Top             =   5940
         Width           =   975
      End
      Begin VB.CommandButton cmdSrcView 
         Appearance      =   0  'Flat
         Caption         =   "&View"
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   51
         ToolTipText     =   "View File or Disk Image file Contents"
         Top             =   6300
         Width           =   585
      End
      Begin VB.CommandButton cmdNewImage 
         Appearance      =   0  'Flat
         Caption         =   "&New Dnn"
         Height          =   315
         Index           =   0
         Left            =   1140
         TabIndex        =   30
         ToolTipText     =   "Create a new/blank CBM Image File"
         Top             =   5940
         Width           =   975
      End
      Begin VB.ComboBox cboFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":53E2C
         Left            =   990
         List            =   "frmMain.frx":53E66
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   480
         Width           =   3225
      End
      Begin VB.FileListBox lstLocal 
         Appearance      =   0  'Flat
         BackColor       =   &H00212226&
         ForeColor       =   &H00E0E0E0&
         Height          =   4905
         Index           =   0
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   12
         Top             =   870
         Width           =   4485
      End
      Begin VB.CommandButton cmdSrcRun 
         Appearance      =   0  'Flat
         Caption         =   "R&un"
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   11
         ToolTipText     =   "Run File or Image using Vice "
         Top             =   5940
         Width           =   1305
      End
      Begin VB.CommandButton cmdSrcRename 
         Appearance      =   0  'Flat
         Caption         =   "R&ename"
         Height          =   315
         Index           =   0
         Left            =   1140
         TabIndex        =   10
         ToolTipText     =   "Rename Selected File(s)"
         Top             =   6300
         Width           =   975
      End
      Begin VB.CommandButton cmdSrcDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Delete selected file(s)"
         Top             =   6300
         Width           =   975
      End
      Begin VB.Label BlockText 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "0"
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
         Index           =   0
         Left            =   3540
         TabIndex        =   116
         Top             =   6390
         Width           =   735
      End
      Begin VB.Label KBText 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "0"
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
         Index           =   0
         Left            =   3540
         TabIndex        =   115
         Top             =   6120
         Width           =   735
      End
      Begin VB.Label lblPathView 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   65
         ToolTipText     =   "Drive and Folder View"
         Top             =   660
         Width           =   255
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blks"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   8
         Left            =   4300
         TabIndex        =   27
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KB"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   6
         Left            =   4305
         TabIndex        =   26
         Top             =   6120
         Width           =   180
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   10
         Left            =   450
         TabIndex        =   22
         Top             =   540
         Width           =   480
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Selected:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   16
         Left            =   3540
         TabIndex        =   14
         Top             =   5880
         Width           =   690
      End
   End
   Begin VB.Frame frX 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "CBM Drive on X-Cable"
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
      Height          =   6735
      Left            =   10740
      TabIndex        =   0
      Top             =   750
      Width           =   5205
      Begin VB.PictureBox cmdXMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4830
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   168
         ToolTipText     =   "Browse for Folder"
         Top             =   120
         Width           =   285
      End
      Begin VB.PictureBox cmdImageMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3570
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   164
         Top             =   480
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
         Left            =   3240
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   163
         ToolTipText     =   "Encoding"
         Top             =   480
         Width           =   285
      End
      Begin VB.PictureBox cmdScale 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   2910
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   162
         ToolTipText     =   "Scale Height"
         Top             =   480
         Width           =   285
      End
      Begin VB.CommandButton cmdDNone 
         Appearance      =   0  'Flat
         Caption         =   "--"
         Height          =   345
         Index           =   2
         Left            =   4710
         TabIndex        =   141
         ToolTipText     =   "Select None"
         Top             =   4260
         Width           =   435
      End
      Begin VB.CommandButton cmdDAll 
         Appearance      =   0  'Flat
         Caption         =   "++"
         Height          =   345
         Index           =   2
         Left            =   4230
         TabIndex        =   140
         ToolTipText     =   "Select ALL files"
         Top             =   4260
         Width           =   465
      End
      Begin VB.PictureBox picDiskID 
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
         Height          =   270
         Index           =   2
         Left            =   2220
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   43
         TabIndex        =   138
         Top             =   510
         Width           =   645
      End
      Begin VB.PictureBox picDiskName 
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
         Height          =   270
         Index           =   2
         Left            =   120
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   135
         TabIndex        =   134
         Top             =   510
         Width           =   2025
      End
      Begin VB.VScrollBar vsImgDir 
         CausesValidation=   0   'False
         Height          =   5010
         Index           =   2
         Left            =   3900
         TabIndex        =   125
         Top             =   870
         Width           =   225
      End
      Begin VB.PictureBox picDir 
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
         Height          =   5010
         Index           =   2
         Left            =   120
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   334
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   251
         TabIndex        =   124
         Top             =   870
         Width           =   3765
      End
      Begin VB.CommandButton cmdXRoot 
         Appearance      =   0  'Flat
         Caption         =   "Root"
         Height          =   360
         Left            =   4200
         TabIndex        =   57
         ToolTipText     =   "Return to Root Partition"
         Top             =   3510
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdXPart 
         Appearance      =   0  'Flat
         Caption         =   "Sel"
         Height          =   360
         Left            =   4740
         TabIndex        =   56
         ToolTipText     =   "Select/View partition"
         Top             =   3510
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton cmdXView 
         Appearance      =   0  'Flat
         Caption         =   "View"
         Height          =   360
         Left            =   4200
         TabIndex        =   50
         ToolTipText     =   "CBM File Viewer"
         Top             =   2790
         Width           =   945
      End
      Begin VB.ComboBox cboXDevNum 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmMain.frx":53FB3
         Left            =   720
         List            =   "frmMain.frx":53FC3
         Style           =   2  'Dropdown List
         TabIndex        =   25
         ToolTipText     =   "Select X Device Unit Number"
         Top             =   90
         Width           =   4035
      End
      Begin VB.CommandButton cmdXReset 
         Appearance      =   0  'Flat
         Caption         =   "Reset"
         Height          =   360
         Left            =   4200
         TabIndex        =   6
         Top             =   5850
         Width           =   945
      End
      Begin VB.CommandButton cmdXDriveStatus 
         Appearance      =   0  'Flat
         Caption         =   "Status"
         Height          =   360
         Left            =   4200
         TabIndex        =   5
         ToolTipText     =   "Get Drive Status"
         Top             =   6270
         Width           =   945
      End
      Begin VB.CommandButton cmdXRefresh 
         Appearance      =   0  'Flat
         Caption         =   "Directory"
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         ToolTipText     =   "Read Disk Directory"
         Top             =   1170
         Width           =   945
      End
      Begin VB.CommandButton cmdXRename 
         Appearance      =   0  'Flat
         Caption         =   "Rename"
         Height          =   360
         Left            =   4200
         TabIndex        =   3
         ToolTipText     =   "Rename selected file(s)"
         Top             =   2370
         Width           =   945
      End
      Begin VB.CommandButton cmdXScratch 
         Appearance      =   0  'Flat
         Caption         =   "Delete"
         Height          =   360
         Left            =   4200
         TabIndex        =   2
         ToolTipText     =   "Scratch (delete) selected file(s)"
         Top             =   1950
         Width           =   945
      End
      Begin VB.Label lblDName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3900
         TabIndex        =   144
         Top             =   540
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Device:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   97
         Top             =   150
         Width           =   540
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Partition:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   19
         Left            =   4230
         TabIndex        =   58
         Top             =   3270
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   20
         Left            =   4230
         TabIndex        =   18
         Top             =   3990
         Width           =   495
      End
      Begin VB.Label DFBlocksFree 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   5940
         Width           =   4005
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Drv Status:"
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
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   6360
         Width           =   840
      End
      Begin VB.Label lblXLastStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1020
         TabIndex        =   15
         ToolTipText     =   "Drive Status"
         Top             =   6345
         Width           =   3105
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Drive:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   21
         Left            =   4260
         TabIndex        =   13
         Top             =   5610
         Width           =   435
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Disk:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   17
         Left            =   4230
         TabIndex        =   8
         Top             =   930
         Width           =   375
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Files:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   18
         Left            =   4230
         TabIndex        =   7
         Top             =   1710
         Width           =   390
      End
   End
   Begin VB.Frame frDestB 
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
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   10740
      TabIndex        =   104
      Top             =   420
      Width           =   5235
      Begin VB.Label lblDstMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Disk Image"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   3
         Left            =   3840
         TabIndex        =   108
         Tag             =   "&H0000C0C0&"
         ToolTipText     =   "Click to View Disk Image Files"
         Top             =   0
         Width           =   1320
      End
      Begin VB.Label lblDstMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "X-Cable"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   107
         Tag             =   "&H00FF0000&"
         ToolTipText     =   "Click to View X-Cable directory"
         Top             =   0
         Width           =   1200
      End
      Begin VB.Label lblDstMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "CBMLink"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   1260
         TabIndex        =   106
         Tag             =   "&H000040C0&"
         ToolTipText     =   "Click to View CBMLink directory"
         Top             =   0
         Width           =   1230
      End
      Begin VB.Label lblDstMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         Caption         =   "Local PC"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   2550
         TabIndex        =   105
         Tag             =   "&H0000C0C0&"
         ToolTipText     =   "Click to View Files"
         Top             =   0
         Width           =   1230
      End
   End
   Begin VB.Label cmdMinimize 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   4320
      TabIndex        =   110
      ToolTipText     =   "Minimize"
      Top             =   -180
      Width           =   165
   End
   Begin VB.Label cmdClose 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4650
      TabIndex        =   109
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   165
   End
   Begin VB.Label lblSrcMode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Disk Image"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   1
      Left            =   2460
      TabIndex        =   55
      Tag             =   "&H00C0C000&"
      ToolTipText     =   "Click to View Disk Image Files"
      Top             =   420
      Width           =   2325
   End
   Begin VB.Label lblSrcMode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Caption         =   "Local PC"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   54
      Tag             =   "&H0000C0C0&"
      ToolTipText     =   "Click to View Files"
      Top             =   420
      Width           =   2355
   End
   Begin VB.Label lblDrag 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "  CBM Transfer"
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
      Height          =   300
      Left            =   0
      TabIndex        =   103
      Top             =   0
      Width           =   4200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2023 Steve J. Gray
' ====================================================
'
' frmMain - The MAIN window. Code execution starts here!
'
' Based on GUI4CBM4WIN. The following (between "/" lines) is the notice
' included with the GUI4CBM4WIN source code:


'/////////////////////////////////////////////////////////////////////////
'   Copyright (C) 2004-2005 Leif Bloomquist
'   Copyright (C) 2006      Wolfgang Moser
'   Copyright (C) 2006      Spiro Trikaliotis
'---------------------------------------------
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
'/////////////////////////////////////////////////////////////////////////


' Steve Gray's notes:
'
' CBM-Transfer (CBMXfer) is based on GUI4CBM4WIN, which was written in VB6 by Leif Bloomquist.
' Eventually GUI4CBM4WIN was converted to VB.NET, which I do not own and find to be bulky and confusing.
' So, I forked the VB6 version and Renamed it to CBM-Transfer to make it easier to say ;-)
' I have greatly expanded the original program, and replaced or re-written the majority of code
' with my own. I added support for CBMLink (which could use IEEE drives before ZoomFloppy),
' for C1541 to work with Disk Images, and the ability to run programs using VICE. I added support
' for Zoomfloppy, NIBTOOLS, IMGCOPY, 1581 directories, and P00 files.
' I added the File Viewer which supports multiple viewing formats, and dual-view mode. I wrote a
' fully-featured 6502 Symbolic Machine Language Disassembler, Font Viewer/Editor, and Screen Designer.
' I also integrated Peter Weighhill's Picture Viewer which was heavily modified to improve its speed
' and work in VB. I have added the Theme features and True CBM Font rendering for Disk Directories,
' Disk Images, BASIC, SEQ and HEX views.

Option Explicit

Dim Drive(31) As String
Dim DragFlag As Boolean                 'Indicases when the form is being dragged
Dim IX As Single, IY As Single          'Window dragging
Dim FX As Single, FY As Single          'Window dragging
Dim TwipX As Single, TwipY As Single    'Screen sizing

Dim RenderY(3)      As Integer          'Y rendering size for each List
Dim EncodeL(3)      As Integer          'Font Encoding for each List
Dim EncodeS(5)      As String           'Encoding Strings for tooltips
Dim TabTheme(3, 5)  As Long             'Tab Theme Colours
Dim FontBuf         As String           'Font Binary for Rendering CBM Fonts
Dim MouseDownI      As Integer          'Mouse Down Index
Dim FontBMPFlag     As Boolean          'Flag to indicate BMP font loaded

'---- INFO: Display Program info and acknowlegements
Private Sub cmdAbout_Click()
    MyMsg "CBM-Transfer V2.03 (Jul 30/2025)" & Cr & _
          "(C)2007-2025 Steve J. Gray" & Cr & Cr & _
          "A front-end for: OpenCBM, VICE, NibTools, and CBMLink" & Cr & Cr & _
          "Based on GUI4CBM4WIN V0.4.1," & Cr & _
          "by Leif Bloomquist, Wolfgang Moser and Spiro Trikaliotis." & Cr & _
          "Viewer includes portions of 'CBM2BMP' code by Peter Weighill"
End Sub

'========================================
' FORM and PROGRAM SUBS
'========================================

'---- FORM: Form is Loaded
' This is the main initialization of the program.
' Global variables are set, the INI file is loaded, Paths are set. Utility EXE files are defined.
' TEMP files are defined. The Path History is loaded, and X-cable Drives are detected.
' The initial state of the GUI is set.

Private Sub Form_Load()
    Dim i As Integer, Tmp As String, Flag As Boolean, xTmp As String

    On Error Resume Next
    
    CurDir = App.Path                                                           'Path to CBMXfer.exe
    ExeDir = AddSlash(CurDir)                                                   'Add a slash
    ThemeDir = ExeDir & "themes\"                                               'Set Theme Directory
    INIFile = ExeDir & "cbmxfer.ini"                                            'Set INI file
    
    SetAllPaths
    
    FontBMPFlag = True                                                          'Assume no theme selected - use internal bitmap
    
    '-- Set Special Paths
    
    Tmp = Environ$("temp")                                                      'Try to use TEMP folder
    If Tmp = "" Then Tmp = Environ$("tmp")                                      'If NOT defined try TMP
    Tmp = Tmp & "\cbmxfer"                                                      'Common start for TEMPFILES
    
    LogFile = ExeDir & "cbmxferlog.txt"                                         'Set Log file path
    CatalogFile = ExeDir & "catalog.txt"                                        'Set the Catalog file path
    HistoryFile = ExeDir & "pathhistory.txt"                                    'Path History location and name
    
    flbThemes.Pattern = "theme-*.bmp"                                           'Theme Pattern
    flbThemes.Path = ThemeDir                                                   'Themes Folder
        
    TEMPFILE1 = Tmp & "out.txt"                                                 'Captured Output from shell
    TEMPFILE2 = Tmp & "err.txt"                                                 'Captured Errors from shell
    TEMPFILE3 = Tmp & "tmp.tmp"                                                 'General-purpose Temp File (ie: for multi-step copies)
        
    '-- Set Initial Prgram Variables
    
    MsgTitle = "CBM Transfer"                                                   'Program Name for Messages
    
    Cr = Chr(13): LF = Chr(10): Qu = Chr(34): Nu = Chr(0): Hx = "&h"            'some common characters
    SrcMode = 0: Layout = 0: Layout2 = 1
    Theme = -1                                                                  'Default Theme number
    
    TwipX = Screen.TwipsPerPixelX
    TwipY = Screen.TwipsPerPixelY
    
    
    '--- Powers of 2 for bitmap operations
    
    For i = 0 To 7: Pow(i) = 2 ^ i: Next                                        'Set Powers of 2
       
    '-- Load Settings, Set Paths, and Theme
    
    If Mid(CurDir, 2, 1) = ":" Then ChDrive Left(CurDir, 1)                     'If CurDir has a drive specification then switch to it
    MyChDir CurDir                                                              'Change to the directory
               
    '---- Load Settings, set paths, load Path History file and Set Theme
    
    DoTheme                                                                     'Set Initial Theme from Internal Bitmap
    
    LoadINI                                                                     'Load the INI file. This will load the user's settings for Utility paths etc
    LoadHistory                                                                 'Load the Path History
    
    '-- Read Encoding Strings from Menu
    
    For i = 1 To frmMenu.mnuEnc.UBound
        EncodeS(i - 1) = frmMenu.mnuEnc(i).Caption                              'Get the string from menu itself
    Next

    '-- Set Encoding Defaults
    
    EncodeL(0) = 4                                                              'Disk Image use ASCII
    EncodeL(1) = 4
    EncodeL(2) = 0                                                              'X Cable, CBM-Link use PETSCII
    EncodeL(3) = 0
    
    For i = 0 To 4: RenderY(i) = 2: Next                                        'Render Y scale factor (based on 8-pixel) for list
    
    SetEncodeDesc                                                               'Set the Tooltips
    
    '---- Startup Drive Selection
    
    DetectXDrives False                                                         'Detect all drives silently
    If UseFirstDrive = True Then
        cboXDevNum.ListIndex = 0                                                'Set device to first entry
        RefreshX
    End If
    
    '---- Set Dropdown defaults
    
    cboLinkDev.ListIndex = 0                                                    'Set CBM Link Drive
    cboFilter(0).ListIndex = 0                                                  'Set file filter default
    cboFilter(1).ListIndex = 0                                                  'Set file filter default
   
    '---- Set Local PC paths
    
    For i = 0 To 1
        If DirExists(LocalDir(i)) = False Then LocalDir(i) = ExeDir
        SetLocalPath i, LocalDir(i)
        lstImageFiles(i).Clear
    Next i
    
    '---- Setup the GUI Layout and Windows
    
    MakeThemeMenu                                                               'Create Popup Theme List Menu
    SetTheme                                                                    'Set the Theme, load bitmap
    SetLayout                                                                   'Set Layout and Visibility
    SetSrcFrame                                                                 'Set the default LEFT view
    SetDstFrame                                                                 'Set the default RIGHT view
        
    If StartDAD = True Then
        frmDAD.Show                                                             'Open the DAD window
        frmMain.WindowState = 1                                                 'Minimize the Main Window
        DoEvents
    End If
    
    Me.Visible = True
    
    DoEvents
        
    If Exists(INIFile) = False Then frmOptions.Show vbModal                       'Show the options window for first run (INI file is not found)
    
End Sub

'---- FORM: Query by Windows when user clicks CLOSE
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    PreExit                                                                     'Do the things we need before closing

End Sub

'---- FORM: End the program
Private Sub Form_Unload(Cancel As Integer)
    
    End                                                                         'END THE PROGRAM! That's it, we're outta here ;-)

End Sub

'---- FORM: Save stuff before program exits
' This stuff gets done before ANY forms are unloaded, otherwise some properties
' needed for INI saving will be unloaded before they can be saved
Sub PreExit()

    SaveHistory                                                                 'Save Path History
    SaveINI                                                                     'Save setting to INI file
    MyChDir ExeDir                                                              'Change back to CBM-Transfer directory

End Sub

'--- FORM: Dragging - Initiate when mouse clicks on Titlebar
' Make sure form is set to SCALEMODE=1
Private Sub lblDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    If Button And 1 Then                                                        '--- Setup form for Dragging
        If DragFlag = False Then
            IX = X: IY = Y                                                      'Remember X and Y
            FX = Me.Left: FY = Me.Top                                           'Remember initial Window position
            DragFlag = True                                                     'Set the flag
        End If
    End If

End Sub

'---- FORM: Form Dragging - Move the Form with the mouse
Private Sub lblDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If DragFlag = True Then
        Me.Move FX + X - IX, FY + Y - IY                                        'Move it!
        FX = Me.Left: FY = Me.Top                                               'Remember new position
        DoEvents
    End If

End Sub

'---- FORM: Dragging - Mouse button released
Private Sub lblDrag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragFlag = False                                                            'Stop the form from moving
End Sub

'---- FORM: Our Custom Close Button
' Since we have no standard Close button we don't get Query_Unload or Form_Unload
Private Sub cmdClose_Click()
    
    PreExit                                                                     'Do the things we need before closing
    End                                                                         'END THE PROGRAM

End Sub

'---- FORM: Our Custom Minimize Button      *****FIX ME*****
' Minimize to system tray.
' Note: We don't yet know how to handle closing from the tray itself. We must un-minimize then use our own close button....
Private Sub cmdMinimize_Click()
    
    frmMain.BorderStyle = 1                                                     'Set the Border
    frmMain.WindowState = 1                                                     'Minimize the Main Window
    
End Sub

'========================================
' GUI SUBS
'========================================
' Subs related to GUI Interaction - ICONS, Frames, Layout etc

'---- GUI: Popup Menu to Change List Encoding
Private Sub cmdEncode_Click(Index As Integer)
    
    MenuForm = 1                                                                'The Form to notify
    MenuNum = Index                                                             'The Menu Number
    
    PopupMenu frmMenu.mnuE                                                      'Popup the Menu
    
End Sub

'---- GUI: Set Encoding Description for buttons
Private Sub SetEncodeDesc()
    Dim i As Integer
    
    For i = 0 To 3
        cmdEncode(i).ToolTipText = "Encoding: " & EncodeS((EncodeL(i)))
    Next i
    
End Sub

'---- GUI: Increase RenderY scale for List
Private Sub cmdScale_Click(Index As Integer)
    RenderY(Index) = RenderY(Index) + 1                                         'Increaase
    If RenderY(Index) > 3 Then RenderY(Index) = 1                               'Set Max Scale Limite
    
    RefreshList Index                                                           'Draw the list
    
End Sub

'---- GUI: Click to Show/Hide RIGHT side
' Changes Window widt
Private Sub lblSizer_Click()
    
    Layout2 = 1 - Layout2                                           'Toggle It
    SetLayout                                                       'Re-arrange GUI

End Sub

'---- GUI: Click to Expand/Contract LEFT tabs View
' Toggles Layout from Single tab to Dual Tabs
Private Sub lblSizer2_Click()
    
    Layout = 1 - Layout                                             'Toggle It
    SetLayout                                                       'Re-arrange GUI

End Sub

'---- GUI: Click to Select the LEFT Frame
Private Sub lblSrcMode_Click(Index As Integer)
    
    SrcMode = Index
    SetSrcFrame

End Sub

'---- GUI: Hi-Light Selected Tab
Sub SetSrcFrame()
    
    If Layout = 0 Then frSrc(0).Visible = False: frDDF(0).Visible = False
    
    Select Case SrcMode
        Case 0: frSrc(0).Visible = True
        Case 1: frDDF(0).Visible = True
    End Select
    
    SetSourceSelector
    
End Sub

'---- GUI: Set Source Selector Tabs
Private Sub SetSourceSelector()
    Dim i As Integer
    
    For i = 0 To 1
        If i = SrcMode Then
            lblSrcMode(i).Font.Bold = True
            lblSrcMode(i).ForeColor = TabTheme(0, i)                    'Selected Foreground colour
            lblSrcMode(i).BackColor = TabTheme(1, i)                   'Selected Background colour
        Else
            lblSrcMode(i).Font.Bold = False
            lblSrcMode(i).ForeColor = TabTheme(2, i)                    'Unselected Foreground colour
            lblSrcMode(i).BackColor = TabTheme(3, i)                    'Unselected Background colour
        End If
    Next i
        

End Sub
'---- GUI: Show and Hide the Drag-and-Drop (DAD) window
Private Sub cmdDAD_Click()
    
    If frmDAD.Visible = False Then frmDAD.Show Else frmDAD.Hide

End Sub

'---- GUI: Click to Change Destination Mode
Private Sub lblDstMode_Click(Index As Integer)
    
    DstMode = Index
    SetDstFrame

End Sub

'---- GUI: Set Destination Frame
Sub SetDstFrame()
    Dim a As Integer
    
    frLink.Visible = False
    frSrc(1).Visible = False
    frX.Visible = False
    frDDF(1).Visible = False
    
    Select Case DstMode
        Case 0: frX.Visible = True
        Case 1: frLink.Visible = True
        Case 2: frSrc(1).Visible = True
        Case 3: frDDF(1).Visible = True
    End Select
    
    SetDestSelector
    DoEvents
End Sub

'---- GUI: Set Destination Selector Tab
Private Sub SetDestSelector()
    Dim i As Integer
    
    For i = 0 To 3
        If i = DstMode Then
            lblDstMode(i).Font.Bold = True
            lblDstMode(i).ForeColor = TabTheme(0, i + 2)                    'Selected Foreground colour
            lblDstMode(i).BackColor = TabTheme(1, i + 2)                    'Selected Background colour
        Else
            lblDstMode(i).Font.Bold = False
            lblDstMode(i).ForeColor = TabTheme(2, i + 2)                    'Unselected Foreground colour
            lblDstMode(i).BackColor = TabTheme(3, i + 2)                    'Unselected Background colour
        End If
    Next i

    DoEvents
    
End Sub

'--- GUI: Click to Show/Hide Directory Paths
Private Sub lblPathView_Click(Index As Integer)
    Dim W1 As Integer, W2 As Integer, L1 As Integer
    
    L1 = dirLocal(Index).Left
    W1 = frSrc(Index).Width - 210
    W2 = drvLocal(Index).Width
    
    If lblPathView(Index).Caption = ">>" Then
        lblPathView(Index).Caption = "<<"
        drvLocal(Index).Visible = True
        dirLocal(Index).Visible = True
        lstLocal(Index).Left = L1 + W2
        lstLocal(Index).Width = W1 - W2
    Else
        lblPathView(Index).Caption = ">>"
        drvLocal(Index).Visible = False
        dirLocal(Index).Visible = False
        lstLocal(Index).Left = L1
        lstLocal(Index).Width = W1
    End If
    
    DoEvents
    
End Sub

'---- GUI: Set the positions and sizes on the form
' This sets the LEFT, CENTRE, and RIGHT section element positions
Sub SetLayout()
    Dim C0 As Single, C1 As Single, C2 As Single, C3 As Single  'Left Edges of "columns"
    Dim W0 As Single, W1 As Single, W2 As Single, WB As Single  'Widths of Elements
    Dim FWid As Single, FHi As Single                           'Form Width and Height
    Dim S As Integer, T As Single, TT As Single, H As Single    'Sizing Variables
    
    S = 90                                                      'Spacing between things
    TT = lblDrag.Height + S                                     'Top of all Buttons below Titlebar dragbar
    T = TT + lblSrcMode(0).Height + S                           'Top of all Frames - touching bottom of Buttons
    H = frSrc(0).Height                                         'Height of all frames
        
    W0 = frSrc(0).Width                                         'Width of LEFT frame
    W1 = frMiddle.Width                                         'Width of MIDDLE
    W2 = frX.Width                                              'Width of RIGHT
    
    WB = (W0 - S) \ 2                                           'Width of Buttons = Half size minus Gap
    
    C0 = S                                                      'LEFT starts at S margin
    C1 = C0                                                     'LEFT defaults to same
    
    If Layout = 1 Then
        C1 = C0 + W0 + S                                        'LEFT Shift over for Dual
        WB = W0                                                 'Width of Buttons Same as frame
    End If
    
    C2 = C1 + W0 + S                                            'MIDDLE position
    C3 = C2 + W1 + S                                            'RIGHT position
        
    frSrc(0).Move C0, T
    frDDF(0).Move C1, T                                         'LEFT
    frMiddle.Move C2, T                                         'MIDDLE
    
    frDestB.Move C3, TT                                         'Right header
    
    frX.Move C3, T                                              'RIGHT
    frSrc(1).Move C3, T                                         'RIGHT
    frLink.Move C3, T                                           'RIGHT
    frDDF(1).Move C3, T                                         'RIGHT
    
    '-- Adjust LEFT Button size/position
    If Layout = 0 Then
        lblSrcMode(0).Move C0, TT, WB                           'LEFT "Local PC" Button
        lblSrcMode(1).Move C0 + WB + S, TT, WB              'LEFT "Disk Image" button
        
    Else
        lblSrcMode(0).Move C0, TT, WB                           'LEFT "Local PC" Button
        lblSrcMode(1).Move C1, TT, WB                           'LEFT "Disk Image" button
    End If

    '-- Calculate FORM Size, and set Sizing arrows
    
    FHi = lblDrag.Height + S + frSrc(0).Height + S + lblSrcMode(0).Height + S            'FORM Height
    
    If Layout2 = 0 Then                                         'Is DEST area Hidden?
        lblSizer.Caption = ">>"                                 'YES: Set sizer to ENLARGE
        FWid = C3                                               '     Width = left of DEST area
    Else
        lblSizer.Caption = "<<"                                 'NO: Set sizer to REDUCE
        FWid = C3 + W2 + S                                      '    Width is Full
    End If

    Select Case Layout
        Case 0                                                  '-- SINGLE Mode
            lblSizer2.Caption = ">>"
            SetSrcFrame                                         'Set SOURCE frame positions
            
        Case 1                                                  '-- DUAL Mode
            lblSizer2.Caption = "<<"
            frSrc(0).Visible = True
            frDDF(0).Visible = True
    End Select

    '-- Set the Form Size and Titlebar controls
    
    Me.Height = FHi                                            'Force HEIGHT of window DEBUG!
    Me.Width = FWid                                            'Set the Main Window width
    
    '-- Set the title bar drag area, close and minimize - after the form width is set
    
    lblDrag.Width = Me.Width                                    'Set the drag area
    cmdMinimize.Left = Me.Width - 800                           'Move the Minimize button
    cmdClose.Left = Me.Width - 400                              'Move the close button
    
    DoEvents
        
End Sub


'========================================
' THEME SUBS
'========================================
' Subs that deal with Theme Loading and Setting


'---- THEME: Change Theme with LEFT/RIGHT Controls
Private Sub cmdThemeSel_Click(Index As Integer)
    
    If Index = 0 Then
        Theme = Theme - 1
    Else
        Theme = Theme + 1
    End If
    
    SetTheme
    
End Sub

'---- THEME: Change CBM Font colours
Private Sub lblFGBG_Click(Index As Integer)
    Dim i As Integer
    
    frmColourPicker.Show vbModal                                                'Set form to show Modal
    
    If PickedColour >= 0 Then
        lblFGBG(Index).BackColor = PickedColour                                 'Set the Box to the picked colour
        If Index = 0 Then
            ThemeListFG = PickedColour                                          'Set Foreground colour
        Else
            ThemeListBG = PickedColour                                          'Set Background colour
        End If
        
        CreateFontPixels ThemeListFG, ThemeListBG                               'Generate new pixel byte bitmap in theme colours
        RenderFont ThemeListFG, ThemeListBG                                     'Render new font in Theme colours
        
        For i = 0 To 3: RefreshList i: Next                                     'Re-draw the lists with the new theme fonts
    End If
    
End Sub
'---- THEME: Build Theme Menu
' This Adds Names from the Theme File List Box to the Theme PopUp Menu
Private Sub MakeThemeMenu()
    Dim i As Integer, MaxTheme As Integer, Tmp As String
    Dim P As Integer
    
    MaxTheme = flbThemes.ListCount - 1
        
    For i = 0 To MaxTheme
        Load frmMenu.mnuTheme.Item(i + 1)                                       'This ADDS a new Menu Item using Item(0) as a template
        Tmp = FileNameOnly(flbThemes.List(i))                                   'Get filename from FLB
        Tmp = Mid(Tmp, 7, Len(Tmp) - 10)                                        'Strip off "theme-" and ".bmp"
        P = InStr(2, Tmp, ".")                                                  'Check for "." for optional font file
        If P > 0 Then Tmp = Left(Tmp, P - 1)
        frmMenu.mnuTheme.Item(i + 1).Caption = UCase(Tmp)                       'Set the Theme Name
    Next i
    
    frmMenu.mnuTheme.Item(0).Visible = False                                    'Hide the first menu item
    
End Sub

'---- THEME: Show Theme PopUp Menu
Private Sub lblTheme_Click()
    
    MenuForm = 1                                                                'Which Form to notify
    MenuNum = 4                                                                 'Which List to use
    
    PopupMenu frmMenu.mnuT                                                      'Show the menu

End Sub

'---- THEME: Set the Theme
' This looks at the THEME variable as an index to the theme list
' If the Theme bitmap exists then it is loaded and the Theme menu checkmark is updated
' If there are no themes it will just use the built-in bitmap and disable the theme selection controls
' Finally, it calls the sub to Extract Theme colours from the bitmap
Private Sub SetTheme()
    Dim ThemeFile As String, ThemeName As String
    Dim FontName As String, Filename As String
    Dim Tmp As String
    Dim i As Integer, MaxTheme As Integer, P As Integer
    Dim Flag As Boolean
    
    Static LastTheme As String, LastFont As String                              'For saving Last loaded Theme
    
    flbThemes.Refresh                                                           'Refresh to discover new themes ##### will need to update menu
    Flag = False                                                                'Assume No themes
    MaxTheme = flbThemes.ListCount - 1                                          'Maximum Theme#
    
    If MaxTheme >= 0 Then
        If Theme > MaxTheme Then Theme = MaxTheme                               'Check of out of bounds
        ThemeFile = flbThemes.List(Theme)
        Filename = ThemeDir & ThemeFile                                         'Point to Theme in folder
        Flag = True                                                             'There are themes available - flag to enable dropdown
    End If
    
    cmdThemeSel(0).Visible = Flag                                               'Show or Hide theme controls
    cmdThemeSel(1).Visible = Flag
    lblTheme.Visible = Flag
    
    If (Filename <> "") And (Exists(Filename) = True) Then
        LastTheme = Filename                                                    'Remember last file. Filename format: theme-{name}{-font}.bmp
        
        picTheme.Cls
        picTheme.Picture = LoadPicture(Filename)                                'Load the new Theme Bitmap file

        ThemeListFG = GetTheme(50, 15)
        ThemeListBG = GetTheme(50, 25)
        
        Tmp = FileNameBase(Filename)                                            'Get filename without path or extension
        Tmp = Mid(Tmp, 7)                                                       'Strip off "theme-"
        FontName = Tmp                                                          'Default to Font with the same name as the theme
        ThemeName = Tmp
        
        P = InStr(2, Tmp, ".")                                                  'Look for a "." for optional Font name at end
        If P > 0 Then
            ThemeName = Left(Tmp, P - 1)
            FontName = Mid(Tmp, P + 1)
        End If
        
        lblTheme.Caption = UCase(ThemeName)                                     'Set the Theme Label
        
        Filename = ThemeDir & "cxfont-" & FontName & ".bin"                     'Look for matchine BINARY Font file
        
        'Debug.Print "Theme     ="; Theme; " "; Tmp
        'Debug.Print "ThemeName ="; ThemeName
        'Debug.Print "FontName  ="; FontName
        
        If Filename <> LastFont Then                                            'Do not re-load same font!
            If Exists(Filename) = True Then                                     'Check if it exists
                If LoadBuffer(FontBuf, Filename) > 0 Then                       'Load the Font Buffer
                    FontBMPFlag = False                                         'FLAG=true to indecate need to render BIN font
                End If
                lblFGBG(0).Visible = True
                lblFGBG(1).Visible = True
                LastFont = Filename                                             'Remember it
            Else
                Filename = ThemeDir & "cxfont-" & FontName & ".bmp"             'Look for matchine BINARY Font file
                If Exists(Filename) = True Then
                    picCBM.Picture = LoadPicture(Filename)                      'Load it
                    FontBMPFlag = True                                          'FLAG=false to indicate BMP font
                    lblFGBG(0).Visible = False
                    lblFGBG(1).Visible = False
                    LastFont = Filename
                End If
            End If
        End If
        
        '-- Set the Theme Menu Checkmark
        For i = 1 To frmMenu.mnuTheme.UBound                                    '-- Uncheck all Menu Items
            frmMenu.mnuTheme.Item(i).Checked = False                            'Un-check
        Next i
        
        frmMenu.mnuTheme.Item(Theme + 1).Checked = True                         'Check the new Theme
    End If
    
    DoTheme                                                                     'Set colours and Icons, Refresh Lists
    
    If FontBMPFlag = False Then                                                 'TRUE means BIN font, FALSE=BMP Font
        CreateFontPixels ThemeListFG, ThemeListBG                               'Generate new pixel byte bitmap in theme colours
        RenderFont ThemeListFG, ThemeListBG                                     'Render new font in Theme colours
    End If
    
    For i = 0 To 3: RefreshList i: Next                                         'Re-draw the lists with the new theme fonts
    frmViewer.SetVTheme                                                         'Update Viewer Theme
    
End Sub

'---- THEME: Do the Theme
' Extracts the colours and icons from the currently loaded Theme Bitmap
Private Sub DoTheme()
    Dim i As Integer, J As Integer, FG As Integer, BG As Integer
    Dim X As Integer, Y As Integer
    
    On Error Resume Next
    
    '-- Get the Theme Colours

    FG = 15: BG = 25                                                    'Y positions of FG and BG
    
    ThemeFG = GetTheme(23, FG)
    ThemeBG = GetTheme(23, BG)
    ThemeTitleFG = GetTheme(32, FG)
    ThemeTitleBG = GetTheme(32, BG)
    ThemeMenuFG = GetTheme(41, FG)
    ThemeMenuBG = GetTheme(41, BG)
    ThemeListFG = GetTheme(50, FG)
    ThemeListBG = GetTheme(50, BG)
    ThemeFrFG = GetTheme(59, FG)
    ThemeFrBG = GetTheme(59, BG)
    ThemeFr2FG = GetTheme(68, FG)
    ThemeFr2BG = GetTheme(68, BG)
    
    '-- Get the Icons
    
    Y = 67                                                              'Most icons top pixel aligned to here
    
    GetIcon cmdCopyLeft, 3, Y                                           'Get Copy-Left Icon
    GetIcon cmdCopyRight, 62, Y                                         'Get Copy-Right Icon
    GetIcon cmdConfig, 121, Y                                           'Get Config Icon
    GetIcon cmdAbout, 161, Y                                            'Get About Icon
    GetIcon cmdHelp, 186, Y                                             'Get Help Icon
    GetIcon cmdDAD, 161, 93                                             'Get DAD Icon
    
    GetIcon cmdXMenu, 288, Y                                            'Get X-cable Options Icon
    
    For i = 0 To 1
        GetIcon cmdBrowse(i), 225, Y                                    'Get Browse For file Icon
        GetIcon cmdPathUp(i), 330, Y                                    'Get Directory Path Up Icon
        GetIcon cmdLocalMenu(i), 288, Y                                 'Get Local Menu Options Icon
    Next i

    For i = 0 To 3
        GetIcon cmdEncode(i), 246, Y                                    'Get Encode Icon
        GetIcon cmdImageMenu(i), 288, Y                                 'Get DiskImage Options Icon
        GetIcon cmdScale(i), 309, Y                                     'Get Height Scale Icon
    Next i
    
    For i = 1 To 3
        cmdImageMenu(i).ToolTipText = cmdImageMenu(0).ToolTipText       'Set Tooltip same as 0
        cmdScale(i).ToolTipText = cmdScale(0).ToolTipText               'Set Tooltip same as 0
    Next
    
    '-- Get Tab Colours
    
    For i = 0 To 3
        For J = 0 To 5
            TabTheme(i, J) = GetTheme(80 + J * 9, 13 + i * 5)           'Get Colours from bitmap (row,col)
        Next J
    Next i
    
    '-- Set the Main Form and Titlebar
    
    frmMain.BackColor = ThemeBG: frmMain.ForeColor = ThemeFG                            'Set Main Window Colours
    frDestB.BackColor = ThemeBG                                                         'Set Destination Selector Frame Background
    lblDrag.ForeColor = ThemeTitleFG: lblDrag.BackColor = ThemeTitleBG
    cmdMinimize.ForeColor = ThemeTitleFG
    cmdClose.ForeColor = ThemeTitleFG
    
    '-- Set Menu Area
        
    frMiddle.BackColor = ThemeMenuBG:   frMiddle.ForeColor = ThemeMenuFG
    
    lblSizer.ForeColor = ThemeMenuFG
    lblSizer2.ForeColor = ThemeMenuFG
    lblTheme.ForeColor = ThemeMenuFG:
        
    '-- Set Frames and Elements inside them
    
    frX.BackColor = ThemeFrBG:                  frX.ForeColor = ThemeFG                 'X-Cable Frame
    frLink.BackColor = ThemeFrBG:               frLink.ForeColor = ThemeFG              'CBM-Link Frame
    
    For i = 0 To 1
        frSrc(i).BackColor = ThemeFrBG:         frSrc(i).ForeColor = ThemeFrFG          'Source Frame
        frDDF(i).BackColor = ThemeFrBG:         frDDF(i).ForeColor = ThemeFrFG          'Destination Frams
        
        lblPathView(i).ForeColor = ThemeFrFG
        
        txtLocalDir(i).BackColor = ThemeListBG: txtLocalDir(i).ForeColor = ThemeListFG  'Local PC
        lstLocal(i).BackColor = ThemeListBG:    lstLocal(i).ForeColor = ThemeListFG
        dirLocal(i).BackColor = ThemeListBG:    dirLocal(i).ForeColor = ThemeListFG
        drvLocal(i).BackColor = ThemeListBG:    drvLocal(i).ForeColor = ThemeListFG
        cboFilter(i).BackColor = ThemeListBG:   cboFilter(i).ForeColor = ThemeListFG
        KBText(i).BackColor = ThemeListBG:      KBText(i).ForeColor = ThemeListFG
        BlockText(i).BackColor = ThemeListBG:   BlockText(i).ForeColor = ThemeListFG
        
        lblDDFile(i).BackColor = ThemeListBG:   lblDDFile(i).ForeColor = ThemeListFG    'Disk Image
        lblExt(i).BackColor = ThemeListBG:      lblExt(i).ForeColor = ThemeListFG
        
        cmdThemeSel(i).ForeColor = ThemeMenuFG                                          'Menu Area
    Next i
    
    For i = 0 To 3
        DFBlocksFree(i).BackColor = ThemeFrBG:  DFBlocksFree(i).ForeColor = ThemeFrFG
    Next i
    
    For i = 0 To 22
        Label(i).ForeColor = ThemeFrFG                                                  'Normal Text Labels
    Next i

    lblXLastStatus.BackColor = ThemeListBG:     lblXLastStatus.ForeColor = ThemeListFG
    lblLinkLastStatus.BackColor = ThemeListBG:  lblLinkLastStatus.ForeColor = ThemeListFG
    cboXDevNum.BackColor = ThemeListBG:         cboXDevNum.ForeColor = ThemeListFG
    cboLinkDev.BackColor = ThemeListBG:         cboLinkDev.ForeColor = ThemeListFG
    
    lblDName.BackColor = ThemeListBG:           lblDName.ForeColor = ThemeListFG
          
    lblFGBG(0).BackColor = ThemeListFG
    lblFGBG(1).BackColor = ThemeListBG
    
    '-- Re-draw elements
    
    SetSourceSelector                                                                       'Set Source Selector Tabs
    SetDestSelector                                                                         'Set Destination Selector Tabs
    
End Sub


'========================================
'  Subs for Program Options
'========================================
' Subs for General Program Options


'---- OPTIONS: Show the CBM-Transfer TXT file documentation, using associated viewer (ie: Notepad)
Private Sub cmdHelp_Click()
    ViewFile ExeDir & "\CBMXfer.txt"
End Sub

'---- OPTIONS: Open the options window
Private Sub cmdConfig_Click()
    
    Call frmOptions.SetTheme
    frmOptions.Show vbModal
    
End Sub


'---- GENERAL: Show Output Results
' Opens program Associated with TXT files to view Output Results
Private Sub cmdResults_Click()
    
    If frmOptions.cbErr.value = vbChecked Then ViewFile TEMPFILE2   'This file is usually empty
    ViewFile TEMPFILE1                                              'This contains the actual data

End Sub


'========================================
' Subs for Popup menu
'========================================

'---- GUI: Show Local Directory Options Menu
Private Sub cmdLocalMenu_Click(Index As Integer)
    MenuForm = 1                'Which Form to notify
    MenuNum = Index
    PopupMenu frmMenu.mnuL

End Sub

'---- GUI: Show Image Options Menu
Private Sub cmdImageMenu_Click(Index As Integer)
    
    MenuForm = 1                'Which Form to notify
    MenuNum = Index             'Which List to use
    
    PopupMenu frmMenu.mnuI      'Show the menu

End Sub

'---- GUI: Handle Menu Selections
' This is called from frmMenu
Sub DoMenu(ByVal Index As Integer)
    Dim Tmp As String
    
    Select Case Index                                       '--- Menu 1 - Local Options
        
        Case 1: ShellExecute hWnd, "open", LocalDir(0), vbNullString, LocalDir(0), 1
        Case 2: ShellExecute hWnd, "open", LocalDir(1), vbNullString, LocalDir(1), 1
        Case 3: SwapDirs
        Case 4: SetLocalPath 0, LocalDir(1)                 'Make LEFT  path = RIGHT path
        Case 5: SetLocalPath 1, LocalDir(0)                 'Make RIGHT path = LEFT path
        Case 6: AddPathHistory LocalDir(0)                  'Add Path to History
        Case 7: RemovePathHistory LocalDir(0)               'Remove Current Path from History
        Case 8:
            KillFile HistoryFile                            'Clear History File
            frmMain.txtLocalDir(0).Clear                    'Clear LEFT List
            frmMain.txtLocalDir(1).Clear                    'Clear RIGHT List
            
        Case 9:
            Tmp = NewFolder(LocalDir(MenuNum))              'Get Name and create folder. Return path if sucessful
            If Tmp <> "" Then
                If MsgBox("Do you want to switch to this directory?", vbYesNo) = vbYes Then SetLocalPath MenuNum, Tmp
            End If
            
            dirLocal(MenuNum).Refresh                       'Refresh the Directory List
            
        '--- Menu 2 - Directory Options
        
        Case 101: SaveDirText                               'Save Directory List as Text
        Case 102: WriteDirTextTo CatalogFile, True          'Add To the Catalog
        Case 103: ViewFile CatalogFile                      'View the Catalog File
        Case 104: ImageValidate MenuNum                     'Validate the Image
        Case 105: ImageBackup MenuNum                       'Backup the Image
        Case 106 To 109: ImageSort MenuNum, Index - 105     'Sort the Directory (Method must be 1 to 4)
        
        '--- Menu 4 - Select Theme
        
        Case 201 To 299
            Theme = Index - 201                             'Which Theme Number is it?
            SetTheme                                        'Set the Theme
            
        '--- Menu 5 - Encoding Options
        
        Case 300 To 399
            EncodeL(MenuNum) = Index - 301                  'Set the Encoding
            SetEncodeDesc                                   'Set the ToolTip
            RefreshList MenuNum                             'RE-draw the list

        '--- Menu 6 - Device Control
        
        Case 400 To 499
            DoDeviceMenu Index - 400                        'Handle Device Options
    End Select
    
End Sub

'=====================================================
' Subs for Local PC /// Src=Left Side, Dst=Right Side
'=====================================================

'---- LOCAL: Swap Left and Right directories
Private Sub SwapDirs()
    Dim Tmp
    
    Tmp = LocalDir(0): LocalDir(0) = LocalDir(1): LocalDir(1) = Tmp
    
    SetLocalPath 0, LocalDir(0)
    SetLocalPath 1, LocalDir(1)

End Sub

'---- LOCAL: Click to Change Drive Letter
Private Sub drvLocal_Change(Index As Integer)
    Dim Tmp As String
    
    Tmp = Left(drvLocal(Index).Drive, 2)                                            'fix for drives that show up with volume label
    If Tmp = "" Then Exit Sub
    
    On Local Error GoTo 0
    
    If DirExists(Tmp) = True Then
        dirLocal(Index).Path = Tmp                                                  'Set directory path.
    Else
        MyMsg "The path: " & drvLocal(Index).Drive & Cr & "Index=" & Str(Index) & Cr & "is not available!"
    End If
    
End Sub

'---- LOCAL: Change Local Directory Path
Private Sub dirLocal_Change(Index As Integer)
    
    SetLocalPath Index, dirLocal(Index).Path        'Set file path.

End Sub

'---- LOCAL: Browse to Change Folder
Private Sub cmdBrowse_Click(Index As Integer)
    Dim Tmp As String, Tmp2 As String
    
    Select Case Index
        Case 0: Tmp2 = "Select LEFT Path:"
        Case 1: Tmp2 = "Select RIGHT Path:"
    End Select
    
    Tmp = GetBrowseDir(Me, Tmp2)                                        'Display Window's built-in Select Path Dialog
    
    If Tmp <> "" Then
        SetLocalPath Index, Tmp                                         'Change to selected Path
        If AddPathFlag = True Then AddPathHistory Tmp                   'Add it to the History
    End If
    
End Sub

'---- LOCAL: Move UP one level in Path
Private Sub cmdPathUp_Click(Index As Integer)
    
    SetLocalPath Index, PathUp(LocalDir(Index))

End Sub

'---- LOCAL: Handle VIEW Button
Private Sub cmdSrcView_Click(Index As Integer)
    CheckSelected Index, 0
End Sub

'---- LOCAL: Handle VIEW 2 Button
Private Sub cmdSrcView2_Click(Index As Integer)
    CheckSelected Index, 1
End Sub

'========================================
' COMMON SUBS
'========================================


'---- LOCAL: Change LocalDir to item clicked on
Private Sub txtLocalDir_Click(Index As Integer)
    Dim Tmp As String
    
    Tmp = txtLocalDir(Index).List(txtLocalDir(Index).ListIndex)
    SetLocalPath Index, Tmp
End Sub

'---- LOCAL: Process keystrokes for Directory Path
Private Sub txtLocalDir_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then SetLocalPath Index, txtLocalDir(Index).Text
    
End Sub

'---- LOCAL: Set LocalPC Directory Path
Private Sub SetLocalPath(ByVal Index As Integer, ByVal SPath As String)
    On Local Error Resume Next
    
    LocalDir(Index) = AddSlash(SPath)
    lstLocal(Index).Path = LocalDir(Index)                      'Set the File List Box Path
    lstLocal(Index).Refresh                                     'Refresh the File List
    txtLocalDir(Index).Text = LocalDir(Index)                   'Show the Path
    txtLocalDir(Index).ToolTipText = LocalDir(Index)            'Set the tooltip
    dirLocal(Index).Path = LocalDir(Index)                      'Set the Directory List Box Path
    drvLocal(Index).Drive = Left(SPath, 2)                      'Set tje Directory List Box Drive Letter

End Sub

'---- LOCAL: Refresh Directory list
Private Sub cmdSrcRefresh_Click(Index As Integer)
    lstLocal(Index).Refresh
End Sub

'---- LOCAL: Delete File
Private Sub cmdSrcDelete_Click(Index As Integer)
    Dim T As Integer, FSel As Integer, Filename As String, OneName As String
    
    '-- Count number of files to delete
    FSel = 0
    For T = 0 To lstLocal(Index).ListCount - 1
        If (lstLocal(Index).Selected(T)) Then FSel = FSel + 1: OneName = lstLocal(Index).List(T) 'Remember filename if single file
    Next T
    
    '-- Prompt
    If FSel = 1 Then Filename = Quoted(OneName) Else Filename = Str(FSel) & " file(s)"  'Single filename or number of files
    If MsgBox("Are you sure you want to delete " & Filename & "?", vbYesNo, "Confirm delete") <> vbYes Then Exit Sub
    
    OneName = UnQuoted(DDFile(Index))                                                               'This is the currently open Disk Image file incase user tried to delete it
    
    '-- Delete them
    For T = 0 To lstLocal(Index).ListCount - 1
        If (lstLocal(Index).Selected(T)) Then
            Filename = LocalDir(Index) & lstLocal(Index).List(T)
            If Filename = OneName Then
                If MsgBox("The file you are trying to delete is open in the disk view!" & Cr & "Are you sure you want to delete " & Filename & "?", vbYesNo, "Confirm delete") <> vbYes Then
                    Filename = "" 'If NO then clear filename so it wont be deleted!
                Else
                    ClearDD Index
                End If
            End If
            
            If Filename <> "" Then KillFile Filename
        End If
    Next T
    
    lstLocal(Index).Refresh
End Sub


'========================================
' GENERAL SUBS
'========================================
' General routines used by multiple sections


'---- GENERAL: Add Path to History List
' Adds a Path to the History if it's not already included
Sub AddPathHistory(ByVal SPath As String)
    Dim a As Integer, Flag As Integer
    
    Flag = True
    
    For a = 0 To txtLocalDir(0).ListCount - 1
        If txtLocalDir(0).List(a) = SPath Then Flag = False: Exit For   'Set Flag if path is already in the list
    Next
    
    If Flag = True Then
        txtLocalDir(0).AddItem SPath                                    'Adds the item to LEFT History dropdown list
        txtLocalDir(1).AddItem SPath                                    'Adds the item to RIGHT History dropdown list
    End If
    
End Sub

'---- GENERAL: Remove Path from History List
Sub RemovePathHistory(ByVal SPath As String)
    Dim a As Integer

    For a = 0 To txtLocalDir(0).ListCount - 1                   'Go through the existing entries
        If txtLocalDir(0).List(a) = SPath Then
            txtLocalDir(0).RemoveItem (a)                       'Remove item from LEFT History dropdown list
            txtLocalDir(1).RemoveItem (a)                       'Remove item from RIGHT History dropdown list
        End If
    Next
    
End Sub

'---- GENERAL: Save Directory Listing to Text File
Private Sub SaveDirText()
        Dim Filename As String, TmpF As String, TmpP As String
        
        On Local Error GoTo DialogError
        
        If MenuNum < 2 Then
            TmpP = LocalDir(MenuNum)                                'Set default path to same as Image Filename
            TmpF = FileBase(UnQuoted(DDFile(MenuNum))) & ".txt"     'If Disk Image then use Image filename with TXT extension
        Else
            TmpP = CurDir
            TmpF = "listing.txt"                                    'If X Drive then use generic filename
        End If
        
        CommonDialog1.CancelError = True
        CommonDialog1.InitDir = TmpP                                'Set Path
        CommonDialog1.Filter = "Text (*.txt)|*.txt|All Files|*.*"   'Set Filter menu
        CommonDialog1.Filename = TmpF                               'Provide default filename
        
        CommonDialog1.ShowSave                                      'Display the dialog
        
        Filename = CommonDialog1.Filename                           'Get Returned filename
        
        If Filename = "" Then Exit Sub                              'Exit if filename not specified
        If Overwrite(Filename) = False Then Exit Sub                'Exit if User doesn't want to overwrite the exitsting file
                
        WriteDirTextTo Filename, False                              'Write it!
                
DialogError:
        Exit Sub
End Sub

'---- GENERAL: Save the disk directory to a file
' Flag: True=Append, False=Create
Private Sub WriteDirTextTo(ByVal Filename As String, Flag As Boolean)
        Dim FIO As Integer, J As Integer
        
        FIO = FreeFile
        
        If Flag = False Then
            Open Filename For Output As FIO                                         'Create a New file
        Else
            Open Filename For Append As FIO                                         'Append to Existing file
            Print #FIO, ""
        End If
        
        If MenuNum < 2 Then
            Print #FIO, "**** FILE: " & FileNameOnly(lblDDFile(MenuNum).Caption)    'Write filename if Image
        End If
        
        Print #FIO, Qu & DiskName(MenuNum) & Qu & " " & DiskID(MenuNum)             'Write Header
        Print #FIO, "========================"
        
        For J = 0 To lstImageFiles(MenuNum).ListCount - 1
            Print #FIO, lstImageFiles(MenuNum).List(J)                              'Write one file entry
        Next J
        
        Print #FIO, DFBlocksFree(MenuNum).Caption                                   'Write Blocks Free + # of files
        Print #FIO, ""                                                              'Write a blank line
        
        Close FIO

End Sub


'======================
' Subs for Disk Images
'======================

'---- DISKIMAGES: Make a NEW Disk Image File
' Asks user for Image name with valid Image Extension (ie: D64)
Private Sub cmdNewImage_Click(Index As Integer)
    Dim Filename As String, Ext As String, P As Integer
    
    frmPrompt.Reply.Text = "new.d64"
    frmPrompt.Ask "Create new Image File", "Enter Image Filename (include correct extension):", 1, False
    If Response = "" Then Exit Sub
    
    Ext = FileExtU(Response)
    If SupportedImg(Ext, True) = False Then MyMsg "You must enter a valid image extension (D64,D71,D81 etc)!": Exit Sub
    Filename = Response
    
    frmPrompt.Reply.Text = "title,id"
    frmPrompt.Ask "Enter Header Info", "Enter Diskname (header title) and Disk ID", 1, False
    If Response = "" Then Exit Sub
    
    If InStr(1, Response, ",") = 0 Then
        MyMsg "You must enter TITLE,ID!"
    Else
        DoCommand CBMC1541, _
                  "-format " & Quoted(Response) & " " & Ext & " " & Quoted(LocalDir(Index) & Filename), _
                  "Creating new " & Ext & " file named '" & Filename & "'"
    End If
    
    lstLocal(Index).Refresh
    
End Sub

'---- DISKIMAGES: Click to View a file inside Disk Image
Private Sub cmdDView_Click(Index As Integer)
    ImageFileView Index
End Sub

'---- DISKIMAGES: View file(s) inside Disk Image
' Goes through Entries and Views the first Selected if it is PRG, SEQ or USR.
Private Sub ImageFileView(Index As Integer)
    Dim T As Integer, Tmp As String, Filename As String, Ext As String, Mode As Integer

    For T = 0 To lstImageFiles(Index).ListCount - 1
        If lstImageFiles(Index).Selected(T) = True Then
            Tmp = lstImageFiles(Index).List(T)                                  'Get the selected line
            Filename = CBMName(Tmp)                                             'Extract filename,p
            Ext = CBMType(Filename)                                             'Get the ",p" or ",s" etc
            Tmp = UCase(Right(Ext, 1))                                          'Get the "P"
            
            KillFile TEMPFILE3
            
            Select Case Tmp
            
                Case "P", "S", "U", "R"
                    
                    'If Right(Filename, 2) = ",r" Then Mid(Filename, Len(Filename) - 1, 2) = ",l" 'Fix for REL files??????????????????????????
                    
                    DoCommand CBMC1541, _
                              DDFile(Index) & " -read " & Quoted(Filename) & " " & Quoted(TEMPFILE3), _
                              "Copying '" & Filename & "' from image..."
                            
                    If Exists(TEMPFILE3) = True Then
                        frmViewer.Show: DoEvents                                'Make it open
                        While frmViewer.ViewerReady = False: Wend               'Wait for the Form to load!
                        DoEvents
                        Mode = 0: If Ext = ",s" Then Mode = 1                   'If ",s" then select SEQ tab
                        frmViewer.ViewIt Mode, TEMPFILE3, Filename, Ext         'View it!
                    Else
                        Warning 1, Filename                                     'Could not extract 'Filename' from image
                    End If
                    Exit For                                                    'Only view one file at a time
                    
                Case Else
                
                    MsgBox "Filetype: " & Ext & " - Sorry only PRG/SEQ/USR/REL files can be viewed."
                    
            End Select
        End If
    Next T
    
End Sub

'---- DISKIMAGES: Delete a File(s) within Disk Image
Private Sub cmdDDelete_Click(Index As Integer)
    Dim T As Integer, FSel As Integer, Filename As String, OneName As String

    '-- Count how many files are selected
    
    FSel = 0
    For T = 0 To lstImageFiles(Index).ListCount - 1
        If lstImageFiles(Index).Selected(T) = True Then FSel = FSel + 1: OneName = CBMName(lstImageFiles(Index).List(T))
    Next T
    If FSel = 0 Then Exit Sub
    
    '-- Prompt
    
    If FSel = 1 Then Filename = Quoted(OneName) Else Filename = Str(FSel) & " file(s)"  'Single filename or number of files
    If MsgBox("Are you sure you want to delete " & Filename & "?", vbYesNo, "Confirm delete") <> vbYes Then Exit Sub
    
    '-- Delete selected files
    
    For T = 0 To lstImageFiles(Index).ListCount - 1
        If lstImageFiles(Index).Selected(T) = True Then
            Filename = CBMName(lstImageFiles(Index).List(T))
            
            DoCommand CBMC1541, _
                      DDFile(Index) & " -delete " & Quoted(Filename), _
                      "Deleting '" & Filename & "..."
        End If
    Next T
    
    ImageRefresh Index
    
    
End Sub

'---- DISKIMAGES: Validate Disk Image
Private Sub ImageValidate(ByVal Index As Integer)
    If Exists(UnQuoted(DDFile(0))) = False Then MyMsg "Select an image first!": Exit Sub
    If MsgBox("Validate Image?", vbYesNo) <> vbYes Then Exit Sub
    
    DoCommand CBMC1541, _
            DDFile(Index) & " -validate", "Validating Image..."
            
    ImageRefresh Index
End Sub

'---- DISKIMAGES: Backup Disk Image
Private Sub ImageBackup(ByVal Index As Integer)
    Dim Filename As String, Filename2 As String
    On Local Error GoTo ImgBakErr
    
    Filename = UnQuoted(DDFile(Index))
    Filename2 = FileBase(Filename) & ".bak"
    
    If Exists(Filename) = False Then MyMsg "Select an image first!": Exit Sub
    If MsgBox("Make Backup of Image File '" & Filename & "'?", vbYesNo) <> vbYes Then Exit Sub
    If Overwrite(Filename2) = False Then Exit Sub
    KillFile Filename2
    FileCopy Filename, Filename2
    lstLocal(Index).Refresh

ImgBakErr:

End Sub

'---- DISKIMAGES: Sort Image Files
' Simple bubble sort. Index = Disk Image
' Method = 1-4, where: 1=Filename A-Z, 2=Filename Z-A, 3=File Size 0-9, 4=File Size 9-0
Private Sub ImageSort(ByVal Index As Integer, ByVal Method As Integer)
    Dim L1 As String, L2 As String                              'List Entries
    Dim F1 As String, F2 As String                              'Fields to Compare
    Dim S1 As Boolean, S2 As Boolean                            'Selected flags
    
    Dim MaxLine As Integer, P As Integer, L As Integer
    Dim J As Integer, K As Integer, Flag As Boolean
       
    MaxLine = lstImageFiles(Index).ListCount - 1: If MaxLine < 2 Then Exit Sub
    
    Select Case Method
        Case 1 To 2: P = 7: L = 16                              'Sort by Name (between quotes
        Case 3 To 4: P = 1: L = 5                               'Sort by Size
    End Select
    
    For J = 0 To MaxLine - 1
        For K = 0 To MaxLine - 1 - J
            L1 = lstImageFiles(Index).List(K)                   'First Entry
            L2 = lstImageFiles(Index).List(K + 1)               'Second Entry
            F1 = Mid(L1, P, L)                                  'Field of First Entry
            F2 = Mid(L2, P, L)                                  'Field of Second Entry
            S1 = lstImageFiles(Index).Selected(K)               'Hilight of First Entry
            S2 = lstImageFiles(Index).Selected(K + 1)           'Hilight of Second Entry
            
            Flag = False                                        'Assume no swap
            
            Select Case Method
                Case 1: If F1 > F2 Then Flag = True             'Check string Greater  (A-Z)
                Case 2: If F1 < F2 Then Flag = True             'Check string Less     (Z-A)
                Case 3: If Val(F1) > Val(F2) Then Flag = True   'Check numeric Greater (0-9)
                Case 4: If Val(F1) < Val(F2) Then Flag = True   'Check numeric Less    (9-0)
            End Select
            
            If Flag = True Then
                lstImageFiles(Index).List(K) = L2               'Swap Entry
                lstImageFiles(Index).List(K + 1) = L1           'Swap Entry
                lstImageFiles(Index).Selected(K) = S2           'Swap Hilighting
                lstImageFiles(Index).Selected(K + 1) = S1       'Swap Hilighting
            End If
        Next K
    Next J

    RefreshList Index
    
End Sub

'---- DISKIMAGES: Click to Rename File(s) within Disk Image
Private Sub cmdDRename_Click(Index As Integer)
    
    ImageFileRename Index

End Sub

'---- DISKIMAGES: Rename File(s) within Disk Image
Sub ImageFileRename(ByVal Index As Integer)
    Dim T As Integer, FSel As Integer, Filename As String, OneName As String

    '-- Count how many are selected
    
    FSel = 0
    
    For T = 0 To lstImageFiles(Index).ListCount - 1
        If lstImageFiles(Index).Selected(T) = True Then
            FSel = FSel + 1
            OneName = ExtractQuotes(lstImageFiles(Index).List(T))
        End If
    Next T
    
    If FSel = 0 Then Exit Sub                                   'Exit if none selected
    
    '-- Confirm Rename
    
    If FSel > 1 Then
        If MsgBox("Are you sure you want to Rename " & Str(FSel) & " file(s)?", vbYesNo, "Confirm Rename") <> vbYes Then Exit Sub
    End If
    
    For T = 0 To lstImageFiles(Index).ListCount - 1
        If lstImageFiles(Index).Selected(T) = True Then
            OneName = ExtractQuotes(lstImageFiles(Index).List(T))   'we just want the filename (see bug below)
            Filename = InputBox("Enter new name:", "Rename File", OneName)
            
            If (OneName <> Filename) And (Filename <> "") Then
                '-- Rename the file
                '   BUG?: If you add ",p" then the file will become "*PRG" then validating may erase it
                DoCommand CBMC1541, _
                          DDFile(Index) & " -rename " & Quoted(OneName) & " " & Quoted(Filename), _
                          "Renaming '" & OneName & "..."
            End If
        End If
    Next T
    
    ImageRefresh Index

End Sub

'---- COMMON: Click to SELECT all Entries in the list
Private Sub cmdDAll_Click(Index As Integer)
    DSelector True, Index
End Sub

'---- COMMON: Click to DE-SELECT all Entries in the list
Private Sub cmdDNone_Click(Index As Integer)
    DSelector False, Index
End Sub

'---- COMMON: Select Entries in List
' B: TRUE=Select ALL, FALSE=Select none
Private Sub DSelector(ByVal B As Boolean, ByVal Index As Integer)
    Dim J As Integer
    
    For J = 0 To lstImageFiles(Index).ListCount - 1
        lstImageFiles(Index).Selected(J) = B
    Next J
    
    RefreshList Index
    
End Sub

'=============================
' Subs for X-Cable/ZoomFloppy
'=============================

'---- XCABLE: X Cable Device Options
Private Sub cmdXMenu_Click()
    
    PopupMenu frmMenu.mnuD      'Show the device menu
    
End Sub

'---- XCABLE: Fetch the drive status strings
Private Sub cmdXDriveStatus_Click()

    GetXDriveStatus

End Sub

'---- XCABLE: Get Drive Status
Private Sub GetXDriveStatus()

    lblXLastStatus.Caption = GetXStatus()
    
End Sub

'---- XCABLE: Get X-Cable Status String
Private Function GetXStatus() As String
    Dim Status As ReturnStringType
    
    Status = DoCommand(CBMCtrl, "status " & DriveNum, "Reading drive status, please wait.")
    GetXStatus = UCase(Status.Output)
    
End Function

'---- XCABLE: Get X-Cable Status Value
Private Function GetXStatusN() As Integer
    
    GetXStatusN = Val(Left(GetXStatus, 2))

End Function

'---- XCABLE: Display X-Cable Status when Error>19
Private Function GetShowXError()
    Dim Tmp As String
   
    Tmp = GetXStatus                                                            'Read the Status
    If Val(Left(Tmp, 2)) > 19 Then MyMsg Tmp                                     'Only numbers >19 are real errors

End Function

'---- XCABLE: Detect Drives via CBMCTRL
' CBMCTRL returns a list of drives with device#'s like this:
' 8:1541<cr>
' 9:1571<cr>
' We look for the ID string for drives 8 to 30 then save them in the array, ie:  Drive(8)="1541"
' Flag=true to display results, False=silent
Public Sub DetectXDrives(ByVal Flag As Boolean)
    Dim What As String, FIO As Integer, D As Integer, Tmp As String, DTmp As String
    Dim Flag2 As Boolean
    
    frmMain.PubDoCommand CBMCtrl, "detect", "Detecting Drives...", False

    If Exists(TEMPFILE1) = False Then Exit Sub
    
    '-- Read in the complete output file
    FIO = FreeFile
    Open TEMPFILE1 For Input As FIO
        What = Input(LOF(FIO), FIO) 'was LOF(1)???
    Close FIO
    
    Flag2 = False
    
    cboXDevNum.Clear                                                    'Clear the Drive list
    
    '-- Read returned string for drives - NOTE: May need to adjust this for drives with parallel cable
    For D = 8 To 30
        Drive(D) = ""                                                   'Clear this drive's info
        DTmp = Format(D) & ":"                                          'Make a string to search with
        Tmp = MyTrim(GetNamedField(What, DTmp))                         'Get the string for the specified drive number
        
        If Tmp <> "" Then
            cboXDevNum.AddItem DTmp & "  " & Tmp
            If InStr(1, Tmp, "*") > 0 Then
                Tmp = "1541": Flag2 = True                                  'Unknown drive. We will assume it's a 1541 clone/compatible
            End If
            Drive(D) = Tmp
        End If
    Next D
    
    cboXDevNum.AddItem "(Re-Scan devices)"
    
    KillTemp                                                            'And delete both temp files, so we're not cluttering things up
    
    If Flag = True Then
        If Flag2 = True Then MyMsg "One or more drives are UNKNOWN." & Cr & "Please submit details to the OPENCBM team!" & Cr & "NOTE: CBM-Transfer treats unknown drives as a 1541!"
        If (What = "") Then What = "No drives found, please check cables and drive power."
        MsgBox What, vbOKOnly, "Drive Detection"
    End If
    
End Sub

'---- XCABLE: Format Floppy Disk
' Checks which drive is being used to determine proper formatting method (opencbm command or via dos command string)
' For 1571 will ask for double-sided format
Private Sub FormatXDrive()
    Dim Status As ReturnStringType, Tmp As String, Flag As Boolean, M1 As String
    
    Tmp = UCase(Drive(DriveNum))    'Drive Type

    If MsgBox("This will erase ALL data on the floppy in Device#" & Format(DriveNum) & Cr & " (" & Tmp & ")" & Cr & Cr & "Are you sure?", vbExclamation Or vbYesNo, "Format Disk") = vbNo Then Exit Sub
    
    frmPrompt.Reply.Text = "new,id"
    frmPrompt.Ask "Format CBM Floppy", "Please Enter Diskname, ID", 1, False
    If Response = "" Then Exit Sub
    
    Flag = True                     'Assume 1541
    
    '-- Determine proper formatting method
    Select Case Tmp
        Case "1540", "1541"
        Case "1571"
            If MsgBox("You have a 1571. Do you want to format double-sided?", vbYesNo, "Format") = vbYes Then
                Flag = False
                Status = DoCommand(CBMCtrl, CMDSTR & DriveNum & " " & Quoted("U0>M1"), "Enabling 1571 mode...")
            End If
        Case Else
            Flag = False
    End Select
    
    M1 = "Formatting floppy disk, please wait."

    If Flag = True Then
        '-- Format a 1541
        Status = DoCommand(CBMOpen & "cbmforng", " -vso " & DriveNum & " " & Quoted(UCase(Response)), M1)
        lblXLastStatus.Caption = UCase(Status.Output)
        Sleep 1000                                                          'Just so message is visible
        GetXDir
    Else
        '-- Format using standard DOS New command
        Status = DoCommand(CBMCtrl, CMDSTR & DriveNum & " " & Quoted("N0:" & UCase(Response)), M1)
        MyMsg "Formatting... You may continue working, however, do not" & Cr & "attempt to access the drive until formatting" & Cr & "is complete. Check drive status when the light goes off."
    End If
        
End Sub

'---- XCABLE: Validate Disk
Private Sub ValidateXDrive()

    DoCommand CBMCtrl, CMDSTR & DriveNum & " " & Quoted("V0:"), "Validating drive, please wait."
    GetXDriveStatus
    
End Sub

'---- XCABLE: Initialize Drive
Private Sub InitXDrive()
    
    DoCommand CBMCtrl, CMDSTR & DriveNum & " I0", "Initializing Drive"
    GetXDriveStatus

End Sub

'---- XCABLE: Click to View File
Private Sub cmdXView_Click()
    XView
End Sub

'---- XCABLE: View a File
Private Sub XView()
    Dim Filename As String, FileIn As String, FileOut As String
    Dim Ext As String, CExt As String, Tmp As String
    Dim T As Integer, Mode As Integer
          
    For T = 0 To lstImageFiles(2).ListCount - 1
        If (lstImageFiles(2).Selected(T)) Then
        
            Tmp = lstImageFiles(2).List(T)                              'The directory Line
            Filename = LCase(ExtractQuotes(Tmp))                        'Get the filename, ie: filename
            Ext = DOSExt(Tmp)                                           'Get DOS extension, ie: SEQ
            CExt = "," & LCase(Left(Ext, 1))                            'DOS extension, ie: ",s"
            FileIn = Filename & CExt                                    'Source File, ie: filename,s
            FileOut = TEMPFILE3                                         'Output File, ie: TEMPFILE3
            
            Select Case UCase(Ext)
                Case "PRG", "SEQ", "USR"
                    
                    KillFile TEMPFILE3                                  'Erase any existing temp file first
                    
                    'NOTE: If filename starts with "--" the parser gets confused, thinking it is a parameter.
                    '      We must fool it by replacing the first "-" with "?"
                    
                    If Left(FileIn, 2) = "--" Then Mid(FileIn, 1, 1) = "?"          'change "--" to "?-"
                    
                    DoCommand CBMCopy, _
                              "--transfer=" & TransferString & " -q -r " & DriveNum & " " & Quoted(FileIn) & _
                              " --output=" & Quoted(FileOut), _
                              "Reading '" & Filename & "' ..."
                    
                    If Exists(TEMPFILE3) = True Then
                        frmViewer.Show: DoEvents                                'Make it open
                        While frmViewer.ViewerReady = False: Wend               'Wait for the Form to load!
                        DoEvents
                        Mode = 0: If UCase(Ext) = "SEQ" Then Mode = 1           'Set TAB for view
                        
                        frmViewer.ViewIt Mode, TEMPFILE3, Filename, Ext
                    Else
                        Warning 3, TEMPFILE3 'could not transfer file
                    End If
                    
                Case "CBM" '1581 Partition
                    XChangePart Filename
                    GetXDir
                    
                Case Else
                    MyMsg "Sorry, can only View PRG, SEQ or USR files!"
            End Select
            Exit For
        End If
    Next T
    
End Sub

'---- XCABLE: Change to specified partition "file" on 1581 drive
Private Sub XChangePart(ByVal Filename As String)
    Dim Tmp As String
    
    DoCommand CBMCtrl, _
            CMDSTR & DriveNum & " " & Quoted("/0:" & UCase(Filename)), "Changing partition"
            
    Tmp = GetXStatus()
    If Left(Tmp, 2) = "77" Then MyMsg "Partition is Illegal!"
    
End Sub

'---- XCABLE: Read Directory
Private Sub cmdXRefresh_Click()
    RefreshX
End Sub

'---- XCABLE: Refresh Directory
Private Sub RefreshX()
    KillTemp                                                'Delete any temp files
    ClearXDir                                               'Clear the directory listing
    DriveNum = Val(cboXDevNum.List(cboXDevNum.ListIndex))   'Set the drive number (Seems to get confused occasionally)
    GetXDir                                                 'Read the Directory
    RefreshList 2
End Sub

'---- XCABLE: Find out what drives are connected and return string for selected device#
Private Sub GetXDevices()

    DetectXDrives False                                      'Get all drives

End Sub

'---- XCABLE: Read the X-Cable Directory, parse it, and fill the file list
Public Sub GetXDir()
    Dim CmdLine As String, temp As String, Temp2 As String, Results As ReturnStringType, J As Integer
    
    On Local Error GoTo GetXErr
    

    lstImageFiles(2).Clear
       
    Label(19).Visible = False: cmdXPart.Visible = False:  cmdXRoot.Visible = False              'Hide 1581 controls
    ClearXDir 'DiskName(2) = "": DiskID(2) = "":  picDiskID(2).ToolTipText = ""                            'Clear old fields
    
    frmWaiting.SetMode ""                                                                       'No progress bar
    
    Results = DoCommand(CBMCtrl, "--raw dir " & DriveNum, "Reading directory, please wait.", False) 'Run the program
    
    Close #1                                                'Make sure File#1 is closed so it can be opened below
    DoEvents                                                '(seems it sometimes doesn't close properly below)
    
    If Exists(TEMPFILE1) = False Then Exit Sub
    
    Open TEMPFILE1 For Input As #1                          'Read in the complete output file
    If EOF(1) Then Close #1: Exit Sub                       'Check for empty file
    
    Do
        Line Input #1, temp                                 'First line is dir. name and ID
    Loop While Left(temp, 7) = "GetProc"                    'Filter out occasional wayward status messages
    
    DiskName(2) = ExtractQuotes(temp)                       'Set the Disk Name
    DiskID(2) = Right$(temp, 5)                             'Set the Disk ID
    picDiskID(2).ToolTipText = DriveModel(temp)             'DriveModel tells you DOS version and Disk Format 2A=1541, 3D=1581
    
    If UCase(Right(temp, 2)) = "3D" Then
        cmdXPart.Visible = True: cmdXRoot.Visible = True    'Enable 1581 buttons for partitions
        Label(19).Visible = True                            'Enable Label above buttons
    End If

    If (Not EOF(1)) Then Line Input #1, temp
    
    While (Not EOF(1))
        If Temp2 <> "" Then lstImageFiles(2).AddItem Temp2  'Add the directory line
        Temp2 = temp                                        'Remember it
        Line Input #1, temp                                 'Get the next line
    Wend
    Close #1
       
    lblXLastStatus.Caption = UCase(temp)                    'The drive status is taken from the last line on stdout
    
    J = lstImageFiles(2).ListCount
    DFBlocksFree(2).Caption = MyTrim(Temp2) & " " & Format(J) & " files."                         'Blocks free from second last line
        
    On Local Error GoTo 0

    RefreshList 2                                           'Render CBM Font listing
    Exit Sub
    
GetXErr:
    If (Err.Number = 53) Or (Err.Number = 55) Then Exit Sub
    
    MyMsg "GetX Error: " & Err.Number & Cr & "[" & temp & "]"
    ClearXDir
    
    Exit Sub
    
End Sub

'---- XCABLE: Rename File(s)
Private Sub cmdXRename_Click()
   Dim T As Integer, Filename As String, NewFilename As String

    For T = 0 To lstImageFiles(2).ListCount - 1
        If (lstImageFiles(2).Selected(T)) Then
            Filename = ExtractQuotes(lstImageFiles(2).List(T))
            
            frmPrompt.Reply.Text = ExtractQuotes(lstImageFiles(2).List(T))
            frmPrompt.Ask "Rename CBM File", "Enter new name for '" & Filename & "'", 1, False
            
            NewFilename = UCase(Response)
            
            If Response <> "" Then
                DoCommand CBMCtrl, _
                          CMDSTR & DriveNum & " " & Quoted("R0:" & NewFilename & "=" & UCase(Filename)), _
                          "Renaming"
                GetShowXError
            Else
                Exit Sub
            End If
        End If
    Next T
    
   GetXDir
End Sub

'---- XCABLE: Reset Bus
Private Sub cmdXReset_Click()

     DoCommand CBMCtrl, _
            "reset", "Resetting BUS, please wait."

End Sub

'---- XCABLE: Delete (Scratch) a file
Private Sub cmdXScratch_Click()
    Dim T As Integer, Filename As String, FSel As Integer, OneName As String
    
    '-- Count how many files are selected
    For T = 0 To lstImageFiles(2).ListCount - 1
        If (lstImageFiles(2).Selected(T)) = True Then FSel = FSel + 1: Filename = Quoted(ExtractQuotes(lstImageFiles(2).List(T)))
    Next T
    If FSel = 0 Then Exit Sub
    
    '-- Prompt
    If FSel > 1 Then Filename = Str(FSel) & " file(s)"      'If more than one file show # of files rather than single filename
    If MsgBox("Are you sure you want to delete " & Filename & "?", vbYesNo, "Confirm delete") <> vbYes Then Exit Sub
    
    '-- Delete selected files

    For T = 0 To lstImageFiles(2).ListCount - 1
        If (lstImageFiles(2).Selected(T)) Then
            Filename = ExtractQuotes(lstImageFiles(2).List(T))   'do not use the ,p or ,s - works without them. adding ,p might delete p.prg?
            
            ' Send the scratch command. NOTE: filename must be converted to uppercase.
            DoCommand CBMCtrl, _
                     CMDSTR & DriveNum & " " & Quoted("S0:" & UCase(Filename)), _
                     "Scratching " & Filename
        End If
    Next T
    
    GetShowXError                                                          'Display any error
    GetXDir
End Sub

'---- XCABLE: Return to ROOT of 1581 Disk directory
Private Sub cmdXRoot_Click()
     Dim Tmp As String
     
     DoCommand CBMCtrl, _
               CMDSTR & DriveNum & " " & Quoted("/"), _
               "Selecting Root Partition, please wait..."
            
     Tmp = GetXStatus()
     If Left(Tmp, 2) = "77" Then MyMsg "Could not select partition!"
     
     GetXDir
End Sub

'---- XCABLE: Select Partition
Private Sub cmdXPart_Click()
    XView
End Sub

'---- XCABLE: Select a Drive
Private Sub cboXDevNum_Click()
    Dim Tmp As String, V As Integer
    
    Tmp = cboXDevNum.List(cboXDevNum.ListIndex)                                 'Get the selected string
    If Left(Tmp, 1) = "(" Then DetectXDrives False: Exit Sub                      'User selected to Re-Scan the devices. Exit
    
    lblDName.Caption = MyTrim(Mid(Tmp, 5))                                      'Set the Model# info and options box
        
    V = Val(Tmp)                                                                'Get the Device#
    
    If V <> DriveNum Then
        DriveNum = V                                                            'Use it
        ClearXDir                                                               'Clear the directory
        If UseFirstDrive = True Then GetXDir
    End If
    
End Sub

'---- XCABLE: Select ALL or NONE for all files in X-cable or Zoomfloppy disk
' Sel: TRUE=All files, FALSE=
Private Sub Selector(ByVal Sel As Boolean)
    Dim J As Integer
    
    For J = 0 To lstImageFiles(2).ListCount - 1
      lstImageFiles(2).Selected(J) = Sel                                         'Set to desired state
    Next J
End Sub

'---- XCABLE: Do Device Menu
' Menu: Initialize, Validate, Format, Change Device#, Double-sided, Detect Drives
Private Sub DoDeviceMenu(ByVal Index As Integer)
    Dim Model As String, Choice As Integer, Tmp As String, TCmd As String, TMsg As String
    Dim Status As ReturnStringType

    If Index = 7 Then DetectXDrives False: Exit Sub                      'Re-Sscan - Always Allowed.
    If DriveNum = 0 Then MsgBox "No Device Selected!": Exit Sub         'All other option require a valid device
    
    Model = UCase(Left(lblDName.Caption, 4))                            'Check drive model#
    
    Select Case Index
        Case 1: InitXDrive                                              'Reset/Inintialize
        Case 2: ValidateXDrive                                          'Validate
        Case 3: FormatXDrive                                            'Format
        Case 4: DeviceXDrive Model                                      'Change Dev#
            
        Case 5, 6                                                       'Set Single/Double sided
            Select Case Model
                Case "1571", "8250", "SFD-": AskSidedMode Model, Index - 4
                Case Else: MsgBox "Sorry, this drive is not changable"
            End Select
            
        Case 7: DetectXDrives False                                      'Re-scan
    End Select
    
End Sub

'---- XCABLE: Set Device#
' Device# memory locations (2 bytes starting with):
' 4040,8x50,D90x0               = $0C (12)
' 2031,1540/1541/1570/1571/1581 = $77 (119)
Private Sub DeviceXDrive(ByVal Model As String)
    Dim TCmd As String, TMsg As String
    Dim LB As String, HB As String                                          'LOW and HIGH byte of memory location in drive
    Dim Tmp As String, V As Integer                                         'Input of new device#
    Dim ND1 As String, ND2 As String                                        'New device# bytes

    TCmd = CMDSTR & DriveNum & " "
    TMsg = "Setting 1571 mode..."
    
    HB = "0"                                                                'Known addressa are all in zero page
    
    Select Case Model
        Case "2031", "1540", "1581"
            LB = "119"                                                      'New-style drives
        Case "2040", "4040", "8050", "8250", "D9060", "D909"
            LB = "12"                                                       'Old-style IEEE drives
        Case Else
            MsgBox "Sorry, I do not recognize this drive model.": Exit Sub
    End Select
    
    Do
        Tmp = InputBox("Changing device " & Str(DriveNum) & Cr & "Enter new Device# from 8 to 30", "Change Device#")
        If Tmp = "" Then Exit Sub
    
        V = Val(Tmp): If V > 7 And V < 31 Then Exit Do
        If V <> DriveNum Then Exit Do
        MsgBox "Enter a different device# from 8 to 30"
    Loop
    
    ND1 = Format(V + 32)
    ND2 = Format(V + 64)
    
    '-- CBMCTRL requires ascii for all parameters and will convert to bytes
    'Fmt:  m-w     lo         hi    2     #+32        #+64
    Tmp = "M-W " & LB & " " & HB & " 2 " & ND1 & " " & ND2
    
    DoCommand CBMCtrl, TCmd & Tmp, TMsg                                     'Change the device#
    DetectXDrives False                                                      'Re-read devices
    
End Sub

'---- XCABLE: Ask user to set Single or Double-sided Mode for Drive
' Model is the string for the model type
' Sides is 1 or 2
Private Sub AskSidedMode(ByVal Model As String, ByVal Sides As Integer)
    Dim Choice As Integer, Tmp As String, TCmd As String, TMsg As String
    Dim Status As ReturnStringType

    If Sides = 0 Then Tmp = "0": TMsg = "Single-Sided"
    If Sides = 1 Then Tmp = "1": TMsg = "Doublee-Sided"
    
    If MsgBox("Set Device#" & Str(DriveNum) & " (" & Model & ")" & Cr & "to " & TMsg & "?", vbOKCancel, "Select Mode") = vbCancel Then Exit Sub
    
    TCmd = CMDSTR & DriveNum & " "
    TMsg = "Setting 1571 mode..."

    Select Case UCase(Left(Model, 4))
        Case "1572"
            Status = DoCommand(CBMCtrl, TCmd & Quoted("U0>M" & Tmp), TMsg)
        Case "8250", "SFD-"
            Status = DoCommand(CBMCtrl, TCmd & Quoted("m-w 172 16 1 1"), TMsg)
            Status = DoCommand(CBMCtrl, TCmd & Quoted("m-w 195 16 1 0"), TMsg)
            Status = DoCommand(CBMCtrl, TCmd & Quoted("u9"), TMsg)
    End Select
    
End Sub


'========================================
'  Subs for Copy Operations
'========================================

'---- COPY: Copy "-->" LEFT to RIGHT; Figure out what type of copy
Private Sub cmdCopyRight_Click()
    If SrcMode = 0 Then
        '-- Source Files showing on left
        Select Case DstMode              'LEFT         RIGHT
            Case 0: Copy_LocalToX 0      'LocalPC0 --> X-Cable
            Case 1: Copy_LocalToLink 0   'LocalPC0 --> Link
            Case 2: Copy_LocalToLocal 0  'LocalPC0 --> LocalPC1
            Case 3: Copy_LocalToImg 0, 1 'LocalPC0 --> Image1
        End Select
    Else
        '-- Disk Image showing on left
        Select Case DstMode              'LEFT       RIGHT
            Case 0: Copy_ImgToX          'Image0 --> X Drive
            Case 1: Copy_ImgToLink       'Image0 --> Link
            Case 2: Copy_ImgToLocal 0, 1 'Image0 --> LocalPC1
            Case 3: Copy_ImgToImg 0, 1   'Image0 --> Image1
        End Select
    End If
End Sub

'---- COPY: Right-to-Left BUTTON - Figure out what type of copy
Private Sub cmdCopyLeft_Click()
    
    If SrcMode = 0 Then
        '-- Local PC 'Source' showing on left
        Select Case DstMode              'LEFT         RIGHT
            Case 0: Copy_XToLocal        'LocalPC0 <-- X-Cable
            Case 1: Copy_LinkToLocal 0   'LocalPC0 <-- Link
            Case 2: Copy_LocalToLocal 1  'LocalPC0 <-- LocalPC1
            Case 3: Copy_ImgToLocal 1, 0 'LocalPC0 <-- Image1
        End Select
    Else
        '-- Image File showing on left
        Select Case DstMode              'LEFT       RIGHT
            Case 0: Copy_XToImg          'Image0 <-- X-Cable
            Case 1: Copy_LinkToImg       'Image0 <-- Link
            Case 2: Copy_LocalToImg 1, 0 'Image0 <-- LocalPC1
            Case 3: Copy_ImgToImg 1, 0   'Image0 <-- Image1
        End Select
    End If

End Sub

'---- COPY: LocalPC to LocalPC
' DD=Direction: 0=left to right, 1=right to left
Private Sub Copy_LocalToLocal(ByVal DD As Integer)
    Dim T As Integer, FilesSelected As Integer, FileBase As String, Filename As String, Filename2 As String
    
    If LocalDir(0) = LocalDir(1) Then MyMsg "Can't copy when Directories are the same!": Exit Sub
    
    FilesSelected = 0

    For T = 0 To lstLocal(DD).ListCount - 1
        If (lstLocal(DD).Selected(T)) Then
            FilesSelected = FilesSelected + 1
            FileBase = lstLocal(DD).List(T)
            Filename = LocalDir(DD) & FileBase
            Filename2 = LocalDir(1 - DD) & FileBase
            If Overwrite(Filename2) = True Then
                FileCopy Filename, Filename2
                If Exists(Filename2) = False Then MyMsg "Sorry, '" & FileBase & "' could not be copied!"
            End If
        End If
    Next T
    
    If FilesSelected > 0 Then lstLocal(1 - DD).Refresh
    
End Sub

'---- COPY: Image to X-Cable
Private Sub Copy_ImgToX()
    Dim T As Integer, Filename As String, FilenameOut As String
    Dim Ext As String, Ext2 As String, CExt As String
    Dim SeqType As String, Tmp As String

    For T = 0 To lstImageFiles(0).ListCount - 1
        If lstImageFiles(0).Selected(T) = True Then
            Tmp = lstImageFiles(0).List(T)                      ' Get the Directory entry string
            Ext = DOSExt(Tmp)                                   ' Get the Extension, ie: PRG, SEQ, USR
            Ext2 = UCase(Left(Ext, 1))                          ' Get the file type, ie: P,S,U
            Filename = CBMName(Tmp)                             ' Get the filename, ie: FILENAME,P
            CExt = CBMExt(Tmp)                                  ' Get type, ie: P,S,U etc
            SeqType = " --file-type " & Ext2                    ' Build file-type string
            
            Select Case Ext2
                Case "P", "S", "U"
                
                    KillFile TEMPFILE3                          ' Erase temporary file
                    
                    DoCommand CBMC1541, _
                              DDFile(0) & " -read " & Quoted(Filename) & " " & Quoted(TEMPFILE3), _
                              "Copying '" & Filename & "' from image..."
                              
                    If Exists(TEMPFILE3) = True Then
                        DoCommand CBMCopy, _
                              "--transfer=" & TransferString & " -q -w " & DriveNum & " " & Quoted(TEMPFILE3) & _
                              " --output=" & Quoted(Filename) & SeqType, _
                              "Copying file to floppy disk as '" & Filename & "'..."
                              
                        GetShowXError
                    Else
                        Warning 2, TEMPFILE3 'Problem! The source file could not be extracted from the image
                    End If
                    
                Case Else
                    MsgBox "Sorry, the file '" & Filename & "' is a type that can not be copied."
                    
            End Select
                
        End If
    Next T

    GetXDir
    
End Sub

'---- COPY: Local to Image
' SrcList: 0=Left 1=Right , DstImage: 0=Left 1=Right
' Special handling for P00,S00,R00,U00 files
Private Sub Copy_LocalToImg(ByVal SrcList As Integer, ByVal DstImage As Integer)
    Dim T As Integer, Filename As String, Filename2 As String, Flag As Boolean
    Dim Base As String, Ext As String, Ext2 As String, Tmp As String, Max As Integer
    
    If DDFile(DstImage) = "" Then MyMsg "Please select the destination image file first!": Exit Sub
    Max = lstLocal(SrcList).ListCount: If Max = 0 Then MyMsg "There are no files in the source list!": Exit Sub
    Flag = False
       
    For T = 0 To Max - 1
        If lstLocal(SrcList).Selected(T) = True Then
            Tmp = lstLocal(SrcList).List(T)     'FILENAME.p00 - Source list entry
            Base = FileBase(Tmp)                'FILENAME     - Filename without extension
            Ext = FileExtU(Tmp)                 'P00          - DOS Extension only uppercase
            Ext2 = CBMExt(Ext)                    ',p           - CBM Extension
            Filename = LocalDir(SrcList) & Tmp  'D:\PATH\FILENAME.p00 -Source filename with path
            Filename2 = LCase(Base & Ext2)      'filename,p   - Filename with CBM extension

            'Copy file depending on EXTension
            Select Case Left(Ext, 2)
                Case "P0", "S0", "R0", "U0"
                    '-- Write P00-type file to image.
                    ' NOTE: Filename is included in source file, so no need to include it,
                    ' however if we don't then we will get P00 files even with S00 extension
                    DoCommand CBMC1541, _
                              DDFile(DstImage) & " -write " & Quoted(Filename), _
                              "Copying '" & Tmp & "' to image..."
                Case Else
                    If SupportedExt(Ext) = True Then
                        '-- Write all other types
                        DoCommand CBMC1541, _
                                DDFile(DstImage) & " -write " & Quoted(Filename) & " " & Quoted(Filename2), _
                                "Copying '" & Tmp & "' to image..."
                    Else
                        If Flag = False Then Flag = True: MyMsg "You can only write PRG,SEQ,ROM,BIN,'P00','S00' or files with NO extension INTO images!" & Cr & "Disk Images are not supported!" 'Warn Once!
                    End If
            End Select
        End If
    Next T

    ImageRefresh DstImage
    
End Sub

'---- COPY: X-Cable to Image
Private Sub Copy_XToImg()
    Dim Filename As String, FileIn As String, C As Integer
    Dim T As Integer
    
    C = 0
    For T = 0 To lstImageFiles(2).ListCount - 1
        If (lstImageFiles(2).Selected(T)) Then
            Filename = LCase(CBMName(lstImageFiles(2).List(T)))      'FILENAME,P (needed for source and destination)
            FileIn = Filename
            If Left(FileIn, 2) = "--" Then Mid(FileIn, 1, 1) = "?"   'Change "--" to "?-"

            KillFile TEMPFILE3                            'delete temp file first
            
            '-- Copy from X to TEMPFILE
            DoCommand CBMCopy, _
                      "--transfer=" & TransferString & " -q -r " & DriveNum & " " & Quoted(FileIn) & " --output=" & Quoted(TEMPFILE3), _
                      "Copying '" & Filename & "' from floppy disk."
                  
            '-- Copy TEMPFILE to Image
            DoCommand CBMC1541, _
                        DDFile(0) & " -write " & Quoted(TEMPFILE3) & " " & Quoted(Filename), _
                        "Copying '" & Filename & "' to image..."
            C = C + 1
        End If
    Next T
    
    If C = 0 Then
        MyMsg "You did not select any files to transfer into the IMAGE." & Cr & "(You can not store a Disk Image INSIDE another Disk Image!)"
    Else
        ImageRefresh 0
    End If
    
End Sub

'---- COPY: CBMLink to Disk Image
' Intermediate file is stored in EXE directory
Private Sub Copy_LinkToImg()
    Dim i As Integer, Tmp As String, Filename As String, Ext As String, Filename2 As String, FilenameOut As String
          
    For i = 0 To lstImageFiles(3).ListCount - 1
        If (lstImageFiles(3).Selected(i)) Then
            Tmp = lstImageFiles(3).List(i)
            Filename = UCase(ExtractQuotes(Tmp)): Ext = DOSExt(Tmp)     'FILENAME and PRG
            Filename2 = CBMName(Tmp)                                 'FILENAME,P
            FilenameOut = ExeDir & UCase(Filename & "." & Ext)          'EXEPATH\FILENAME.PRG
                                    
            '-- Read Link file. File is written to EXE directory.
            DoCommand CBMLink, LinkCStr & " -fr " & Quoted(Filename), _
                      "Copying '" & Filename & "' from remote floppy disk."
                      
            '-- Write File from EXE directory to Image File
            If Exists(FilenameOut) = True Then
                DoCommand CBMC1541, _
                        DDFile(0) & " -write " & Quoted(FilenameOut) & " " & Quoted(Filename2), _
                        "Copying '" & Filename & "' to image..."
                KillFile FilenameOut                                    'Delete local copy of file
            End If
        End If
    Next i
    
    ImageRefresh 0
End Sub

'---- COPY: Image to CBMLINK
'
Private Sub Copy_ImgToLink()
    MyMsg "Sorry, IMG to Link Not available!"
End Sub

'---- COPY: CBMLINK to LocalPC
'
Private Sub Copy_LinkToLocal(Index As Integer)
    Dim T As Integer, FilesSelected As Integer, Filename As String, FExt As String, FExt2 As String, FilenameOut As String
    Dim Tmp As String ', Response As ReturnStringType
    
    FilesSelected = 0
          
    For T = 0 To lstImageFiles(3).ListCount - 1
        If (lstImageFiles(3).Selected(T)) Then
            Tmp = lstImageFiles(3).List(T)                                           'Get list entry string
            Filename = UCase(ExtractQuotes(Tmp))                            'FILENAME
            FExt = DOSExt(Tmp)                                              'PRG       - Get CBM Filetype (extension)
            FilenameOut = LocalDir(Index) & UCase(Filename & "." & FExt)    'D:\PATH\FILENAME.PRG
            
            '-- Transfer file to EXE directory
            DoCommand CBMLink, _
                      LinkCStr & " -fr " & Quoted(Filename), _
                      "Copying '" & Filename & "' from floppy disk."
                      
            '-- Move the file to proper LocalPC destination
            If Exists(ExeDir & Filename) = True Then
                Name ExeDir & Filename As FilenameOut                       'Move the file
                lstLocal(Index).Refresh
            End If
            
            FilesSelected = FilesSelected + 1
        End If
    Next T
    
    '-- No Files were selected, make a Disk Image (D80) instead.
    If FilesSelected = 0 Then
        If ConfirmD64 = True Then
            If MsgBox("No files selected. Do you want to make an image of this floppy disk?", vbQuestion Or vbYesNo, "Create Disk Image") = vbNo Then Exit Sub
        End If
        
        frmPrompt.Reply.Text = RTrim(DiskName(Index)) & ".D80"
        frmPrompt.Ask "Create Disk Image", "Please Enter Image Filename:", 1, False
        If Response = "" Then Exit Sub
        
        '-- Read DISK Image file. File is written to EXE directory
        DoCommand CBMLink, _
                  LinkCStr & " -dr" & Format(LinkDrive) & " " & Response, _
                  "Creating disk image, please wait..."
                  
        '-- Copy the file to LocalPC destination folder
        If Exists(ExeDir & Response) = True Then
                Name ExeDir & Response As LocalDir(Index) & Response             'Move the file
                lstLocal(Index).Refresh
        End If
    End If

    lstLocal(Index).Refresh

End Sub

'---- COPY: LocalPC to CBMLink
'
Private Sub Copy_LocalToLink(Index As Integer)
    Dim T As Integer, FilesSelected As Integer, Filename As String, Ext As String
    
    FilesSelected = 0
      
    For T = 0 To lstLocal(Index).ListCount - 1
        If (lstLocal(Index).Selected(T)) Then
            FilesSelected = FilesSelected + 1
            Filename = lstLocal(Index).List(T)
            Ext = FileExtU(Filename)
            
            Select Case Ext
                Case "D64", "D71", "D40", "D80", "D81", "D82"   'Make Disk from Image
                    WriteImageToLink Filename, False
                    Exit Sub                                    'Exit so only 1 image is copied!
                Case Else                                       'Copy a File
                    TransferToLink LCase(Filename)
                    lstLocal(Index).Selected(T) = False         'de-select it
            End Select
        End If
    Next T

    If (FilesSelected > 0) Then GetXDir
End Sub

'---- COPY: Inside Image to LocalPC
' SrcImg: 0=left, 1=right Image
' DstPC : 0=left, 1=right LocalPC folder
' updated for P00 files - Mar 11/2016
Private Sub Copy_ImgToLocal(ByVal SrcImg As Integer, ByVal DstPC As Integer)
    Dim T As Integer, Filename As String, Filename2 As String, Filename3 As String
    Dim FilenameOut As String, Ext As String, Tmp As String
    
    If P00Flag = True Then MyChDir LocalDir(DstPC)                              'C1541.EXE writes P00 files to the Dst directory!!
    
    For T = 0 To lstImageFiles(SrcImg).ListCount - 1
        If lstImageFiles(SrcImg).Selected(T) = True Then
            Tmp = lstImageFiles(SrcImg).List(T)
            Filename = CBMName(Tmp)                                             ' FILENAME,P
            Filename2 = DOSName(Tmp)                                            ' FILENAME.PRG
            
            If P00Flag = True Then
                '-- Write P00 files to dest. P00 filename will be created automatically in CURRENT directory (hence the CD command above)
                '   BUG!: C1541.EXE always seems to write "p" files, even with SEQ source file, and then
                '         writing it back to an image looses the SEQ and instead creates a PRG file!
                DoCommand CBMC1541, _
                          DDFile(SrcImg) & " -p00save 1 -read " & Quoted(Filename), _
                          "Copying '" & Filename & "' from image..."
            Else
                FilenameOut = LocalDir(DstPC) & MakePCName(Filename2)                   'Make output filename. Edit DOS name if required
                
                '-- Write normal files to dest
                DoCommand CBMC1541, _
                          DDFile(SrcImg) & " -read " & Quoted(Filename) & " " & Quoted(FilenameOut), _
                          "Copying '" & Filename & "' from image..."
                    
                If Exists(FilenameOut) = False Then Warning 1, FilenameOut              'File was not copied
                
            End If

        End If
    Next T
    
    MyChDir ExeDir              'Change back to EXE directory
    lstLocal(DstPC).Refresh
    
End Sub

'---- COPY: Inside Image to Image
'
Private Sub Copy_ImgToImg(ByVal SrcImg As Integer, DstImg As Integer)
    Dim T As Integer, Filename As String
    
    If lstImageFiles(SrcImg).ListCount = 0 Then MyMsg "Image not loaded or has no entries!": Exit Sub
    If DDFile(0) = DDFile(1) Then MyMsg "Can't copy when the same Image is loaded on both sides!": Exit Sub
    
    For T = 0 To lstImageFiles(SrcImg).ListCount - 1
        If lstImageFiles(SrcImg).Selected(T) = True Then
            Filename = CBMName(lstImageFiles(SrcImg).List(T))
            
            KillFile TEMPFILE3  'use temp file as inbetween. Delete it to start
            
            '-- Copy file from Source Image
            DoCommand CBMC1541, _
                      DDFile(SrcImg) & " -read " & Quoted(Filename) & " " & Quoted(TEMPFILE3), _
                      "Copying '" & Filename & "' from image..."
            
            If Exists(TEMPFILE3) = True Then
                '-- File was copied, to temp dir, so now copy it to dest image
                DoCommand CBMC1541, _
                          DDFile(DstImg) & " -write " & Quoted(TEMPFILE3) & " " & Quoted(Filename), _
                          "Copying '" & Filename & "' to image..."
            Else
                Warning 2, TEMPFILE3 'The source file could not be extracted
            End If
        End If
    Next T
    
    KillFile TEMPFILE3
    ImageRefresh DstImg
    
End Sub

'---- COPY: X-Cable to LocalPC
'
Private Sub Copy_XToLocal()
    Dim Filename As String, FileIn As String, FileOut As String
    Dim Ext As String, CExt As String, Tmp As String
    Dim Filename2

    Dim T As Integer, FilesSelected As Integer, FilesCopied As Integer
    
    If DiskID(2) = "" Then GetXDir                                              'Added for batch
    
    FilesSelected = 0                                                           'Count of how many files selected to copy
    FilesCopied = 0                                                             'Count of actual copied files.
    
    '-- Figure out how many files are selected
    For T = 0 To lstImageFiles(2).ListCount - 1
        If (lstImageFiles(2).Selected(T)) Then FilesSelected = FilesSelected + 1       'It is selected
    Next T
    
    '-- Copy Selected, and count them
    If FilesSelected > 0 Then
        For T = 0 To lstImageFiles(2).ListCount - 1
            If (lstImageFiles(2).Selected(T)) Then
            
                Tmp = lstImageFiles(2).List(T)                                  'The directory Line
                Filename = LCase(ExtractQuotes(Tmp))                            'Get the filename, ie: filename
                Ext = LCase(DOSExt(Tmp))                                        'Get DOS extension, ie: SEQ
                CExt = "," & Left(Ext, 1)                                       'DOS extension, ie: ",s"
                FileIn = Filename & CExt                                        'Source File, ie: filename,s

                If Left(FileIn, 2) = "--" Then Mid(FileIn, 1, 1) = "?"          'Change "--" to "?-"

                FileOut = LocalDir(0) & MakePCName(Filename) & "." & Ext        'Output filename PATH\FILENAME.EXT
                
                Select Case UCase(Ext)
                    Case "PRG", "SEQ", "USR"
                        If Overwrite(FileOut) = True Then
                            DoCommand CBMCopy, _
                                  "--transfer=" & TransferString & " -q -r " & DriveNum & " " & Quoted(FileIn) & " --output=" & Quoted(FileOut), _
                                  "Copying '" & Filename & "' from floppy disk."
                                  
                            If Exists(FileOut) = False Then
                                Warning 5, Filename                             'File not copied
                            Else
                                FilesCopied = FilesCopied + 1                   'File copied ok - count it
                                lstImageFiles(2).Selected(T) = False            'De-select the file
                            End If
                        Else
                            FilesCopied = FilesCopied + 1                       'No overwrite - dont count towards files not copied
                        End If
                        
                    Case Else
                    
                        MyMsg "Sorry, '" & Filename & "' is a " & Ext & " file which is unsupported." & Cr & _
                                "It is recommended that you create a Disk Image."
                End Select
            End If
        Next T
    End If
    
    '-- Check results from looking for selected files and copting them
    
    If FilesCopied < FilesSelected Then
        If MsgBox("Some Files were not copied!" & Cr & _
        "Would you like to make an image of this disk?", vbYesNo, "Query") = vbYes Then FilesSelected = 0
    End If
    
    '-- If no files were selected, or some files were not copied then make an image
    
    If (FilesSelected = 0) Then
        If ConfirmD64 = True Then
            If MsgBox("No files selected.  Do you want to make an image of this floppy disk?", vbQuestion Or vbYesNo, "Create an Image") = vbNo Then Exit Sub
        End If
    
        If UseBatch = True Then
            frmBatch.Show
        Else
            MakeXDiskImage                                      'No Files were selected, so image the disk D64/G64/NIB etc.
        End If
    End If
    
    RefreshList 2                                               'Redraw source list as some files may now be deselected
    lstLocal(0).Refresh                                         'Refresh the Local directory
    
End Sub

'---- COPY: LocalPC to X-Cable
'
Private Sub Copy_LocalToX(ByVal Index As Integer)
    Dim T As Integer, ImgFlag As Boolean, C As Integer, DS As String
    Dim FilesSelected As Integer, ImagesSelected As Integer, Filename As String, Ext As String, FileOut As String
        
    FilesSelected = 0: ImagesSelected = 0: ImgFlag = False
    
    '-- Check files selected to determine operation(s)
    
    For T = 0 To lstLocal(Index).ListCount - 1
        If (lstLocal(Index).Selected(T)) Then
            Ext = FileExtU(lstLocal(Index).List(T))
            If SupportedImg(Ext, True) = True Then
                ImagesSelected = ImagesSelected + 1
            Else
                FilesSelected = FilesSelected + 1
            End If
        End If
    Next T
    
    '-- Check if mix of types is selected. Allow only one type
    
    If IgnoreD = False Then
        If (ImagesSelected > 0) And (FilesSelected > 0) Then MyMsg "Sorry, you have a mix of Images and files selected. Select only one type!": Exit Sub
    End If
    
    '-- Is it batch write mode?
    
    If (IgnoreD = False) And (ImagesSelected > 1) Then
        If MsgBox("Batch Image writing mode. Imaging will overwrite all contents on disk! " & Cr & "Please insert the FIRST disk then press OK to Start, or CANCEL to Abort!", vbOKCancel, "BATCH WRITE") = vbCancel Then Exit Sub
    End If
    
    '-- Process each selected file
    
    C = 0
    For T = 0 To lstLocal(Index).ListCount - 1
        If lstLocal(Index).Selected(T) = True Then
            Filename = lstLocal(Index).List(T)
            Ext = FileExtU(Filename)
            FileOut = LocalDir(0) & Filename
                               
            If (IgnoreD = False) And (SupportedImg(Ext, True) = True) Then
            
                '---- handle image files here
                If ImagesSelected = 1 Then
                    If MsgBox("Please insert FORMATTED DISK and click OK, or CANCEL to abort", vbOKCancel, "WRITE IMAGE") = vbCancel Then Exit For
                End If
            
                If ImgFlag = True Then
                    If MsgBox("Please insert NEXT DISK and click OK, or CANCEL to abort", vbOKCancel, "BATCH WRITE") = vbCancel Then Exit For
                End If
                
                ImgFlag = True
                
                WriteDFileToX Ext, FileOut, ImgFlag
                                
                'Select Case Ext
                '    Case "D64", "D71"   'Make Disk from D64 or D71
                '        If (UseNIB = True) And (WriteD64 = True) Then
                '            WriteNIBtoX FileOut, ImgFlag
                '        Else
                '            WriteImageToX FileOut, ImgFlag
                '        End If
                '    Case "NIB", "NBZ", "G64", "G71"
                '        WriteNIBtoX FileOut, ImgFlag
                '    Case "D80", "D81", "D82"
                '        WriteImageToX FileOut, ImgFlag
                'End Select
            
            Else
                
                '-- Handle regular files here
                
                TransferToX FileOut                         'Transfer the file. NOTE: Clears disk status
                C = C + 1
                
                '-- Check if there were errors (ie: write protect)
                
                If LastCMDError <> "" Then
                    If C < FilesSelected Then
                        If MsgBox("Error: " & LastCMDError & Cr & "Do you want to skip remaining files?", vbYesNo, "Continue") = vbYes Then Exit For
                    Else
                        MyMsg "Error: " & LastCMDError
                    End If
                End If
            End If
        End If
        lstLocal(Index).Selected(T) = False: DoEvents 'Deselect the file
    Next T

    If (FilesSelected > 0) Then GetXDir
    
End Sub

'---- XCABLE: Write a single supported Disk Image File to X
' EXT = Disk Extension, ImgFlag = Flag to ???????
Private Sub WriteDFileToX(ByVal Ext As String, FileOut As String, ImgFlag As Boolean)

                Select Case Ext
                    Case "D64", "D71"   'Make Disk from D64 or D71
                        If (UseNIB = True) And (WriteD64 = True) Then
                            WriteNIBtoX FileOut, ImgFlag
                        Else
                            WriteImageToX FileOut, ImgFlag
                        End If
                    Case "NIB", "NBZ", "G64", "G71"
                        WriteNIBtoX FileOut, ImgFlag
                    Case "D80", "D81", "D82"
                        WriteImageToX FileOut, ImgFlag
                End Select

End Sub

'---- XCABLE: Create a Disk Image from X-Cable Disk
'
Public Sub MakeXDiskImage()
    Dim Filename As String, Ext As String, FilenameOut As String, Ostr As String
    Dim Tmp As String, TmpP As VbMsgBoxResult, TmpExt As String, TempNIB As Boolean
    Dim X0 As String, X1 As String, X2 As String, NibTmp As String, OpFlag As Boolean
    
    X0 = ""
    X1 = "d64copy"
    X2 = "imgcopy"
    TempNIB = UseNIB
    
    '-- Check Disk Format using DiskID string
    Tmp = UCase(Mid(DiskID(2), 4, 2))
    Select Case Tmp
        Case "2A": TmpExt = "D64": X0 = X1
        Case "2C": TmpExt = "D80": X0 = X2: TempNIB = False         'Can't NIB D80
        Case "3D", "1D": TmpExt = "D81": X0 = X2: TempNIB = False   'Can't NIB D81
        Case Else
            '-- Handle Unknown Disk ID - Could be corrupt disk?
            
            TmpExt = "D64": X0 = X1     'default to D64 using D64COPY

            If (TempNIB = False) And (IgnoreBadID = False) Then
                TmpP = MsgBox("The source disk ID (" & Tmp & ") is unknown. This could be a corrupt disk, copy-protected disk, or unsupported format." & Cr & _
                "Do you want to try imaging with NIBTOOLS?" & Cr & "( Yes=NIBTOOLS, No=D64COPY, Cancel=Do Not Image )", vbYesNoCancel, "Warning!")
                Select Case TmpP
                    Case vbYes: TempNIB = True
                    Case vbCancel: Exit Sub
                End Select
            End If
    End Select
    
    '-- Prompt for NIB/IMGCOPY if Prompt option set
    If (TempNIB = True) And (NIBPrompt = True) Then
        TmpP = MsgBox("Do you want to use NIBTools", vbYesNoCancel, "User Option")
        If TmpP = vbNo Then TempNIB = False
        If TmpP = vbCancel Then Exit Sub
    End If
  
    '-- Determine if using normal copier or NIBTOOLS
    If (TempNIB = False) Then
        '-- Create image of disk using D64COPY.EXE or IMGCOPY
        If UseBatch = False Then
            frmPrompt.Reply.Text = RTrim(FixPCName(DiskName(2), "")) & "." & TmpExt
            frmPrompt.Ask "Create Dxx", "Please Enter Image Filename:", 1, False
            If Response = "" Then Exit Sub
            FilenameOut = LocalDir(0) & Response
        Else
            FilenameOut = LocalDir(0) & BatchFilename
        End If
        
        Ostr = "": Ext = FileExtU(Response): Tmp = X0
        
        Select Case Ext
            Case "D64": Ostr = ""                                       'Check for 1541 image
            Case "D71": Ostr = "-2 "                                    'Check for 1571 image and add option string
            Case "D80": Ostr = "-d8050 --error-map=never "              'Check for 8050 image
            Case "D81": Ostr = "-d1581 --error-map=never "              'Check for 1581 image and add option string
            Case "D82": Ostr = "-d1001 -2 --error-map=never "           'Check for 8250/SFD image and add option string
        End Select
        
        If Overwrite(FilenameOut) = True Then
            KillFile FilenameOut
            frmWaiting.SetMode Ext
            
            DoCommand CBMOpen & Tmp, _
                      Ostr & "--transfer=" & TransferString & " " & NoWarpString & " " & Format(DriveNum) & " " & Quoted(FilenameOut), _
                      "Creating " & Ext & " image, please wait."
        End If
        
    Else
        '-- Create image of disk using NIBREAD.EXE
        If UseBatch = False Then
            frmPrompt.Reply.Text = RTrim(DiskName(2))
            frmPrompt.Ask "Create NIB", "Please Enter Filename:  (Do NOT include an extension!)", 1, False
            If Response = "" Then Exit Sub
            Response = FileBase(Response)
        Else
            Response = FileBase(BatchFilename)
        End If

        Ext = ".nib": If UseNBZ = True Then Ext = ".nbz"
        
        Filename = LocalDir(0) & Response
        FilenameOut = Quoted(Filename & Ext)
        
        If Overwrite(Filename & Ext) = True Then
            KillFile Filename & Ext
            
            '-- We always need a NIB file to make G64 or D64
            frmWaiting.SetMode "nibread"
            NibTmp = NIBstr: If UseNibCustom = True Then NibTmp = frmOptions.txtNibRead.Text        'Std or Custom NIB options?
            
            DoCommand CBMNib & "nibread", "-D" & Format(DriveNum) & " " & NibTmp & " " & FilenameOut, _
                      "Creating " & Ext & " file, please wait." 'Fix: Add space after NibTmp!
                    
            If Exists(Filename & Ext) = True Then
                NibTmp = NIBstr: If UseNibCustom = True Then NibTmp = frmOptions.txtNibConv.Text    'Std or Custom NIB options?
                
                frmWaiting.SetMode "nibconv"
                '-- Convert NIB to G64
                If CreateG64 = True Then
                    OpFlag = True
                    If NIBPrompt = True Then If MsgBox("Do you want to convert to G64?", vbYesNo) = vbNo Then OpFlag = False
                    If OpFlag = True Then
                        DoCommand CBMNib & "nibconv", NibTmp & " " & FilenameOut & " " & Quoted(Filename & ".g64"), _
                              "Converting " & Ext & " to G64"
                    End If
                End If
                
                '-- Convert NIB to D64
                If CreateD64 = True Then
                    OpFlag = True
                    If NIBPrompt = True Then If MsgBox("Do you want to convert to D64?", vbYesNo) = vbNo Then OpFlag = False
                    If OpFlag = True Then
                        DoCommand CBMNib & "nibconv", NibTmp & " " & FilenameOut & " " & Quoted(Filename & ".d64"), _
                              "Converting " & Ext & " to D64"
                    End If
                End If
                
                If CreateNIB = False Then KillFile FilenameOut
            Else
                MsgBox "Problem! The file '" & FilenameOut & "' was not created!", vbExclamation, "Warning!"
            End If
        End If
    End If
    
End Sub

'---- COPY: Transfer specified file to X-Cable or Zoomfloppy disk
' Filename should have path included and be PRG or SEQ extension
Private Sub TransferToX(ByVal Filename As String)
    Dim FilenameOut As String, Ext As String, SeqType As String
    
    FilenameOut = FileNameOnly(Filename): Ext = FileExtU(FilenameOut)
    FilenameOut = FileBase(FilenameOut)
    SeqType = " --file-type " & Left(Ext, 1)
    
    DoCommand CBMCopy, _
              "--transfer=" & TransferString & " -q -w " & DriveNum & " " & Quoted(Filename) & _
              " --output=" & Quoted(FilenameOut) & SeqType, _
              "Copying '" & Filename & "' to floppy disk as '" & FilenameOut & "'"

    'NOTES: CBMCOPY will display any error messages in it's output, which clears the DISK STATUS

End Sub

'---- COPY: Transfer a file to CBMLink device
Private Sub TransferToLink(ByVal Filename As String)
    Dim FilenameOut As String, FPath As String

    FPath = FilePath(Filename)
    FilenameOut = FileNameOnly(Filename)
        
    MyChDir FPath
    
    DoCommand CBMLink, _
              LinkCStr & " -fw " & FilenameOut, _
              "Copying '" & Filename & "' via link as '" & UCase(FilenameOut) & "'"

End Sub

'---- COPY: Write Disk Image (D64, D71, D80 etc) to X-cable using D64copy or ImgCopy
'
Public Sub WriteImageToX(ByVal Filename As String, ByVal NoWarn As Boolean)
    Dim Ext As String, Opt As String, Tmp As String, Tmp2 As String
    
    Tmp = "d64copy": Tmp2 = "imgcopy"
    Opt = "": Ext = FileExtU(Filename)
    
    If NoWarn = False Then If MsgBox("This will overwrite ALL data on the floppy disk! Are you sure?", vbExclamation Or vbYesNo, "Write " & Ext & " to Disk") = vbNo Then Exit Sub
    
    '-- Select option string for specified image format
    Select Case Ext
        Case "D64": Opt = ""                        'Use D64COPY
        Case "D71": Opt = "-2 "                     'Use D64COPY
        Case "D80": Opt = "-d8050 ": Tmp = Tmp2     'Use IMGCOPY
        Case "D81": Opt = "-d1581 ": Tmp = Tmp2     'Use IMGCOPY
        Case "D82": Opt = "-d8250 ": Tmp = Tmp2     'Use IMGCOPY
    End Select
    
    frmWaiting.SetMode Ext
    
    Tmp = MakeUPath(0, Tmp)
    
    DoCommand Tmp, _
              Opt & "--transfer=" & TransferString & " " & NoWarpString & " " & Quoted(Filename) & " " & Format(DriveNum), _
              "Creating disk from " & Ext & " image, please wait."
            
    GetXDir
End Sub

'---- COPY: Write NIB File to X-Cable
'
Public Sub WriteNIBtoX(ByVal Filename As String, ByVal NoWarn As Boolean)
    Dim Ext As String, NibTmp As String
    
    If UseNIB = False Then MyMsg "You must select the NIB option to write this file to disk.": Exit Sub
    If NoWarn = False Then If MsgBox("This will overwrite ALL data on the floppy disk! Are you sure?", vbExclamation Or vbYesNo, "Write " & Ext & " to Disk") = vbNo Then Exit Sub
        
    Ext = FileExtU(Filename)
    
    frmWaiting.SetMode "nib"
    
    NibTmp = NIBstr: If UseNibCustom = True Then NibTmp = frmOptions.txtNibWrite.Text    'Std or Custom NIB options?
    DoCommand CBMNib & "nibwrite", _
              " -D" & Format(DriveNum) & " " & NibTmp & " " & Quoted(Filename), _
              "Creating disk from " & Ext & " image, please wait."
    
    GetXDir
End Sub

'---- COPY: Write Disk Image to CBMLink
'
Public Sub WriteImageToLink(d64file As String, ByVal NoWarn As Boolean)
    Dim Ext As String
    
    Ext = FileExtU(d64file)
    
    If NoWarn = False Then If MsgBox("This will overwrite ALL data on Destination unit!" & Cr & "(disk must already be formatted!)" & Cr & " Are you sure?", vbExclamation Or vbYesNo, "Write " & Right(Ext, 3) & " to Disk") = vbNo Then Exit Sub
    
    frmWaiting.SetMode CBMLink
    
    'Usage: CBMLINK.EXE -c serial 19200,com1 -d 8 -dw0 image.d80
    DoCommand CBMLink, LinkCStr & " -dw" & Format(LinkDrive) & " " & LocalDir(0) & d64file, _
              "Writing " & Ext & " image to drive, please wait..."
    
    GetXDir
End Sub

'========================================
' Subs For CBM-LINK
'========================================

'---- CBMLINK: Pick a new Drive# Unit#
Private Sub cboLinkDev_Click()
    SetLinkString
End Sub

'---- CBMLINK: Make Drive Selection string
Public Sub SetLinkString()
    Dim N As Integer
    
    N = cboLinkDev.ListIndex
    LinkUnit = Int(N / 2) + 8                   'Drive#
    LinkDrive = N Mod 2                         'Unit#
    
    ClearLinkDir
    LinkCStr = "-c " & frmOptions.txtConStr.Text & " -d " & LinkUnit
    
End Sub

'---- CBMLINK: Click to Get Directory
Private Sub cmdLinkDir_Click()
    GetLinkDir
End Sub

'---- CBMLINK: Read Directory
Private Sub GetLinkDir()
    Dim CmdLine As String, temp As String, Temp2 As String, Results As ReturnStringType
    
    On Local Error GoTo GetLinkErr:
    
    ClearLinkDir
    
    'Run the program
    Results = DoCommand(CBMLink, LinkCStr & " -dd $" & Format(LinkDrive) & ":*", "Reading directory, please wait...", False)
    
    Close #1
    Open TEMPFILE1 For Input As #1                          'Read in the complete output file
    If EOF(1) Then Exit Sub                                 'Check for empty file
    
    Line Input #1, temp                                     'First line is dir. name and ID
    DiskName(3) = UCase(ExtractQuotes(temp))                'Extract Diskname
    DiskID(3) = UCase(Right$(temp, 5))                      'Extract ID
    picDiskID(3).ToolTipText = DriveModel(temp)             'Drive/dos

    Line Input #1, temp                                     'Get next line
    
    While (Not EOF(1))
        lstImageFiles(3).AddItem temp                       'Add the file entry
        Line Input #1, temp                                 'Get next line
    Wend
    
    DFBlocksFree(3) = UCase(temp)                           'The Blocks free is taken from the last line
    Close #1
    
    RefreshList 3
    
    Exit Sub
    
GetLinkErr:
    If Not (Err.Number = 53) Then
        MyMsg "GetLink Error: " & Err.Number & Cr & "[" & temp & "]"
        ClearLinkDir
    End If
    Exit Sub
    
End Sub

'---- CBMLINK: Format Drive
Private Sub cmdLinkFormat_Click()
    Dim Status As ReturnStringType
    
    If MsgBox("This will erase ALL data on the floppy disk.  Are you sure?", vbExclamation Or vbYesNo, "Format Disk") = vbNo Then Exit Sub
    
    frmPrompt.Reply.Text = "new,id"                                            'Provide a default string
    frmPrompt.Ask "Format CBM Floppy", "Please Enter Diskname, ID", 1, False   'Get Disk name and ID
    If Response = "" Then Exit Sub                                          'Exit if null response

    Status = DoCommand(CBMLink, LinkCStr & " -dc N" & Format(LinkDrive) & ":" & UCase(Response), "Formatting floppy disk, please wait.")
    
    lblLinkLastStatus.Caption = UCase(Status.Output)
    Sleep 1000                                                              'Just so message is visible

End Sub

'---- CBMLINK: Initialize Drive
Private Sub cmdLinkInit_Click()

    DoCommand CBMLink, _
            LinkCStr & " -dc I" & Format(LinkDrive), "Initializing Drive..."
            
    GetLinkDir

End Sub

'---- CBMLINK: Validate Drive
Private Sub cmdLinkValidate_Click()

    DoCommand CBMLink, _
            LinkCStr & " -dc V" & Format(LinkDrive), "Validating Drive..."
            
    GetLinkDir
    
End Sub

'---- CBMLINK: Rename file(s)
Private Sub cmdLinkRename_Click()
   Dim T As Integer, Filename As String

    For T = 0 To lstImageFiles(3).ListCount - 1
        If (lstImageFiles(3).Selected(T)) Then
            Filename = ExtractQuotes(lstImageFiles(3).List(T))
            frmPrompt.Reply.Text = Filename
            frmPrompt.Ask "Rename CBM File", "Enter new name for '" & Filename & "'", 1, False
            
            If Response Then
                DoCommand CBMLink, _
                          LinkCStr & " -dc R" & Format(LinkDrive) & ":" & Response & "=" & Filename, _
                          "Renaming file"
            Else
                Exit Sub
            End If
        End If
    Next T
    
   GetLinkDir

End Sub

'---- CBMLINK: Reset CBM-Link Drive
Private Sub cmdLinkReset_Click()
    DoCommand CBMLink, _
              LinkCStr & " -dc UJ", _
              "Resetting drives, please wait."
End Sub

'---- CBMLINK: Scratch file(s)
Private Sub cmdLinkScratch_Click()
    Dim T As Integer, Filename As String
    
    For T = 0 To lstImageFiles(2).ListCount - 1
        If (lstImageFiles(2).Selected(T)) Then
            Filename = ExtractQuotes(lstImageFiles(2).List(T))
            DoCommand CBMLink, _
                      "LinkStr &  -dc S" & Format(LinkDrive) & ":" & Filename & Qu, _
                      "Scratching " & Filename
        End If
    Next T

    GetLinkDir

End Sub

'---- CBMLINK: Get Drive Status
Private Sub cmdLinkStatus_Click()
    Dim Results As ReturnStringType

    Results = DoCommand(CBMLink, LinkCStr & " -ds", "Reading drive status, please wait.")
    lblLinkLastStatus.Caption = Results.Output

End Sub


'========================================
' DISK IMAGE Subs
'========================================

'---- DISKIMAGE: Click to Refresh List
Private Sub cmdImageRefresh_Click(Index As Integer)
    ImageRefresh Index
End Sub

'---- DISKIMAGE: Refresh List
Private Sub ImageRefresh(Index As Integer)
    
    GetImageDir Index, DDFile(Index)
    RefreshList Index
    
End Sub

'---- DISKIMAGE: Clear Disk Image display
Private Sub ClearDD(ByVal Index As Integer)

    DDFile(Index) = ""                          'Clear filename
    lblDDFile(Index).Caption = ""               'Clear the filename field
    lblDDFile(Index).ToolTipText = ""           'Clear tooltip
    lstImageFiles(Index).Clear                  'Clear file list
    DiskName(Index) = ""
    DiskID(Index) = ""
    lblExt(Index).Caption = ""
    DFBlocksFree(Index).Caption = ""

End Sub

'---- DISKIMAGE: Handle automatic viewing of Disk Image when LocalPC entry is selected
' The LEFT disk image must be visible, and the RIGHT disk image must NOT
Private Sub lstLocal_Click(Index As Integer)
    Dim Filename As String, Ext As String, P As Integer
    
    If Layout = 1 Then
        P = lstLocal(Index).ListIndex
        Filename = LocalDir(Index) & lstLocal(Index).List(P)
        Ext = FileExtU(Filename)
        
        If SupportedImg(Ext, False) = True Then
            If Index = 0 And DstMode <> 3 Then SelectImage Filename, Index
        End If
    End If
    
    CalcBlocks (Index)
End Sub

'---- Handle viewing file when entry double-clicked
Private Sub lstLocal_DblClick(Index As Integer)
    If Index = 0 Then
        Call cmdSrcView_Click(Index)  'LEFT side
    Else
        Call cmdSrcView2_Click(Index) 'RIGHT side
    End If
End Sub

'---- GENERAL: Calculate Size of Selected file in BLOCKS
Private Sub CalcBlocks(ByVal Index As Integer)
    Dim T As Integer, Bytes As Double, Flag As Boolean
    
    Bytes = 0: Flag = False
    
    For T = 0 To lstLocal(Index).ListCount - 1
        If (lstLocal(Index).Selected(T)) Then
            If Exists(LocalDir(Index) & lstLocal(Index).List(T)) = False Then
                Flag = True                                         'File does not exist anymore.. Flag it
            Else
                Bytes = Bytes + FileLen(LocalDir(Index) & lstLocal(Index).List(T))  'Add it up
            End If
        End If
    Next T
    
    KBText(Index).Caption = Format(Bytes / 1024, "0.0")
    BlockText(Index).Caption = Format(Bytes / 254, "0") '254 Bytes per C= Block
                    
    If Flag = True Then lstLocal(Index).Refresh    'We found a missing file, so refresh the list
End Sub

'---- LOCAL: Refresh File List
Private Sub cmdLocalRefresh_Click(Index As Integer)
    lstLocal(Index).Refresh
End Sub

'---- LOCAL: Rename files on LocalPC
Private Sub cmdSrcRename_Click(Index As Integer)
    Dim T As Integer, Count As Integer, Flag As Boolean, CFlag As Integer, InFile As String, OutFile As String
    
    Count = 0: Flag = False
    
    For T = 0 To lstLocal(Index).ListCount - 1
        If (lstLocal(Index).Selected(T)) Then Count = Count + 1: If Count > 1 Then Flag = True: Exit For
    Next
        
    CFlag = 1: If Flag = True Then CFlag = 2 'Set 'Cancel All' button for multiple files
    
    For T = 0 To lstLocal(Index).ListCount - 1
        If (lstLocal(Index).Selected(T)) Then
            frmPrompt.Reply.Text = lstLocal(Index).List(T)
            frmPrompt.Ask "Rename File", "Enter new name for '" & lstLocal(Index).List(T) & "'", CFlag, False
            
            If Response = "***" Then Exit For 'Cancel All was selected
            
            If Response <> "" Then
                InFile = LocalDir(Index) & lstLocal(Index).List(T)
                OutFile = LocalDir(Index) & Response
                If InFile <> OutFile Then
                    If Exists(OutFile) = False Then
                        Name InFile As OutFile
                        If Exists(OutFile) = False Then MyMsg "Sorry, couldn't rename the file!"
                    Else
                        MyMsg "Can't rename " & InFile & Cr & "There is already a file called " & OutFile & " in the directory!"
                    End If
                End If
            End If
        End If
    Next T
    
    lstLocal(Index).Refresh
End Sub


'========================================
' VICE SUBS
'========================================
' Subs that deal with VICE EMULATION

'---- VICE: Run a PRG file with VICE
Private Sub cmdDRun_Click(Index As Integer)
    Dim T As Integer, Filename As String, Ext As String
    
    For T = 0 To lstImageFiles(Index).ListCount - 1
        If lstImageFiles(Index).Selected(T) = True Then
            Filename = LCase(ExtractQuotes(lstImageFiles(Index).List(T)))
            Ext = UCase(DOSExt(lstImageFiles(Index).List(T)))
            
            Select Case Ext
                Case "PRG", "RG<", "P00", "P01", "P02"  'Only adding P00-02...this could be a potential problem
                    RunVice frmOptions.cboPRG.ListIndex, DDFile(Index), Filename 'RUN with VICE!
                Case Else
                    MyMsg "Sorry, you can only run PRG or P00 files!"
            End Select
            Exit For
        End If
    Next T

End Sub


'---- VICE: Search File List for Image or File to Run via VICE
Private Sub cmdSrcRun_Click(Index As Integer)
    Dim T As Integer, Filename As String, Ext As String, Tmp As String
  
 
    For T = 0 To lstLocal(Index).ListCount - 1
        If (lstLocal(Index).Selected(T)) Then
            Filename = LocalDir(Index) & lstLocal(Index).List(T)
            Ext = FileExtU(lstLocal(Index).List(T))

            Select Case Ext
                Case "D64", "X64", "G64":   RunVice frmOptions.cbo64.ListIndex, Filename, ""
                Case "D71", "D81":          RunVice frmOptions.cbo71.ListIndex, Filename, ""
                Case "D80", "D82":          RunVice frmOptions.cbo80.ListIndex, Filename, ""
                Case "", "PRG":             RunVicePRG Filename
            End Select
            
            Exit Sub 'Stop so only the first file is executed
        End If
    Next T
End Sub

'---- VICE: Run a single PRG file in VICE
' Runs a file from local PC
Public Sub RunVicePRG(ByVal Filename As String)
    Dim J As Integer, LA As Long
    
    '---Check PRG option (specified or by load address)
    J = frmOptions.cboPRG.ListIndex                                             'Selected EMU for PRG files
    If J = 0 Then Exit Sub
    
    If frmOptions.OptPRGMode(1).value = True Then
        '-- Use program Load Address to select emulator.
        '   Note: VIC-20 and TED can have same load address. TODO: Allow selection when multiple choices
        LA = GetLoadAddress(Filename)
        J = GetMachine(LA)
        
        If J < 2 Then
            frmViceSelect.Show vbModal                                  'Waits here until Selection is made or Form is closed
            J = frmViceSelect.EmuNum                                    'User selection or 0=cancelled
        End If
    End If
    
    If J > 0 Then RunVice J, "", Filename                               'If not cancelled then Run Vice

End Sub

'---- VICE: Run VICE with specified Emulator, Disk Image and Filename
Public Sub RunVice(ByVal Emu As Integer, ByVal DName As String, FName As String)
    Dim Tmp As String, VPath As String
    
    If (UseVice = False) Or (Emu = 0) Then Exit Sub
    
    If Emu = 1 Then
        frmViceSelect.Show vbModal                                                  'Ask for emulator here
        Emu = frmViceSelect.EmuNum                                                  'Selected emulation
        If Emu = 0 Then Exit Sub                                                    'No selection
    End If
    
    VPath = CBMVICE & ViceEXE(Emu) & ".exe"                                         'Build path to selected VICE Executable
        
    If Exists(VPath) = False Then MyMsg "Vice Executable#" & Str(Emu) & " ('" & VPath & "') not found!": Exit Sub
    
    Tmp = UnQuoted(DName)
    If Tmp <> "" Then
        If FName <> "" Then Tmp = Tmp & ":" & FName
    Else
        Tmp = FName
    End If
      
    Tmp = VPath & " -autostart " & Quoted(Tmp)

    If PreviewCheck = True Then
        If MsgBox("Command:" & Cr & Cr & Tmp & Cr & Cr & "OK to continue?", vbYesNo) = vbNo Then Exit Sub
    End If

    Shell Tmp, vbNormalFocus
End Sub


'---- GENERAL: Search File List for selected File, figure out best way to view it
' This will open an image, the viewer, or the windows program associated with an unrecognized file type as appropriate

Private Sub CheckSelected(Index As Integer, Target As Integer)
  Dim T As Integer, V As Integer, FLen As Long
  Dim Filename As String, Ext As String, FilenameOut As String
  Dim FileB As String
  Dim NibTmp As String, ViewSel As Integer
  
  ViewSel = -1  'Nothing
  
    For T = 0 To lstLocal(Index).ListCount - 1
        If (lstLocal(Index).Selected(T)) Then
            Filename = LocalDir(Index) & lstLocal(Index).List(T)
            FilenameOut = FileBase(Filename) & ".d64"
                        
            Ext = FileExtU(lstLocal(Index).List(T))
            Select Case Ext
                Case "D64", "X64", "G64", "G71", "D71", "D80", "D81", "D82", "D2M", "D4M", "DNP"
                    SelectImage Filename, Target
                    
                Case "NIB", "NBZ"
                    If MsgBox("Do you want to convert this " & Ext & " to D64 to view the contents?", vbYesNo, "Convert?") = vbYes Then
                        frmWaiting.SetMode "nibconv"
                        NibTmp = NIBstr: If UseNibCustom = True Then NibTmp = frmOptions.txtNibConv.Text    'Std or Custom NIB options?
                        DoCommand CBMNib & "nibconv", NibTmp & " " & Quoted(Filename) & " " & Quoted(FilenameOut), _
                                  "Converting " & Filename & " to D64"
                        If Exists(FilenameOut) Then SelectImage FilenameOut, Target
                    End If
                    
                Case "", "PRG"
                    ViewSel = 0
                    
                Case "SEQ"
                    ViewSel = 1
                    
                Case "BIN", "ROM", "PRG", "ASM-PROJ"
                    ViewSel = 2                                             'Assume no PROJ file
                    FileB = FileBase(Filename)                              'Get BASE filename
                    
                    If Exists(FileB & ".asm-proj") = True Then ViewSel = 4  'Got an ASM-PROJ File!
                    
                    If Ext = "ASM-PROJ" Then                                'Got an ASM-PROJ, so find matching ROM or BIN
                        Filename = ""                                       'Assume no matching binary
                        Ext = "BIN": FilenameOut = FileB & ".bin"           'Check for .BIN
                        If Exists(FilenameOut) Then Filename = FilenameOut
                        Ext = "ROM": FilenameOut = FileB & ".rom"           'Check for .ROM
                        If Exists(FilenameOut) Then Filename = FilenameOut
                        Ext = "PRG": FilenameOut = FileB & ".prg"           'Check for .PRG
                        If Exists(FilenameOut) Then Filename = FilenameOut

                        If Filename <> "" Then ViewSel = 4                  'Use ASM tab
                    End If
                    
                Case "ART", "CDU", "KOA", "GEO", "P00", "S00"
                    ViewSel = 5
                    
                Case Else
                    V = MsgBox("Unknown file type. Open with associated WINDOWS app?" & Cr & "YES=Windows, NO=CBM-Transfer Viewer", vbYesNoCancel, "Unknown File type")
                    If V = vbYes Then ViewFile Filename: Exit Sub
                    ViewSel = 0
                    
            End Select
            
            If (ViewSel >= 0) And (Filename <> "") Then
                    frmViewer.Show: DoEvents                                'Make it open
                    While frmViewer.ViewerReady = False: Wend               'Wait for the Form to load!
                    If ViewSel = 4 Then
                        'We have an ASM-PROJ file so load it
                        FileB = FileBase(Filename) & ".asm-proj"
                        frmViewer.LoadProjFile FileB
                    End If

                    frmViewer.ViewIt ViewSel, Filename, Filename, Ext       'Open the Viewer with the best mode
            End If
            Exit Sub 'Stop so only the first file is executed
        End If
    Next T
End Sub

'---- Refresh the directory - only if automatic refresh hasn't been turned off
Private Sub RefreshXDir()
    If AutoRefreshDir Then
        GetXDir
    Else
        ClearXDir
    End If
End Sub

'---- Clear the XCable directory listing, because contents have changed
Private Sub ClearXDir()
    
    ClearDir 2                                                                      'Clear X Files Directory
    lblXLastStatus.Caption = ""                                                     'Clear Drive Status

End Sub

'---- Clear the CBMLink directory listing, because contents have changed
Private Sub ClearLinkDir()
    
    ClearDir 3
    lblLinkLastStatus.Caption = ""

End Sub

'---- Clear All Listing entries, Disk Name and ID, and picturebox
Private Sub ClearDir(ByVal Index As Integer)

    lstImageFiles(Index).Clear                                                      'Clear the File list
    
    picDir(Index).Cls                                                               'Clear the directory picture
    picDiskName(Index).Cls                                                          'Clear the Disk Name picture
    picDiskID(Index).Cls                                                            'Clear the Disk ID picture
    
    DiskName(Index) = ""                                                            'Clear the Disk Name
    DiskID(Index) = ""                                                              'Clear the Disk ID
    DFBlocksFree(Index).Caption = ""                                                'Clear Blocks free
    
End Sub


'---- Set File Filter for PC directory listing
Private Sub cboFilter_Click(Index As Integer)
    Dim Tmp As String
    Static CustFilt(1) As String
    
    If cboFilter(Index).ListIndex = 17 Then
        If CustFilt(Index) = "" Then CustFilt(Index) = "*.*"
        Tmp = InputBox("Enter filter:", "Custom Filter", CustFilt(Index))
        If Tmp <> "" Then
            CustFilt(Index) = Tmp: lstLocal(Index).Pattern = Tmp
            cboFilter(Index).List(17) = "CUSTOM [" & CustFilt(Index) & "]"
        End If
    Else
        lstLocal(Index).Pattern = FilterString(cboFilter(Index).ListIndex)
    End If
    
    lstLocal(Index).Refresh

End Sub

'---- Return Filter string for given Index
Private Function FilterString(ByVal N As Integer) As String
    Dim FX As String
    Select Case N
        Case 1: FX = "*.D64;*.D71;*.D80;*.D81;*.D82;*.NIB;*.G64;*.G71;*.X64;*.D1M;*.D2M;*.D4M"
        Case 2: FX = "*.NIB;*.NBZ;*.G64;*.G71;*.D64"
        Case 3: FX = "*.D80;*.D82"
        Case 4: FX = "*.PRG"
        Case 5: FX = "*.SEQ"
        Case 6: FX = "*.TXT"
        Case 7: FX = "*.D64"
        Case 8: FX = "*.D71"
        Case 9: FX = "*.D80"
        Case 10: FX = "*.D81"
        Case 11: FX = "*.D82"
        Case 12: FX = "*.G64;*.G71"
        Case 13: FX = "*.NIB;*.NBZ"
        Case 14: FX = "*.D1M;*.D2M;*.D4M"
        Case 15: FX = "*.BIN;*.ROM;*.ASM-PROJ"
        Case 16: FX = "*.ART;*.CDU;*.GEO;*.KOA"
        Case Else: FX = "*.*"
    End Select

    FilterString = FX

End Function

'---- DISKIMAGE: Select and View Disk Image File
Private Sub SelectImage(ByVal Filename As String, Index As Integer)
    
    DDFile(Index) = Quoted(Filename)                    'Remember the Filename
    If Index = 0 Then SrcMode = 1: SetSrcFrame          'Change LEFT view only if VIEW button
    If Index = 1 Then DstMode = 3: SetDstFrame          'Change RIGHT view
    lblDDFile(Index).Caption = Filename                 'Set the filename field
    lblDDFile(Index).ToolTipText = DDFile(Index)        'Set tooltip
    GetImageDir Index, DDFile(Index)
    
    RefreshList Index
    
End Sub

'---- DISKIMAGE: Read Directory
' This calls the C1541 program to read the directory of the DISK IMAGE file
' Output of the C1541 program is saved to TEMPFILE1.
' It then parses the output file and loads it into the LIST
'
Private Sub GetImageDir(Index As Integer, ByVal Filename As String)
    Dim temp As String, Temp2 As String, Results As ReturnStringType
    Dim P As Integer, PP As Integer, Terminator As String, J As Integer

    On Local Error GoTo GIError
             
    Results = DoCommand(CBMC1541, Quoted(Filename) & " -list", "", False) 'Run the program
    If Exists(TEMPFILE1) = False Then Exit Sub
    
    lstImageFiles(Index).Clear
    
    Close 1
    Open TEMPFILE1 For Input As #1
    
    If EOF(1) Then Exit Sub     'Check for empty file
    
    'NOTE: Early versions of C1541 used only the LF terminator which means that the entire file
    '      would load in all at once. Newer versions use CRLF and must be loaded in line-by-line.
    While Not EOF(1)
        Input #1, Temp2              'Output is in one long string. Must Parse!...
        temp = temp & Temp2 & LF
    Wend
    Close #1
    
    Terminator = LF
        
    '-- Throw away extraneous strings containing "GetProc" etc
    PP = 1
    Do
        P = InStr(PP, temp, Terminator): If P = 0 Then Exit Do
        Temp2 = Mid(temp, PP, P - PP): PP = P + 1
    Loop While Left(Temp2, 1) > "9"
       
    '-- Get the Disk Name and Disk ID
    DiskName(Index) = ExtractQuotes(Temp2)
    DiskID(Index) = Right$(Temp2, 5)
    
    picDiskID(Index).ToolTipText = DriveModel(Temp2)
    lblExt(Index).Caption = FileExtU(Filename)
    
    '-- Now parse remaining entries
    Do
        P = InStr(PP, temp, Terminator): If P = 0 Then Exit Do
        Temp2 = Mid(temp, PP, P - PP): PP = P + 1
        If InStr(1, Temp2, "blocks free", vbTextCompare) = 0 Then lstImageFiles(Index).AddItem Temp2 Else Exit Do 'Lowercase
    Loop
    
    J = lstImageFiles(Index).ListCount
    DFBlocksFree(Index).Caption = MyTrim(Temp2) & " " & Format(J) & " files"
    Exit Sub
    
GIError:
    If Not (Err.Number = 53) Then
        MyMsg "GetImage Error: " & Err.Number & Cr & "[" & temp & "]"
    End If
    Exit Sub

End Sub

'---- GENERAL: Prompt to Create a New Folder, then Make it if it doesn't already exist
Private Function NewFolder(ByVal RootPath As String) As String
    Dim DirName As String
    
    frmPrompt.Ask "Make Directory", "Enter Directory Name:", 1
    If Response = "" Then Exit Function                     'Check for null string
    
    DirName = RootPath & Response                           'Make path+filename
    NewFolder = ""                                          'Assume failure
    
    If DirExists(DirName) = False Then
        MkDir DirName                                       'Try to create folder
        If DirExists(DirName) = True Then
            NewFolder = DirName                             'It worked so return full path
        Else
            MsgBox "Couldn't creat the new folder! Check your directory permissions."   'Error!
        End If
    Else
        MyMsg "Can't create! There is a already a Directory called '" & DirName & "'!"  'Already exists
    End If

End Function

'---- GENERAL: Load Source Path Drop-down list (Path History)
Public Sub LoadHistory()
    Dim FIO As Integer, Tmp As String, LastTmp As String
        
    If Exists(HistoryFile) = False Then Exit Sub
    
    FIO = FreeFile
    Open HistoryFile For Input As FIO
    
    txtLocalDir(0).Clear                                                'Clear both history lists
    txtLocalDir(1).Clear
    
    LastTmp = ""                                                        'Remembers the previous entry to allow duplicate removal
    
    While Not EOF(FIO)
        Line Input #FIO, Tmp                                            'Read the path string
        If Tmp <> LastTmp Then                                          'Entry is NOT a duplicate. so add it
            txtLocalDir(0).AddItem Tmp
            txtLocalDir(1).AddItem Tmp
            LastTmp = Tmp                                               'Remember it for next path
        End If
    Wend
    
    Close FIO
    
End Sub

'---- GENERAL: Save Source Path Drop-down list (Path History)
Public Sub SaveHistory()
    Dim FIO As Integer, Tmp As String, a As Integer
      
    KillFile HistoryFile
    
    FIO = FreeFile
    Open HistoryFile For Output As FIO
    
    For a = 0 To txtLocalDir(0).ListCount - 1
        Print #FIO, txtLocalDir(0).List(a)
        'Debug.Print "Path>>>" & txtLocalDir(0).List(a)
    Next
    Close FIO
    
End Sub

'========================================
' DRAG and DROP functions
'========================================

'---- DRAGANDDROP: Change to Dropped Directory Path
Private Sub txtLocalDir_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Tmp As String
    
    If Data.GetFormat(vbCFFiles) Then
        Dim vFn As Variant
        For Each vFn In Data.Files
            Tmp = PathOnly(vFn): If Tmp <> "" Then SetLocalPath Index, Tmp             'Get path and use it if valid
        Next
    End If

End Sub

'---- DRAGANDDROP: Accept Dropped Image files
Private Sub picDir_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ext As String, Filename As String, N As Integer, C As Integer
    
    N = 0: C = 0
      
    If Data.GetFormat(vbCFFiles) Then
        Dim vFn As Variant
        
        For Each vFn In Data.Files
            N = N + 1
            Filename = (vFn)                                                'vFn is name of file dropped
            Ext = FileExtU(Filename)                                        'Get the Extension
            
            Select Case Index                                               '-- Check the destination type
                Case 0, 1                                                   'Disk Image
                    If SupportedImg(Ext, False) = True Then                 'Is it a supported Image for reading?
                        SelectImage Filename, Index                         'Yes
                        C = C + 1
                        Exit For                                            'Only process one
                    End If
                    
                    If SupportedExt(Ext) = True Then
                        C = C + 1
                        DoCommand CBMC1541, _
                          DDFile(Index) & " -write " & Quoted(Filename), _
                          "Copying " & FileNameOnly(Filename) & "' to image..."
                    End If
                    
                Case 2 'X-Cable
                    If SupportedExt(Ext) = True Then
                        C = C + 1
                        TransferToX Filename         'vFn is name of file dropped
                    End If
                    
                    If SupportedImg(Ext, True) = True Then
                        If MsgBox("Confirm Write of Disk Image: " & FileNameOnly(Filename) & Cr & "to X-Cable Drive", vbYesNo, "Continue") = vbYes Then
                            C = C + 1
                            WriteDFileToX Ext, Filename, True
                            Exit For
                        End If
                    End If

                
                Case 3 'CBM-Link
                    MsgBox "Sorry, Drag and Drop not supported for CBM-Link!": Exit For
            End Select
        Next vFn
    End If
    
    If Index < 3 Then
        If C < N Then MyMsg "Copied " & Str(C) & " of " & Str(N) & " files." & Cr & "Some files were not copied. Only Disk Image files, or Files with Supported CBM extensions can be dropped!"
    End If
    
    '--- Get updated content
    Select Case Index
        Case 0, 1:      GetImageDir Index, DDFile(Index)
        Case 2:         RefreshX
    End Select
    
    RefreshList Index
    
End Sub

'---- DRAGANDDROP: Provide drag and drop feedback to source
Private Sub picDir_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

    '0=do not allow drop, 1=inform source that data will be copied
    If Data.GetFormat(vbCFFiles) Then Effect = 1 Else Effect = 0
    
End Sub


'=============================================
' SHELL COMMANDS
'=============================================

'---- SHELL: Do Command
'
' This function checks if there is another command in progress and if so exits immediately.
' Check to make sure the DOS command EXE exists and if not displays message then exits.
' Checks of PREVIEW is enable and if so displays the command string
' Opens a WAITING status window if needed.
' Builds a Command Line string with all arguments. paths, and redirection
' Logs the command if enabled.
' Calls the SHELLWAIT routine to actually execute the command
' Reads the results
' Sets LastCMDError variable that can be checked by calling routine

' **** This function must be private, because of the return type.
'
Private Function DoCommand(Action As String, Args As String, WaitMessage As String, Optional DeleteOutFile As Boolean = True) As ReturnStringType
    Dim CmdLine As String, CmdLine2 As String, ErrorString As String, FIO As Integer
    Static InProgress As Boolean, Tmp As String
    
    If (InProgress) Then Exit Function
        
    Close
    If UCase(Right(Action, 4)) <> ".EXE" Then Action = Action & ".exe"
    
    If Exists(Action) = False Then
        MyMsg "The UTILITY: " & FileBase(Action) & Cr & _
        " was not found! Please copy it to the CBM-Transfer directory (legacy)," & Cr & _
        "or specify the correct path in the Config! (recommended)"
        Exit Function
    End If
      
    If PreviewCheck = True Then
        If MsgBox("Requested command:" & Cr & Cr & Action & " " & Args & Cr & Cr & "OK to continue?", vbYesNo) = vbNo Then Exit Function
    End If
    
    KillTemp 'And delete both temp files, so we're not cluttering things up
    
    '-- Flag that the background process is starting.
    InProgress = True
    
    '-- Display Dialog if WaitMessage is specified
    If WaitMessage <> "" Then
        frmWaiting.SetMode ""
        frmWaiting.Show vbModeless, frmMain
        frmWaiting.lblMsg = WaitMessage
    End If
    
    '-- Build command-line string
    'cmd /c is needed in order to have a shell write to a file (long, complicated explanation)
    '1> redirects stdout to a file, and 2> redirects stderr to a file (Win2K/XP only)
    'All these quotes [chr$(34)] are needed to handle spaces.  So you get: cmd /c ""path\command" args "files""
    
    CmdLine = Qu & Action & Qu & " " & Args
    CmdLine2 = "cmd /c " & Qu & CmdLine & Qu & " 1>" & TEMPFILE1 & " 2>" & TEMPFILE2
    
    'Debug.Print "DoCmd: "; CmdLine
    
    '-- This is where the commands are run. We are stuck here until it completes....
    
    KillFlag = False                            'Rest Kill Flag
    ShellWait CmdLine2, vbHide                  '******************* RUN THE COMMAND!!!!!
    DoEvents
        
    '-- Read in the output file
    
    FIO = FreeFile
    Open TEMPFILE1 For Input As FIO
        If (Not EOF(FIO)) Then Line Input #FIO, DoCommand.Output
    Close FIO
    
    '-- Read in the error file
    
    FIO = FreeFile
    Open TEMPFILE2 For Input As FIO
        If Not EOF(FIO) Then ErrorString = Input$(LOF(FIO), FIO)
    Close FIO
        
    '-- Display Error message
    
    If ErrorString <> "" Then
        cmdResults.ToolTipText = ErrorString          'Set Last Result as Tooltip
    End If
    
    '-- Save Error Message so Calling routine can decide what to do
    
    LastCMDError = ErrorString                  'Remember the error results
    
    '-- Close Waiting window and allow other commands to execute
    
    InProgress = False: frmWaiting.Hide         'We are done, so clear InProgress Flag and hide the dialog
    
    '-- Log Commands
    
    If LogAll = True Then
        Tmp = "=== " & Date & ", " & Time & Cr & "CMD: " & CmdLine2 & Cr & "ERR: " & ErrorString & Cr & "OUT: " & DoCommand.Output & Cr
        LogIt Tmp
    End If
    
End Function

'---- SHELL: PubDoCommand
'
Public Function PubDoCommand(Action As String, Args As String, WaitMessage As String, Optional DeleteOutFile As Boolean = True) As String
    Dim Returns As ReturnStringType
    
    Returns = DoCommand(Action, Args, WaitMessage, DeleteOutFile)
    PubDoCommand = Returns.Output
End Function

'---- SHELL: Wait for process to finish
' Found on Experts Exchange website. Written by, and Thanks to, vinnyd79 for this function!
Private Function ShellWait(PathName, Optional WindowStyle As VbAppWinStyle = vbNormalFocus) As Double

    Dim hProcess As Long, RetVal As Long

    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(PathName, WindowStyle))
    Do
        GetExitCodeProcess hProcess, RetVal
        DoEvents: Sleep 100
    Loop While (RetVal = STILL_ACTIVE) And (KillFlag = False)

    KillFlag = False    'Reset Kill Flag
    
End Function

'---- SHELL: Log Commandline string
Public Sub LogIt(ByVal Tmp As String)
    Dim FIO As Integer
    
    FIO = FreeFile
    
    Open LogFile For Append As FIO
        Print #FIO, Tmp
    Close FIO
    
End Sub

'========================================
' GRAPHICAL/BITMAP ROUTINES
'========================================

'---- GRAPHICAL: Draw CBM Listing using PETSCII encoding and real CBM Fonts
' Must have CBM Font loaded into 'picCBM' PictureBox.
' - Each "line" (set) must contain 256 characters with each character exactly 8x8 pixels
' - There are 6 sets: petscii uppper,lower, screen upper,lower,asci superpet and ascii
'
' SrcList.. is a ListBox as the source
' DstList.. is a PictureBox to render to
' Index.... is the first entry to display at the top
' CSet..... is the Set to display  (0=Upper,1=Lower,2=ScreenUpper,3=ScreenLower,4=SuperPET,5=ASCII)
' ZoomX/Y.. are the Zoom factors IE: 2/2=square (40 col), 1/2=Tall (80-col) etc
'
Public Sub DrawCBM(ByRef SrcList As ListBox, ByRef picD As PictureBox, ByRef vScroll As VScrollBar, ByVal CSet As Integer, ByVal ZoomX As Integer, ByVal ZoomY As Integer)
    Static Busy As Boolean
    
    Dim SrcIndex As Integer, MaxIndex As Integer                                        'Source Index and Max Index
    Dim DstW As Integer, DstH As Integer
    Dim DstX As Integer, DstY As Integer                                                'Destination co-ordinates
    Dim ChrW As Integer, ChrH As Integer                                                'FONT Character Width and Height (normally 8x8)
    Dim SetY As Integer, SetR As Integer, SetC As Integer                               'FONT Set Y Top, Rows and Characters per set
    Dim FR   As Integer, FC   As Integer                                                'FONT Character Row, Col
    Dim FY   As Integer, FX   As Integer                                                'FONT Character X/Y Co-ordinate
    Dim ZW   As Integer, ZH   As Integer                                                'Rendered Character size with zoom factor
    Dim i As Integer, J As Integer, P As Integer, Ch As Integer, Tmp As String          'Work variables
    Dim Sel As Boolean                                                                  'Flag if list entry is selected
    
    SrcIndex = vScroll.value: MaxIndex = SrcList.ListCount - 1                          'Get Top of list and list size
    
    picD.BackColor = ThemeListBG                                                     'Set background colour
    
    If SrcIndex > MaxIndex Then Exit Sub                                                'Check if out of range
    If ZoomX = 0 Or ZoomY = 0 Then Exit Sub                                             'Shouldn't happen.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz
    
    Busy = True
    picD.Visible = False

    
    DstW = picD.Width: DstH = picD.ScaleHeight                                    'Destination dimensions
    ChrW = 8: ChrH = 8                                                                  'CBM font should be 8x8 - perhaps future support for CBM-II 8x14 fonts?
    ZW = ChrW * ZoomX: ZH = ChrH * ZoomY                                                'Size of character with Zoom factor
    DstX = 0: DstY = 0                                                                  'Start at Top left of picturebox
    SetR = 2: SetC = 128                                                                'Layout of CBM FONT Bitmap - 2x128 (full set)
        
    SetY = CSet * SetR * ChrH                                                           'Y positon of Top of Selected SET
    DstY = 0                                                                            'Start Destination Y co-ordinate
    
    Do
        Tmp = SrcList.List(SrcIndex)                                                    'Get a PETSCII character from the line
        Sel = SrcList.Selected(SrcIndex)                                                'Is it selected?
        DstX = 0                                                                        'Start of Line
        
        For i = 1 To Len(Tmp)
            Ch = Asc(Mid(Tmp, i, 1))                                                    'Get the CBM character value

            FR = Int(Ch / SetC)                                                         'Character ROW
            FC = Ch Mod SetC                                                            'Character COL
            FY = SetY + FR * ChrH                                                       'Character Y co-ordinate
            FX = FC * ChrW                                                              'Character X co-ordinate
            
            If Sel = True Then
                picD.PaintPicture picCBM.Image, DstX, DstY, ZW, ZH, FX, FY, ChrW, ChrH, vbNotSrcCopy   'Invert it!
            Else
                picD.PaintPicture picCBM.Image, DstX, DstY, ZW, ZH, FX, FY, ChrW, ChrH, vbSrcCopy              'Copy it!
            End If
                        
            DstX = DstX + ZW                                                            'Move to the next position
            If DstX > DstW Then Exit For                                                'Stop rendering if past right boundry
            
        Next i
        
        SrcIndex = SrcIndex + 1: If SrcIndex > MaxIndex Then Exit Do                    'Exit loop if Index is too high
        DstY = DstY + ZH: If DstY > DstH Then Exit Do                                   'Exit loop if gone beyond bottom of list

    Loop
    
    picD.Visible = True                                                              'Show the list
    
    DoEvents                                                                            'scrolling is jerky without this!
    
    Busy = False
End Sub

'---- GRAPHICAL: Write CBM text to any Picturebox
' Must have CBM Font loaded into 'picCBM' PictureBox.
' - Each "line" (set) must contain 128 characters with each character exactly 8x8 pixels
' - There are 6 sets: petscii uppper,lower, screen upper,lower,asci superpet and ascii
'
' Txt...... is the source string
' DstPic... is a PictureBox to render to
' CSet..... is the Set to display
' ZoomX/Y.. are the Zoom factors IE: 2/2=square (40 col), 1/2=Tall (80-col) etc

Public Sub WriteCBM(ByVal Txt As String, ByRef picD As PictureBox, ByVal CSet As Integer, ByVal ZoomX As Integer, ByVal ZoomY As Integer)
    Static Busy As Boolean
    
    Dim DstW As Integer, DstH As Integer
    Dim DstX As Integer, DstY As Integer                                                'Destination co-ordinates
    Dim ChrW As Integer, ChrH As Integer                                                'FONT Character Width and Height (normally 8x8)
    Dim SetY As Integer, SetR As Integer, SetC As Integer                               'FONT Set Y Top, Rows and Characters per set
    Dim FR   As Integer, FC   As Integer                                                'FONT Character Row, Col
    Dim FY   As Integer, FX   As Integer                                                'FONT Character X/Y Co-ordinate
    Dim ZW   As Integer, ZH   As Integer                                                'Rendered Character size with zoom factor
    Dim i As Integer, J As Integer, P As Integer, Ch As Integer                         'Work variables
    Dim Tmp As String                                                                   'Work variables
    Dim Sel As Boolean                                                                  'Flag if list entry is selected
       
    Busy = True
    picD.Visible = False
    
    picD.BackColor = ThemeListBG                                                      'Set background colour
    picD.ForeColor = ThemeListFG                                                      'Set foreground colour
    picD.Cls
    
    DstW = picD.Width: DstH = picD.ScaleHeight                                      'Destination dimensions
    ChrW = 8: ChrH = 8                                                                  'CBM font should be 8x8 - perhaps future support for CBM-II 8x14 fonts?
    ZW = ChrW * ZoomX: ZH = ChrH * ZoomY                                                'Size of character with Zoom factor
    DstX = 0: DstY = 0                                                                  'Start at Top left of picturebox
    SetR = 2: SetC = 128                                                                'Layout of CBM FONT Bitmap - 2x128 (full set)
        
    SetY = CSet * SetR * ChrH                                                           'Y positon of Top of Selected SET
    DstX = 0: DstY = 0                                                                  'Start Destination X,Y co-ordinates
    
    For i = 1 To Len(Txt)
        
        Ch = Asc(Mid(Txt, i, 1))                                                        'Get the CBM character value

        FR = Int(Ch / SetC)                                                             'Character ROW
        FC = Ch Mod SetC                                                                'Character COL
        FY = SetY + FR * ChrH                                                           'Character Y co-ordinate
        FX = FC * ChrW                                                                  'Character X co-ordinate
        
        picD.PaintPicture picCBM.Image, DstX, DstY, ZW, ZH, FX, FY, ChrW, ChrH, vbSrcCopy 'copy it
        
        DstX = DstX + ZW                                                                'Move to the next position
        If DstX > DstW Then Exit For                                                    'Stop rendering if past right boundry
    
    Next i
        
    picD.Visible = True
    DoEvents                                                                            'scrolling is jerky without this!
    Busy = False
    
End Sub

'---- GRAPHICAL: Refresh List Pictures
' Refreshes specified Directory List - Writes NAME and DISK ID and re-draws Directory and Re-Calculate scrollbar
Private Sub RefreshList(ByVal Index As Integer)

    WriteCBM DiskName(Index), picDiskName(Index), EncodeL(Index), 1, 2                              'Draw Disk Name to picturebox
    WriteCBM DiskID(Index), picDiskID(Index), EncodeL(Index), 1, 2                                  'Draw Disk ID to picturebox
    CalcScroll Index                                                                                'Calculate Scrollbar
    DrawCBM lstImageFiles(Index), picDir(Index), vsImgDir(Index), EncodeL(Index), 1, RenderY(Index) 'Draw the files list

End Sub

'---- GRAPHICAL: Get an Icon from Theme bitmap into a Picturebox
Public Sub GetIcon(ByRef Icon As PictureBox, ByVal X As Integer, ByVal Y As Integer)
    Dim W As Integer, H As Integer
    
    W = Icon.ScaleWidth                                                 'Get the width and height from the Icon
    H = Icon.ScaleHeight
    
    'Icon.PaintPicture picCBM.Picture, 0, 0, W, H, X, Y, W, H            'Copy the area from the Theme bitmap
    Icon.PaintPicture picTheme.Picture, 0, 0, W, H, X, Y, W, H            'Copy the area from the Theme bitmap
    DoEvents
    
End Sub

'---- GRAPHICAL: Get Theme Colour at co-ordinates X,Y
Public Function GetTheme(ByVal X As Integer, Y As Integer) As Long

    GetTheme = picTheme.Point(X, Y)

End Function

'---- GRAPHICAL: Handle Mouse Clicks on Directory List
' Used for Item Selection. When NOT SHIFTED we need to save the Index for the SHIFT-Selection operation
Private Sub picDir_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    If Button = 1 Then                                          'Left-Button
        If Shift <> 1 Then                                      'NOT SHIFTED
            MouseDownI = vsImgDir(Index).value + Int(Y / 16)    'Remember Index of Down
        End If
    Else
        MouseDownI = -1                                         'Ignore NOT Left-button clicks
    End If
    
End Sub

'---- GRAPHICAL: Select Items in GRAPHICAL Directory List
' Works similar to Windows normal selection for file lists:
' - Click without a KEY to select/deselect - previouse selected items are de-selected
' - Hold SHIFT to select a range
' - Hold CTRL  to add to selected items
Private Sub picDir_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim N As Integer, N1 As Integer, N2 As Integer, i As Integer, Max As Integer
    Dim Sel As Boolean
    
    If Button <> 1 Then Exit Sub                                'Not Left-Mouse Button so exit
    If MouseDownI = -1 Then Exit Sub                            'Down was different button
    
    N = vsImgDir(Index).value + Int(Y / (RenderY(Index) * 8))   'New Item Index - calculate bases on height of characters
    Max = lstImageFiles(Index).ListCount - 1                    'Max Index
    If Max < 0 Then Exit Sub
    If N > Max Then N = Max
    
    Sel = lstImageFiles(Index).Selected(N)                      'New Item Selected Status
    
    Select Case Shift
        Case 0                                                  '--- No key: De-Select all files then select new one
            DSelector False, Index                              'De-Select files
            lstImageFiles(Index).Selected(N) = Not Sel          'Invert Selected State
            
        Case 1                                                  '--- Shift Key: Select a range
            N1 = MouseDownI                                     'Mouse Down Index
            N2 = N                                              'New Index
            If N1 > N2 Then i = N1: N1 = N2: N2 = i             'Swap to put lowest first
            If N2 > Max Then N2 = Max                           'Make sure not to go past end of list
            
            For i = N1 To N2
                lstImageFiles(Index).Selected(i) = True         'Make it selected
            Next i
            
        Case 2                                                  '--- Ctrl key: Do not De-Select any, toggle new one
            If N <= Max Then
                lstImageFiles(Index).Selected(N) = Not Sel      'Invert Selected State
            End If
    End Select
        
    RefreshList Index

End Sub

'---- GRAPHICAL: Double-click Items in GRAPHICAL Directory List
Private Sub picDir_DblClick(Index As Integer)
    
    lstImageFiles(Index).Selected(MouseDownI) = True        'Make sure double-clicked entry is selected
    
    Select Case Index
        Case 0, 1: ImageFileView Index
        Case 2: XView
        Case 3:
    End Select
End Sub


'---- GRAPHICAL: Scroll the directory listing
Private Sub vsImgDir_Change(Index As Integer)
    RefreshList Index
End Sub

'---- GRAPHICAL: Scroll the directory listing
Private Sub vsImgDir_Scroll(Index As Integer)
    RefreshList Index
End Sub

'-- GRAPHICAL: Calculate scrollbar values
Private Sub CalcScroll(ByVal Index As Integer)
    Dim L As Integer, M As Integer, N As Integer
    Dim vH As Integer
    
    vH = vsImgDir(Index).Height / (15 * 8 * RenderY(Index))     'Number of visible entries in list based on height of list and scalefactor
    
    N = lstImageFiles(Index).ListCount                          'Number of Entries in list
    M = N - vH                                                  'Subtract how many fit in listing
    
    If M < 0 Then
        M = 0: L = vH                                           'Doesn't fill the window?
    Else
        L = Int(N / vH) * 100: If L > vH Then L = vH            'Scrollbar size
    End If
    
    '-- Set the Scrollbar
    
    vsImgDir(Index).Max = M                                    'Set Scrollbar Range
    vsImgDir(Index).Min = 0
    vsImgDir(Index).SmallChange = 1
    vsImgDir(Index).LargeChange = L                            'Set Scrollbar size
    If vsImgDir(Index).value > M Then vsImgDir(Index).value = 0

End Sub

 
'---- FONT: Render Font for CBM Directory Lists
' Font File must be in FontBuf string.
' Render a CBM-Transfer BINARY font file (no Load Address) for use in Rendering Directory Lists etc
' FG/BG are render colours
Public Sub RenderFont(ByVal FG As Long, ByVal BG As Long)
    Dim J As Integer, V As Integer, P As Integer
    Dim X As Integer, Y As Integer                                      'Co-ordinates
    Dim TopX As Integer, TopY As Integer                                'Top-left for character set
    Dim R As Integer, C As Integer                                      'Row and Col
    Dim SrcX As Integer, VLen As Integer
    
    Dim CCZ As Integer, RRZ As Integer, YYZ As Integer                  'To help speed up drawing
    Dim FH As Integer, FW As Integer
    
    VLen = Len(FontBuf)                                                 'Re-calculate font size buffer in case it has been modified by CUT or INSERT
   
    FW = 8: FH = 8                                                      'Chr Width/Height in pixels
    P = 1: SrcX = 0                                                     'Height constant for 1 pixel and L
    
    C = 0: R = 0: X = 0: Y = 0: TopX = 0: TopY = 0                      'Init Variables for Top-Left of bitmap
    
    picCBM.BackColor = BG                                                 'Set BG colour
    picCBM.Cls                                                            'Clear the bitmap
    'picCBM.Visible = False                                                'Hide the bitmap so drawing is faster
    
    For J = 1 To VLen
        V = Asc(Mid(FontBuf, J, 1))                                        'Get the BYTE value as Y offset into Pixel bitmap
        
        '----paintpicture {srcimg},destX,destY,destW,destH ,srcX,srcY,srcW,srcH,mode
        picCBM.PaintPicture Pix.Image, TopX, TopY + Y, FW, P, SrcX, V, FW, P  'blit the pixels
        
        Y = Y + 1                                                       'Next scanline
        
        If Y = FH Then                                                  '-- Reached character height
            Y = 0: C = C + 1                                            'Reset Y and go to Next Column
            If C >= 128 Then C = 0: R = R + 1: TopY = R * FH            'Go to Start of next Character Row - recalc Top Y
            TopX = C * FW                                               'Next Character Column - pre-calc to speed up draw
        End If

    Next J
   
    'picCBM.Visible = True
End Sub

'---- FONT: Create Pixels
' Creates a bitmap containing the pixel representation for all values from 0 to 255, using the specified Fore and Back-ground colours.
' These pixels will be blitted to the font bitmap one scanline x 8 pixels at a time
Private Sub CreateFontPixels(ByVal FG As Long, ByVal BG As Long)
    Dim J As Integer, K As Integer
    
    Pix.ForeColor = FG
    Pix.BackColor = BG
    Pix.Cls
        
    '-- Create a 2-colour bitmap with pixels to match binary representation of value (row=value,cols 0 to 7=pixel)
    For J = 0 To 255
        For K = 0 To 7
            If (J And Pow(K)) Then Pix.PSet (7 - K, J)
        Next K
    Next J

End Sub

