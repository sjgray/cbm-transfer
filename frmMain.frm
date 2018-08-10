VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CBM Transfer"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15870
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   15870
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frDDF 
      Caption         =   "Disk Image File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Index           =   1
      Left            =   4860
      TabIndex        =   124
      Top             =   7200
      Visible         =   0   'False
      Width           =   4695
      Begin VB.ListBox lstImageFiles 
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
         Height          =   4785
         Index           =   1
         ItemData        =   "frmMain.frx":0442
         Left            =   120
         List            =   "frmMain.frx":0444
         MultiSelect     =   2  'Extended
         OLEDropMode     =   1  'Manual
         TabIndex        =   133
         Top             =   1020
         Width           =   4455
      End
      Begin VB.CommandButton cmdDAll 
         Caption         =   "++"
         Height          =   345
         Index           =   1
         Left            =   900
         TabIndex        =   132
         ToolTipText     =   "Select ALL files"
         Top             =   6300
         Width           =   345
      End
      Begin VB.CommandButton cmdDNone 
         Caption         =   "--"
         Height          =   345
         Index           =   1
         Left            =   1260
         TabIndex        =   131
         ToolTipText     =   "Select None"
         Top             =   6300
         Width           =   315
      End
      Begin VB.CommandButton cmdDDelete 
         Caption         =   "&Del"
         Height          =   345
         Index           =   1
         Left            =   1680
         TabIndex        =   130
         ToolTipText     =   "Delete the selected file(s)"
         Top             =   6300
         Width           =   465
      End
      Begin VB.CommandButton cmdDView 
         Caption         =   "&View"
         Height          =   345
         Index           =   1
         Left            =   3300
         TabIndex        =   129
         ToolTipText     =   "View selected file"
         Top             =   6300
         Width           =   555
      End
      Begin VB.CommandButton cmdDRun 
         Caption         =   "&Run"
         Height          =   345
         Index           =   1
         Left            =   3900
         TabIndex        =   128
         ToolTipText     =   "Run selected file in Vice"
         Top             =   6300
         Width           =   645
      End
      Begin VB.CommandButton cmdImageRefresh 
         Caption         =   "Refresh"
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   127
         ToolTipText     =   "Delete the selected file(s)"
         Top             =   6300
         Width           =   705
      End
      Begin VB.CommandButton cmdDskEd 
         Caption         =   "&Edit"
         Height          =   345
         Index           =   1
         Left            =   2730
         TabIndex        =   126
         Top             =   6300
         Width           =   465
      End
      Begin VB.CommandButton cmdDRename 
         Caption         =   "Ren"
         Height          =   345
         Index           =   1
         Left            =   2190
         TabIndex        =   125
         ToolTipText     =   "Rename file(s)"
         Top             =   6300
         Width           =   465
      End
      Begin VB.Image cmdImageMenu 
         Height          =   255
         Index           =   1
         Left            =   4290
         Picture         =   "frmMain.frx":0446
         Top             =   660
         Width           =   255
      End
      Begin VB.Label txtImageHeader 
         BackColor       =   &H00FF8383&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   138
         Top             =   660
         Width           =   2025
         WordWrap        =   -1  'True
      End
      Begin VB.Label txtImageID 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8383&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Index           =   1
         Left            =   2190
         TabIndex        =   137
         Top             =   660
         Width           =   765
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filename:"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   136
         Top             =   300
         Width           =   675
      End
      Begin VB.Label lblDDFile 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "- No Image Loaded -"
         Height          =   315
         Index           =   1
         Left            =   810
         TabIndex        =   135
         Top             =   240
         Width           =   3750
      End
      Begin VB.Label DFBlocksFree 
         BackColor       =   &H00FF8383&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   134
         Top             =   5880
         Width           =   4455
      End
   End
   Begin VB.Frame frSrc 
      Caption         =   "Directory on Local PC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Index           =   1
      Left            =   120
      TabIndex        =   102
      Top             =   7200
      Width           =   4695
      Begin VB.ComboBox cboFilter 
         Height          =   315
         Index           =   1
         ItemData        =   "frmMain.frx":07FC
         Left            =   960
         List            =   "frmMain.frx":0836
         Style           =   2  'Dropdown List
         TabIndex        =   121
         Top             =   600
         Width           =   2835
      End
      Begin VB.CommandButton cmdSrcBrowse 
         Caption         =   "..."
         Height          =   255
         Index           =   1
         Left            =   3870
         TabIndex        =   120
         ToolTipText     =   "Select Folder"
         Top             =   630
         Width           =   375
      End
      Begin VB.ComboBox txtLocalDir 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         ItemData        =   "frmMain.frx":0983
         Left            =   390
         List            =   "frmMain.frx":0985
         OLEDropMode     =   1  'Manual
         Sorted          =   -1  'True
         TabIndex        =   119
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton cmdSrcDelete 
         Caption         =   "&Delete"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   114
         ToolTipText     =   "Delete selected file(s)"
         Top             =   6300
         Width           =   1065
      End
      Begin VB.CommandButton cmdSrcRename 
         Caption         =   "R&ename"
         Height          =   315
         Index           =   1
         Left            =   1260
         TabIndex        =   113
         ToolTipText     =   "Rename Selected File(s)"
         Top             =   6300
         Width           =   1065
      End
      Begin VB.CommandButton cmdSrcRun 
         Caption         =   "R&un"
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   112
         ToolTipText     =   "Run File or Image using Vice "
         Top             =   5940
         Width           =   1065
      End
      Begin VB.TextBox KBText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3540
         TabIndex        =   110
         Text            =   "0"
         Top             =   6060
         Width           =   705
      End
      Begin VB.TextBox BlockText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3540
         TabIndex        =   109
         Text            =   "0"
         Top             =   6360
         Width           =   705
      End
      Begin VB.CommandButton cmdNewImage 
         Caption         =   "&New Dnn"
         Height          =   315
         Index           =   1
         Left            =   1260
         TabIndex        =   108
         ToolTipText     =   "Create a new/blank CBM Image File"
         Top             =   5940
         Width           =   1065
      End
      Begin VB.CommandButton cmdSrcView 
         Caption         =   "&View"
         Height          =   315
         Index           =   1
         Left            =   2370
         TabIndex        =   107
         ToolTipText     =   "View File or Disk Image file Contents"
         Top             =   6300
         Width           =   675
      End
      Begin VB.CommandButton cmdSrcRefresh 
         Caption         =   "Re&fresh"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   106
         ToolTipText     =   "Refresh Directory"
         Top             =   5940
         Width           =   1065
      End
      Begin VB.DriveListBox drvLocal 
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   105
         Top             =   1020
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.DirListBox dirLocal 
         Height          =   4365
         Index           =   1
         Left            =   150
         TabIndex        =   104
         Top             =   1410
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdSrcView2 
         Caption         =   "&2"
         Height          =   345
         Index           =   1
         Left            =   3090
         TabIndex        =   103
         ToolTipText     =   "View File or Disk Image file Contents"
         Top             =   6270
         Width           =   375
      End
      Begin VB.FileListBox lstLocal 
         Height          =   4770
         Index           =   1
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   111
         Top             =   1020
         Width           =   4490
      End
      Begin VB.Image cmdLocalMenu 
         Height          =   255
         Index           =   1
         Left            =   4320
         Picture         =   "frmMain.frx":0987
         Top             =   630
         Width           =   255
      End
      Begin VB.Image cmdPathUp 
         Height          =   270
         Index           =   1
         Left            =   120
         Picture         =   "frmMain.frx":0D3D
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show:"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   122
         Top             =   660
         Width           =   450
      End
      Begin VB.Label lblSrcSel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected:"
         Height          =   195
         Index           =   1
         Left            =   3540
         TabIndex        =   118
         Top             =   5820
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "KB"
         Height          =   195
         Index           =   1
         Left            =   4300
         TabIndex        =   117
         Top             =   6120
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Blks"
         Height          =   195
         Index           =   1
         Left            =   4300
         TabIndex        =   116
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label lblPathView 
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
         Index           =   1
         Left            =   180
         TabIndex        =   115
         ToolTipText     =   "Drive and Folder View"
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.Frame frDestB 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   10620
      TabIndex        =   90
      Top             =   60
      Width           =   5175
      Begin VB.Label lblDstMode 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Disk Image"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   3900
         TabIndex        =   99
         Tag             =   "&H0000C0C0&"
         ToolTipText     =   "Click to View Disk Image Files"
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label lblDstMode 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X-Cable"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   -60
         TabIndex        =   93
         Tag             =   "&H00FF0000&"
         ToolTipText     =   "Click to View X-Cable directory"
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label lblDstMode 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CBMLink"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   1260
         TabIndex        =   92
         Tag             =   "&H000040C0&"
         ToolTipText     =   "Click to View CBMLink directory"
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label lblDstMode 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local PC"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   2580
         TabIndex        =   91
         Tag             =   "&H0000C0C0&"
         ToolTipText     =   "Click to View Files"
         Top             =   0
         Width           =   1260
      End
   End
   Begin VB.Frame frMiddle 
      BorderStyle     =   0  'None
      Height          =   6555
      Left            =   9600
      TabIndex        =   81
      Top             =   480
      Width           =   915
      Begin VB.CommandButton cmdCopyRight 
         Height          =   735
         Left            =   60
         Picture         =   "frmMain.frx":10DF
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Transfer"
         Top             =   2160
         Width           =   795
      End
      Begin VB.CommandButton cmdCopyLeft 
         Height          =   735
         Left            =   60
         Picture         =   "frmMain.frx":2C21
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "If no files are selected, this button will create a Disk Image of the entire disk."
         Top             =   3000
         Width           =   795
      End
      Begin VB.CommandButton About 
         Caption         =   "&About"
         Height          =   345
         Left            =   60
         TabIndex        =   86
         ToolTipText     =   "Show Program Information"
         Top             =   5280
         Width           =   795
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "&Options"
         Height          =   345
         Left            =   60
         TabIndex        =   85
         ToolTipText     =   "Set Program Options"
         Top             =   480
         Width           =   795
      End
      Begin VB.CommandButton cmdResults 
         Caption         =   "Results"
         Height          =   345
         Left            =   60
         TabIndex        =   84
         ToolTipText     =   "Show Output from Last Command"
         Top             =   1080
         Width           =   795
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "&Help"
         Height          =   345
         Left            =   60
         TabIndex        =   83
         ToolTipText     =   "Show Help File"
         Top             =   5880
         Width           =   795
      End
      Begin VB.CommandButton cmdDAD 
         Caption         =   "DAD"
         Height          =   345
         Left            =   60
         TabIndex        =   82
         ToolTipText     =   "Open Drag and Drop Window"
         Top             =   4320
         Width           =   795
      End
      Begin VB.Label lblR 
         Alignment       =   2  'Center
         Caption         =   "*"
         Height          =   165
         Left            =   330
         TabIndex        =   141
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblSizer 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   600
         TabIndex        =   89
         ToolTipText     =   "Show/Hide Pane"
         Top             =   6360
         Width           =   225
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8460
      Top             =   -90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frDDF 
      Caption         =   "Disk Image File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Index           =   0
      Left            =   4860
      TabIndex        =   29
      Top             =   420
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdDRename 
         Caption         =   "Ren"
         Height          =   345
         Index           =   0
         Left            =   2190
         TabIndex        =   123
         ToolTipText     =   "Rename file(s)"
         Top             =   6300
         Width           =   465
      End
      Begin VB.CommandButton cmdDskEd 
         Caption         =   "&Edit"
         Height          =   345
         Index           =   0
         Left            =   2730
         TabIndex        =   101
         Top             =   6300
         Width           =   465
      End
      Begin VB.CommandButton cmdImageRefresh 
         Caption         =   "Refresh"
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   78
         ToolTipText     =   "Delete the selected file(s)"
         Top             =   6300
         Width           =   705
      End
      Begin VB.CommandButton cmdDRun 
         Caption         =   "&Run"
         Height          =   345
         Index           =   0
         Left            =   3900
         TabIndex        =   70
         ToolTipText     =   "Run selected file in Vice"
         Top             =   6300
         Width           =   645
      End
      Begin VB.CommandButton cmdDView 
         Caption         =   "&View"
         Height          =   345
         Index           =   0
         Left            =   3300
         TabIndex        =   45
         ToolTipText     =   "View selected file"
         Top             =   6300
         Width           =   555
      End
      Begin VB.CommandButton cmdDDelete 
         Caption         =   "&Del"
         Height          =   345
         Index           =   0
         Left            =   1680
         TabIndex        =   44
         ToolTipText     =   "Delete the selected file(s)"
         Top             =   6300
         Width           =   465
      End
      Begin VB.CommandButton cmdDNone 
         Caption         =   "--"
         Height          =   345
         Index           =   0
         Left            =   1260
         TabIndex        =   42
         ToolTipText     =   "Select None"
         Top             =   6300
         Width           =   315
      End
      Begin VB.CommandButton cmdDAll 
         Caption         =   "++"
         Height          =   345
         Index           =   0
         Left            =   900
         TabIndex        =   41
         ToolTipText     =   "Select ALL files"
         Top             =   6300
         Width           =   345
      End
      Begin VB.ListBox lstImageFiles 
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
         Height          =   4785
         Index           =   0
         ItemData        =   "frmMain.frx":4763
         Left            =   120
         List            =   "frmMain.frx":4765
         MultiSelect     =   2  'Extended
         OLEDropMode     =   1  'Manual
         TabIndex        =   30
         Top             =   1020
         Width           =   4455
      End
      Begin VB.Image cmdImageMenu 
         Height          =   255
         Index           =   0
         Left            =   4260
         Picture         =   "frmMain.frx":4767
         Top             =   660
         Width           =   255
      End
      Begin VB.Label DFBlocksFree 
         BackColor       =   &H00FF8383&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   5880
         Width           =   4455
      End
      Begin VB.Label lblDDFile 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "- No Image Loaded -"
         Height          =   315
         Index           =   0
         Left            =   810
         TabIndex        =   34
         Top             =   240
         Width           =   3750
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filename:"
         Height          =   195
         Index           =   9
         Left            =   105
         TabIndex        =   33
         Top             =   300
         Width           =   675
      End
      Begin VB.Label txtImageID 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8383&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Index           =   0
         Left            =   2190
         TabIndex        =   32
         Top             =   660
         Width           =   765
      End
      Begin VB.Label txtImageHeader 
         BackColor       =   &H00FF8383&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   660
         Width           =   2025
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame frLink 
      Caption         =   "CBM Drive via CBMLink"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   10620
      TabIndex        =   46
      Top             =   7200
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdLinkScratch 
         Caption         =   "Scratch"
         Height          =   360
         Left            =   4095
         TabIndex        =   62
         ToolTipText     =   "Scratch (delete) selected file(s)"
         Top             =   3990
         Width           =   975
      End
      Begin VB.CommandButton cmdLinkRename 
         Caption         =   "Rename"
         Height          =   360
         Left            =   4095
         TabIndex        =   61
         ToolTipText     =   "Rename selected file(s)"
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdLinkStatus 
         Caption         =   "Status"
         Height          =   360
         Left            =   4095
         TabIndex        =   60
         Top             =   6270
         Width           =   975
      End
      Begin VB.CommandButton cmdLinkReset 
         Caption         =   "Reset"
         Height          =   360
         Left            =   4095
         TabIndex        =   59
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton cmdLinkFormat 
         Caption         =   "Format"
         Height          =   360
         Left            =   4095
         TabIndex        =   58
         ToolTipText     =   "Format disk in Floppy drive"
         Top             =   1905
         Width           =   975
      End
      Begin VB.CommandButton cmdLinkInit 
         Caption         =   "Initialize"
         Height          =   360
         Left            =   4095
         TabIndex        =   57
         ToolTipText     =   "Reset the Drive"
         Top             =   2295
         Width           =   975
      End
      Begin VB.CommandButton cmdLinkValidate 
         Caption         =   "Validate"
         Height          =   360
         Left            =   4095
         TabIndex        =   56
         ToolTipText     =   "Perform Disk Validation"
         Top             =   2685
         Width           =   975
      End
      Begin VB.CommandButton cmdLinkAll 
         Caption         =   "All"
         Height          =   360
         Left            =   4095
         TabIndex        =   55
         ToolTipText     =   "Select ALL files"
         Top             =   5085
         Width           =   420
      End
      Begin VB.CommandButton cmdLinkNone 
         Caption         =   "None"
         Height          =   360
         Left            =   4545
         TabIndex        =   54
         ToolTipText     =   "Select None"
         Top             =   5085
         Width           =   540
      End
      Begin VB.CommandButton cmdLinkDir 
         Caption         =   "Directory"
         Height          =   375
         Left            =   4065
         TabIndex        =   52
         Top             =   1305
         Width           =   975
      End
      Begin VB.ComboBox cboLinkDev 
         Height          =   315
         ItemData        =   "frmMain.frx":4B1D
         Left            =   720
         List            =   "frmMain.frx":4B39
         Style           =   2  'Dropdown List
         TabIndex        =   51
         ToolTipText     =   "Select X Device Unit Number"
         Top             =   225
         Width           =   1230
      End
      Begin VB.ListBox lstLink 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   4560
         ItemData        =   "frmMain.frx":4B80
         Left            =   120
         List            =   "frmMain.frx":4B82
         MultiSelect     =   2  'Extended
         OLEDropMode     =   1  'Manual
         TabIndex        =   47
         Top             =   1005
         Width           =   3885
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Device:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   140
         Top             =   270
         Width           =   555
      End
      Begin VB.Label lblLinkLastStatus 
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8383&
         Height          =   285
         Left            =   120
         TabIndex        =   67
         Top             =   6345
         Width           =   3855
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Drive Status:"
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
         Index           =   18
         Left            =   120
         TabIndex        =   66
         Top             =   6105
         Width           =   2175
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
         Height          =   195
         Index           =   17
         Left            =   4125
         TabIndex        =   65
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drive:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   4095
         TabIndex        =   64
         Top             =   5640
         Width           =   525
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   4125
         TabIndex        =   63
         Top             =   4845
         Width           =   615
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disk:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   4080
         TabIndex        =   53
         Top             =   1065
         Width           =   450
      End
      Begin VB.Label lblLinkDiskName 
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   120
         TabIndex        =   50
         Top             =   630
         Width           =   2775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLinkDiskID 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   2940
         TabIndex        =   49
         Top             =   630
         Width           =   855
      End
      Begin VB.Label lblLinkBlocksFree 
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   5640
         Width           =   3885
      End
   End
   Begin VB.Frame frX 
      Caption         =   "CBM Drive on X-Cable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   10620
      TabIndex        =   0
      ToolTipText     =   "Reset Drive"
      Top             =   420
      Width           =   5175
      Begin VB.CommandButton cmdXRoot 
         Caption         =   "Root"
         Height          =   360
         Left            =   4080
         TabIndex        =   76
         ToolTipText     =   "Return to Root Partition"
         Top             =   4500
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdXPart 
         Caption         =   "Sel"
         Height          =   360
         Left            =   4620
         TabIndex        =   75
         ToolTipText     =   "Select/View partition"
         Top             =   4500
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton cmdXView 
         Caption         =   "View"
         Height          =   360
         Left            =   4080
         TabIndex        =   68
         ToolTipText     =   "CBM File Viewer"
         Top             =   3780
         Width           =   975
      End
      Begin VB.ComboBox cboXDevNum 
         Height          =   315
         ItemData        =   "frmMain.frx":4B84
         Left            =   720
         List            =   "frmMain.frx":4B94
         Style           =   2  'Dropdown List
         TabIndex        =   38
         ToolTipText     =   "Select X Device Unit Number"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdXNone 
         Caption         =   "None"
         Height          =   360
         Left            =   4515
         TabIndex        =   27
         ToolTipText     =   "Select None"
         Top             =   5220
         Width           =   540
      End
      Begin VB.CommandButton cmdXAll 
         Caption         =   "All"
         Height          =   360
         Left            =   4065
         TabIndex        =   26
         ToolTipText     =   "Select ALL files"
         Top             =   5220
         Width           =   420
      End
      Begin VB.CommandButton cmdXValidate 
         Caption         =   "Val"
         Height          =   360
         Left            =   4600
         TabIndex        =   24
         ToolTipText     =   "Perform Disk Validation"
         Top             =   1800
         Width           =   435
      End
      Begin VB.CommandButton cmdXInit 
         Caption         =   "Init"
         Height          =   360
         Left            =   4080
         TabIndex        =   17
         ToolTipText     =   "Reset the Drive"
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton cmdXFormat 
         Caption         =   "Format"
         Height          =   360
         Left            =   4080
         TabIndex        =   8
         ToolTipText     =   "Format disk in Floppy drive"
         Top             =   2220
         Width           =   975
      End
      Begin VB.ListBox lstXFiles 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   4560
         ItemData        =   "frmMain.frx":4BA6
         Left            =   120
         List            =   "frmMain.frx":4BA8
         MultiSelect     =   2  'Extended
         OLEDropMode     =   1  'Manual
         TabIndex        =   7
         Top             =   1020
         Width           =   3855
      End
      Begin VB.CommandButton cmdXReset 
         Caption         =   "Reset"
         Height          =   360
         Left            =   4080
         TabIndex        =   6
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton cmdXDriveStatus 
         Caption         =   "Status"
         Height          =   360
         Left            =   4080
         TabIndex        =   5
         ToolTipText     =   "Get Drive Status"
         Top             =   6300
         Width           =   975
      End
      Begin VB.CommandButton cmdXRefresh 
         Caption         =   "Directory"
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         ToolTipText     =   "Read Disk Directory"
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdXRename 
         Caption         =   "Rename"
         Height          =   360
         Left            =   4080
         TabIndex        =   3
         ToolTipText     =   "Rename selected file(s)"
         Top             =   2940
         Width           =   975
      End
      Begin VB.CommandButton cmdXScratch 
         Caption         =   "Scratch"
         Height          =   360
         Left            =   4080
         TabIndex        =   2
         ToolTipText     =   "Scratch (delete) selected file(s)"
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Device:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   139
         Top             =   300
         Width           =   555
      End
      Begin VB.Label lblDName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1740
         TabIndex        =   79
         ToolTipText     =   "Drive Model# (click to re-scan)"
         Top             =   240
         Width           =   3345
      End
      Begin VB.Label lblXPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Partition:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4080
         TabIndex        =   77
         Top             =   4260
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblXSel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4080
         TabIndex        =   28
         Top             =   4980
         Width           =   615
      End
      Begin VB.Label lblXBlocksFree 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Left            =   120
         TabIndex        =   25
         Top             =   5640
         Width           =   3855
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Drive Status:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   6090
         Width           =   2175
      End
      Begin VB.Label lblXLastStatus 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   6330
         Width           =   3855
      End
      Begin VB.Label lblXDrive 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drive:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4080
         TabIndex        =   18
         Top             =   5640
         Width           =   525
      End
      Begin VB.Label lblXDiskName 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   660
         Width           =   2055
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblXDiskID 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   2220
         TabIndex        =   11
         Top             =   660
         Width           =   765
      End
      Begin VB.Label lblXDisk 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disk:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4080
         TabIndex        =   10
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label lblXFiles 
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
         Height          =   195
         Left            =   4080
         TabIndex        =   9
         Top             =   2700
         Width           =   465
      End
   End
   Begin VB.Frame frSrc 
      Caption         =   "Directory on Local PC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   4695
      Begin VB.CommandButton cmdSrcView2 
         Caption         =   "&2"
         Height          =   345
         Index           =   0
         Left            =   3090
         TabIndex        =   100
         ToolTipText     =   "View File or Disk Image file Contents"
         Top             =   6270
         Width           =   375
      End
      Begin VB.DirListBox dirLocal 
         Height          =   4365
         Index           =   0
         Left            =   150
         TabIndex        =   96
         Top             =   1410
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.DriveListBox drvLocal 
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   95
         Top             =   1020
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox txtLocalDir 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":4BAA
         Left            =   390
         List            =   "frmMain.frx":4BAC
         OLEDropMode     =   1  'Manual
         Sorted          =   -1  'True
         TabIndex        =   80
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton cmdSrcRefresh 
         Caption         =   "Re&fresh"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   72
         ToolTipText     =   "Refresh Directory"
         Top             =   5940
         Width           =   1065
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   3870
         TabIndex        =   71
         ToolTipText     =   "Select Folder"
         Top             =   630
         Width           =   375
      End
      Begin VB.CommandButton cmdSrcView 
         Caption         =   "&View"
         Height          =   315
         Index           =   0
         Left            =   2370
         TabIndex        =   69
         ToolTipText     =   "View File or Disk Image file Contents"
         Top             =   6300
         Width           =   675
      End
      Begin VB.CommandButton cmdNewImage 
         Caption         =   "&New Dnn"
         Height          =   315
         Index           =   0
         Left            =   1260
         TabIndex        =   43
         ToolTipText     =   "Create a new/blank CBM Image File"
         Top             =   5940
         Width           =   1065
      End
      Begin VB.ComboBox cboFilter 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":4BAE
         Left            =   960
         List            =   "frmMain.frx":4BE8
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   600
         Width           =   2835
      End
      Begin VB.TextBox BlockText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3540
         TabIndex        =   21
         Text            =   "0"
         Top             =   6360
         Width           =   705
      End
      Begin VB.TextBox KBText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3540
         TabIndex        =   20
         Text            =   "0"
         Top             =   6060
         Width           =   705
      End
      Begin VB.FileListBox lstLocal 
         Height          =   4770
         Index           =   0
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   16
         Top             =   1020
         Width           =   4490
      End
      Begin VB.CommandButton cmdSrcRun 
         Caption         =   "R&un"
         Height          =   315
         Index           =   0
         Left            =   2400
         TabIndex        =   15
         ToolTipText     =   "Run File or Image using Vice "
         Top             =   5940
         Width           =   1065
      End
      Begin VB.CommandButton cmdSrcRename 
         Caption         =   "R&ename"
         Height          =   315
         Index           =   0
         Left            =   1260
         TabIndex        =   14
         ToolTipText     =   "Rename Selected File(s)"
         Top             =   6300
         Width           =   1065
      End
      Begin VB.CommandButton cmdSrcDelete 
         Caption         =   "&Delete"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Delete selected file(s)"
         Top             =   6300
         Width           =   1065
      End
      Begin VB.Image cmdLocalMenu 
         Height          =   255
         Index           =   0
         Left            =   4320
         Picture         =   "frmMain.frx":4D35
         Top             =   630
         Width           =   255
      End
      Begin VB.Image cmdPathUp 
         Height          =   270
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":50EB
         Top             =   270
         Width           =   240
      End
      Begin VB.Label lblPathView 
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
         Index           =   0
         Left            =   180
         TabIndex        =   97
         ToolTipText     =   "Drive and Folder View"
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Blks"
         Height          =   195
         Index           =   0
         Left            =   4300
         TabIndex        =   40
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "KB"
         Height          =   195
         Index           =   0
         Left            =   4300
         TabIndex        =   39
         Top             =   6120
         Width           =   210
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show:"
         Height          =   195
         Index           =   10
         Left            =   480
         TabIndex        =   35
         Top             =   660
         Width           =   450
      End
      Begin VB.Label lblSrcSel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected:"
         Height          =   195
         Index           =   0
         Left            =   3540
         TabIndex        =   19
         Top             =   5820
         Width           =   675
      End
   End
   Begin VB.Label lblSizer2 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   4920
      TabIndex        =   98
      ToolTipText     =   "Dual-view"
      Top             =   120
      Width           =   225
   End
   Begin VB.Label lblReminder 
      AutoSize        =   -1  'True
      Caption         =   " (Selects SOURCE for transfer)"
      Height          =   195
      Left            =   5160
      TabIndex        =   94
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblSrcMode 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Disk Image"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   2520
      TabIndex        =   74
      Tag             =   "&H00C0C000&"
      ToolTipText     =   "Click to View Disk Image Files"
      Top             =   60
      Width           =   2325
   End
   Begin VB.Label lblSrcMode 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local PC"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   73
      Tag             =   "&H0000C0C0&"
      ToolTipText     =   "Click to View Files"
      Top             =   60
      Width           =   2325
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2017 Steve J. Gray
' ====================================================
'
' frmMain - The MAIN window. Code execution starts here!
'
' Based on GUI4CBM4WIN. The following (between "/" lines) is the notice
' included with the GUI4CBM4WIN source code:
'
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
'
' Steve Gray's notes:
'
' CBM-Transfer (CBMXfer) is based on GUI4CBM4WIN, which was written in VB6 by Leif Bloomquist.
' Eventually GUI4CBM4WIN was converted to VB.NET, which I do not own and find to be bulky and confusing.
' So, I forked the VB6 version and Renamed it to CBM-Transfer to make it easier to say ;-)
' I have greatly expanded the original program, and replaced or re-written the majority of code
' with my own. I added support for CBMLink (which could use IEEE drives before ZoomFloppy),
' for C1541 to work with Disk Images, and the ability to run programs using VICE. I added support
' for Zoomfloppy, NIBTOOLS, IMGCOPY, 1581 directories, and P00 files.
' I added the File Viewer which supports multiple viewing formats, including a fully-featured
' 6502 symbolic machine language disassembler, with dual-view mode.

Option Explicit
Dim Drive(31) As String

'---- Display Program info and acknowlegements
Private Sub About_Click()
    MyMsg "CBM-Transfer  V1.10 (Aug 9/2018)" & Cr & _
          "(C)2007-2018 Steve J. Gray" & Cr & Cr & _
          "A front-end for: OpenCBM, VICE, NibTools, and CBMLink" & Cr & Cr & _
          "Based on GUI4CBM4WIN V0.4.1," & Cr & _
          "by Leif Bloomquist, Wolfgang Moser and Spiro Trikaliotis." & Cr & _
          "Viewer includes portions of 'CBM2BMP' code by Peter Weighill"
End Sub

'---- Program Initialization
Private Sub Form_Load()
    Dim i As Integer, Tmp As String, Flag As Boolean, xTmp As String
    
    On Error Resume Next
    
    MsgTitle = "CBM Transfer"
    Cr = Chr(13): LF = Chr(10): Qu = Chr(34): Nu = Chr(0): Hx = "&h"    'some common characters
    SrcMode = 0: Layout = 0: Layout2 = 1
    
    '-- OpenCBM Command Strings
    CBMCtrl = "cbmctrl"
    CBMCopy = "cbmcopy"
    CBMC1541 = "c1541"
    CMDSTR = "command "
    '--
    CurDir = App.Path
    ExeDir = AddSlash(CurDir)
        
    If Mid(CurDir, 2, 1) = ":" Then ChDrive Left(CurDir, 1)
    MyChDir CurDir
    
    LoadINI
    
    LogFile = CurDir & "\cbmxferlog.txt"
    
    Tmp = Environ$("temp"): If Tmp = "" Then Tmp = Environ$("tmp")
    Tmp = Tmp & "\cbmxfer"
    
    TEMPFILE1 = Tmp & "out.txt"                     'Captured Output from shell
    TEMPFILE2 = Tmp & "err.txt"                     'Captured Errors from shell
    TEMPFILE3 = Tmp & "tmp.tmp"                     'General-purpose Temp File (ie: for multi-step copies)
    
    PathFile = ExeDir & "pathhistory.txt"
    LoadHistory
    
    cboXDevNum.ListIndex = frmOptions.cboDriveNum.ListIndex
    cboLinkDev.ListIndex = 0
    cboFilter(0).ListIndex = 0
    cboFilter(1).ListIndex = 0
       
    
    If frmOptions.cbUseCBMFont.value = vbChecked Then
        Tmp = "C64 User Mono"
        lstImageFiles(0).Font.Name = Tmp: lstImageFiles(0).Font.Size = 8
        lstImageFiles(1).Font.Name = Tmp: lstImageFiles(1).Font.Size = 8
        lstXFiles.Font.Name = Tmp: lstXFiles.Font.Size = 8
        lstLink.Font.Name = Tmp: lstLink.Font.Size = 8
    End If
    
    For i = 0 To 1
        If DirExists(LocalDir(i)) = False Then LocalDir(i) = AddSlash(App.Path)
        SetLocalPath i, LocalDir(i)
        lstImageFiles(i).Clear
    Next i
        
    SetLayout
    
    SetSrcFrame
    SetDstFrame
    
    If CheckEXE = True Then
        Flag = False
        Tmp = "The following files are missing:"
        xTmp = "cbmlink.exe": If Exists(ExeDir & xTmp) = False Then Flag = True: Tmp = Tmp & "  " & xTmp
        xTmp = "cbmctrl.exe": If Exists(ExeDir & xTmp) = False Then Flag = True: Tmp = Tmp & "  " & xTmp
        xTmp = "c1541.exe":   If Exists(ExeDir & xTmp) = False Then Flag = True: Tmp = Tmp & "  " & xTmp
        If Flag = True Then MyMsg Tmp & Cr & Cr & "Please copy them to:" & Cr & ExeDir
    End If
    
    If StartDAD = True Then
        frmDAD.Show
        frmMain.WindowState = 1
        DoEvents
    End If
    
End Sub

'---- Save stuff before program exits
' This stuff gets done before ANY forms are unloaded, otherwise some properties
' needed for INI saving will be unloaded before they can be saved
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveHistory
    SaveINI
    MyChDir ExeDir
End Sub

'---- End the program
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

'---- Set the positions and sizes on the form
Sub SetLayout()
    Dim c0 As Single, C1 As Single, C2 As Single, C3 As Single
    Dim W0 As Single, W1 As Single
    
    W0 = frSrc(0).Width + 30
    W1 = frMiddle.Width + 30
    
    c0 = frSrc(0).Left
    C1 = c0: If Layout = 1 Then C1 = C1 + W0
    C2 = C1 + W0
    C3 = C2 + W1
    
    frDDF(0).Move C1, frSrc(0).Top              'LEFT
    frMiddle.Left = C2                          'MIDDLE
    frDestB.Left = C3                           'Right header
    frX.Left = C3                               'RIGHT
    frSrc(1).Move C3, frX.Top                   'RIGHT
    frLink.Move C3, frX.Top                     'RIGHT
    frDDF(1).Move C3, frX.Top                   'RIGHT
    
    Me.Height = 7715

    If Layout2 = 0 Then
        lblSizer.Caption = ">>"
    Else
        lblSizer.Caption = "<<"
    End If

    Select Case Layout
        Case 0:
            lblSizer2.Caption = ">>"
            lblReminder.Visible = False
            
            SetSrcFrame
    
            If Layout2 = 0 Then
                Me.Width = 5850                 'Set main window size
            Else
                Me.Width = 11145                'Set main window size
            End If
            
        Case 1
            lblSizer2.Caption = "<<"
            lblReminder.Visible = True
    
            frSrc(0).Visible = True
            frDDF(0).Visible = True
            
            Me.Width = 11145                    'Set main window size
    
            If Layout2 = 0 Then
                Me.Width = 10600                'Set main window size
            Else
                Me.Width = 15860                'Set main window size
            End If
    End Select

End Sub

'============================
'  Subs for Program Options
'============================

'---- Show the CBM-Transfer TXT file documentation, using associated viewer (ie: Notepad)
Private Sub cmdHelp_Click()
    ViewFile ExeDir & "\CBMXfer.txt"
End Sub

'---- Open the options window
Private Sub cmdOptions_Click()
    frmOptions.Show vbModal
End Sub

'---- Swap Left and Right directories
Private Sub SwapDirs()
    Dim Tmp
    
    Tmp = LocalDir(0): LocalDir(0) = LocalDir(1): LocalDir(1) = Tmp
    
    SetLocalPath 0, LocalDir(0)
    SetLocalPath 1, LocalDir(1)

End Sub

'---- Show the results TXT files with associated viewer
Private Sub cmdResults_Click()
    If frmOptions.cbErr.value = vbChecked Then ViewFile TEMPFILE2 'This file is usually empty
    ViewFile TEMPFILE1  'This contains the actual data
End Sub

'---- Special options for specific disk types
Private Sub lblDName_Click()
    DetectDrives True
    GetXDevices
    If Left(lblDName.Caption, 4) = "1571" Then Ask1571Mode
    If lblDName.Caption = "8250" Then Ask8050Mode
    If lblDName.Caption = "SFD-1001" Then Ask8050Mode
End Sub

'---- Ask user for 1571 mode (single or double-sided)
Private Sub Ask1571Mode()
    Dim Choice As Integer, Tmp As String, TCmd As String, TMsg As String
    Dim Status As ReturnStringType

    Choice = MsgBox("1571 Drive. Do you want to use Double-sided mode?", vbYesNoCancel, "Select Mode")
    Tmp = "0": If Choice = vbYes Then Tmp = "1"
    TCmd = CMDSTR & DriveNum & " "
    TMsg = "Setting 1571 mode..."

    If Choice <> vbCancel Then Status = DoCommand(CBMCtrl, TCmd & Quoted("U0>M" & Tmp), TMsg)

End Sub

'---- Ask user for 8050 mode (single sided)
' If 8050 mode is wanted it must send commands to set specific locations in the Drive's memory
Private Sub Ask8050Mode()
    Dim Choice As Integer, Tmp As String, TCmd As String, TMsg As String
    Dim Status As ReturnStringType ''

    Choice = MsgBox("8250/SFD. Do you want to use 8050 mode?", vbYesNoCancel, "Select Mode")
    
    Tmp = "0": If Choice = vbYes Then Tmp = "1"
    TCmd = CMDSTR & DriveNum & " "
    TMsg = "Setting 8050 mode..."
    
    If Choice <> vbCancel Then
        Status = DoCommand(CBMCtrl, TCmd & Quoted("m-w 172 16 1 1"), TMsg)
        Status = DoCommand(CBMCtrl, TCmd & Quoted("m-w 195 16 1 0"), TMsg)
        Status = DoCommand(CBMCtrl, TCmd & Quoted("u9"), TMsg)
    End If

End Sub

Private Sub lblDstMode_Click(Index As Integer)
    DstMode = Index
    SetDstFrame
End Sub

Private Sub lblPathView_Click(Index As Integer)

    If lblPathView(Index).Caption = ">>" Then
        lblPathView(Index).Caption = "<<"
        drvLocal(Index).Visible = True
        dirLocal(Index).Visible = True
        lstLocal(Index).Move 2360, 1020, 2245, 4770
    Else
        lblPathView(Index).Caption = ">>"
        drvLocal(Index).Visible = False
        dirLocal(Index).Visible = False
        lstLocal(Index).Move 120, 1020, 4490, 4770
    End If
    
    DoEvents
    
End Sub

Private Sub lblSizer_Click()
    Layout2 = 1 - Layout2
    SetLayout
End Sub

Private Sub lblSizer2_Click()
    Layout = 1 - Layout
    SetLayout
End Sub

'---- Show and Hide the Drag-and-Drop (DAD) window
Private Sub cmdDAD_Click()
    If frmDAD.Visible = False Then frmDAD.Show Else frmDAD.Hide
End Sub

'====================
' Subs for Popup menu
'====================
Private Sub cmdLocalMenu_Click(Index As Integer)
    MenuNum = Index
    PopupMenu frmMenu.mnuPopup
End Sub

Private Sub cmdImageMenu_Click(Index As Integer)
    MenuNum = Index
    PopupMenu frmMenu.mnuSave
End Sub

'---- Handle Menu Selections
' This is called from frmMenu
Sub DoMenu(ByVal Index As Integer)

    Select Case Index
        '--- Menu 1
        Case 1: ShellExecute hWnd, "open", LocalDir(0), vbNullString, LocalDir(0), 1
        Case 2: ShellExecute hWnd, "open", LocalDir(1), vbNullString, LocalDir(1), 1
        Case 3: SwapDirs
        Case 4: SetLocalPath 0, LocalDir(1)
        Case 5: SetLocalPath 1, LocalDir(0)
        Case 6: AddPathHistory LocalDir(0)
        Case 7: RemovePathHistory LocalDir(0)
        Case 8:
            KillFile PathHistory
            frmMain.txtLocalDir(0).Clear
            frmMain.txtLocalDir(1).Clear
            
        Case 9: NewFolder LocalDir(MenuNum): dirLocal(MenuNum).Refresh
            
        '--- Menu 2
        Case 101: SaveDirText
        Case 102: AddToCatalog
        Case 103: ViewFile ExeDir & "\catalog.txt"
        Case 104: ImageValidate MenuNum
        Case 105: ImageBackup MenuNum
    End Select
    
End Sub

Private Sub SaveDirText()
        Dim Filename As String
        
        On Local Error GoTo DialogError
        
        CommonDialog1.CancelError = True
        CommonDialog1.InitDir = LocalDir(MenuNum)
        CommonDialog1.Filter = "Text (*.txt)|*.txt|All Files|*.*"
        CommonDialog1.Filename = FileBase(UnQuoted(DDFile(MenuNum))) & ".txt"
        CommonDialog1.ShowSave
        
        Filename = CommonDialog1.Filename
        
        If Filename = "" Then Exit Sub
        If Overwrite(Filename) = False Then Exit Sub
        WriteDirTextTo Filename, False
        
DialogError:
        Exit Sub
End Sub

Private Sub AddToCatalog()
    WriteDirTextTo ExeDir & "\catalog.txt", True
End Sub

'---- Save the disk directory to a file
' Flag: True=Append, False=Create
Private Sub WriteDirTextTo(ByVal Filename As String, Flag As Boolean)
        Dim FIO As Integer, j As Integer
        
        FIO = FreeFile
        
        If Flag = False Then
            Open Filename For Output As FIO
        Else
            Open Filename For Append As FIO
            Print #FIO, ""
        End If
        Print #FIO, "**** FILE: " & FileNameOnly(lblDDFile(MenuNum).Caption)
        Print #FIO, Qu & txtImageHeader(MenuNum).Caption & Qu & " " & txtImageID(MenuNum).Caption
        Print #FIO, "========================"
        
        For j = 0 To lstImageFiles(MenuNum).ListCount - 1
            Print #FIO, lstImageFiles(MenuNum).List(j)
        Next j
        
        Print #FIO, DFBlocksFree(MenuNum).Caption
        Print #FIO, ""
        Close FIO

End Sub

'=====================================================
'  Subs for Local PC /// Src=Left Side, Dst=Right Side
'=====================================================

Private Sub drvLocal_Change(Index As Integer)
    dirLocal(Index).Path = drvLocal(Index).Drive   'Set directory path.
End Sub

Private Sub dirLocal_Change(Index As Integer)
    SetLocalPath Index, dirLocal(Index).Path        'Set file path.
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    Dim Tmp As String
    Tmp = GetBrowseDir(Me, "Select Source Path:")
    If Tmp <> "" Then SetLocalPath Index, Tmp
End Sub

'---- Move UP one level in Path
Private Sub cmdPathUp_Click(Index As Integer)
    SetLocalPath Index, PathUp(LocalDir(Index))
End Sub

Private Sub lblSrcMode_Click(Index As Integer)
    SrcMode = Index
    SetSrcFrame
End Sub

Private Sub lstImageFiles_DblClick(Index As Integer)
    Call ImageFileView(Index)
End Sub

'---- Change LocalDir to item clicked on
Private Sub txtLocalDir_Click(Index As Integer)
    Dim Tmp As String
    
    Tmp = txtLocalDir(Index).List(txtLocalDir(Index).ListIndex)
    SetLocalPath Index, Tmp
End Sub

'---- Process keystrokes for Directory Path
Private Sub txtLocalDir_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SetLocalPath Index, txtLocalDir(Index).Text
End Sub

'---- Set LocalPC Directory Path
Private Sub SetLocalPath(ByVal Index As Integer, ByVal SPath As String)
    On Local Error Resume Next
    LocalDir(Index) = AddSlash(SPath)
    lstLocal(Index).Path = LocalDir(Index)
    lstLocal(Index).Refresh
    txtLocalDir(Index).Text = LocalDir(Index)
    txtLocalDir(Index).ToolTipText = LocalDir(Index)
    dirLocal(Index).Path = LocalDir(Index)
    
    If PathHistory = True Then AddPathHistory SPath
End Sub

'---- Add Path to History List
Sub AddPathHistory(ByVal SPath As String)
    Dim a As Integer, Flag As Integer
    Flag = True
    
    For a = 0 To txtLocalDir(0).ListCount - 1
        If txtLocalDir(0).List(a) = SPath Then Flag = False: Exit For
    Next
    
    If Flag = True Then
        txtLocalDir(0).AddItem SPath 'Adds the item to the list
        txtLocalDir(1).AddItem SPath 'Adds the item to the list
    End If
    
End Sub

'---- Remove Path from History List
Sub RemovePathHistory(ByVal SPath As String)
    Dim a As Integer

    For a = txtLocalDir(0).ListCount - 1 To 0 Step -1
        If txtLocalDir(0).List(a) = SPath Then
            txtLocalDir(0).RemoveItem (a)
            txtLocalDir(1).RemoveItem (a)
        End If
    Next
    
End Sub

'---- Hilight Selected Tab
Sub SetSrcFrame()
    Dim a As Integer
    
    If Layout = 0 Then frSrc(0).Visible = False: frDDF(0).Visible = False
        
    For a = 0 To 1
        lblSrcMode(a).Font.Bold = False
        lblSrcMode(a).ForeColor = vbBlack
    Next a
        
    Select Case SrcMode
        Case 0: frSrc(0).Visible = True
        Case 1: frDDF(0).Visible = True
    End Select
    
    lblSrcMode(SrcMode).Font.Bold = True
    lblSrcMode(SrcMode).ForeColor = vbWhite
    
End Sub

'---- Make a NEW Disk Image File
Private Sub cmdNewImage_Click(Index As Integer)
    Dim Filename As String, Ext As String, p As Integer
    
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

Private Sub cmdDView_Click(Index As Integer)
    ImageFileView Index
End Sub

'========================
'  Subs for Disk Images
'========================
Private Sub ImageFileView(Index As Integer)
    Dim T As Integer, Tmp As String, Filename As String, Ext As String

    For T = 0 To lstImageFiles(Index).ListCount - 1
        If lstImageFiles(Index).Selected(T) = True Then
            Tmp = lstImageFiles(Index).List(T)
            Filename = CBMName(Tmp)                     'FILENAME,p
            Ext = CBMType(Filename)                     'Get the ",p" or ",s" etc
            
            KillFile TEMPFILE3
        
            DoCommand CBMC1541, _
                      DDFile(Index) & " -read " & Quoted(Filename) & " " & Quoted(TEMPFILE3), _
                      "Copying '" & Filename & "' from image..."
                      
            If Exists(TEMPFILE3) = True Then
                frmViewer.ViewIt 0, TEMPFILE3, Filename, Ext
            Else
                MyMsg "Could not extract " & Quoted(Filename) & " from image."
            End If
            Exit For
        End If
    Next T
    
End Sub

'---- Delete a File(s) within Disk Image
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

'---- Validate Disk Image
Private Sub ImageValidate(ByVal Index As Integer)
    If Exists(UnQuoted(DDFile(0))) = False Then MyMsg "Select an image first!": Exit Sub
    If MsgBox("Validate Image?", vbYesNo) <> vbYes Then Exit Sub
    DoCommand CBMC1541, DDFile(Index) & " -validate", "Validating Image..."
    ImageRefresh Index
End Sub

'---- Backup Disk Image
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

'---- Rename File(s) within Disk Image
Private Sub cmdDRename_Click(Index As Integer)
    ImageFileRename Index
End Sub

'---- Rename File(s) within Disk Image
Sub ImageFileRename(ByVal Index As Integer)
    Dim T As Integer, FSel As Integer, Filename As String, OneName As String

    FSel = 0
    For T = 0 To lstImageFiles(Index).ListCount - 1
        If lstImageFiles(Index).Selected(T) = True Then FSel = FSel + 1: OneName = ExtractQuotes(lstImageFiles(Index).List(T))
    Next T
    If FSel = 0 Then Exit Sub
    
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

'---- Run a PRG file with VICE
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

'---- Edit the Disk Image File
Private Sub cmdDskEd_Click(Index As Integer)
    Dim Filename As String, Tmp As String, Ext As String
    
    Filename = lblDDFile(0).Caption                         'The filename
    If Left(Filename, 1) <> "-" Then
        Ext = FileExt(Filename)                             'Extension of disk image
        Tmp = ExeDir & "image-" & Ext & ".txt"              'The control file for this disk image type
        If Exists(Tmp) = True Then
            MyMsg "This is a preview of the Disk Editor." & Cr & "** WORK IN PROGRESS! - WRITING IS DISABLED ** "
            frmDiskEd.Show                                  'Display the form
            Call frmDiskEd.LoadImg(Filename)                'Tell it what filename to edit
        Else
            MyMsg "Sorry, " & UCase(Ext) & " images are not supported at this time!"
        End If
    End If
End Sub

Private Sub cmdDAll_Click(Index As Integer)
    DSelector True, Index
End Sub

Private Sub cmdDNone_Click(Index As Integer)
    DSelector False, Index
End Sub

Private Sub DSelector(ByVal b As Boolean, ByVal Index As Integer)
    Dim j As Integer
    For j = 0 To lstImageFiles(Index).ListCount - 1
        lstImageFiles(Index).Selected(j) = b
    Next j
End Sub

Private Sub cmdImageRefresh_Click(Index As Integer)
    ImageRefresh Index
End Sub

Private Sub ImageRefresh(Index As Integer)
    GetImageDir Index, DDFile(Index)
End Sub

Private Sub cmdLinkDir_Click()
    GetLinkDir
End Sub


'====================
'  Subs For CBMLink
'====================

'---- Pick a new Drive# Unit#
Private Sub cboLinkDev_Click()
    SetLinkString
End Sub

'---- Make CBM-Link Drive Selection string
Public Sub SetLinkString()
    Select Case cboLinkDev.ListIndex
        Case 0: CBMUnit = 8: CBMDrive = 0
        Case 1: CBMUnit = 8: CBMDrive = 1
        Case 2: CBMUnit = 9: CBMDrive = 0
        Case 3: CBMUnit = 9: CBMDrive = 1
        Case 4: CBMUnit = 10: CBMDrive = 0
        Case 5: CBMUnit = 10: CBMDrive = 1
        Case 6: CBMUnit = 11: CBMDrive = 0
        Case 7: CBMUnit = 11: CBMDrive = 1
    End Select
    
    ClearLinkDir
    LinkCStr = "-c " & frmOptions.txtConStr.Text & " -d " & CBMUnit
End Sub

'---- Read CBM-Link Directory
Private Sub GetLinkDir()
    Dim CmdLine As String, temp As String, Temp2 As String, Results As ReturnStringType
    
    On Local Error GoTo GetLinkErr:
    
    ClearLinkDir
    
    'Run the program
    Results = DoCommand(CBMLink, LinkCStr & " -dd $" & Format(CBMDrive) & ":*", "Reading directory, please wait...", False)
    
    Close #1
    Open TEMPFILE1 For Input As #1                          'Read in the complete output file
    If EOF(1) Then Exit Sub                                 'Check for empty file
    
    Line Input #1, temp                                     'First line is dir. name and ID
    lblLinkDiskName.Caption = UCase(ExtractQuotes(temp))    'Extract Diskname
    lblLinkDiskID.Caption = UCase(Right$(temp, 5))          'Extract ID
    lblLinkDiskID.ToolTipText = DiskID(temp)                'Drive/dos

    Line Input #1, temp                                     'Get next line
    
    While (Not EOF(1))
        lblLinkBlocksFree.Caption = temp
        Line Input #1, temp                                 'Get next line
    Wend
    
    lblLinkBlocksFree.Caption = UCase(temp)                 'The drive status is taken from the last line on stdout
    Close #1
    Exit Sub
    
GetLinkErr:
    If Not (Err.Number = 53) Then
        MyMsg "GetLink Error: " & Err.Number & Cr & "[" & temp & "]"
        ClearLinkDir
    End If
    Exit Sub
    
End Sub

'Format CBM Link-Drive
Private Sub cmdLinkFormat_Click()
    Dim Status As ReturnStringType
    
    If MsgBox("This will erase ALL data on the floppy disk.  Are you sure?", vbExclamation Or vbYesNo, "Format Disk") = vbNo Then Exit Sub
    
    frmPrompt.Reply.Text = "new,id"                                            'Provide a default string
    frmPrompt.Ask "Format CBM Floppy", "Please Enter Diskname, ID", 1, False   'Get Disk name and ID
    If Response = "" Then Exit Sub                                          'Exit if null response

    Status = DoCommand(CBMLink, LinkCStr & " -dc N" & Format(CBMDrive) & ":" & UCase(Response), "Formatting floppy disk, please wait.")
    lblLinkLastStatus.Caption = UCase(Status.Output)
    Sleep 1000                                                              'Just so message is visible

End Sub

'---- Initialize CBM-Link Drive
Private Sub cmdLinkInit_Click()
    DoCommand CBMLink, LinkCStr & " -dc I" & Format(CBMDrive), "Initializing Drive..."
    GetLinkDir
End Sub

'---- Validate CBM-Link Drive
Private Sub cmdLinkValidate_Click()
    DoCommand CBMLink, LinkCStr & " -dc V" & Format(CBMDrive), "Validating Drive..."
    GetLinkDir
End Sub

'---- Rename CBM-Link file(s)
Private Sub cmdLinkRename_Click()
   Dim T As Integer, Filename As String

    For T = 0 To lstLink.ListCount - 1
        If (lstLink.Selected(T)) Then
            Filename = ExtractQuotes(lstLink.List(T))
            frmPrompt.Reply.Text = Filename
            frmPrompt.Ask "Rename CBM File", "Enter new name for '" & Filename & "'", 1, False
            
            If Response Then
                DoCommand CBMLink, _
                          LinkCStr & " -dc R" & Format(CBMDrive) & ":" & Response & "=" & Filename, _
                          "Renaming file"
            Else
                Exit Sub
            End If
        End If
    Next T
    
   GetLinkDir
End Sub

'---- Reset CBM-Link Drive
Private Sub cmdLinkReset_Click()
    DoCommand CBMLink, _
              LinkCStr & " -dc UJ", _
              "Resetting drives, please wait."
End Sub

'---- Scratch CBM-Link file(s)
Private Sub cmdLinkScratch_Click()
    Dim T As Integer, Filename As String
    
    For T = 0 To lstXFiles.ListCount - 1
        If (lstXFiles.Selected(T)) Then
            Filename = ExtractQuotes(lstXFiles.List(T))
            DoCommand CBMCtrl, _
                      "LinkStr &  -dc S" & Format(CBMDrive) & ":" & Filename & Qu, _
                      "Scratching " & Filename
        End If
    Next T

    GetLinkDir
End Sub

'---- Get CBM-Link Drive Status
Private Sub cmdLinkStatus_Click()
    Dim Results As ReturnStringType

    Results = DoCommand(CBMLink, LinkCStr & " -ds", "Reading drive status, please wait.")
    lblLinkLastStatus.Caption = Results.Output
End Sub

'---- Select ALL entries in CBMLink list
Private Sub cmdLinkAll_Click()
    LSelector (True)
End Sub

'---- De-Select ALL entries in CBMLink list
Private Sub cmdLinkNone_Click()
    LSelector (False)
End Sub

'---- Select or De-Select ALLall Entries in CBMLink list
Private Sub LSelector(ByVal b As Boolean)
    Dim j As Integer
    
    For j = 0 To lstLink.ListCount - 1
      lstLink.Selected(j) = b
    Next j
End Sub

'====================
'  Subs for X-Cable
'====================

'---- Fetch the drive status strings.
Private Sub cmdXDriveStatus_Click()
    GetDriveStatus
End Sub

'---- Get Drive Status
Private Sub GetDriveStatus()
    lblXLastStatus.Caption = GetXStatus()
End Sub

'---- Get X-cable Status String
Private Function GetXStatus() As String
    Dim Status As ReturnStringType
    
    Status = DoCommand(CBMCtrl, "status " & DriveNum, "Reading drive status, please wait.")
    GetXStatus = UCase(Status.Output)
End Function

'---- Get X-cable Status Value
Private Function GetXStatusN() As Integer
    GetXStatusN = Val(Left(GetXStatus, 2))
End Function

'---- Detect Drives via CBMCTRL
' CBMCTRL returns a list of drives with device#'s like this:
' 8:1541<cr>
' 9:1571<cr>
' Flag=true to display results, False=silent
Public Sub DetectDrives(ByVal Flag As Boolean)
    Dim What As String, FIO As Integer, D As Integer
    
    If Exists(ExeDir & "cbmctrl.exe") = False Then Exit Sub
    
    frmMain.PubDoCommand CBMCtrl, "detect", "Detecting Drives...", False

    If Exists(TEMPFILE1) = False Then Exit Sub
    
    '-- Read in the complete output file
    FIO = FreeFile
    Open TEMPFILE1 For Input As FIO
        What = Input(LOF(FIO), FIO) 'was LOF(1)???
    Close FIO
    
    '-- Read returned string for drives - NOTE: May need to adjust this for drives with parallel cable
    For D = 8 To 30: Drive(D) = MyTrim(GetNamedField(What, Format(D) & ":")): Next D
   
    KillTemp                    'And delete both temp files, so we're not cluttering things up
    
    If Flag = True Then
        If (What = "") Then What = "No drives found, please check CBM-Transfer directory paths!"
        MsgBox What, vbOKOnly, "Drive Detection"
    End If
End Sub

'---- Format a floppy on X-Cable or ZoomFloppy
' Checks which drive is being used to determine proper formatting method (opencbm command or via dos command string)
' For 1571 will ask for double-sided format
Private Sub cmdXFormat_Click()
    Dim Status As ReturnStringType, Tmp As String, Flag As Boolean, M1 As String
    
    If MsgBox("This will erase ALL data on the floppy disk! Are you sure?", vbExclamation Or vbYesNo, "Format Disk") = vbNo Then Exit Sub
    
    GetXDevices
    
    frmPrompt.Reply.Text = "new,id"
    frmPrompt.Ask "Format CBM Floppy", "Please Enter Diskname, ID", 1, False
    If Response = "" Then Exit Sub
    
    Tmp = UCase(Drive(DriveNum))    'Drive Type
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
        Status = DoCommand("cbmforng", " -vso " & DriveNum & " " & Quoted(UCase(Response)), M1)
        lblXLastStatus.Caption = UCase(Status.Output)
        Sleep 1000                                                          'Just so message is visible
        GetXDir
    Else
        '-- Format using standard DOS New command
        Status = DoCommand(CBMCtrl, CMDSTR & DriveNum & " " & Quoted("N0:" & UCase(Response)), M1)
        MyMsg "Formatting... You may continue working, however, do not attempt to access" & Cr & "the drive until formatting is complete. Check drive status when the light goes off."
    End If
        
End Sub

'---- Initialize X-cable or Zoomfloppy Drive
Private Sub cmdXInit_Click()
    InitXDrive
End Sub

'---- Initialize X-cable or Zoomfloppy Drive
Private Sub InitXDrive()
    DoCommand CBMCtrl, CMDSTR & DriveNum & " I0", "Initializing Drive"
    cmdXDriveStatus_Click
End Sub

Private Sub cmdXView_Click()
    XView
End Sub

'---- View a file from inside a Disk Image
Private Sub XView()
    Dim T As Integer, Filename As String, Ext As String
          
    For T = 0 To lstXFiles.ListCount - 1
        If (lstXFiles.Selected(T)) Then
            
            Filename = LCase(ExtractQuotes(lstXFiles.List(T)))
            Ext = DOSExt(lstXFiles.List(T))                     'Get CBM Filetype (extension)
                        
            Select Case UCase(Ext)
                Case "PRG", "SEQ"
                    DoCommand CBMCopy, _
                              "--transfer=" & TransferString & " -q -r " & DriveNum & " " & Quoted(Filename) & _
                              " --output=" & Quoted(TEMPFILE3), _
                              "Reading '" & Filename & "' ..."
                    frmViewer.ViewIt 0, TEMPFILE3, Filename, Ext
                    
                Case "CBM" '1581 Partition
                    XChangePart Filename
                    GetXDir
                    
                Case Else
                    MyMsg "Sorry, can only View PRG or SEQ files!"
            End Select
            Exit For
        End If
    Next T
    
End Sub

'---- Change to specified partition "file" on 1581 drive
Private Sub XChangePart(ByVal Filename As String)
    Dim Tmp As String
    
    DoCommand CBMCtrl, CMDSTR & DriveNum & " " & Quoted("/0:" & UCase(Filename)), "Changing partition"
    Tmp = GetXStatus()
    If Left(Tmp, 2) = "77" Then MyMsg "Partition is Illegal!"
    
End Sub

'---- Read Directory from X Cable drive
Private Sub cmdXRefresh_Click()
    ClearXDir
    GetXDevices
    GetXDir
End Sub

'---- Find out what drives are connected and return string for selected device#
Private Sub GetXDevices()
    DetectDrives False                      'Get all drives
    lblDName.Caption = Drive(DriveNum)      'Show the drive string for device#
End Sub

'---- Read the X-Cable Directory, parse it, and fill the file list
Public Sub GetXDir()
    Dim CmdLine As String, temp As String, Temp2 As String, Results As ReturnStringType
    
    On Local Error GoTo GetXErr
    
    lstXFiles.Clear
       
    lblXPart.Visible = False: cmdXPart.Visible = False:  cmdXRoot.Visible = False               'Hide 1581 controls
    lblXDiskName.Caption = "": lblXDiskID.Caption = "":  lblXDiskID.ToolTipText = ""            'Clear old fields
    frmWaiting.SetMode ""                                                                       'No progress bar
    
    Results = DoCommand(CBMCtrl, "dir " & DriveNum, "Reading directory, please wait.", False) 'Run the program
    
    Close #1                                                'Make sure File#1 is closed so it can be opened below
                                                            '(seems it sometimes doesn't close properly below)
    
    If Exists(TEMPFILE1) = False Then Exit Sub
    
    Open TEMPFILE1 For Input As #1                          'Read in the complete output file
    If EOF(1) Then Close #1: Exit Sub                       'Check for empty file
    
    Do
        Line Input #1, temp                                 'First line is dir. name and ID
    Loop While Left(temp, 7) = "GetProc"                    'Filter out occasional wayward status messages
    
    lblXDiskName.Caption = ExtractQuotes(temp)              'Set the Disk Name
    lblXDiskID.Caption = Right$(temp, 5)                    'Set the Disk ID
    lblXDiskID.ToolTipText = DiskID(temp)                   'DiskID tells you DOS version and Disk Format 2A=1541, 3D=1581

    If UCase(Right(temp, 2)) = "3D" Then
        lblXPart.Visible = True: cmdXPart.Visible = True: cmdXRoot.Visible = True   'Enable 1581 buttons for partitions
    End If

    If (Not EOF(1)) Then Line Input #1, temp
    lblXBlocksFree.Caption = temp
    
    While (Not EOF(1))
        If Temp2 <> "" Then lstXFiles.AddItem Temp2         'Add the directory line
        Temp2 = temp                                        'Remember it
        Line Input #1, temp                                 'Get the next line
    Wend
    Close #1
    
    lblXLastStatus.Caption = UCase(temp)                    'The drive status is taken from the last line on stdout
    lblXBlocksFree = Temp2                                  'Blocks free from second last line
    
    Exit Sub
    
GetXErr:
    If Not (Err.Number = 53) Then
        MyMsg "GetX Error: " & Err.Number & Cr & "[" & temp & "]"
        ClearXDir
    End If
    Exit Sub
End Sub

'---- Rename files in X-cable or Zoomfloppy File List
Private Sub cmdXRename_Click()
   Dim T As Integer, Filename As String, NewFilename As String

    For T = 0 To lstXFiles.ListCount - 1
        If (lstXFiles.Selected(T)) Then
            Filename = ExtractQuotes(lstXFiles.List(T))
            
            frmPrompt.Reply.Text = ExtractQuotes(lstXFiles.List(T))
            frmPrompt.Ask "Rename CBM File", "Enter new name for '" & Filename & "'", 1, False
            
            NewFilename = UCase(Response)
            
            If Response <> "" Then
                DoCommand CBMCtrl, _
                          CMDSTR & DriveNum & " " & Quoted("R0:" & NewFilename & "=" & UCase(Filename)), _
                          "Renaming"
            Else
                Exit Sub
            End If
        End If
    Next T
    
   GetXDir
End Sub

'---- Reset X-cable or Zoomfloppy Drive
Private Sub cmdXReset_Click()
     DoCommand CBMCtrl, "reset", "Resetting drives, please wait."
End Sub

'---- Delete (Scratch) a file on X-cable or Zoomfloppy
Private Sub cmdXScratch_Click()
    Dim T As Integer, Filename As String, FSel As Integer, OneName As String
    
    '-- Count how many files are selected
    For T = 0 To lstXFiles.ListCount - 1
        Filename = CBMName(lstXFiles.List(T))
        If (lstXFiles.Selected(T)) Then FSel = FSel + 1: OneName = Filename
    Next T
    If FSel = 0 Then Exit Sub
    
    '-- Prompt
    If FSel = 1 Then Filename = Quoted(OneName) Else Filename = Str(FSel) & " file(s)"  'Single filename or number of files
    If MsgBox("Are you sure you want to delete " & Filename & "?", vbYesNo, "Confirm delete") <> vbYes Then Exit Sub
    
    '-- Delete selected files

    For T = 0 To lstXFiles.ListCount - 1
        If (lstXFiles.Selected(T)) Then
            Filename = CBMName(lstXFiles.List(T))
            
            DoCommand CBMCtrl, _
                      CMDSTR & DriveNum & " " & Quoted("S0:" & UCase(Filename)), _
                      "Scratching " & Filename
        End If
    Next T

    GetXDir
End Sub

'---- Do a Disk Validation on X-Cable or Zoomfloppy Drive
Private Sub cmdXValidate_Click()
     DoCommand CBMCtrl, _
               CMDSTR & DriveNum & " " & Quoted("V0:"), _
               "Validating drive, please wait."
End Sub

'---- Return to ROOT of 1581 Disk directory
Private Sub cmdXRoot_Click()
     Dim Tmp As String
     
     DoCommand CBMCtrl, _
               CMDSTR & DriveNum & " " & Quoted("/"), _
               "Selecting Root Partition, please wait..."
            
     Tmp = GetXStatus()
     If Left(Tmp, 2) = "77" Then MyMsg "Could not select partition!"
     
     GetXDir
End Sub

Private Sub cmdXPart_Click()
    XView
End Sub

Private Sub cmdXAll_Click()
    Selector (True)
End Sub

Private Sub cmdXNone_Click()
    Selector (False)
End Sub

Private Sub cboXDevNum_Click()
    DriveNum = cboXDevNum.ListIndex + 8
    ClearXDir
    GetXDevices
End Sub

'---- Select ALL or NONE for all files in X-cable or Zoomfloppy disk
Private Sub Selector(ByVal b As Boolean)
    Dim j As Integer
    
    For j = 0 To lstXFiles.ListCount - 1
      lstXFiles.Selected(j) = b             'Set to desired state
    Next j
End Sub

'============================
'  Subs for Copy Operations
'============================

'---- Copy "-->" LEFT to RIGHT; Figure out what type of copy
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

'---- Copy "<--" RIGHT to LEFT; Figure out what type of copy
Private Sub cmdCopyLeft_Click()
    
    If SrcMode = 0 Then
        '-- Source Files showing on left
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

'---- Copy LocalPC to LocalPC
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

'---- Copy Image to X-Cable
Private Sub Copy_ImgToX()
    Dim T As Integer, Filename As String, FilenameOut As String, Ext As String, SeqType As String, Tmp As String

    For T = 0 To lstImageFiles(0).ListCount - 1
        If lstImageFiles(0).Selected(T) = True Then
            Tmp = lstImageFiles(0).List(T)
            Filename = LCase(ExtractQuotes(Tmp)): Ext = DOSExt(Tmp)
            FilenameOut = MakePCName(Filename & "." & Ext)
            
            If FilenameOut <> "" Then
                SeqType = "": If Ext = "SEQ" Then SeqType = " --file-type S"
                            
                KillFile TEMPFILE3
                
                DoCommand CBMC1541, _
                          DDFile(0) & " -read " & Quoted(Filename) & " " & Quoted(TEMPFILE3), _
                          "Copying '" & Filename & "' from image..."
                          
                DoCommand CBMCopy, _
                          "--transfer=" & TransferString & " -q -w " & DriveNum & " " & Quoted(TEMPFILE3) & _
                          " --output=" & Quoted(Filename) & SeqType, _
                          "Copying file to floppy disk as '" & Filename & "'..."
            End If
        End If
    Next T
    
    'KillTemp
    GetXDir
    
End Sub

'---- Copy Local to Image
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
                        If Flag = False Then Flag = True: MyMsg "You can only write PRG,SEQ,ROM,BIN,'P00','S00' or files with NO extension INTO images!" 'Warn Once!
                    End If
            End Select
        End If
    Next T

    ImageRefresh DstImage
    
End Sub

Private Sub Copy_XToImg()
    Dim T As Integer, Filename As String
         
    For T = 0 To lstXFiles.ListCount - 1
        If (lstXFiles.Selected(T)) Then
            Filename = CBMName(lstXFiles.List(T))      'FILENAME,P (needed for source and destination)
                        
            KillFile TEMPFILE3                            'delete temp file first
            
            '-- Copy from X to TEMPFILE
            DoCommand CBMCopy, _
                      "--transfer=" & TransferString & " -q -r " & DriveNum & " " & Quoted(Filename) & " --output=" & Quoted(TEMPFILE3), _
                      "Copying '" & Filename & "' from floppy disk."
                  
            '-- Copy TEMPFILE to Image
            DoCommand CBMC1541, _
                        DDFile(0) & " -write " & Quoted(TEMPFILE3) & " " & Quoted(Filename), _
                        "Copying '" & Filename & "' to image..."
        End If
    Next T
    
    ImageRefresh 0
End Sub

'---- Copy from CBMLink to Disk Image
' Intermediate file is stored in EXE directory
Private Sub Copy_LinkToImg()
    Dim i As Integer, Tmp As String, Filename As String, Ext As String, Filename2 As String, FilenameOut As String
          
    For i = 0 To lstLink.ListCount - 1
        If (lstLink.Selected(i)) Then
            Tmp = lstLink.List(i)
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

Private Sub Copy_ImgToLink()
    MyMsg "Sorry, IMG to Link Not available!"
End Sub

Private Sub Copy_LinkToLocal(Index As Integer)
    Dim T As Integer, FilesSelected As Integer, Filename As String, FExt As String, FExt2 As String, FilenameOut As String
    Dim Tmp As String ', Response As ReturnStringType
    
    FilesSelected = 0
          
    For T = 0 To lstLink.ListCount - 1
        If (lstLink.Selected(T)) Then
            Tmp = lstLink.List(T)                                           'Get list entry string
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
        
        frmPrompt.Reply.Text = RTrim(lblXDiskName.Caption) & ".D80"
        frmPrompt.Ask "Create Disk Image", "Please Enter Image Filename:", 1, False
        If Response = "" Then Exit Sub
        
        '-- Read DISK Image file. File is written to EXE directory
        DoCommand CBMLink, _
                  LinkCStr & " -dr" & Format(CBMDrive) & " " & Response, _
                  "Creating disk image, please wait..."
                  
        '-- Copy the file to LocalPC destination folder
        If Exists(ExeDir & Response) = True Then
                Name ExeDir & Response As LocalDir(Index) & Response             'Move the file
                lstLocal(Index).Refresh
        End If
    End If

    lstLocal(Index).Refresh

End Sub

'---- Copy from LocalPC to CBMLink
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

'---- Copy file inside Image to LocalPC Folder
' SrcImg: 0=left, 1=right Image
' DstPC : 0=left, 1=right LocalPC folder
' updated for P00 files - Mar 11/2016
Private Sub Copy_ImgToLocal(ByVal SrcImg As Integer, ByVal DstPC As Integer)
    Dim T As Integer, Filename As String, Filename2 As String, Filename3 As String
    Dim FilenameOut As String, Ext As String, Tmp As String
    
    If P00Flag = True Then MyChDir LocalDir(DstPC)          'C1541.Exe writes P00 files to the Dst directory!!
    
    For T = 0 To lstImageFiles(SrcImg).ListCount - 1
        If lstImageFiles(SrcImg).Selected(T) = True Then
            Tmp = lstImageFiles(SrcImg).List(T)
            Filename = CBMName(Tmp)  ' FILENAME,P
            Filename2 = DOSName(Tmp) ' FILENAME.PRG
            
            If P00Flag = True Then
                '-- Write P00 files to dest. P00 filename will be created automatically in CURRENT directory (hence the CD command above)
                '   BUG!: C1541.EXE always seems to write "p" files, even with SEQ source file, and then
                '         writing it back to an image looses the SEQ and instead creates a PRG file!
                DoCommand CBMC1541, _
                          DDFile(SrcImg) & " -p00save 1 -read " & Quoted(Filename), _
                          "Copying '" & Filename & "' from image..."
            Else
                FilenameOut = LocalDir(DstPC) & MakePCName(Filename2)
                '-- Write normal files to dest
                DoCommand CBMC1541, _
                          DDFile(SrcImg) & " -read " & Quoted(Filename) & " " & Quoted(FilenameOut), _
                          "Copying '" & Filename & "' from image..."
                    
                If Exists(FilenameOut) = False Then
                    MsgBox "Problem! File does not exist at destination!" & Cr & _
                        "It appears that the source file could not be extracted from the image.", vbExclamation, "Warning!"
                End If
            End If

        End If
    Next T
    
    MyChDir ExeDir              'Change back to EXE directory
    lstLocal(DstPC).Refresh
    
End Sub

'---- Copy from Image to Image
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
            
            If Exists(TEMPFILE3) = False Then
                MsgBox "Problem! The source file could not be extracted from the image.", vbExclamation, "Warning!"
            Else
                '-- File was copied, to temp dir, so now copy it to dest image
                DoCommand CBMC1541, _
                          DDFile(DstImg) & " -write " & Quoted(TEMPFILE3) & " " & Quoted(Filename), _
                          "Copying '" & Filename & "' to image..."
            End If
        End If
    Next T
    
    KillFile TEMPFILE3
    ImageRefresh DstImg
    
End Sub

'---- Copy from X-cable or Zoomfloppy to LocalPC
Private Sub Copy_XToLocal()
    Dim T As Integer, FilesSelected As Integer, Filename As String, FilenameOut As String, Tmp As String

    If lblXDiskID.Caption = "" Then GetXDir                             'Added for batch
    FilesSelected = 0
          
    '-- Copy Selected, and count them
    For T = 0 To lstXFiles.ListCount - 1
        If (lstXFiles.Selected(T)) Then
            Tmp = lstXFiles.List(T)
            Filename = CBMName(Tmp)                                  'ie: FILENAME,P
            FilenameOut = LocalDir(0) & DOSName(Tmp)                 'Output filename PATH\FILENAME.EXT
            If Overwrite(FilenameOut) = True Then
                DoCommand CBMCopy, _
                      "--transfer=" & TransferString & " -q -r " & DriveNum & " " & Quoted(Filename) & " --output=" & Quoted(FilenameOut), _
                      "Copying '" & Filename & "' from floppy disk."
            End If
            FilesSelected = FilesSelected + 1
        End If
    Next T
    
    '-- If no files were selected then make an image
    If (FilesSelected = 0) Then
        If ConfirmD64 = True Then
            If MsgBox("No files selected.  Do you want to make an image of this floppy disk?", vbQuestion Or vbYesNo, "Create an Image") = vbNo Then Exit Sub
        End If
    
        If UseBatch = True Then
            frmBatch.Show
        Else
            MakeXDiskImage      'No Files were selected, so image the disk D64/G64/NIB etc.
        End If
    End If
    
    lstLocal(0).Refresh
    
End Sub

'---- Copy selected files or Disk Images to X-cable
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
                
                Select Case Ext
                    Case "D64", "D71"   'Make Disk from D64 or D71
                        If (UseNIB = True) And (WriteD64 = True) Then
                            WriteNIBtoX FileOut, ImgFlag
                        Else
                            WriteImageToX FileOut, ImgFlag
                        End If
                    Case "NIB", "NBZ", "G64"
                        WriteNIBtoX FileOut, ImgFlag
                    Case "D80", "D81", "D82"
                        WriteImageToX FileOut, ImgFlag
                End Select
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

'---- Create a Disk Image from X-Cable/Zoomfloppy Disk
Public Sub MakeXDiskImage()
    Dim Filename As String, Ext As String, FilenameOut As String, Ostr As String
    Dim Tmp As String, TmpP As VbMsgBoxResult, TmpExt As String, TempNIB As Boolean
    Dim X0 As String, X1 As String, X2 As String, NibTmp As String
    
    X0 = ""
    X1 = "d64copy"
    X2 = "imgcopy"
    TempNIB = False
    
    '-- Check Disk Format using DiskID string
    Tmp = UCase(Mid(lblXDiskID, 4, 2))
    Select Case Tmp
        Case "2A": TmpExt = "D64": X0 = X1
        Case "2C": TmpExt = "D80": X0 = X2
        Case "3D", "1D": TmpExt = "D81": X0 = X2
        Case Else
            TmpExt = "D64": X0 = X1     'default to D64 using D64COPY

            If (UseNIB = False) And (IgnoreBadID = False) Then
                TmpP = MsgBox("The source disk ID (" & Tmp & ") is unknown. This could be a corrupt disk, copy-protected disk, or unsupported format." & Cr & _
                "Do you want to try imaging with NIBTOOLS?" & Cr & "( Yes=NIBTOOLS, No=D64COPY, Cancel=Do Not Image )", vbYesNoCancel, "Warning!")
                Select Case TmpP
                    Case vbYes: TempNIB = True
                    'Case vbNo: TmpExt = "D64": X0 = X1
                    Case vbCancel: Exit Sub
                End Select
            End If
    End Select
    
    '-- Create image of disk using D64COPY.EXE or IMGCOPY
    If (UseNIB = False) And (TempNIB = False) Then
        If UseBatch = False Then
            frmPrompt.Reply.Text = RTrim(FixPCName(lblXDiskName.Caption, "")) & "." & TmpExt
            frmPrompt.Ask "Create Dxx", "Please Enter Image Filename:", 1, False
            If Response = "" Then Exit Sub
            FilenameOut = LocalDir(0) & Response
        Else
            FilenameOut = LocalDir(0) & BatchFilename
        End If
        
        Ostr = "": Ext = FileExtU(Response): Tmp = X0
        
        Select Case Ext
            Case "D64": Ostr = ""                               'Check for 1541 image
            Case "D71": Ostr = "-2 "                            'Check for 1571 image and add option string
            Case "D80": Ostr = "-d8050 --error-map=never "      'Check for 8050 image
            Case "D81": Ostr = "-d1581 --error-map=never "      'Check for 1581 image and add option string
            Case "D82": Ostr = "-d1001 -2 --error-map=never "   'Check for 8250/SFD image and add option string
        End Select
        
        If Overwrite(FilenameOut) = True Then
            KillFile FilenameOut
            frmWaiting.SetMode Ext
            DoCommand Tmp, _
                      Ostr & "--transfer=" & TransferString & " " & NoWarpString & " " & Format(DriveNum) & " " & Quoted(FilenameOut), _
                      "Creating " & Ext & " image, please wait."
        End If
        
    Else
        '-- Create image of disk using NIBREAD.EXE
        If UseBatch = False Then
            frmPrompt.Reply.Text = RTrim(lblXDiskName.Caption)
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
            NibTmp = NIBstr: If UseNibCustom = True Then NibTmp = frmOptions.txtNibRead.Text    'Std or Custom NIB options?
            
            DoCommand "nibread", "-D" & Format(DriveNum) & " " & NibTmp & FilenameOut, _
                      "Creating " & Ext & " file, please wait."
                    
            If Exists(Filename & Ext) = True Then
                NibTmp = NIBstr: If UseNibCustom = True Then NibTmp = frmOptions.txtNibConv.Text    'Std or Custom NIB options?
                
                frmWaiting.SetMode "nibconv"
                '-- Convert NIB to G64
                If CreateG64 = True Then
                    DoCommand "nibconv", NibTmp & " " & FilenameOut & " " & Quoted(Filename & ".g64"), _
                              "Converting " & Ext & " to G64"
                End If
                
                '-- Convert NIB to D64
                If CreateD64 = True Then
                    DoCommand "nibconv", NibTmp & " " & FilenameOut & " " & Quoted(Filename & ".d64"), _
                              "Converting " & Ext & " to D64"
                End If
                
                If CreateNIB = False Then KillFile FilenameOut
            Else
                MsgBox "Problem! The file '" & FilenameOut & "' was not created!", vbExclamation, "Warning!"
            End If
        End If
    End If
    
End Sub

'---- Transfer specified file to X-Cable or Zoomfloppy disk
' Filename should have path included and be PRG or SEQ extension
Private Sub TransferToX(ByVal Filename As String)
    Dim FilenameOut As String, Ext As String, SeqType As String
    
    FilenameOut = FileNameOnly(Filename): Ext = FileExtU(FilenameOut)
        
    Select Case Ext
        Case "PRG": FilenameOut = FileBase(FilenameOut): SeqType = ""
        Case "SEQ": FilenameOut = FileBase(FilenameOut): SeqType = " --file-type S"
    End Select
    
    DoCommand CBMCopy, _
              "--transfer=" & TransferString & " -q -w " & DriveNum & " " & Quoted(Filename) & _
              " --output=" & Quoted(FilenameOut) & SeqType, _
              "Copying '" & Filename & "' to floppy disk as '" & FilenameOut & "'"
'NOTES: CBMCOPY will display any error messages in it's output, which clears the DISK STATUS

End Sub

'---- Transfer a file to CBMLink device
Private Sub TransferToLink(ByVal Filename As String)
    Dim FilenameOut As String, FPath As String

    FPath = FilePath(Filename)
    FilenameOut = FileNameOnly(Filename)
        
    MyChDir FPath
    
    DoCommand CBMLink, _
              LinkCStr & " -fw " & FilenameOut, _
              "Copying '" & Filename & "' via link as '" & UCase(FilenameOut) & "'"

End Sub

'---- Write Disk Image (D64, D71, D80 etc) to X-cable using D64copy or ImgCopy
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
    
    DoCommand Tmp, _
              Opt & "--transfer=" & TransferString & " " & NoWarpString & " " & Quoted(Filename) & " " & Format(DriveNum), _
              "Creating disk from " & Ext & " image, please wait."
            
    GetXDir
End Sub

Public Sub WriteNIBtoX(ByVal Filename As String, ByVal NoWarn As Boolean)
    Dim Ext As String, NibTmp As String
    
    If UseNIB = False Then MyMsg "You must select the NIB option to write this file to disk.": Exit Sub
    If NoWarn = False Then If MsgBox("This will overwrite ALL data on the floppy disk! Are you sure?", vbExclamation Or vbYesNo, "Write " & Ext & " to Disk") = vbNo Then Exit Sub
        
    Ext = FileExtU(Filename)
    
    frmWaiting.SetMode "nib"
    
    NibTmp = NIBstr: If UseNibCustom = True Then NibTmp = frmOptions.txtNibWrite.Text    'Std or Custom NIB options?
    DoCommand "nibwrite", _
              " -D" & Format(DriveNum) & " " & NibTmp & " " & Quoted(Filename), _
              "Creating disk from " & Ext & " image, please wait."
    
    GetXDir
End Sub

'---- Write Disk Image to CBMLink
' Example usage: CBMLINK.EXE -c serial 19200,com1 -d 8 -dw0 image.d80
Public Sub WriteImageToLink(d64file As String, ByVal NoWarn As Boolean)
    Dim Ext As String
    
    Ext = FileExtU(d64file)
    
    If NoWarn = False Then If MsgBox("This will overwrite ALL data on Destination unit!" & Cr & "(disk must already be formatted!)" & Cr & " Are you sure?", vbExclamation Or vbYesNo, "Write " & Right(Ext, 3) & " to Disk") = vbNo Then Exit Sub
    
    frmWaiting.SetMode CBMLink
    
    DoCommand CBMLink, LinkCStr & " -dw" & Format(CBMDrive) & " " & LocalDir(0) & d64file, _
              "Writing " & Ext & " image to drive, please wait..."
    
    GetXDir
End Sub

'---- Log Commandline string
Public Sub LogIt(ByVal Tmp As String)
    Dim FIO As Integer
    
    FIO = FreeFile
    Open LogFile For Append As FIO
    Print #FIO, Tmp
    Close FIO
End Sub

'---- Refresh Local Directory list
Private Sub cmdSrcRefresh_Click(Index As Integer)
    lstLocal(Index).Refresh
End Sub

'---- Delete file from LocalPC
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
    
    '-- Delete them
    For T = 0 To lstLocal(Index).ListCount - 1
        If (lstLocal(Index).Selected(T)) Then KillFile LocalDir(Index) & lstLocal(Index).List(T)
    Next T
    
    lstLocal(Index).Refresh
End Sub

'---- Handle automatic viewing of Disk Image when LocalPC entry is selected
' The LEFT disk image must be visible, and the RIGHT disk image must NOT
Private Sub lstLocal_Click(Index As Integer)
    Dim Filename As String, Ext As String, p As Integer
    
    If Layout = 1 Then
        p = lstLocal(Index).ListIndex
        Filename = LocalDir(Index) & lstLocal(Index).List(p)
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

'---- Calculate Size of Selected file in BLOCKS
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
    
    KBText(Index).Text = Format(Bytes / 1024, "0.0")
    BlockText(Index).Text = Format(Bytes / 254, "0") '254 Bytes per C= Block
                    
    If Flag = True Then lstLocal(Index).Refresh    'We found a missing file, so refresh the list
End Sub

'---- Refresh File list on LocalPC
Private Sub cmdLocalRefresh_Click(Index As Integer)
    lstLocal(Index).Refresh
End Sub

'---- Rename files on LocalPC
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

'---- Search File List for Image or File to Run via VICE
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

'---- Run a single PRG file in VICE
' Runs a file from local PC
Public Sub RunVicePRG(ByVal Filename As String)
    Dim j As Integer, LA As Long
    
    '---Check PRG option (specified or by load address)
    j = frmOptions.cboPRG.ListIndex 'Selected EMU for PRG files
    If j = 0 Then Exit Sub
    
    If frmOptions.OptPRGMode(1).value = True Then
        '-- Use program Load Address to select emulator.
        '   Note: VIC-20 and TED can have same load address. TODO: Allow selection when multiple choices
        LA = GetLoadAddress(Filename)
        j = GetMachine(LA)
        
        If j < 2 Then
            frmViceSelect.Show vbModal
            j = frmViceSelect.EmuNum
        End If
    End If
    
    If j > 0 Then RunVice j, "", Filename

End Sub

'---- Run VICE with specified Emulator, Disk Image and Filename
Public Sub RunVice(ByVal Emu As Integer, ByVal DName As String, FName As String)
    Dim Tmp As String, VPath As String
    
    If (UseVice = False) Or (Emu = 0) Then Exit Sub
    
    If Emu = 1 Then
        frmViceSelect.Show vbModal  'Ask for emulator here
        Emu = frmViceSelect.EmuNum  'Selected emulation
        If Emu = 0 Then Exit Sub    'No selection
    End If
    
    VPath = VicePath & ViceEXE(Emu) & ".exe"
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

'---- Handle View Buttons
Private Sub cmdSrcView_Click(Index As Integer)
    CheckSelected Index, 0
End Sub
Private Sub cmdSrcView2_Click(Index As Integer)
    CheckSelected Index, 1
End Sub

'---- Search File List for selected File, figure out best way to view it
Private Sub CheckSelected(Index As Integer, Target As Integer)
  Dim T As Integer, Filename As String, Ext As String, FilenameOut As String, V As Integer, FLen As Long
  Dim NibTmp As String
  
    For T = 0 To lstLocal(Index).ListCount - 1
        If (lstLocal(Index).Selected(T)) Then
            Filename = LocalDir(Index) & lstLocal(Index).List(T)
            FilenameOut = FileBase(Filename) & ".d64"
            
            Ext = FileExtU(lstLocal(Index).List(T))
            Select Case Ext
                Case "D64", "X64", "G64", "D71", "D80", "D81", "D82", "D2M", "D4M", "DNP"
                    SelectImage Filename, Target
                    
                Case "NIB", "NBZ"
                    If MsgBox("Do you want to convert this " & Ext & " to D64 to view the contents?", vbYesNo, "Convert?") = vbYes Then
                        frmWaiting.SetMode "nibconv"
                        NibTmp = NIBstr: If UseNibCustom = True Then NibTmp = frmOptions.txtNibConv.Text    'Std or Custom NIB options?
                        DoCommand "nibconv", NibTmp & " " & Quoted(Filename) & " " & Quoted(FilenameOut), _
                                  "Converting " & Filename & " to D64"
                        If Exists(FilenameOut) Then SelectImage FilenameOut, Target
                    End If
                    
                Case "", "PRG", "SEQ", "BIN", "ROM"
                    frmViewer.Show
                    frmViewer.ViewIt 0, Filename, Filename, Ext
                    
                Case "ART", "CDU", "KOA", "GEO", "P00", "S00"
                    frmViewer.Show
                    frmViewer.ViewIt 5, Filename, Filename, Ext
                    
                Case Else
                    V = MsgBox("Unknown file type. Open with associated WINDOWS app?" & Cr & "YES=Windows, NO=CBM-Transfer Viewer", vbYesNoCancel, "Unknown File type")
                    Select Case V
                        Case vbYes: ViewFile Filename
                        Case vbNo
                            frmViewer.Show
                            frmViewer.ViewIt 0, Filename, Filename, Ext
                    End Select
            End Select
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
    lstXFiles.Clear
    lblXDiskName.Caption = ""
    lblXDiskID.Caption = ""
    lblXBlocksFree.Caption = ""
    lblXLastStatus.Caption = ""
End Sub

'---- Clear the CBMLink directory listing, because contents have changed
Private Sub ClearLinkDir()
    lstLink.Clear
    lblLinkDiskName.Caption = ""
    lblLinkDiskID.Caption = ""
    lblLinkLastStatus.Caption = ""
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
Private Function FilterString(ByVal n As Integer) As String
    Dim FX As String
    Select Case n
        Case 1: FX = "*.D64;*.D71;*.D80;*.D81;*.D82;*.NIB;*.G64;*.X64;*.D1M;*.D2M;*.D4M"
        Case 2: FX = "*.NIB;*.NBZ;*.G64;*.D64"
        Case 3: FX = "*.D80;*.D82"
        Case 4: FX = "*.PRG"
        Case 5: FX = "*.SEQ"
        Case 6: FX = "*.TXT"
        Case 7: FX = "*.D64"
        Case 8: FX = "*.D71"
        Case 9: FX = "*.D80"
        Case 10: FX = "*.D81"
        Case 11: FX = "*.D82"
        Case 12: FX = "*.G64"
        Case 13: FX = "*.NIB;*.NBZ"
        Case 14: FX = "*.D1M;*.D2M;*.D4M"
        Case 15: FX = "*.BIN;*.ROM;*.ASM-PROJ"
        Case 16: FX = "*.ART;*.CDU;*.GEO;*.KOA"
        Case Else: FX = "*.*"
    End Select

    FilterString = FX

End Function

'---- Set Destination Frame TAB
Sub SetDstFrame()
    Dim a As Integer
    
    frLink.Visible = False
    frSrc(1).Visible = False
    frX.Visible = False
    frDDF(1).Visible = False
    
    For a = 0 To 2
        lblDstMode(a).Font.Bold = False
        lblDstMode(a).ForeColor = vbBlack
    Next a
            
    Select Case DstMode
        Case 0: frX.Visible = True
        Case 1: frLink.Visible = True
        Case 2: frSrc(1).Visible = True
        Case 3: frDDF(1).Visible = True
    End Select
    
    lblDstMode(DstMode).Font.Bold = True
    lblDstMode(DstMode).ForeColor = vbWhite
    DoEvents
    
End Sub

'---- Select and View Disk Image File
Private Sub SelectImage(ByVal Filename As String, Index As Integer)
    DDFile(Index) = Quoted(Filename)                'Remember the Filename
    If Index = 0 Then SrcMode = 1: SetSrcFrame      'Change LEFT view only if VIEW button
    If Index = 1 Then DstMode = 3: SetDstFrame      'Change RIGHT view
    lblDDFile(Index).Caption = Filename             'Set the filename field
    lblDDFile(Index).ToolTipText = DDFile(Index)    'Set tooltip
    GetImageDir Index, DDFile(Index)

End Sub

'---- Read Disk Image Directory
Private Sub GetImageDir(Index As Integer, ByVal Filename As String)
    Dim temp As String, Temp2 As String, Results As ReturnStringType
    Dim p As Integer, PP As Integer

    On Local Error GoTo GIError
             
    Results = DoCommand(CBMC1541, Quoted(Filename) & " -list", "", False) 'Run the program
    If Exists(TEMPFILE1) = False Then Exit Sub
    
    lstImageFiles(Index).Clear
    
    Close 1
    Open TEMPFILE1 For Input As #1
    
    If EOF(1) Then Exit Sub     'Check for empty file
    
    Input #1, temp              'Output is in one long string. Must Parse!...
    Close #1
    
    '-- Throw away extraneous strings containing "GetProc" etc
    PP = 1
    Do
        p = InStr(PP, temp, LF): If p = 0 Then Exit Do
        Temp2 = Mid(temp, PP, p - PP): PP = p + 1
    Loop While Left(Temp2, 1) > "9"
       
    txtImageHeader(Index).Caption = ExtractQuotes(Temp2)
    txtImageID(Index).Caption = Right$(Temp2, 5)
    txtImageID(Index).ToolTipText = DiskID(Temp2)

    '-- Now parse remaining entries
    Do
        p = InStr(PP, temp, LF): If p = 0 Then Exit Do
        Temp2 = Mid(temp, PP, p - PP): PP = p + 1
        If InStr(1, Temp2, "blocks free", vbTextCompare) = 0 Then lstImageFiles(Index).AddItem Temp2 Else Exit Do 'Lowercase
    Loop
    
    DFBlocksFree(Index).Caption = Temp2
    Exit Sub
    
GIError:
    If Not (Err.Number = 53) Then
        MyMsg "GetImage Error: " & Err.Number & Cr & "[" & temp & "]"
    End If
    Exit Sub

End Sub

'---- Prompt to Create a New Folder, then Make it if it doesn't already exist
Private Sub NewFolder(ByVal RootPath As String)
    Dim DirName As String
    
    frmPrompt.Ask "Make Directory", "Enter Directory Name:", 1
    If Response = "" Then Exit Sub                          'Check for null string
    
    DirName = RootPath & Response
    
    If DirExists(DirName) = False Then
        MkDir DirName
    Else
        MyMsg "Can't create! There is a already a Directory called '" & DirName & "'!"
    End If

End Sub

'---- Load Source Path Drop-down list (Path History)
Public Sub LoadHistory()
    Dim FIO As Integer, Tmp As String, LastTmp As String
        
    If Exists(PathFile) = False Then Exit Sub
    
    FIO = FreeFile
    Open PathFile For Input As FIO
    
    txtLocalDir(0).Clear
    LastTmp = ""
    
    While Not EOF(FIO)
        Line Input #FIO, Tmp                'Read the path string
        If Tmp <> LastTmp Then
            txtLocalDir(0).AddItem Tmp
            txtLocalDir(1).AddItem Tmp
            LastTmp = Tmp                   'Add it, unless it's the same as the previous
        End If
    Wend
    Close FIO
    
End Sub

'---- Save Source Path Drop-down list (Path History)
Public Sub SaveHistory()
    Dim FIO As Integer, Tmp As String, a As Integer
        
    KillFile PathFile
    
    FIO = FreeFile
    Open PathFile For Output As FIO
    
    For a = 0 To txtLocalDir(0).ListCount - 1
        Print #FIO, txtLocalDir(0).List(a)
    Next
    Close FIO
    
End Sub

'========================
' DRAG and DROP functions
'========================

'---- Change to Dropped Directory Path
Private Sub txtLocalDir_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Tmp As String
    
    If Data.GetFormat(vbCFFiles) Then
        Dim vFn As Variant
        For Each vFn In Data.Files
            Tmp = PathOnly(vFn): If Tmp <> "" Then SetLocalPath Index, Tmp             'Get path and use it if valid
        Next
    End If

End Sub

'---- Transfer dropped files to directory list
Private Sub lstXFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Data.GetFormat(vbCFFiles) Then
       Dim vFn As Variant
       For Each vFn In Data.Files
         TransferToX (vFn)    'vFn is name of file dropped
       Next vFn
    End If
    RefreshXDir
End Sub

'---- Provide drag and drop feedback to source
Private Sub lstXFiles_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    '0=do not allow drop, 1=inform source that data will be copied
    If Data.GetFormat(vbCFFiles) Then Effect = 1 Else Effect = 0
End Sub

'---- Accept Dropped Image files
Private Sub lstImageFiles_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ext As String, Filename As String
      
    If Data.GetFormat(vbCFFiles) Then
        Dim vFn As Variant
        For Each vFn In Data.Files
            'vFn is name of file dropped
            Filename = (vFn)
            Ext = FileExtU(Filename)
            Select Case Ext
                Case "D64", "X64", "G64", "D71", "D80", "D81", "D82", "D1M", "D2M", "D3M"
                    SelectImage Filename, Index: Exit For
                Case Else: MyMsg "Sorry, only Disk Image files can be dropped!"
            End Select
        Next vFn
    End If
    
End Sub

'---- Provide drag and drop feedback to source
Private Sub lstImageFiles_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    '0=do not allow drop, 1=inform source that data will be copied
    If Data.GetFormat(vbCFFiles) Then Effect = 1 Else Effect = 0
End Sub


'=============================================
' SHELL COMMANDS - The real work is done here!
'=============================================

'This function must be private, because of the return type.
Private Function DoCommand(Action As String, Args As String, WaitMessage As String, Optional DeleteOutFile As Boolean = True) As ReturnStringType
    Dim CmdLine As String, CmdLine2 As String, ErrorString As String, FIO As Integer
    Static InProgress As Boolean
    
    If (InProgress) Then Exit Function
        
    Close
    If UCase(Right(Action, 4)) <> ".EXE" Then Action = Action & ".exe"
    
    If Exists(ExeDir & Action) = False Then MyMsg "A required file is missing! Please copy '" & Action & "' to the CBM-Transfer directory!": Exit Function
      
    If PreviewCheck = True Then
        If MsgBox("Requested command:" & Cr & Cr & Action & " " & Args & Cr & Cr & "OK to continue?", vbYesNo) = vbNo Then Exit Function
    End If
    
    KillTemp 'And delete both temp files, so we're not cluttering things up
    
    '-- Flag that the background process is starting.
    InProgress = True
    
    '-- Display Dialog if WaitMessage is specified
    If WaitMessage <> "" Then
        frmWaiting.Show vbModeless, frmMain
        frmWaiting.Label = WaitMessage
    End If
    
    '-- Build command-line string
    'cmd /c is needed in order to have a shell write to a file (long, complicated explanation)
    '1> redirects stdout to a file, and 2> redirects stderr to a file (Win2K/XP only)
    'All these quotes [chr$(34)] are needed to handle spaces.  So you get: cmd /c ""path\command" args "files""
    
    CmdLine = Qu & ExeDir & Action & Qu & " " & Args
    CmdLine2 = "cmd /c " & Qu & CmdLine & Qu & " 1>" & TEMPFILE1 & " 2>" & TEMPFILE2
 
    If LogAll = True Then LogIt CmdLine2        'Log all commands
    
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
        lblR.ToolTipText = ErrorString          'Set Last Result as Tooltip
    End If
    
    LastCMDError = ErrorString                  'Remember the error results
    InProgress = False: frmWaiting.Hide         'We are done, so clear InProgress Flag and hide the dialog
    
End Function

Public Function PubDoCommand(Action As String, Args As String, WaitMessage As String, Optional DeleteOutFile As Boolean = True) As String
    Dim Returns As ReturnStringType
    
    Returns = DoCommand(Action, Args, WaitMessage, DeleteOutFile)
    PubDoCommand = Returns.Output
End Function

'---- Wait for process to finish
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

