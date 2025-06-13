VERSION 5.00
Begin VB.Form frmDiskEd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Disk Image Editor"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   409
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   989
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frSector 
      Height          =   1185
      Left            =   5310
      TabIndex        =   17
      Top             =   420
      Width           =   9435
      Begin VB.CommandButton cmdBlockOps 
         Caption         =   "Restore"
         Height          =   585
         Index           =   4
         Left            =   8490
         TabIndex        =   32
         Top             =   510
         Width           =   825
      End
      Begin VB.CommandButton cmdBlockOps 
         Caption         =   "Fill"
         Height          =   285
         Index           =   3
         Left            =   7620
         TabIndex        =   31
         Top             =   810
         Width           =   825
      End
      Begin VB.CommandButton cmdBlockOps 
         Caption         =   "Zero"
         Height          =   285
         Index           =   2
         Left            =   7620
         TabIndex        =   30
         Top             =   510
         Width           =   825
      End
      Begin VB.CommandButton cmdBlockOps 
         Caption         =   "Paste"
         Height          =   285
         Index           =   1
         Left            =   8490
         TabIndex        =   29
         Top             =   180
         Width           =   825
      End
      Begin VB.CommandButton cmdBlockOps 
         Caption         =   "Copy"
         Height          =   285
         Index           =   0
         Left            =   7620
         TabIndex        =   28
         Top             =   180
         Width           =   825
      End
      Begin VB.TextBox txtLTrack 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2100
         TabIndex        =   23
         Text            =   "00"
         Top             =   480
         Width           =   405
      End
      Begin VB.TextBox txtLSector 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2100
         TabIndex        =   22
         Text            =   "00"
         Top             =   810
         Width           =   405
      End
      Begin VB.TextBox txtCurSector 
         Height          =   285
         Left            =   870
         TabIndex        =   21
         Text            =   "00"
         Top             =   810
         Width           =   405
      End
      Begin VB.TextBox txtCurTrack 
         Height          =   285
         Left            =   870
         TabIndex        =   20
         Text            =   "00"
         Top             =   480
         Width           =   405
      End
      Begin VB.Label lblNav 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Error Map"
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
         Height          =   465
         Index           =   11
         Left            =   5760
         TabIndex        =   46
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lblNav 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "TT/SS"
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
         Height          =   465
         Index           =   10
         Left            =   5130
         TabIndex        =   45
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lblNav 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "SET"
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
         Height          =   255
         Index           =   9
         Left            =   5130
         TabIndex        =   44
         Top             =   300
         Width           =   585
      End
      Begin VB.Label lblNav 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "TT/SS"
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
         Height          =   465
         Index           =   8
         Left            =   4500
         TabIndex        =   43
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lblNav 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "SET"
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
         Height          =   255
         Index           =   7
         Left            =   4500
         TabIndex        =   42
         Top             =   300
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "QUICK NAV:"
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
         Left            =   2610
         TabIndex        =   41
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label lblNav 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "First BAM"
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
         Height          =   465
         Index           =   6
         Left            =   3870
         TabIndex        =   40
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lblNav 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "First DIR"
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
         Height          =   465
         Index           =   5
         Left            =   3240
         TabIndex        =   39
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lblNav 
         Alignment       =   2  'Center
         BackColor       =   &H00C000C0&
         Caption         =   "TT/SS LINK"
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
         Height          =   465
         Index           =   4
         Left            =   2610
         TabIndex        =   38
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lblNav 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   37
         Top             =   810
         Width           =   225
      End
      Begin VB.Label lblNav 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1590
         TabIndex        =   36
         Top             =   810
         Width           =   225
      End
      Begin VB.Label lblNav 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   1590
         TabIndex        =   35
         Top             =   510
         Width           =   225
      End
      Begin VB.Label lblNav 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   34
         Top             =   510
         Width           =   225
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "LINK:"
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
         Left            =   2010
         TabIndex        =   27
         Top             =   210
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CURRENT:"
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
         Left            =   150
         TabIndex        =   26
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "T:"
         Height          =   195
         Left            =   1935
         TabIndex        =   25
         Top             =   510
         Width           =   150
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "S:"
         Height          =   195
         Left            =   1935
         TabIndex        =   24
         Top             =   840
         Width           =   150
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "SECTOR:"
         Height          =   195
         Left            =   60
         TabIndex        =   19
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "TRACK:"
         Height          =   195
         Left            =   30
         TabIndex        =   18
         Top             =   510
         Width           =   765
      End
   End
   Begin VB.PictureBox picV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   5310
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   609
      TabIndex        =   16
      Top             =   1710
      Width           =   9165
   End
   Begin VB.VScrollBar vsV 
      Height          =   3855
      LargeChange     =   2
      Left            =   14460
      Max             =   20
      TabIndex        =   14
      Top             =   1710
      Width           =   285
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Changes"
      Height          =   285
      Left            =   13260
      TabIndex        =   13
      Top             =   60
      Width           =   1515
   End
   Begin VB.CheckBox cbShAll 
      Caption         =   "ALL"
      Height          =   285
      Left            =   60
      TabIndex        =   9
      Top             =   5940
      Width           =   705
   End
   Begin VB.VScrollBar vsDir 
      Height          =   5145
      Left            =   4890
      Max             =   20
      TabIndex        =   8
      Top             =   420
      Width           =   255
   End
   Begin VB.PictureBox picED 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   5310
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   0
      ToolTipText     =   "Editing Box: ENTER=Done, ESC=Abort, END=Toggle Case"
      Top             =   5610
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.PictureBox picFree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   60
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   7
      Top             =   5640
      Width           =   4815
   End
   Begin VB.PictureBox picID 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3990
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   6
      Top             =   60
      Width           =   525
   End
   Begin VB.PictureBox picHeader 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   60
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   254
      TabIndex        =   5
      Top             =   60
      Width           =   3840
   End
   Begin VB.PictureBox PicFiles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   5145
      Left            =   60
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   4
      Top             =   420
      Width           =   4845
   End
   Begin VB.PictureBox picDOS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4620
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   3
      Top             =   60
      Width           =   525
   End
   Begin VB.PictureBox picFontSet 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2205
      Left            =   5070
      ScaleHeight     =   2175
      ScaleWidth      =   4335
      TabIndex        =   2
      Top             =   6240
      Width           =   4365
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   16860
      Top             =   120
   End
   Begin VB.Label lblKeystroke 
      AutoSize        =   -1  'True
      Caption         =   "Key"
      Height          =   195
      Left            =   14070
      TabIndex        =   48
      Top             =   5700
      Width           =   270
   End
   Begin VB.Label lblKeyD 
      AutoSize        =   -1  'True
      Caption         =   "KeyD"
      Height          =   195
      Left            =   13590
      TabIndex        =   47
      Top             =   5730
      Width           =   390
   End
   Begin VB.Label lblDebug 
      Caption         =   "Label7"
      Height          =   375
      Left            =   7560
      TabIndex        =   33
      Top             =   5670
      Width           =   5025
   End
   Begin VB.Label lblView 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Error Map"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   10260
      TabIndex        =   15
      Tag             =   "&H00FF0000&"
      ToolTipText     =   "Click to View X-Cable directory"
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label lblView 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sector"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   2
      Left            =   8610
      TabIndex        =   12
      Tag             =   "&H00FF0000&"
      ToolTipText     =   "Click to View X-Cable directory"
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label lblView 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Block Avail Map"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   6960
      TabIndex        =   11
      Tag             =   "&H00FF0000&"
      ToolTipText     =   "Click to View X-Cable directory"
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label lblView 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Directory Entry"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   5310
      TabIndex        =   10
      Tag             =   "&H00FF0000&"
      ToolTipText     =   "Click to View X-Cable directory"
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label lblX 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12660
      TabIndex        =   1
      Top             =   5610
      Width           =   795
   End
End
Attribute VB_Name = "frmDiskEd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2017 Steve J. Gray
' ====================================================
'
' frmDiskEd - Disk Image File Editor
'
' Edit Disk Images (D64, D71, D80 etc). View Sectors.
'
' NOTE: This feature is incomplete! Viewing is limited, and writing is not supported.

Dim Blink As Boolean
Dim FontSet As Integer
Dim LastRow As Integer, LastCol As Integer, LastW As Integer, LastH As Integer 'Save printing position
Dim GetKey As Integer, GetCode As Integer
Dim ViewMode As Integer

Dim DIFilename As String                        'Disk Image Filename
Dim TPos(160)                                   'File Offsets for track starts - must add sector offset
Dim DI As DskImg                                'The parameters for the disk image

Dim Hdr As Header1Type                          'The Header Info (will be re-dim'd later for proper format)
Dim DEntry(244) As DirEntryType                      'Directory Entry (will be re-dim'd later for proper format)
Dim bam As BAM4Type                             'BAM Track entry (will be re-dim'd later for proper format)

Dim BBuf As String * 256                        'Disk Sector Buffer
Dim CBuf As String * 256                        'Copy Buffer
Dim UBuf As String * 256                        'UNDO Buffer

Dim TrackNum As Integer, SectorNum As Integer   'Current Track and Sector
Dim LinkT As Integer, LinkS As Integer          'Sector Link - first two bytes of sector point to next sector
Dim DFIO As Integer                             'Disk File# - remains open for all subs
Dim UserT1 As Integer, UserS1 As Integer        'User Block1
Dim UserT2 As Integer, UserS2 As Integer        'User Block1

'---- Load the Form
Private Sub Form_Load()
    On Error Resume Next
    
    '-- Load FONTs - Each character is 8x8pixels. Two 256-character fonts arranged in 32x16 grid
    '  (graphics set followed by text set)
    picFontSet.Picture = LoadPicture(ExeDir & "\font-c64.bmp")
    
    Me.Show
End Sub

'---- Unload Form
Private Sub Form_Unload(Cancel As Integer)
    Close DFIO                                          'Close disk image file
End Sub

'---- Load a Disk Image
' This routine is called from the main form with the name of the disk imaged passed to it
Public Sub LoadImg(ByVal Filename As String)
    Dim Ext As String, DParamFile, FIO As Integer, FLen As Long

    If Exists(Filename) = False Then Exit Sub
    DIFilename = Filename                               'Remember the Disk Image filename
    Me.Caption = "Disk Image Editor: " & DIFilename     'Set the Titlebar
    
    DFIO = FreeFile
    Open Filename For Binary As DFIO                    'Open the Disk Image File - DO NOT CLOSE FILE!!!!!
    FLen = LOF(DFIO): DI.FileSize = FLen                'Get the File Size
        
    Ext = FileExt(Filename)                             'Get Disk Image Type (Extension)
    DParamFile = ExeDir & "image-" & Ext & ".txt"       'Filename for Paramters
    
    If Exists(DParamFile) = True Then
        LoadParams DParamFile                           'Load the Parameters
        ReadDir                                         'Read Directory
        ViewMode = 2                                    'TEMP: sector editor
        SelectTab                                       'Select TAB
        TrackNum = DI.DirT: SectorNum = DI.DirS         'TEMP: Directory Track and Sector
        ChangeTS                                        'ValidateTS, ReadBlock and Update View
    Else
        MyMsg "Fatal Error: Can't load Disk Image Parameter File!"
        Unload Me                                       'Ooops, no go!
    End If
End Sub

'---- Read Directory and Header, then Display It
Private Sub ReadDir()

End Sub

'---- Handle clicking of View Tab Buttons
Private Sub lblView_Click(Index As Integer)
    ViewMode = Index
    SelectTab
End Sub

'---- Set View Tab Button Hilighting
Private Sub SelectTab()
    Dim a As Integer
    
    For a = 0 To 3
        lblView(a).Font.Bold = False
        lblView(a).ForeColor = vbBlack
    Next a
    
    If DI.MaxErr = 0 Then
        lblView(3).Enabled = False
        lblNav(11).Visible = False
    Else
        lblView(3).Enabled = True 'Show or Hide Error block tab
        lblNav(11).Visible = True
    End If
    
    lblView(ViewMode).Font.Bold = True
    lblView(ViewMode).ForeColor = vbWhite
    DoEvents
    
    '-- Hide Elements
    picV.Visible = False
    vsV.Visible = False
    frSector.Visible = False
    
    '-- Unhide Elements and set parameters
    Select Case ViewMode
        Case 0
        Case 1
        Case 2 'Sector
            picV.Visible = True         'Show output area
            vsV.Min = 1: vsV.Max = 17   'Set scrollbar range
            vsV.Visible = True          'Show the scrollbar
            frSector.Visible = True     'Show Info Frame
        Case 3
    End Select
    
End Sub

'---- Update the Current view
Private Sub UpdateView()
    Select Case ViewMode
        Case 0: ViewDEnt    'Directory Entry
        Case 1: ViewBAM     'BAM
        Case 2: ViewSector  'Sector
        Case 3: ViewError   'Error
    End Select
    DoEvents
End Sub

'---- View Directory Entry
Private Sub ViewDEnt()

End Sub

'---- View BAM
Private Sub ViewBAM()

End Sub

'---- View Sector
Private Sub ViewSector()
    Dim i As Integer, j As Integer, n As Integer, bv As String * 1, hv As String
    Dim Out As String, Out2 As String, TopRow As Integer

    TopRow = vsV.value - 1: lblX.Caption = Str(TopRow)                          'Read Scrollbar value for top line
    txtCurTrack.Text = Str(TrackNum)                                            'Track
    txtCurSector.Text = Str(SectorNum)                                          'Sector
    n = GetBV(1): LinkT = n: txtLTrack.Text = Str(n)                            'Link-to Track
    n = GetBV(2): LinkS = n: txtLSector.Text = Str(n)                           'Link-to Track
    If LinkT = 0 Then lblNav(4).Visible = False Else lblNav(4).Visible = True   'Show or Hide Link Button
       
    For i = 0 To 15
        n = (TopRow + i) * 8                                                    'Calculate Top Offset
        Out = MyHex(n, 2) & ": ": Out2 = ""                                     'Set initial output strings
        For j = 0 To 7
            bv = GetBC(n + j + 1)                                               'Get the byte
            hv = Asc(bv)                                                        'Hex value
            Out = Out & MyHex(hv, 2) & " "                                      'Add to hex output string
            Out2 = Out2 & bv                                                    'Add to cbm output string
        Next j
        CBMPrint Out, i, 0, 0, 36, 2, 0, picV                                   'Print the hex values
        CBMPrint Out2, i, 29, 0, 36, 2, 0, picV                                 'Print the cbm characters
    Next i

End Sub

'---- View Error Block (if Disk Image has one)
Private Sub ViewError()

End Sub

'---- Navigate Sector View
Private Sub lblNav_Click(Index As Integer)
    Select Case Index
        Case 0: TrackNum = TrackNum + 1                                 'Track UP
        Case 1: TrackNum = TrackNum - 1                                 'Track DOWN
        Case 2: SectorNum = SectorNum + 1                               'Sector UP
        Case 3: SectorNum = SectorNum - 1                               'Sector DOWN
        Case 4: If LinkT > 0 Then TrackNum = LinkT: SectorNum = LinkS   'Jump to Link (Track must be >0)
        Case 5: TrackNum = DI.DirT: SectorNum = DI.DirS                 'Jump to First Directory Block
        Case 6: TrackNum = DI.BAMT: SectorNum = DI.BAMS                 'Jump to First BAM Block
        Case 7
            UserT1 = TrackNum: UserS1 = SectorNum                       'Set User Jump1
            lblNav(8).Caption = TTSS(UserT1, UserS1)
        Case 8: TrackNum = UserT1: SectorNum = UserS1                   'Jump to User Block
        Case 9
            UserT2 = TrackNum: UserS2 = SectorNum                       'Set User Jump2
            lblNav(10).Caption = TTSS(UserT2, UserS2)
        Case 10: TrackNum = UserT2: SectorNum = UserS2                  'Jump to User Block
        Case 11 'Error Map

    End Select
        
    ValidateTS                                                          'Correct T or S if needed
    GetBlock                                                            'Read the block (TrackNum,SectorNum)
    UpdateView                                                          'Display it
End Sub

'---- Block Operations
Private Sub cmdBlockOps_Click(Index As Integer)
    Dim Tmp As String, n As Integer
    Select Case Index
        Case 0: CBuf = BBuf                         'Copy Buffer
        Case 1: BBuf = CBuf                         'Paste Buffer
        Case 2: BBuf = String(256, Chr(0))          'Zero Buffer
        Case 3
            Tmp = InputBox("Enter DECIMAL value to fill:", "Fill Block", "00")
            If Tmp <> "" Then
                n = Int(Val(Tmp))                   'Get value
                BBuf = String(256, Chr(n))          'Fill it
            End If
        Case 4: BBuf = UBuf                         'Restore Buffer
    End Select
    
    UpdateView                                      'Display it
End Sub

'---- Keypress on CBM EDIT field
Private Sub picED_KeyDown(KeyCode As Integer, Shift As Integer)
    GetCode = KeyCode
    lblKeyD.Caption = Str(KeyCode)
End Sub

Private Sub picEd_KeyPress(KeyAscii As Integer)
    GetKey = KeyAscii                                       'Pass ASCII value to global variable
    lblKeystroke.Caption = Str(KeyAscii)                    'Display it for debugging
End Sub

Private Sub picED_LostFocus()
    GetKey = 27                                             'If you click somewhere else assume ESC
End Sub

Private Sub picHeader_Click()
    Dim Txt As String
    Txt = CBMEdit(picHeader, "header", 0, 15, 2, 0, 0, 0)   'dummy
End Sub

Private Sub picID_Click()
    Dim Txt As String
    Txt = CBMEdit(picID, "id", 0, 1, 2, 0, 0, 0)            'dummy
End Sub

Private Sub picDOS_Click()
    Dim Txt As String
    Txt = CBMEdit(picDOS, "2a", 0, 1, 2, 0, 0, 0)           'dummy
End Sub

'---- Blinking Cursor timing
' Every time the timer event fires we flip the Blink flag
' Timer interval determines the speed. Currently 500ms.
Private Sub Timer1_Timer()
    Blink = Not Blink
End Sub

'---- CBM Print Routine for ASCII/SCREEN
' Mode selects format for Txt parameter: 0=ASCII, 1=SCREEN
' picFontSet is picturebox containing CBM font bitmap in 8x8 format with 32 character wide lines
' picP is the picturebox that you want to print to (byref)
' Row,Col are positions calculated with zoom factored in (negative value means use last position)
' MaxR,MaxC is the size of the printing box              (negative value means use last position)
'
Sub CBMPrint(ByVal Txt As String, ByVal Row As Integer, ByVal Col As Integer, ByVal MaxR As Integer, ByVal MaxC As Integer, ByVal Zoom As Integer, ByVal Mode As Integer, ByRef picP As PictureBox)
    Dim Ch As Integer, i As Integer, Z As Integer, ZZ As Integer, FS As Integer, RR As Integer
    Dim R As Integer, C As Integer, SR As Integer, SC As Integer, R2 As Integer, C2 As Integer
        
    R = Row: If Row < 0 Then R = LastRow                                    'Row# to start printing. If Row is negative then use Last Row
    C = Col: If Col < 0 Then C = LastCol                                    'Col# to start printing. If Col is negative then use last Col
    If MaxR < 0 Then MaxR = LastH
    If MaxC < 0 Then MaxC = LastW
    FS = FontSet * 64                                                       'Font Set selection
    ZZ = 8: Z = Zoom * ZZ                                                   'Zoom
    RR = R * Z: cc = C * Z                                                  'Row/Col Pixel for start position
    
    For i = 1 To Len(Txt)
        Ch = Asc(Mid(Txt, i, 1))                                            'Character to print
        
        If Mode = 0 Then
            Select Case Ch
                Case 64 To 127: Ch = Ch - 64
                'Case 96 To 127: Ch = Ch - 96                                'Convert to Screen Code
            End Select
        End If
        
        SR = Ch \ 32: SC = Ch Mod 32                                        'Source Row,Col for character in Font
        R2 = SR * ZZ + FS: C2 = SC * ZZ                                     'Position in Font
        picP.PaintPicture picFontSet.Image, cc, RR, Z, Z, C2, R2, ZZ, ZZ    'Blit it (with zoom)
        C = C + 1: cc = cc + Z                                              'Next Col
        If C > MaxC Then
            C = 0: R = R + 1: If R > MaxR Then Exit For                     'Next line, exit when at BOTTOM
            RR = R * Z: cc = C * Z
        End If
    Next i
    LastRow = R: LastCol = C: LastH = MaxR: LastW = MaxC                    'Remember position and size
    
End Sub

'---- CBM Field Edit Routine for ASCII/SCREEN
' Mode selects format for Txt parameter: 0=ASCII, 1=SCREEN
' Uses picED picturebox. Position/size the box before calling this edit routine
' MaxR,MaxC are size of edit box and determines max size for editing
' Fmt is field type 0=Any,1=Numeric,2=Hex
' XFlag determines if cursor movement out of field acts as <CR> (ie: ends input). END key toggle case
' * This routine uses it's own picturebox to do the editing. It moves itself to match the position and
'   size of the specified picturebox.
Private Function CBMEdit(ByRef picT As PictureBox, ByVal Txt As String, MaxR As Integer, MaxC As Integer, Zoom As Integer, Mode As Integer, Fmt As Integer, XFlag As Integer) As String
    Dim Ch As Integer, i As Integer, Z As Integer, ZZ As Integer, FS As Integer, RR As Integer
    Dim R As Integer, C As Integer, SR As Integer, SC As Integer, R2 As Integer, C2 As Integer
    Dim CursorR As Integer, CursorC As Integer, LastBlink As Boolean
    Dim StrPos As Integer
    
    'SR = Screen.TwipsPerPixelX: SC = Screen.TwipsPerPixelY
    picED.Move picT.Left, picT.Top, picT.Width, picT.Height
    
    StrPos = (MaxR + 1) * (MaxC + 1)
    Txt = Left(Txt + Space(StrPos), StrPos)             'Pad string to max length
    FS = FontSet * 64: ZZ = 8: Z = Zoom * ZZ            'Font Set selection and Zoom
    CBMPrint Txt, 0, 0, MaxR, MaxC, Zoom, Mode, picED   'Display the string to be edited (usually you will use 0,0 for Row,Col)
    
    GetCode = 0: GetKey = 0                             'Reset key input
    picED.Visible = True                                'Show the edit field
    DoEvents
    picED.SetFocus                                      'Make it target of keystrokes
    LastBlink = False                                   'Remember Blinking Cursor Flag so we only draw when it changes
    CursorR = 0: CursorC = 0                            'Cursor Row,Col
    
    Do
        '-- Process new KeyCodes (cursor keys etc)
        If GetCode > 0 Then
            Select Case GetCode
                Case 37: CursorC = CursorC - 1: If (CursorC < 0) And (XFlag = 1) Then Exit Do       'LEFT  /Check if moving outside field
                Case 38: CursorR = CursorR - 1: If (CursorR < 0) And (XFlag = 1) Then Exit Do       'UP    /Other routines can check 'GetCode'
                Case 39: CursorC = CursorC + 1: If (CursorC > MaxC) And (XFlag = 1) Then Exit Do    'RIGHT /To act on cause of exit
                Case 40: CursorR = CursorR + 1: If (CursorR > MaxR) And (XFlag = 1) Then Exit Do    'DOWN
                Case 35: FontSet = 1 - FontSet: FS = FontSet * 64                                   'END KEY=Toggle Case
            End Select
            If CursorC < 0 Then CursorC = MaxC: CursorR = CursorR - 1     'Previous line
            If CursorC > MaxC Then CursorC = 0: CursorR = CursorR + 1     'Next line
            If CursorR < 0 Then CursorR = 0                               'Stop at TOP
            If CursorR > MaxR Then CursorR = MaxR                         'Stop at BOTTOM
            CBMPrint Txt, 0, 0, MaxR, MaxC, Zoom, Mode, picED             'Display the string to be edited
            GetCode = 0: Blink = True                                     'Clear Code and force Blink
            StrPos = CursorR * (MaxC + 1) + CursorC + 1                   'Position in string to edit
            'Debug.Print "@"; CursorR; ","; CursorC; " StrPos="; StrPos; " > "; Asc(Mid(Txt, StrPos, 1))
        End If
     
        
        '-- Do Cursor Blinking - blink speed is set in the Timer1 interval to 500ms.
        
        If Blink <> LastBlink Then
            '-- Calculate Cursor position
            RR = CursorR * Z: cc = CursorC * Z                  'Row/Col Pixel
            StrPos = CursorR * (MaxC + 1) + CursorC + 1         'Position in string to edit
            Ch = Asc(Mid(Txt, StrPos, 1))                       'Character at cursor
            If Mode = 0 Then
                Select Case Ch
                    Case 96 To 127: Ch = Ch - 96                'Convert to Screen Code
                End Select
            End If
            If Blink = True Then Ch = Ch Xor 128                'Reverse it
            SR = Ch \ 32: SC = Ch Mod 32                        'Calc R,C in Fontset
            R2 = SR * ZZ + FS: C2 = SC * ZZ                     'Calc pixel position
            picED.PaintPicture picFontSet.Image, cc, RR, Z, Z, C2, R2, ZZ, ZZ    'Blit it (with zoom)
            LastBlink = Blink                                   'Remember blink state so we don't continually blit
        End If
        
        
        '-- Process new KeyPresses
        If GetKey > 0 Then
            Ch = GetKey                                         'Character that is typed (comes from picED Keypress event
            
            Debug.Print "KEY="; Ch
            
            Select Case Ch
                Case 13: Exit Do                                '<CR> to exit
                Case 27: Txt = "": Exit Do                      '<ESC> to cancel. Return empty string
                Case 64 To 95: If Mode = 0 Then Ch = Ch - 64   '???? test
                Case 96 To 127: If Mode = 0 Then Ch = Ch - 96   'Convert to Screen Code
            End Select
            
             Debug.Print "CONVERTEDKEY="; Ch
             
            Mid(Txt, StrPos, 1) = Chr(Ch)                       'Store it in the string
            CBMPrint Txt, 0, 0, MaxR, MaxC, Zoom, Mode, picED   'Display the string to be edited
            CursorC = CursorC + 1                               'Advance the cursor
            If CursorC > MaxC Then
                If CursorR < MaxR Then CursorR = CursorR + 1    'Move to next line
                CursorC = 0                                     'Start a leftmost column
            End If
            GetKey = 0: Blink = True: LastBlink = False         'Mark it as processed, force cursor blink
        End If
        DoEvents                                                'Otherwise program will appear dead
    Loop
    
    picED.Visible = False                                       'Hide the field
    GetKey = 0                                                  'Clear <CR> code
    CBMEdit = Txt                                               'Return result
    
End Function

'---- Load Disk Paramters from specified parameter file
' This fills the Disk Image Parameter structure, compares file lenghts to determine
' if extended tracks and/or error block is present, and calculates track positions
Public Sub LoadParams(ByVal Filename As String)
    Dim FIO As Integer, buf As String, i As Integer, Tmp As String, Tmp2 As String
    Dim p As Long, V As Integer
        
    FIO = FreeFile
    Open Filename For Input As FIO
        BufLen = intLOF(FIO)                                    'Get the length
        buf = Input(BufLen, FIO)                                'Read to string
    Close FIO

    With DI
        .Desc = GetNamedField(buf, "DESC=")
        .SectSize = GetNamedV(buf, "SECTSIZE=")
        .SectMin = GetNamedV(buf, "SECTMIN=")
        .SectMax = GetNamedV(buf, "SECTMAX=")
        .SectMap = GetNamedField(buf, "SECTMAP=")
        .HeaderT = GetNamedV(buf, "HEADERT=")
        .HeaderS = GetNamedV(buf, "HEADERS=")
        .DirT = GetNamedV(buf, "DIRT=")
        .DirS = GetNamedV(buf, "DIRS=")
        .DirSize = GetNamedV(buf, "DIRSIZE=")
        .BAMT = GetNamedV(buf, "BAMT=")
        .BAMS = GetNamedV(buf, "BAMS=")
        .BAMPos = GetNamedV(buf, "BAMPOS=")
        .BAMSize = GetNamedV(buf, "BAMSIZE=")
        .MaxFiles = GetNamedV(buf, "MAXFILES=")
    
        '-- Check File Size to different track/error variations
        For i = 4 To 1 Step -1
            Tmp2 = "SIZE" & Format(i)
            .MaxTrack = GetNamedV(buf, Tmp2 & "T=") '# Tracks for current size
            .MaxErr = GetNamedV(buf, Tmp2 & "E=")   '# bytes for error block
            Tmp = GetNamedV(buf, Tmp2 & "=")        'Current Size to compare with
            If Val(Tmp) = .FileSize Then Exit For   'Found a match!
        Next i
    
        '-- Calculate Track Start Positions
        p = 1                                                           'Position of Track1, Sector 1
        For i = 1 To .MaxTrack
            TPos(i) = p                                                 'Save it to array
            p = p + ((.SectMin + Val(Mid(.SectMap, i, 1)))) * .SectSize 'Calc next track position
        Next i
    End With
End Sub

'---- Read Sector to BBuf using TrackNum,SectorNum and copy to UNDO
Private Sub GetBlock()
    Dim p As Long
    
    p = TPos(TrackNum) + (SectorNum * DI.SectSize)          'Calculate position of sector
    Get #DFIO, p, BBuf                                     'Get Sector into buffer
    UBuf = BBuf
    lblDebug.Caption = "GetBlock=" & Format(TrackNum) & "/" & Format(SectorNum)
End Sub

'---- Read Specified Sector to BBuf (no undo)
Private Sub GetTS(ByVal T As Integer, ByVal S As Integer)
    Dim p As Long
    
    p = TPos(T) + (S * DI.SectSize)         'Calculate position of sector
    Get #DFIO, p, BBuf                     'Get Sector into buffer
End Sub

'---- Put Buffer to Sector
Private Sub PutTS(ByVal T As Integer, ByVal S As Integer)
    Dim p As Long
    
    p = TPos(T) + (S * DI.SectSize)         'Calculate position of sector
    Put #DFIO, p, BBuf                     'Put buffer to Sector/File
End Sub

'---- Get Block buffer byte Value
Private Function GetBV(ByVal Offset As Integer) As Integer
    GetBV = Asc(Mid(BBuf, Offset, 1))
End Function

'---- Get Block buffer byte Character/String
Private Function GetBC(ByVal Offset As Integer) As String
    GetBC = Mid(BBuf, Offset, 1)
End Function

'---- Validate Track and Sector Range
Private Sub ValidateTS()
    Dim MaxSector As Integer
    If TrackNum < 1 Then TrackNum = 1
    If TrackNum > DI.MaxTrack Then TrackNum = DI.MaxTrack
    
    MaxSector = DI.SectMin + Val(Mid(DI.SectMap, TrackNum, 1))
    If SectorNum < 0 Then SectorNum = 0
    If SectorNum > MaxSector Then SectorNum = MaxSector
End Sub

'---- Change T or S
Private Sub ChangeTS()
    ValidateTS
    GetBlock
    lblDebug.Caption = "ChangeTS-Got Block=" & TTSS(TrackNum, SectorNum)
    UpdateView
End Sub

'---- Manual Input of Track#
Private Sub txtCurTrack_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TrackNum = Val(txtCurTrack.Text): ChangeTS
End Sub

'---- Manual Input of Sector#
Private Sub txtCurSector_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SectorNum = Val(txtCurSector.Text): ChangeTS
End Sub

Private Sub vsV_Change()
    UpdateView
End Sub

Private Sub vsV_scroll()
    UpdateView
End Sub

Private Function TTSS(ByVal TT As Integer, ByVal SS As Integer) As String
   TTSS = Format(TT, "00") & "/" & Format(SS, "00")
End Function
