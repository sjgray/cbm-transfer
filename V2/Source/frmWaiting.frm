VERSION 5.00
Begin VB.Form frmWaiting 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Working..."
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox cmdDetails 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7230
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   8
      ToolTipText     =   "Show/Hide Details"
      Top             =   960
      Width           =   285
   End
   Begin VB.CommandButton cmdABORT 
      Appearance      =   0  'Flat
      Caption         =   "ABORT"
      Height          =   345
      Left            =   6810
      TabIndex        =   6
      Top             =   60
      Width           =   735
   End
   Begin VB.Timer LEDTimer 
      Interval        =   250
      Left            =   8880
      Top             =   840
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblElapsed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3600
      TabIndex        =   7
      Top             =   1100
      Width           =   405
   End
   Begin VB.Label lblFSize 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7020
      TabIndex        =   5
      Top             =   1100
      Width           =   90
   End
   Begin VB.Label lblOut 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   60
      TabIndex        =   4
      Top             =   2640
      Width           =   7275
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   1380
      Width           =   7575
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   1100
      Width           =   90
   End
   Begin VB.Shape shBar 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   60
      Top             =   660
      Width           =   7395
   End
   Begin VB.Shape shProgress 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   60
      Top             =   660
      Width           =   7485
   End
   Begin VB.Shape LED 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   195
      Left            =   30
      Top             =   60
      Width           =   255
   End
   Begin VB.Label lblMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Please wait."
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   420
      TabIndex        =   0
      Top             =   60
      Width           =   6270
   End
End
Attribute VB_Name = "frmWaiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2021 Steve J. Gray
' ====================================================
'
' frmWaiting - Status and Progress Window
'
' Based on GUI4CBM4WIN. The following (between "/" lines) is the notice
' included with the GUI4CBM4WIN source code:
'
'/////////////////////////////////////////////////////////////////////////

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
Public ProgFIO As Integer, BarWid As Integer
Public FMaxLen As Long, FLen As Long, LastLen As Long, ProgBuf As String
Public StartTime As Date

'---- Form Load
Private Sub Form_Load()
    On Error Resume Next
    
    Me.AlwaysOnTop = True           'Force to top
    
    BarWid = shProgress.Width       'Set Progress bar width
    ProgFIO = 0                     'File# for output file to monitor to calculate progress
    StartTime = Now                 'Initial save of the current time for the Elapsed timer
    
'    SetTheme                        'Set the Theme Colours
    
End Sub

Public Property Let AlwaysOnTop(ByVal bState As Boolean)
  Dim lFlag As Long
  
  If bState Then lFlag = HWND_TOPMOST Else lFlag = HWND_NOTOPMOST
  IsOnTop = bState
  SetWindowPos Me.hWnd, lFlag, 0&, 0&, 0&, 0&, SWP_NOSIZE
  
End Property

'---- Toggle Details section
Private Sub cmdDetails_Click()
    If frmWaiting.Height <= 1725 Then
        frmWaiting.Height = 2115
    Else
        frmWaiting.Height = 1725
    End If
    DoEvents
End Sub

'There's no "elegant" way to abort a running cbm4win process, short of killing the PID, so this is left for future...
Private Sub Cancel_Click()
    Me.Hide
End Sub

Private Sub Form_GotFocus()
    StartTime = Now                                             'Reset Timer
End Sub

'---- Set the Progress Bar Mode
Public Sub SetMode(ByVal ModeStr As String, Optional ByVal FMax As Long)
    Dim Flag As Boolean
       
    SetTheme                                                    'Set Theme colours
    
    Flag = True                                                 'Assume we know size of output file
    
    Select Case UCase(ModeStr)                                  'The string determines the max size of the output file to be parsed
        Case "D64": FMax = 30900
        Case "D71": FMax = 60000
        Case "D80": FMax = 103244
        Case "D81": FMax = 196921
        Case "D82": FMax = 206488
        Case "NIB", "NIBREAD", "NIBWRITE": FMax = 1950
        Case "NIBCONV": FMax = 4300
        Case CBMLink: FMax = 100
        Case Else: Flag = False: If FMax = 0 Then FMax = 30900
    End Select
    
    FMaxLen = CLng(FMax \ 100&)
    
    lblStatus.Visible = Flag                                    'Set Status       visibility
    shProgress.Visible = Flag                                   'Set Progress%    visibility
    shBar.Visible = Flag                                        'Set Progress BAR visibility
    
    shBar.Width = 15                                            'Set bar width to minimum size to start
    
    lblDetails.Caption = "No progress details are available"    'Set Details message
    
    ProgBuf = ""
    StartTime = Now                                             'Record time so we can calculate elapsed
    
End Sub

'---- Blink the "LED" indicator, plus read output
Private Sub LEDTimer_Timer()
    Dim Tmp As String, PB As Integer, P As Integer
    
    On Local Error GoTo PErr
    
    If LED.BackColor = vbRed Then
        LED.BackColor = vbBlack
    Else
        LED.BackColor = vbRed
    End If
    
    lblElapsed.Caption = Format(Now - StartTime, "hh:mm:ss")
    
    If ProgFIO = 0 Then
        ProgFIO = FreeFile                                      '
        Open TEMPFILE1 For Input As ProgFIO
        lblStatus.Caption = "Working..."
        LastLen = 0
    End If
    
    If ProgFIO > 0 Then
        '-- First Method: Compare output file length to known Max length (not including errors)
    
        If FMaxLen = 0 Then FMaxLen = 309
        
        FLen = LOF(ProgFIO)                                     'Get File Length
        PB = Int(FLen / FMaxLen): If PB > 100 Then PB = 100     'Calc % progress
        
        lblStatus.Caption = Str(PB) & " %"                      'Set progress percentage
        shBar.Width = (PB / 100) * BarWid                       'Set progress bar width
        lblFSize.Caption = "Size=" & Str(FLen)                  'Set Size
        
        '-- Second Method: Read output file contents
        ' output format is:  <CR><sector#>:<sectormap> nnn%    <sector#>/<totalsectors><CR>
        '
        
        If (FLen - LastLen) > 80 Then
            Tmp = Input$(80, ProgFIO)                           'Read another chunk of 80 characters
            LastLen = FLen                                      'Updated length read
            ProgBuf = ProgBuf & Tmp                             'Add the new bytes to the buffer
            lblOut.Caption = ProgBuf
        End If
        
        P = InStr(1, ProgBuf, Cr)                               'Look for a <CR>
        If P > 0 Then
            If P > 1 Then ProgBuf = Mid(ProgBuf, P)             'Discard bytes before <CR>
            P = InStr(2, ProgBuf, Cr)                           'Check for next <CR>
            If P > 0 Then
                lblDetails.Caption = Mid(ProgBuf, 2, P - 1)     'Take the part between the <CR>
                ProgBuf = Mid(ProgBuf, P)                       'Discard the above, leaving Buffer starting with <CR>
            End If
        End If
        
    End If
    
    Exit Sub
    
PErr:
    Close ProgFIO: ProgFIO = 0
    Resume Next
End Sub

'---- Handle Abort
' Sets KillFlag to TRUE and displays message to user
Private Sub cmdABORT_Click()

    KillFlag = True
    
    MyMsg "You may need to unplug the Zoom Floppy cable in order to reset it!"

End Sub

'---- Set Theme
Private Sub SetTheme()
    
    Me.BackColor = ThemeListBG
    Me.ForeColor = ThemeListFG
    
    shProgress.BackColor = ThemeFrBG
    shBar.BackColor = ThemeFrBG
    
    lblMsg.BackColor = ThemeListBG: lblMsg.ForeColor = ThemeListFG
    lblDetails.BackColor = ThemeBG: lblDetails.ForeColor = ThemeFG
    
    lblStatus.ForeColor = ThemeFG
    lblFSize.ForeColor = ThemeFG
    lblElapsed.ForeColor = ThemeFG
    
    frmMain.GetIcon cmdDetails, 309, 67                                         'Get Height Scale/Show-Hide Icon
    
    DoEvents
End Sub
