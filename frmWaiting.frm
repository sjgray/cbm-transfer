VERSION 5.00
Begin VB.Form frmWaiting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Working..."
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   7620
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdABORT 
      Caption         =   "ABORT"
      Height          =   345
      Left            =   6840
      TabIndex        =   6
      Top             =   120
      Width           =   705
   End
   Begin VB.Timer LEDTimer 
      Interval        =   300
      Left            =   8880
      Top             =   840
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image cmdDetails 
      Height          =   255
      Left            =   7320
      Picture         =   "frmWaiting.frx":0000
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label lblF 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   7020
      TabIndex        =   5
      Top             =   1020
      Width           =   90
   End
   Begin VB.Label lblOut 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "."
      Height          =   2115
      Left            =   60
      TabIndex        =   4
      Top             =   2640
      Width           =   7275
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDetails 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   1380
      Width           =   7515
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   90
   End
   Begin VB.Shape shBar 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   60
      Top             =   720
      Width           =   15
   End
   Begin VB.Shape shProgress 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   60
      Top             =   720
      Width           =   7515
   End
   Begin VB.Shape LED 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   60
      Top             =   60
      Width           =   255
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait."
      Height          =   465
      Left            =   420
      TabIndex        =   0
      Top             =   60
      Width           =   6120
   End
End
Attribute VB_Name = "frmWaiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2017 Steve J. Gray
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

Private Sub cmdABORT_Click()
    KillFlag = True
    MyMsg "You may need to unplug the Zoom Floppy cable in order to reset it!"
End Sub

'---- Form Load
Private Sub Form_Load()
    On Error Resume Next
    
    Me.AlwaysOnTop = True           'Force to top
    BarWid = shProgress.Width       'Set Progress bar width
    ProgFIO = 0                     'File# for output file to monitor to calculate progress
End Sub

Public Property Let AlwaysOnTop(ByVal bState As Boolean)
  Dim lFlag As Long
  
  If bState Then lFlag = HWND_TOPMOST Else lFlag = HWND_NOTOPMOST
  IsOnTop = bState
  SetWindowPos Me.hWnd, lFlag, 0&, 0&, 0&, 0&, SWP_NOSIZE
End Property

'---- Toggle Details section
Private Sub cmdDetails_Click()
    If frmWaiting.Height <= 1755 Then
        frmWaiting.Height = 2250
    Else
        frmWaiting.Height = 1755
    End If
    DoEvents
End Sub

'There's no "elegant" way to abort a running cbm4win process, short of killing the PID, so this is left for future...
Private Sub Cancel_Click()
    Me.Hide
End Sub

'---- Set the Progress Bar Mode
Public Sub SetMode(ByVal ModeStr As String, Optional ByVal FMax As Long)
    Dim Flag As Boolean
    
    Flag = True
    Select Case UCase(ModeStr)      'The string determines the max size of the output file to be parsed
        Case "D64": FMax = 30900
        Case "D71": FMax = 60000
        Case "D80": FMax = 103244
        Case "D81": FMax = 196921
        Case "D82": FMax = 206488
        Case "NIB", "NIBREAD", "NIBWRITE": FMax = 1950
        Case "NIBCONV": FMax = 4300
        Case "CBMLINK": FMax = 100
        Case Else: Flag = False: If FMax = 0 Then FMax = 30900
    End Select
    
    FMaxLen = CLng(FMax \ 100&)
    
    shProgress.Visible = Flag       'Show Progress%
    shBar.Visible = Flag            'Show Progress BAR
    shBar.Width = 15                'Set bar width
    lblStatus.Visible = Flag
    lblDetails.Caption = "No progress details are available"
    ProgBuf = ""
End Sub

'---- Blink the "LED" indicator, plus read output
Private Sub LEDTimer_Timer()
    Dim Tmp As String, PB As Integer, p As Integer
    
    On Local Error GoTo PErr 'Resume Next
    
    If LED.BackColor = vbRed Then
        LED.BackColor = vbBlack
    Else
        LED.BackColor = vbRed
    End If
    
    If ProgFIO = 0 Then
        ProgFIO = FreeFile                                      '
        Open TEMPFILE1 For Input As ProgFIO
        lblStatus.Caption = "Working..."
        LastLen = 0
    End If
    
    If ProgFIO > 0 Then
        '-- First Method: Compare output file length to known Max length (not including errors)
    
        If FMaxLen = 0 Then FMaxLen = 309
        FLen = LOF(ProgFIO): lblF.Caption = "Size=" & Str(FLen)
        
        PB = Int(FLen / FMaxLen): If PB > 100 Then PB = 100     'Calc %
        lblStatus.Caption = Str(PB) & " %"                      'Set progress percentage
        shBar.Width = (PB / 100) * BarWid                       'Set progress bar width
    
        '-- Second Method: Read output file contents
        ' output format is:  <CR><sector#>:<sectormap> nnn%    <sector#>/<totalsectors><CR>
        '
        
        If (FLen - LastLen) > 80 Then
            Tmp = Input$(80, ProgFIO)                           'Read another chunk of 80 characters
            LastLen = FLen                                      'Updated length read
            ProgBuf = ProgBuf & Tmp                             'Add the new bytes to the buffer
            lblOut.Caption = ProgBuf
        End If
        
        p = InStr(1, ProgBuf, Cr)                               'Look for a <CR>
        If p > 0 Then
            If p > 1 Then ProgBuf = Mid(ProgBuf, p)             'Discard bytes before <CR>
            p = InStr(2, ProgBuf, Cr)                           'Check for next <CR>
            If p > 0 Then
                lblDetails.Caption = Mid(ProgBuf, 2, p - 1)     'Take the part between the <CR>
                ProgBuf = Mid(ProgBuf, p)                       'Discard the above, leaving Buffer starting with <CR>
            End If
        End If
        
    End If
    
    Exit Sub
    
PErr:
    Close ProgFIO: ProgFIO = 0
    Resume Next
End Sub
