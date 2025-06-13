VERSION 5.00
Begin VB.Form frmBatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Batch Imaging"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4980
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRetry 
      Caption         =   "Retry"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2340
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read Disk"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtFilename 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   660
      Width           =   3675
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3780
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblOp2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "msgs"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   6
      Top             =   1140
      Width           =   4815
   End
   Begin VB.Label lblOp 
      BackColor       =   &H00FF8080&
      Caption         =   "Batch Imaging"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   4815
   End
   Begin VB.Label lblFE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2017 Steve J. Gray
' ====================================================
'
' frmBatch - Batch Processing
'
' Batch Mode: 0=Manual, 1=Label, 2=Numbered
'
Dim Fmt As String, LastD As Integer, LastS As Integer

Private Sub cmdCancel_Click()
    frmBatch.Hide
    SaveINI
End Sub

Private Sub cmdRead_Click()
    frmMain.GetXDir
    If BatchMode = 1 Then txtFilename.Text = DiskName(2) 'disk label
End Sub

Private Sub cmdRetry_Click()
    txtFilename.Text = BatchFilename
    DiskNum = LastD: DiskSide = LastS
End Sub

Private Sub cmdStart_Click()
    BatchFilename = txtFilename.Text
    LastD = DiskNum: LastS = DiskSide
    
    If Batch2Sided = True Then
        DiskSide = DiskSide + 1: If DiskSide = 3 Then DiskSide = 1: DiskNum = DiskNum + 1
    Else
        DiskNum = DiskNum + 1
    End If
    MakeNext
    
    lblOp.Caption = "Imaging to: " & BatchFilename & " ..."
    lblOp2.Caption = "While disk is imaging you may edit the filename for the next disk."
    lblFE.Caption = "Next Filename:"
    
    cmdRead.Enabled = False
    cmdStart.Enabled = False
    cmdCancel.Enabled = False
    
    If BatchMode = 2 Then frmMain.GetXDir
    frmMain.MakeXDiskImage
    frmMain.lstLocal(0).Refresh
    
    lblOp.Caption = "Imaging complete. Swap disks then press START..."
    lblOp2.Caption = "If there are no more disks click DONE."
    lblFE.Caption = "Filename:"

    cmdRead.Enabled = True
    cmdStart.Enabled = True
    cmdCancel.Enabled = True

End Sub

Private Sub Form_Load()
    On Error Resume Next
    InitBatch
End Sub

Sub InitBatch()
    DiskSide = 1
    DiskNum = Val(frmOptions.txtStartNum.Text)      'Get starting disk#
    Fmt = frmOptions.txtBatchFN.Text                'Get format string
    
    lblFE.Caption = "Filename:"
    
    Select Case BatchMode
        Case 0:
            txtFilename.Text = ""
            lblOp.Caption = "Ready to start batch imaging with manual filenames. Edit the filename before pressing START!"
        Case 1:
            txtFilename.Text = DiskName(2)         'Get label
            lblOp.Caption = "Ready to start batch imaging using disk label."
        Case 2:
            MakeNext
            lblOp.Caption = "Ready to start batch imaging using numbered filenames..."
    End Select
    
    lblOp2.Caption = "Click START when ready.  Press DONE when all disks are done."
    
End Sub

Private Sub MakeNext()
    txtFilename.Text = BatchName(DiskNum, DiskSide, Fmt)
End Sub

Private Sub SetTheme()
' ThemeBG=Title/Background
' ThemeFrBG=Frames Background
' ThemeListBG=Listbox Background
' ThemeListFG=Listbox Foreground
' ThemeFG=Text Labels

    frmBatch.BackColor = ThemeBG
    lblOp.BackColor = ThemeListBG: lblOp.ForeColor = ThemeFG
    lblOp2.BackColor = ThemeFrBG: lblOp2.ForeColor = ThemeFG
    txtFilename.BackColor = ThemeListBG: txtFilename.ForeColor = ThemeListFG
    lblFE.ForeColor = ThemeFG
    
    DoEvents
End Sub
