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
      Height          =   495
      Left            =   2340
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read Disk"
      Height          =   495
      Left            =   60
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtFilename 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   660
      Width           =   3675
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Done"
      Height          =   495
      Left            =   3780
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblOp2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "msgs"
      Height          =   675
      Left            =   60
      TabIndex        =   6
      Top             =   1140
      Width           =   4815
   End
   Begin VB.Label lblOp 
      BackColor       =   &H00FF8080&
      Caption         =   "Batch Imaging"
      Height          =   435
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   4815
   End
   Begin VB.Label lblFE 
      Alignment       =   1  'Right Justify
      Caption         =   "Filename:"
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
' Batch Mode: 0=Manual, 1=Label, 2=Numbered
'
Dim DiskNum As Integer, DiskSide As Integer, Fmt As String
Dim LastD As Integer, LastS As Integer

Private Sub cmdCancel_Click()
    frmBatch.Hide
End Sub

Private Sub cmdRead_Click()
    frmMain.GetXDir
    If BatchMode = 1 Then txtFilename.Text = frmMain.lblXDiskName 'disk label
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
            txtFilename.Text = frmMain.lblXDiskName         'Get label
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
