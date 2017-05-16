VERSION 5.00
Begin VB.Form frmDAD 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CBM-Transfer"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   Picture         =   "frmDAD.frx":0000
   ScaleHeight     =   1230
   ScaleWidth      =   1845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblCX 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      ToolTipText     =   "Drag a supported file here!!!"
      Top             =   0
      Width           =   1875
   End
   Begin VB.Label lblStat 
      BackColor       =   &H8000000D&
      Caption         =   "Ready!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   1020
      Width           =   1875
   End
End
Attribute VB_Name = "frmDAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2017 Steve J. Gray
' ====================================================
'
' frmDAD - Drag and Drop Window
'
' Accepts specific files dropped into the window for disk imaging
' or file viewing

Dim BusyFlag As Boolean

Private Sub Form_Load()
    On Error Resume Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If BusyFlag = True Then Cancel = True
End Sub

Private Sub lblCX_Click()
    If BusyFlag = True Then Exit Sub
    
    If Trim(frmMain.lblXDiskID.Caption) <> "" Then
        frmMain.MakeXDiskImage
    Else
        MsgBox "No disk. Insert disk and try again!"
    End If
End Sub

Private Sub lblCX_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Filename As String, FExt As String
  ' Idea For future:
  ' Use Shift value (ie: state of the SHIFT, CTRL, and ALT keys when they are depressed) for additional options
    
  If BusyFlag = True Then Exit Sub
  
  BusyFlag = True: lblStat.Caption = "Busy..."
  
  If Data.GetFormat(vbCFFiles) Then
     Dim vFn As Variant
     For Each vFn In Data.Files
        Filename = vFn                                                  'The Name of the file that was dropped
        FExt = FileExtU(Filename)                                       'The Extension of the file (uppercase)
        
        Select Case FExt
            Case "D64", "D71"                                           'Make Disk from D64 or D71
                If UseNIB = True Then
                    frmMain.WriteNIBtoX vFn, False                      'Using NIBTools
                Else
                    frmMain.WriteImageToX vFn, False                    'Using OpenCBM
                End If
                
            Case "NIB", "NBZ", "G64"
                frmMain.WriteNIBtoX vFn, False                          'Using NIBTools
                
            Case "D80", "D81", "D82"
                frmMain.WriteImageToX vFn, False                        'Using OpenCBM
            
            Case "", "PRG", "SEQ", "BIN", "ART", "CDU", "KOA", "GEO"
                frmViewer.ViewIt 0, Filename, Filename, FExt            'Display with built-in Viewer
                Exit For
        End Select
     Next
  End If
  
  BusyFlag = False: lblStat.Caption = "DAD is Ready."
    
End Sub

Private Sub lblCX_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
  '0=do not allow drop, 1=inform source that data will be copied
  If BusyFlag = True Then Effect = 0: Exit Sub
  
  If Data.GetFormat(vbCFFiles) Then Effect = 1 Else Effect = 0
End Sub

