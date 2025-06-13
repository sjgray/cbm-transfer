VERSION 5.00
Begin VB.Form frmDAD 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CBM Transfer"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1785
   FillStyle       =   0  'Solid
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   80
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   119
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCX 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   0
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ready!"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   960
      Width           =   1800
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

Private Sub Form_Activate()
    On Error Resume Next
    
    '---- Set the Theme
        
    Me.ForeColor = ThemeFG                                      'Set Foreground Colour
    Me.BackColor = ThemeBG                                      'Set Background Colour
    
    lblStat.ForeColor = ThemeFG                                 'Set Status Colour
    frmMain.GetIcon picCX, 230, 1                               'Get the DAD Icon
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If BusyFlag = True Then Cancel = True
End Sub

Private Sub picCX_Click()
    If BusyFlag = True Then Exit Sub
    
    If Trim(DiskID(2)) <> "" Then
        frmMain.MakeXDiskImage
    Else
        MyMsg "No disk. Insert disk and try again!"
    End If
End Sub

Private Sub picCX_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub picCX_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
  '0=do not allow drop, 1=inform source that data will be copied
  If BusyFlag = True Then Effect = 0: Exit Sub
  
  If Data.GetFormat(vbCFFiles) Then Effect = 1 Else Effect = 0
End Sub

