VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColourPicker 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colour Picker"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2790
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1770
      TabIndex        =   17
      Top             =   1230
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   405
      Left            =   1770
      TabIndex        =   16
      Top             =   750
      Width           =   885
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   645
      Index           =   16
      Left            =   1740
      TabIndex        =   18
      ToolTipText     =   "Click here for more colours"
      Top             =   30
      Width           =   915
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   15
      Left            =   1320
      TabIndex        =   15
      Top             =   1290
      Width           =   375
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   14
      Left            =   900
      TabIndex        =   14
      Top             =   1290
      Width           =   375
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   13
      Left            =   480
      TabIndex        =   13
      Top             =   1290
      Width           =   375
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   12
      Left            =   60
      TabIndex        =   12
      Top             =   1290
      Width           =   375
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   11
      Left            =   1320
      TabIndex        =   11
      Top             =   870
      Width           =   375
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   10
      Left            =   900
      TabIndex        =   10
      Top             =   870
      Width           =   375
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   9
      Left            =   480
      TabIndex        =   9
      Top             =   870
      Width           =   375
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   8
      Left            =   60
      TabIndex        =   8
      Top             =   870
      Width           =   375
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   7
      Top             =   450
      Width           =   375
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   6
      Left            =   900
      TabIndex        =   6
      Top             =   450
      Width           =   375
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   450
      Width           =   375
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   60
      TabIndex        =   4
      Top             =   450
      Width           =   375
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Top             =   30
      Width           =   375
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   900
      TabIndex        =   2
      Top             =   30
      Width           =   375
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   30
      Width           =   375
   End
   Begin VB.Label CBox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   375
   End
End
Attribute VB_Name = "frmColourPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2017 Steve J. Gray
' ====================================================
'
' frmColourPicker - Colour Picker
'
' Replacement Colour Picker that uses the Commodore VIC-II palette.
' Click in the Selected Colour Box to use the standard VB Colour Picker

'---- Load the form and Set colours
Private Sub Form_Load()
    Dim i As Integer
    On Error Resume Next
    
    For i = 0 To 15
        CBox(i).BackColor = C64Colour(i)
    Next
    
End Sub

'---- Pick a colour
Private Sub CBox_Click(Index As Integer)
    If Index = 16 Then
        PickedColour = PickColor()                      'Use commondialog colour picker
        If PickedColour >= 0 Then CBox(16).BackColor = PickedColour
    Else
        CBox(16).BackColor = CBox(Index).BackColor      'Set chosen colour
    End If
    
End Sub

'---- Cancel - Return -1
Private Sub cmdCancel_Click()
    PickedColour = -1
    Unload Me
End Sub

'---- OK - Return PickedColour
Private Sub cmdOK_Click()
    PickedColour = CBox(16).BackColor
    Unload Me
End Sub

Private Function PickColor() As Long
    On Local Error GoTo NoPick
    
    CommonDialog1.CancelError = True
    CommonDialog1.ShowColor
    PickColor = CommonDialog1.Color
    Exit Function
    
NoPick:
    PickColor = -1
    
End Function
