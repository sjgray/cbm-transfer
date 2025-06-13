VERSION 5.00
Begin VB.Form frmViceSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Vice Emulator "
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "SuperCPU 64"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   5
      Left            =   1320
      TabIndex        =   9
      Top             =   810
      Width           =   3705
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "PLUS/4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   8
      Left            =   2580
      TabIndex        =   8
      Top             =   420
      Width           =   1185
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "DTV"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   2
      Left            =   3840
      TabIndex        =   7
      Top             =   420
      Width           =   1185
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "P500"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   7
      Left            =   1320
      TabIndex        =   6
      Top             =   420
      Width           =   1185
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "CBM-II"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   6
      Left            =   60
      TabIndex        =   5
      Top             =   420
      Width           =   1185
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "PET"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   9
      Left            =   60
      TabIndex        =   4
      Top             =   30
      Width           =   1185
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "VIC-20"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   4
      Left            =   1320
      TabIndex        =   3
      Top             =   30
      Width           =   1185
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "C128"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   3
      Left            =   3840
      TabIndex        =   2
      Top             =   30
      Width           =   1185
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "C64 sc"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   2580
      TabIndex        =   1
      Top             =   30
      Width           =   1185
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "C64"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   810
      Width           =   1185
   End
End
Attribute VB_Name = "frmViceSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2021 Steve J. Gray
' ====================================================
'
' frmViceSelect - PopUp Window to select which VICE emulator to Run

Public EmuNum As Integer

'---- Load the form
' Checks which VICE Emulators are available
Private Sub Form_Load()
    Dim i As Integer, VExe As String
    
    For i = lbVSel.lbound To lbVSel.UBound
        VExe = CBMVICE & ViceEXE(i + 2) & ".exe"                              'Build path to VICE Executable
        
        If Exists(VExe) = True Then
            lbVSel(i).ForeColor = vbWhite
            lbVSel(i).Enabled = True
            lbVSel(i).ToolTipText = VExe
        Else
            lbVSel(i).ForeColor = vbBlack
            lbVSel(i).BackColor = RGB(64, 64, 64)
            lbVSel(i).Enabled = False
            ' lbVSel(i).ToolTipText = "Un-Available"  'tooltips do not display if disabled
        End If
    Next i
    
End Sub

'---- Unload the Form - Return a 0
Private Sub Form_Unload(Cancel As Integer)
    
    EmuNum = 0

End Sub

'---- Select VICE Emulation - Return Selected Emulation#
Private Sub lbVSel_Click(Index As Integer)
    
    EmuNum = Index + 2                                          'Return Index+2
    Me.Hide                                                     'Hides the form and allows caller to continue
End Sub
