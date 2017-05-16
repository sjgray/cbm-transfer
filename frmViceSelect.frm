VERSION 5.00
Begin VB.Form frmViceSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Vice Emulator "
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "PLUS/4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   2640
      TabIndex        =   9
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "DTV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   2640
      TabIndex        =   8
      Top             =   420
      Width           =   1215
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "P500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   1380
      TabIndex        =   7
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "CBM-II"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "PET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "VIC-20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   1380
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "C128"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "C64 sc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   1380
      TabIndex        =   2
      Top             =   420
      Width           =   1215
   End
   Begin VB.Label lbVSel 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "C64"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Which VICE Emulator do you want to run?"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   2985
   End
End
Attribute VB_Name = "frmViceSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2017 Steve J. Gray
' ====================================================
'
' frmViceSelect - PopUp Window to select which VICE emulator to Run

Public EmuNum As Integer

Private Sub Form_Unload(Cancel As Integer)
    EmuNum = 0
End Sub

Private Sub lbVSel_Click(Index As Integer)
    EmuNum = Index + 2
    Me.Hide
End Sub
