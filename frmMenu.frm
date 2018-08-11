VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   0  'None
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   225
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup Menu"
      Begin VB.Menu mnu 
         Caption         =   "Explore &LEFT Directory"
         Index           =   1
      End
      Begin VB.Menu mnu 
         Caption         =   "Explore &RIGHT Directory"
         Index           =   2
      End
      Begin VB.Menu mnu 
         Caption         =   "S&wap LEFT and RIGHT Paths"
         Index           =   3
      End
      Begin VB.Menu mnu 
         Caption         =   "S&et LEFT Path same as RIGHT"
         Index           =   4
      End
      Begin VB.Menu mnu 
         Caption         =   "Set RIGHT Path same as LEFT"
         Index           =   5
      End
      Begin VB.Menu mnu 
         Caption         =   "Add Current Path to History"
         Index           =   6
      End
      Begin VB.Menu mnu 
         Caption         =   "Remove Current Path from History"
         Index           =   7
      End
      Begin VB.Menu mnu 
         Caption         =   "&Clear History"
         Index           =   8
      End
      Begin VB.Menu mnu 
         Caption         =   "Create &New Folder"
         Index           =   9
      End
   End
   Begin VB.Menu mnuSave 
      Caption         =   "Popup Save"
      Begin VB.Menu mnupsa 
         Caption         =   "Save as Text"
         Index           =   1
      End
      Begin VB.Menu mnupsa 
         Caption         =   "Add to catalog"
         Index           =   2
      End
      Begin VB.Menu mnupsa 
         Caption         =   "Show Catalog"
         Index           =   3
      End
      Begin VB.Menu mnupsa 
         Caption         =   "Validate Image"
         Index           =   4
      End
      Begin VB.Menu mnupsa 
         Caption         =   "Backup Image"
         Index           =   5
      End
   End
   Begin VB.Menu mnuFont 
      Caption         =   "FontV"
      Begin VB.Menu mnuF 
         Caption         =   "Toggle Multi-Colour"
         Index           =   1
      End
      Begin VB.Menu mnuF 
         Caption         =   "Toggle Border"
         Index           =   2
      End
      Begin VB.Menu mnuF 
         Caption         =   "Save as Bitmap..."
         Index           =   3
      End
      Begin VB.Menu mnuF 
         Caption         =   "Edit Mode"
         Index           =   4
      End
      Begin VB.Menu mnuF 
         Caption         =   "Save Font As..."
         Index           =   5
      End
      Begin VB.Menu mnuF 
         Caption         =   "Save Range As..."
         Index           =   6
      End
      Begin VB.Menu mnuF 
         Caption         =   "Convert"
         Index           =   7
         Begin VB.Menu mnuConvert 
            Caption         =   "Convert 8x16 to 8x8 font"
            Index           =   0
         End
         Begin VB.Menu mnuConvert 
            Caption         =   "Convert 8x8 to 8x16 font"
            Index           =   1
         End
         Begin VB.Menu mnuConvert 
            Caption         =   "Convert 5x7 sideways font"
            Index           =   2
         End
         Begin VB.Menu mnuConvert 
            Caption         =   "Convert 5x7 upright font"
            Index           =   3
         End
         Begin VB.Menu mnuConvert 
            Caption         =   "Convert 8x14 (EGA ) font"
            Index           =   4
         End
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2017 Steve J. Gray
' ====================================================
'
' frmMenu - Menu selection Dispatch

Option Explicit

'Respond to menu selections - convert to button number for main form dispatcher
Private Sub mnu_Click(Index As Integer)
    Call frmMain.DoMenu(Index)
End Sub

Private Sub mnuConvert_Click(Index As Integer)
    Call frmViewer.DoFMenu(Index + 100)
End Sub

Private Sub mnuF_Click(Index As Integer)
    Call frmViewer.DoFMenu(Index)
End Sub

'Respond to menu selections - convert to button number for main form dispatcher
Private Sub mnupsa_Click(Index As Integer)
    Call frmMain.DoMenu(Index + 100)
End Sub
