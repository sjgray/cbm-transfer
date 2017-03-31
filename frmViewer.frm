VERSION 5.00
Begin VB.Form frmViewer 
   Caption         =   "CBM File Viewer"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   561
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   726
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CBM File Viewer (C)2007 Steve J. Gray
'===============
Public LastFilename As String

Private Sub cboViewMode_Click()
    ViewIt LastFilename
End Sub

Private Sub Form_Load()
    CreatePixels
    Me.Show
    DoEvents
    cboViewMode.ListIndex = 0
End Sub

Public Sub xViewIt(ByVal Filename As String)
    Dim ViewMode As Integer
    
    If Exists(Filename) = False Then Exit Sub
    
    LastFilename = Filename
    
    ViewMode = cboViewMode.ListIndex    'Type of File to display

    Select Case ViewMode
        Case 1: ShowFont Filename, 8
        Case 2: ShowFont Filename, 16
    End Select
         
 End Sub
 
 'Display font
 Public Sub ShowFont(ByVal Filename As String, ByVal FH As Integer)
    Dim FIO As Integer, BufLen As Long, Buf As String
    Dim J As Integer, K As Integer, X As Integer, Y As Integer, V As Integer
    Dim R As Integer, C As Integer, MaxR As Integer, MaxH As Integer
    
    X = 0: Y = 0: MaxR = 256: MaxC = 256
    
    FIO = FreeFile
    Open Filename For Binary As FIO: BufLen = LOF(FIO)
        Buf = Input(BufLen, FIO) 'read to string
    Close FIO
    
    picV.Cls
    
    For J = 1 To BufLen
        V = Asc(Mid(Buf, J, 1))
        picV.PaintPicture Pix.Image, C, R + Y, , , 0, V, 8, 1
        Y = Y + 1
        If Y = FH Then Y = Y - FH: C = C + 8: If C >= MaxC Then C = 0: R = R + FH
    Next J
    
    
End Sub

Public Sub CreatePixels()
    Dim J As Integer, K As Integer, Power(7) As Integer
    
    For J = 0 To 7: Power(J) = 2 ^ J: Next  'Init Powers of 2 array
    
    Pix.Cls
    
    'Create a bitmap with pixels to match binary representation of value (row=value,cols 0 to 7=pixel)
    For J = 0 To 255
        For K = 0 To 7
            If (J And Power(K)) Then Pix.PSet (7 - K, J)
        Next K
    Next J
    
End Sub

