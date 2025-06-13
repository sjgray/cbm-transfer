VERSION 5.00
Begin VB.Form frmPrompt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Title Here"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5790
   ControlBox      =   0   'False
   Icon            =   "frmPrompt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelAll 
      Caption         =   "Cancel All"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1230
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   315
      Left            =   4920
      TabIndex        =   4
      Top             =   705
      Width           =   735
   End
   Begin VB.TextBox Reply 
      Height          =   285
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   4725
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1230
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   1230
      Width           =   1335
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Question"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2017 Steve J. Gray
' ====================================================
'
' frmPrompt - Prompt for Info
'
' Call 'Ask' with titlebar and msg test. ClearLast will clear edit box
' OK will exit with edit box as-is. Cancel will clear edit box

Private Sub Form_Load()
    On Error Resume Next
End Sub

Private Sub Form_Activate()
    DoEvents
    Reply.SetFocus
End Sub

'---- Prompt User
Public Sub Ask(ByVal Title As String, ByVal Msg As String, ByVal CFlag As Integer, Optional ClearLast = True)
    Response = ""
    
    frmPrompt.Caption = Title                       'Set TITLE BAR
    Label.Caption = Msg
    
    cmdCancel.Visible = False                       'Hide CANCEL button
    cmdCancelAll.Visible = False                    'Hide CANCEL ALL button
    
    If CFlag > 0 Then cmdCancel.Visible = True      'Enable CANCEL button
    If CFlag = 2 Then cmdCancelAll.Visible = True   'Enable CANCEL ALL button
    
    If (ClearLast) Then Reply.Text = ""             'Clear last response
    Me.Show vbModal                                 'Show the Prompt
    
End Sub

'---- Process OK button
Private Sub cmdOK_Click()
    Response = Trim(Reply.Text)                     'Return the user's input
    Me.Hide
End Sub

'---- Process CANCEL button
Private Sub cmdCancel_Click()
    Reply.Text = ""
    Response = ""                                   'Return NULL
    Me.Hide
End Sub

'---- Process CANCEL ALL button
Private Sub cmdCancelAll_Click()
    Reply.Text = ""
    Response = "***"                                'Special string to indicate cancelling batch operation
    Me.Hide
End Sub

'---- Process CLEAR button
Private Sub cmdClear_Click()
    Reply.Text = ""                                 'Clear current input
End Sub

'---- Handle Keypresses to Input box
Private Sub Reply_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13) Then KeyCode = 0: cmdOK_Click            'Enable Enter Key
End Sub
