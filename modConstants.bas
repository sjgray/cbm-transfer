Attribute VB_Name = "Constants"
' Copyright (C) 2004-2005 Leif Bloomquist
' Copyright (C) 2006      Wolfgang Moser
'
' This software Is provided 'as-is', without any express or implied
' warranty. In no event will the authors be held liable for any damages
' arising from the use of this software.
'
' Permission is granted to anyone to use this software for any purpose,
' including commercial applications, and to alter it and redistribute it
' freely, subject to the following restrictions:
'
'     1. The origin of this software must not be misrepresented; you must
'        not claim that you wrote the original software. If you use this
'        software in a product, an acknowledgment in the product
'        documentation would be appreciated but is not required.
'
'     2. Altered source versions must be plainly marked as such, and must
'        not be misrepresented as being the original software.
'
'     3. This notice may not be removed or altered from any source
'        distribution.
'

Option Explicit

'====== Windows API Functions

Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByRef pidl As Long) As Long


Public Declare Function ShellExecute _
    Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    
Public Declare Sub SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long)

Public Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "Kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)


'====== Constants for APIs

Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const INIFILE = "CBMXfer.ini"

Public Const BIF_RETURNONLYFSDIRS = &H1

Public Const HWND_TOPMOST = -&H1
Public Const HWND_NOTOPMOST = -&H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2


'====== Type Decl for APIs

Public Type ReturnStringType
    Output As String
    Errors As String
End Type

'Used for browsing directories
Public Type SHITEMID
  cb      As Long
  abID    As Byte
End Type

'Used for browsing directories
Public Type ITEMIDLIST
  mkid    As SHITEMID
End Type

'Used for browsing directories
Public Type BROWSEINFO
  hOwner          As Long
  pidlRoot        As Long
  pszDisplayName  As String
  lpszTitle       As String
  ulFlags         As Long
  lpfn            As Long
  lParam          As Long
  iImage          As Long
End Type


'Used for screen functions - Popup menu
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

