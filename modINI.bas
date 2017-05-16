Attribute VB_Name = "modINI"
' CBM-Transfer - Copyright (C) 2007-2017 Steve J. Gray
' ====================================================
'
' modINI - Module with INI file routines
'
' Based on GUI4CBM4WIN. The following (between "/" lines) is the notice
' included with the GUI4CBM4WIN source code:
'
'/////////////////////////////////////////////////////////////////////////
'
'INI Routines
'============
' Copyright (C) 2004-2005 Leif Bloomquist
' Copyright (C) 2006      Wolfgang Moser
' Copyright (C) 2007-2017 Steve J. Gray
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

Dim INIBuf As String

'---- Load the INI file
Public Sub LoadINI()
    Dim FIO As Integer, Filename As String, Tmp As String, j As Integer, V As Integer
    Dim LastSrc As String, LastDst As String
    
    'Set Defaults
    frmOptions.DefaultSrcPath.Text = CurDir
    TransferString = "auto"
    AutoRefreshDir = True
    
    'Load the parameters from the ini file, overriding the defaults
    Filename = ExeDir & INIFILE
    If Exists(Filename) = True Then
        FIO = FreeFile
        Open Filename For Input As FIO
        INIBuf = Input(LOF(FIO), FIO) 'Load INI to buffer
        Close FIO
        
        With frmOptions
            '---- Path Options

            LastSrc = AddSlash(INIStr("SrcPath"))
            LastDst = AddSlash(INIStr("DstPath"))
            LocalDir(0) = AddSlash(INIStr("DefaultSrcPath")): .DefaultSrcPath.Text = LocalDir(0)
            LocalDir(1) = AddSlash(INIStr("DefaultDstPath")): .DefaultDstPath.Text = LocalDir(1)
            
            PathHistory = INIbool("PathHistory"):       .cbPathHistory.value = B2V(PathHistory)
            UseLP = INIbool("UseLastPaths"):            .cbLastPaths.value = B2V(UseLP)
            
            If UseLP = True Then
                If LastSrc <> "" Then LocalDir(0) = LastSrc
                If LastDst <> "" Then LocalDir(1) = LastDst
            End If
            
            '---- General Options

            DriveNum = INInum("DriveNum"):              If DriveNum > 7 Then .cboDriveNum.ListIndex = DriveNum - 8
            NoWarpString = INIStr("NoWarpString"):      .CheckNoWarpMode.value = B2V((NoWarpString = "--no-warp"))
            TransferString = INIStr("TransferString"):  UpdateTransferSelect
            AutoRefreshDir = INIbool("AutoRefreshDir"): .cbAutoRefreshDir.value = B2V(AutoRefreshDir)
            PreviewCheck = INIbool("PreviewCheck"):     .cbPreview.value = B2V(PreviewCheck)
            P00Flag = INIbool("WriteP00"):              .cbP00.value = B2V(P00Flag)
            ConfirmD64 = INIbool("ConfirmD64"):         .cbConfirmCreate.value = B2V(ConfirmD64)
            IgnoreD = INIbool("IgnoreD"):               .cbIgnoreD.value = B2V(IgnoreD)
            DstMode = INInum("DestMode"):               .cboDefDst.ListIndex = DstMode
            Tmp = INIStr("LinkCStr"):                   .txtConStr.Text = Tmp
            Tmp = INIbool("ShowErr"):                   .cbErr.value = B2V(Tmp)
            StartDAD = INIbool("StartDAD"):             .cbDAD.value = B2V(StartDAD)
            IgnoreBadID = INIbool("IgnoreBadID"):       .cbIgnoreBadID.value = B2V(IgnoreBadID)
            
            Layout = INInum("Layout")
            Layout2 = INInum("Layout2")
            
            LogAll = INIbool("LogAll"):                 .cbLog.value = B2V(LogAll)
            CheckEXE = INIbool("CheckEXE"):             .cbCheckEXE.value = B2V(CheckEXE)
            
            '---- Bad Filename Options optFNMode

            FNChr = INIStr("FNChr"):                    .txtFNChr.Text = FNChr
            FNEdit = INIbool("FNEdit"):                 .cbFNEdit.value = B2V(FNEdit)
            V = INInum("FNMode"):                       If V < 4 Then .optFNMode(V).value = True: FNMode = V
                        
            '---- VICE Options

            UseVice = INIbool("UseVice"):               .cbUseVice.value = B2V(UseVice)
            VicePath = INIStr("VicePath"):              .txtVicePath = VicePath
            V = INInum("Vice64"):                       .cbo64.ListIndex = V
            V = INInum("Vice71"):                       .cbo71.ListIndex = V
            V = INInum("Vice80"):                       .cbo80.ListIndex = V
            V = INInum("VicePRG"):                      .cboPRG.ListIndex = V
            
            V = INInum("VicePrgMode"):                  If V < 8 Then .OptPRGMode(V).value = True
            
            '---- Nibtools Options

            UseNIB = INIbool("EnableNIB"):             .cbUseNib.value = B2V(UseNIB)
            UseNBZ = INIbool("UseNBZ"):                .cbNBZ.value = B2V(UseNBZ)

            Tmp = INIbool("NibSE"):                    .cbNibSE.value = B2V(Tmp)
            Tmp = INIStr("NibSTrk"):                   .txtNibSTrk.Text = Tmp
            Tmp = INIStr("NibETrk"):                   .txtNibETrk.Text = Tmp
            Tmp = INIStr("NibOpt"):                    .txtNibOpt.Text = Tmp
            Tmp = INIStr("NibRetries"):                .txtRetries.Text = Tmp

            CreateNIB = INIbool("CreateNIB"):          .cbCreateNIB.value = B2V(CreateNIB)
            CreateG64 = INIbool("CreateG64"):          .cbCreateG64.value = B2V(CreateG64)
            CreateD64 = INIbool("CreateD64"):          .cbCreateD64.value = B2V(CreateD64)
            WriteD64 = INIbool("WriteD64"):            .cbWriteD64.value = B2V(WriteD64)

            .cbRetries.value = B2V(INIbool("NibEnRetry"))

            For j = 0 To 7
                .cbNibArg(j).value = B2V(INIbool("NibArg" & Format(j)))
            Next j
            
            UseNibCustom = INIbool("NibCustom"):        .cbNibCustom.value = B2V(UseNibCustom)
            
            Tmp = INIStr("NibRead"):                    .txtNibRead.Text = Tmp
            Tmp = INIStr("NibWrite"):                   .txtNibWrite.Text = Tmp
            Tmp = INIStr("NibConv"):                    .txtNibConv.Text = Tmp
            
            '---- Batch Options

            UseBatch = INIbool("UseBatch"):             .cbUseBatch.value = B2V(UseBatch)
            BatchMode = INInum("BatchMode"):            .optBatchMode(BatchMode).value = True
            Batch2Sided = INIbool("Batch2Sided"):       .cbDouble.value = B2V(Batch2Sided)
            Tmp = INIStr("BatchStart"):                 If Tmp <> "" Then .txtStartNum.Text = Tmp
            Tmp = INIStr("BatchFN"):                    If Tmp <> "" Then .txtBatchFN.Text = Tmp
            
            Tmp = INIbool("LogLabels"):                 .cbLogLabels.value = B2V(Tmp)
            Tmp = INIbool("LogContents"):               .cbLogContents.value = B2V(Tmp)
            DiskNum = INInum("DiskNum"):                .txtStartNum.Text = Format(DiskNum)
            DiskSide = INInum("DiskSide")
            
            '-- Font options

            UseCBMFont = INIbool("UseCBMFont"):          .cbUseCBMFont.value = B2V(UseCBMFont)
            
        End With

        Close #1
        frmOptions.SetConfigOptions                     'build nibstr and fnchar variable
        
    Else
        frmOptions.Show vbModal                         'Show the options window for first run (INI file is not found)
    End If
    Exit Sub
    
LoadINIError:
    Close #1
    MyMsg "Conguration file is corrupt! [" & Err.Description & "]" & Cr & "It will be deleted."
    KillFile Filename   'Delete it!!!!
    Exit Sub
End Sub

'---- Write the INI file
Public Sub SaveINI()
    Dim DirTemp As String, Filename As String, j As Integer
    
    On Local Error GoTo SaveINIError
    
    DirTemp = CurDir    'Remember which directory we're in
    Filename = AddSlash(ExeDir) & INIFILE
    
    Close 1: Open Filename For Output As #1
    
    With frmOptions
        '---- Path Options
        
        PutINIValue "SrcPath", LocalDir(0)
        PutINIValue "DstPath", LocalDir(1)
            
        PutINIValue "DefaultSrcPath", AddSlash(.DefaultSrcPath.Text)
        PutINIValue "DefaultDstPath", AddSlash(.DefaultDstPath.Text)
        PutINIValue "PathHistory", .cbPathHistory.value
        PutINIbool "UseLastPaths", UseLP
        
        '---- General Options
        
        PutINIValue "DriveNum", .cboDriveNum.ListIndex + 8
        PutINIValue "NoWarpString", NoWarpString
        PutINIValue "TransferString", TransferString
        PutINIValue "AutoRefreshDir", .cbAutoRefreshDir.value
        PutINIValue "PreviewCheck", .cbPreview.value
        PutINIValue "WriteP00", .cbP00.value
        PutINIValue "ConfirmD64", .cbConfirmCreate.value
        PutINIValue "DestMode", .cboDefDst.ListIndex
        PutINIValue "LinkCStr", .txtConStr.Text
        PutINIValue "ConfirmD64", .cbConfirmCreate.value
        PutINIValue "IgnoreD", .cbIgnoreD.value
        PutINIbool "LogAll", LogAll
        PutINIbool "CheckEXE", CheckEXE
        PutINIbool "ShowErr", .cbErr.value
        PutINIbool "StartDAD", StartDAD
        PutINIbool "IgnoreBadID", IgnoreBadID
        PutINIValue "Layout", Layout
        PutINIValue "Layout2", Layout2
        
        '---- Filename Options
        
        PutINIValue "FNChr", FNChr
        PutINIValue "FNEdit", FNEdit
        For j = 0 To 2
            If .optFNMode(j).value = True Then PutINIValue "FNMode", j
        Next j
        
        '---- VICE Options
        
        PutINIbool "UseVice", UseVice
        PutINIValue "VicePath", VicePath
        PutINIValue "Vice64", .cbo64.ListIndex
        PutINIValue "Vice71", .cbo71.ListIndex
        PutINIValue "Vice80", .cbo80.ListIndex
        PutINIValue "VicePRG", .cboPRG.ListIndex
        
        For j = 0 To 1
            If .OptPRGMode(j).value = True Then PutINIValue "VicePrgMode", j
        Next j
        
        '---- NibTools Options
        
        PutINIbool "EnableNIB", UseNIB
        PutINIbool "UseNBZ", UseNBZ
        PutINIValue "NibSE", .cbNibSE.value
        PutINIValue "NibSTrk", .txtNibSTrk.Text
        PutINIValue "NibETrk", .txtNibETrk.Text
        PutINIValue "NibOpt", .txtNibOpt.Text
        PutINIValue "NibRetries", .txtRetries.Text
        PutINIValue "NibEnRetry", .cbRetries.value
        
        PutINIbool "CreateNIB", CreateNIB
        PutINIbool "CreateG64", CreateG64
        PutINIbool "CreateD64", CreateD64
        PutINIbool "WriteD64", WriteD64
        
        For j = 0 To 7
            PutINIValue "NibArg" & Format(j), .cbNibArg(j).value
        Next j
        
        PutINIbool "NibCustom", UseNibCustom
        PutINIValue "NibRead", .txtNibRead.Text
        PutINIValue "NibWrite", .txtNibWrite.Text
        PutINIValue "NibConv", .txtNibConv.Text
        
        '---- Batch Options
        
        PutINIbool "UseBatch", UseBatch
        PutINIValue "BatchMode", BatchMode
        PutINIValue "Batch2Sided", Batch2Sided
        PutINIValue "BatchStart", .txtStartNum.Text
        PutINIValue "BatchFN", .txtBatchFN.Text
        PutINIValue "LogLabels", .cbLogLabels
        PutINIValue "LogContents", .cbLogContents
        PutINIValue "DiskNum", DiskNum
        PutINIValue "DiskSide", DiskSide

        '---- Font Options
        
        PutINIValue "UseCBMFont", .cbUseCBMFont
    End With
    
    Close #1
    
    Exit Sub

SaveINIError:
    Close #1
    MyMsg "SaveINI(): " & Err.Description & " (" & Err.Number & ")"
    Exit Sub
End Sub

'---- Get variable and convert to number
Public Function INInum(ByVal varname As String) As Integer
    INInum = Val(INIStr(varname))
End Function

'---- Get variable and convert to boolean
Public Function INIbool(ByVal varname As String) As Boolean
    Dim Tmp As String
    
    INIbool = False
    Tmp = UCase(INIStr(varname)) 'first find the string from the INI file
    
    Select Case Tmp
        Case "1", "TRUE", "WAHR", "VRAI", "CIERTO": INIbool = True 'TRUE in various languages (for multi-lingual versions of windows)
    End Select
End Function

'---- Get variable and return as string
Public Function INIStr(varname As String) As String
    Dim Tmp As String, p As Integer, p2 As Integer

    INIStr = "": Tmp = varname & "="
    p = InStr(1, INIBuf, Tmp, vbTextCompare): If p = 0 Then Exit Function
    p = p + Len(Tmp): p2 = InStr(p, INIBuf, Chr(13))
    If p2 > 0 Then INIStr = Mid(INIBuf, p, p2 - p)
    
End Function

'---- Write string to INI file as named variable
Private Sub PutINIValue(valname As String, value As Variant)
    Print #1, valname & "=" & CStr(value)
End Sub

'---- Write string to INI file as named variable
Private Sub PutINIbool(valname As String, value As Boolean)
    Dim V As String
    V = "0": If value = True Then V = "1"
    Print #1, valname & "=" & V
End Sub

'---- Convert boolean to value
Private Function B2V(ByVal State As Boolean) As Integer
    B2V = -1 * State
End Function

'---- Set Transfer mode's radio selection to current mode
Private Sub UpdateTransferSelect()
      Select Case TransferString
        Case "original": frmOptions.optXMode(0).value = True
        Case "serial1":  frmOptions.optXMode(1).value = True
        Case "serial2":  frmOptions.optXMode(2).value = True
        Case "parallel": frmOptions.optXMode(3).value = True
        Case "auto":     frmOptions.optXMode(4).value = True
    End Select
End Sub
