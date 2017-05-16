Attribute VB_Name = "modCommon"
' CBM-Transfer - Copyright (C) 2007-2017 Steve J. Gray
' ====================================================
'
' modCommon - Module with Common Subroutines and Functions
'
' FUNCTION NAME     DESCRIPTION
' ----------------- -----------
' MyMsg............ Display MsgBox with default Title
' MyChDir.......... Changes current path
' DirExists........ Checks for existance of a Directory
' Exists........... Checks for existance of a File
' Overwrite........ Prompts for overwriting a file if it exists
' FileExt.......... Returns the EXTension part of the filename
' FileExtU......... Returns the EXTension part of the filename UPPERCASED
' FilePATH......... Returns the PATH without filename
' FileBase......... Returns the Path+Filename without the EXT
' FileNameOnly..... Returns the Filename with EXT
' PathOnly......... Returns the PATH when supplied either a valid directory path, or a filename
' PathUp........... Returns the parent PATH
' AddSlash......... Adds a slash to the end of a Path if it does not have one
' NoSlash.......... Removes trailing slash
' KillFile......... Checks if a file exists, then deletes it
' KillTemp......... Deletes files in the TEMP directory
' SupportedImg..... Checks if known Disk Image type (ie: D64,D80, etc)
' SupportedExt..... Checks if Extension is a valid CBM file (PRG,SEQ,ROM,BIN,or NO EXTENSION)
' DOSName.......... Returns DOS Name (FILENAME.EXT) from CBM Directory Listing
' CBMName.......... Returns CBM Name (FILENAME,P) from CBM Directory Listing
' DOSExt........... Returns CBM file extension (PRG,SEQ,USR,REL,DEL,CBM)
' CBMExt........... Converts DOSExt to CBM File type (PRG=,p  SEQ=,s  USR=,u  REL=,r  DEL=,d CBM=,c)
' CBMType.......... Extracts the CBM Type (",p" , ",s" etc ) from CBM filename ("filename,p")
' Reverse.......... Reverses the CASE of PETSCII text
' CheckPCFilename.. Check for a valid PC filename
' MakePCName....... Converts to a valid PC filename automatically or via prompting
' FixPCName........ Replaces BAD characters with a specifed replacement
' intLOF........... Find LOF but restrict to integer variable size
' Quoted........... Adds Quotes around string
' UnQuoted......... Removes surrounding quotes
' ExtractQuotes.... Returns the quoted string, when located mid-string
' MyDec............ Convert HHHH hex value to decimal
' MyHex............ Convert decimal to fixed-length HEX string with leading zeros
' MyTrim........... Removes leading and trailing spaces
' GetBrowseDir..... Prompt for new directory using "Browse for Folder" popup
' GetLoadAddress... Reads CBM Load Address from specified file
' ViceEXE.......... Returns VICE executable from list
' GetMachine....... Finds target machine using Commodore file Load Address
' DiskID........... Converts Commodore DOS Disk ID to model number string
' ViewFile......... Opens app associated with the file type (ie: Notepad for TXT files)
' BatchName........ Makes filename given template. Uses *, % and ^ as substitution characters
' GetNamedField.... Retrieve Named Field from <CR> delimited Record
' GetNamedV........ Retrieve Named Field as above, and convert to numeric value
' GetField......... Retrieve Field from Comma-Delimited Record
' GetDField........ Retrieve Field from Delimited Record with specified Delimiter (default=TAB)
'

'---- MessageBox popup with default title
Public Sub MyMsg(ByVal Tmp As String)
    MsgBox Tmp, , MsgTitle
End Sub

'----  Changes current path
Public Sub MyChDir(ByVal Path As String)
    On Local Error GoTo MCDErr
    ChDrive Left(Path, 1)
    ChDir Path
MCDErr:
End Sub

'----  Check if Directory exists. If so, returns TRUE
Public Function DirExists(ByVal DirName As String) As Boolean
    On Local Error GoTo ExErr1
    DirExists = False
    If DirName <> "" Then If Dir(DirName, vbDirectory) <> "" Then DirExists = True
ExErr1:
End Function

'---- Check if a file exists. If so, returns TRUE
Function Exists(ByVal Filename As String) As Boolean
    Dim FIO As Integer
    
    On Local Error GoTo NoFile
    Exists = False
    FIO = FreeFile: Open Filename For Input As FIO
    Exists = True
NoFile:
    Close FIO
    DoEvents
End Function

'---- Checks for file and prompts to Overwrite if necessary
' Returns TRUE if file does NOT exist, or it EXISTS and user says YES.
' Returns FALSE if file EXISTS but user says NO.
Public Function Overwrite(ByVal Filename As String) As Boolean
    
    Overwrite = True 'assume ok to replace
    
    If Exists(Filename) = True Then
        If MsgBox("The file '" & Filename & "' already exists!" & Cr & "Replace it?", vbYesNo, "Overwrite File") = vbNo Then Overwrite = False
    End If
End Function

'---- Return file Extension
Public Function FileExt(ByVal Filename As String) As String
    Dim p As Integer
    
    p = InStrRev(Filename, ".")
    If p > 0 Then FileExt = Mid(Filename, p + 1) Else FileExt = ""
End Function

'---- Return file Extension Uppercased
Public Function FileExtU(ByVal Filename As String) As String
    Dim p As Integer
    
    p = InStrRev(Filename, ".")
    If p > 0 Then FileExtU = UCase(Mid(Filename, p + 1)) Else FileExtU = ""
End Function

'---- Return Path without filename
Public Function FilePath(ByVal Filename As String) As String
    Dim p As Integer
    
    p = InStrRev(Filename, "\")
    If p > 0 Then FilePath = Left(Filename, p - 1) Else FilePath = ""
    
End Function

'---- Return Filename without Extension (do not remove path if included)
Public Function FileBase(ByVal Filename As String) As String
    Dim p As Integer
    
    p = InStrRev(Filename, ".")
    If p > 0 Then FileBase = Left(Filename, p - 1) Else FileBase = Filename
    
End Function

'---- Return Filename without Path
Public Function FileNameOnly(ByVal Filename As String) As String
    Dim p As Integer
    
    p = InStrRev(Filename, "\")
    If p > 0 Then FileNameOnly = Mid(Filename, p + 1) Else FileNameOnly = Filename
    
End Function

'---- Return PATH only
' Tmp must contain a path.
' - If Tmp is a DIRECTORY, then it is returned
' - If Tmp is a FILE, then it's path is returned
' Returns path without ending "\"
Public Function PathOnly(ByVal FileSpec As String) As String
    Dim p As Integer, Tmp As String
    
    p = InStr(FileSpec, ".")
    If p = 0 Then
        '-- File has no extension
        If DirExists(FileSpec) = True Then PathOnly = FileSpec: Exit Function
    Else
        '-- File has extension
        If Exists(FileSpec) = True Then
            Tmp = FilePath(FileSpec)        'If it's a file, then extract the path from it
            If DirExists(Tmp) Then
                PathOnly = Tmp              'Yes, return path of file
            Else
                PathOnly = ""               'No, not DIR or FILE... hmmm
            End If
        End If
    End If
End Function
'---- Return the Path one level up from specified path. Includes ending \
Public Function PathUp(ByVal Path As String) As String
    Dim p As Integer
    
    PathUp = Path
    If Len(Path) > 3 Then
        p = InStrRev(Path, "\", Len(Path) - 1)
        If p > 0 Then PathUp = Left(Path, p)
    End If
    
End Function

'---- Checks end of path for \ and adds if not found
Public Function AddSlash(Path As String) As String
    If Not (Right$(Path, 1) = "\") Then
            Path = Path & "\"
    End If
    
    AddSlash = Path
End Function

'---- Removes the trailing \ from a filename
Public Function NoSlash(ByVal Filename As String) As String
    If Right(Filename, 1) = "\" Then
        NoSlash = Left(Filename, Len(Filename) - 1)
    Else
        NoSlash = Filename
    End If
End Function

'---- Delete a file if the file exists
Public Sub KillFile(ByVal Filename As String)
    On Local Error Resume Next
    Kill Filename
End Sub

'---- Deletes all temporary files
Public Function KillTemp()
    KillFile TEMPFILE1
    KillFile TEMPFILE2
End Function

'---- Check if supported Image Extension
' Flag: False=reading, True=writing
Public Function SupportedImg(ByVal Ext As String, ByVal WriteFlag As Boolean) As Boolean
    Select Case UCase(Ext)
        Case "D64", "D71", "D80", "D81", "D82", "G64", "X64": SupportedImg = True
        Case "NIB", "NBZ": If WriteFlag = True Then SupportedImg = True Else SupportedImg = False
        Case "D1M", "D2M", "D4M": If WriteFlag = True Then SupportedImg = False Else SupportedImg = True
        Case Else: SupportedImg = False
    End Select
    
End Function

'---- Check if supported Image Extension
' Flag: False=reading, True=writing
Public Function SupportedExt(ByVal Ext As String) As Boolean
    Select Case UCase(Ext)
        Case "PRG", "SEQ", "ROM", "BIN", "": SupportedExt = True
        Case Else: SupportedExt = False
    End Select
    
End Function

'---- Return DOS filename from CBM Directory Entry
' entry : 123 "filename"  prg<
' output: filename.prg
Public Function DOSName(ByVal Str As String) As String
    Dim Filename As String, Ext As String

    Filename = ExtractQuotes(Str)       'Get Filename
    Ext = DOSExt(Str)                   'Get Extension (PRG,SEQ etc)
    DOSName = Filename & "." & Ext      'Combine and return
End Function


'---- Return CBM filename from CBM Directory Entry
' entry : 123 "filename"  prg<
' output: filename,p
Public Function CBMName(ByVal Str As String) As String
    Dim Filename As String, Ext As String, Ext2 As String

    Filename = ExtractQuotes(Str)           'Get Filename
    Ext2 = DOSExt(Str)
    Ext = CBMExt(Ext2)                      'Get Extension (",p" or ",s" etc)
    CBMName = Filename & Ext                'Combine and return
End Function

'---- Return DOS File Extension from Directory Line
' ie: PRG,SEQ,USR,REL,DEL,CBM also handles locked files ending with "<"
Public Function DOSExt(ByVal Str As String) As String

    If Right(Str, 1) = "<" Then Str = Left(Str, Len(Str) - 1)   'Remove trailing "<" character
    DOSExt = Right(Replace(Str, " ", ""), 3)                    'Return last 3 characters
    
End Function

'---- Convert File Type to CBMDOS extension
' ie: PRG=,p  SEQ=,s  USR=,u  REL=,r  CBM=,c
Public Function CBMExt(ByVal Str As String) As String
    Dim Tmp As String
    
    Tmp = UCase(Str): If Left(Tmp, 1) = "." Then Tmp = Mid(Tmp, 2) 'remove "." if included and convert to uppercase
    
    Select Case Tmp
        Case "PRG", "P00", "P01": CBMExt = ",p"
        Case "SEQ", "S00", "S01": CBMExt = ",s"
        Case "USR", "U00", "U01": CBMExt = ",u"
        Case "REL", "R00", "R01": CBMExt = ",r"
        Case "CBM": CBMExt = ",c"
        Case "DEL": CBMExt = ",d"
        Case Else: CBMExt = ",p"                  'Default for files with no extension
    End Select
End Function

'---- Extract type from full CBM NAME
' ie: "filename,p" -> ",p"
'     "filename,s" -> ",s"
Public Function CBMType(ByVal Str As String) As String
    Dim p As Integer
    
    p = InStr(1, Str, ",")
    If p > 0 Then CBMType = Mid(Str, p) Else CBMType = ""
    
End Function

'---- Reverse Case of PETSCII Text (Mostly for original PET BASIC 1 text strings
Public Function Reverse(ByVal n As Integer) As Integer
    Select Case n
        Case 65 To 90: n = n + 32
        Case 97 To 122: n = n - 32
    End Select
    
    Reverse = n
End Function

'---- Validate PC filename - Check for invalid characters
Public Function CheckPCFilename(ByVal Filename As String) As Boolean
    Dim Bad As String, j As Integer
    CheckPCFilename = True 'assume all okay
        
    If Left(Filename, 1) = " " Then CheckPCFilename = False: Exit Function
    
    Bad = "/\:*?<>|" & Qu
    
    For j = 1 To Len(Bad)
        If InStr(1, Filename, Mid(Bad, j, 1), vbtext) > 0 Then CheckPCFilename = False: Exit For
    Next j

End Function

'---- Fix CBM Filename to be PC friendly according to Option MODE - updated feb 1/2011
' This will automatically fix filenames, or it will prompt for manual entry
Public Function MakePCName(ByVal Filename As String) As String
    Dim j As Integer, EdFlag As Boolean, OldName As String
    
    OldName = Filename
    If CheckPCFilename(Filename) = False Then
        'Bad filename! What do we do?
        Select Case FNMode
            Case 1: Filename = FixPCName(Filename, "")
            Case 2: Filename = FixPCName(Filename, FNChr)
        End Select
        
        If (FNMode = 0) Or (FNEdit = True) Then
            'Edit the filename
            frmPrompt.Reply.Text = Filename
            frmPrompt.Ask "Rename File", "The file '" & OldName & "' contains illegal characters. Please enter a new name:", 1, False
            Filename = Response
        End If
    End If
    
    MakePCName = Filename
End Function

'---- Takes a PETSCII Filename and looks for invalid DOS File system characters. Replaces them with specified character RStr.
Public Function FixPCName(ByVal Filename As String, ByVal RStr As String) As String
    Dim j As Integer, Tmp As String, Tmp2 As String, Bad As String, Flag As Boolean
    
    Tmp = "": Bad = "/\:*?<>|" & Qu
    Flag = False 'flag for spaces at beginning of filename
    
    For j = 1 To Len(Filename)
        Tmp2 = Mid(Filename, j, 1)
        If Tmp2 <> " " Or Flag = True Then
            If InStr(1, Bad, Tmp2) = 0 Then Tmp = Tmp & Tmp2 Else Tmp = Tmp & RStr
            Flag = True
        End If
    Next j

    FixPCName = Tmp

End Function

'---- Returns the 'integer max' Length of a file
Public Function intLOF(ByVal FIO As Integer) As Integer
    intLOF = 32766: If LOF(FIO) < 32766 Then intLOF = LOF(FIO) 'Should be 32767. Why 32766 ? Perhaps it overflows somewhere?
End Function

'---- Surround string with Quotes
Public Function Quoted(ByVal Str As String) As String
    Dim Tmp As String
    
    Tmp = Str
    If Left(Tmp, 1) <> Qu Then Tmp = Qu & Tmp & Qu
    Quoted = Tmp
End Function

'---- Remove Quotes from string
Public Function UnQuoted(ByVal Str As String) As String
    Dim Tmp As String
    
    Tmp = Str
    If (Left(Tmp, 1) = Qu) And (Right(Tmp, 1) = Qu) Then Tmp = Mid(Tmp, 2, Len(Tmp) - 2)
    UnQuoted = Tmp
End Function

'---- Extract string from inbetween quotes mid-string
Function ExtractQuotes(FullString As String) As String
    Dim Quote1 As Integer, Quote2 As Integer
    
    On Local Error GoTo QuoteError
    
    Quote1 = InStr(FullString, Qu)
    Quote2 = InStr(Quote1 + 1, FullString, Qu)
    
    If Quote1 = 0 Then
        ExtractQuotes = FullString                                              'No Quotes found
    Else
        If Quote2 = 0 Then
            ExtractQuotes = Mid(FullString, Quote1)                             'One Quote found
        Else
            ExtractQuotes = Mid$(FullString, Quote1 + 1, Quote2 - Quote1 - 1)   'Two Quotes found
        End If
    End If
    Exit Function
    
QuoteError:
     MyMsg "Extract Quote Error: " & Err.Number & Cr & "[" & FullString & "]"
    
End Function

'---- Convert decimal value to fixed-length HEX value with leading zeros
Function MyDec(ByVal h As String) As Long
    MyDec = CLng(Hx & h)
End Function

'---- Convert decimal value to fixed-length HEX value with leading zeros
Function MyHex(ByVal n As Single, D As Integer) As String
    MyHex = Right("00000000" & Hex(n), D)
End Function

'---- Trims spaces from Beginning and End of string
Function MyTrim(ByVal Str As String) As String
    MyTrim = LTrim(RTrim(Str))
End Function

'---- Display "Browse for folder" window with message header
Public Function GetBrowseDir(ThaForm As Form, Msg As String) As String
            
    GetBrowseDir = vbGetBrowseDirectory(ThaForm.hWnd, Msg)
    
End Function

Public Function vbGetBrowseDirectory(ThaForm As Long, Msg As String) As String

    Dim BI As BROWSEINFO
    Dim IDL As ITEMIDLIST
    
    Dim R As Long, pidl As Long, tmpPath As String, pos As Integer
    
    BI.hOwner = ThaForm
    BI.pidlRoot = 0&
    BI.lpszTitle = Msg
    BI.ulFlags = BIF_RETURNONLYFSDIRS
    
   'get the folder
    pidl = SHBrowseForFolder(BI)
    
    tmpPath = Space$(512)
    R = SHGetPathFromIDList(ByVal pidl, ByVal tmpPath)
      
    If R Then
        pos = InStr(tmpPath, Chr(0))
        tmpPath = Left(tmpPath, pos - 1)
        vbGetBrowseDirectory = tmpPath
    Else
        vbGetBrowseDirectory = ""
    End If

End Function

'---- Reads the first two bytes of a file and calculates the Commodore Load Address
Public Function GetLoadAddress(ByVal Filename As String) As Long
    Dim FIO As Integer, Tmp As String
    
    GetLoadAddress = 0
    If Exists(Filename) = True Then
        FIO = FreeFile
        Open Filename For Binary As FIO
        Tmp = Input(2, FIO)
        Close FIO
        GetLoadAddress = Asc(Mid(Tmp, 1, 1)) + Asc(Mid(Tmp, 2, 1)) * 256
    End If
End Function

'---- Returns the VICE emulator for specified dropdown index
Public Function ViceEXE(ByVal Index As Integer) As String
    '0=none, 1=ask me, 2 to 10 are valid
    Select Case Index
        Case 2: ViceEXE = "x64"
        Case 3: ViceEXE = "x64sc"
        Case 4: ViceEXE = "x64dtv"
        Case 5: ViceEXE = "x128"
        Case 6: ViceEXE = "xvic"
        Case 7: ViceEXE = "xpet"
        Case 8: ViceEXE = "xcbm2"
        Case 9: ViceEXE = "xcbm5x0"
        Case 10: ViceEXE = "xplus4"
        Case Else: ViceEXE = ""
    End Select
End Function

'---- Converts a Commodore file Load Address to associated computer family or model
Public Function GetMachine(ByVal LA As Long) As Integer
    Dim n As Integer
    
    Select Case LA
        Case 2049:       n = 2 'C64
        '                n = 3 'C64sc as of VICE 2.3
        '                n = 4 'C64DTV
        Case 7169:       n = 5 'C128 Basic 7 [Also? 16385 - C128 mode++]
        Case 4097, 4609: n = 6 'Vic20
        Case 1024, 1025: n = 7 'PET (1025 conflict with VIC-20 +3K)
        Case 3:          n = 8 'CBM2
        'Case 3:         N = 9 'CBM2 P500 as of VICE 2.4
        Case 8193:       n = 10 'C16/Plus4 (Also 4097 which conflicts with VIC-20)  [Also? 8193 - Plus/4-C16++]
        Case Else:       n = 0 'Unknown
    End Select
    
    GetMachine = n
End Function

'---- Converts Commodore DOS Disk ID to model number string
Public Function DiskID(ByVal Tmp As String) As String
    Select Case UCase(Right(Tmp, 2))
        Case "2A": DiskID = "1540/1541/1570/1571"
        Case "2C": DiskID = "8050/8250/SFD"
        Case "3D", "1D": DiskID = "1581"
        Case Else: DiskID = Qu & Tmp & Qu & " is an unknown ID!"
    End Select
End Function

'---- Opens the specified File with associated application (ie: notepad for TXT files)
Public Sub ViewFile(ByVal Filename As String)
    Dim hWnd As Long
    
    ShellExecute hWnd, "open", Filename, vbNullString, ExeDir, 1
End Sub

'---- Makes filename given template. Uses *, % and ^ as substitution characters
Public Function BatchName(ByVal Num As Integer, ByVal Side As Integer, FStr As String) As String
    Dim p As Integer, p2 As Integer, L As Integer
    
    BatchName = FStr
    L = Len(FStr): p = InStr(1, FStr, "#"):    If p = 0 Then Exit Function
    
    p2 = 1
    Do While p2 < L
        If Mid(FStr, p + p2, 1) <> "#" Then Exit Do
        p2 = p2 + 1
    Loop
    
    Mid(BatchName, p, p2) = Right("000000" & Format(Num), p2)
    p = InStr(1, BatchName, "*"): If p > 0 Then Mid(BatchName, p, 1) = Format(Side)
    p = InStr(1, BatchName, "%"): If p > 0 Then Mid(BatchName, p, 1) = Chr(96 + Side)
    p = InStr(1, BatchName, "^"): If p > 0 Then Mid(BatchName, p, 1) = Chr(64 + Side)
End Function

'---- Retrieve Named Field 'FS' from string
'String contains multiple <CR> delimited lines (could be an entire text file)
'Note: string to search must end with <CR>!
Public Function GetNamedField(ByVal Tmp As String, FS As String) As String
    Dim p As Integer, PP As Integer, L As Integer, Tmp2 As String
        
    L = Len(FS)                 'Length of Field String
    p = InStr(1, Tmp, FS)       'Look for the string
    Tmp2 = ""
    
    If p > 0 Then
        p2 = InStr(p + L, Tmp, Cr) 'Now look for carriage return
        If p2 > 0 Then Tmp2 = Mid(Tmp, p + L, p2 - p - L)
    End If
    
    GetNamedField = Tmp2
    
End Function

'---- Get Named Value
' Uses GetNamedField then converts to double
Public Function GetNamedV(ByVal Tmp As String, FS As String) As Double
    Dim Tmp2 As String
    
    Tmp2 = GetNamedField(Tmp, FS)           'Get Named Field as string
    GetNamedV = Val(Tmp2)                   'Convert it to a Value
End Function

'---- Retrieve Field number 'n' from record string 'Tmp'. Record is comma-delimited
' There must not be any commas in a field. It treats a NULL string between two commas as a NULL field.
Public Function GetField(ByVal Tmp As String, n As Integer) As String
    Dim C As Integer, p As Integer, p2 As Integer, Comma As String, T2 As String
    
    Comma = ",": p2 = 1: C = 1

    Do
        p = InStr(p2, Tmp, Comma)           'Look for the Comma
        If p = 0 Then Exit Do               'None, then exit
        If p > 0 And C = n Then Exit Do     'We found the last record (no comma after it)
        p2 = p + 1: C = C + 1               'Move the start, count the comma
    Loop
    
    If p = 0 Then T2 = Mid(Tmp, p2) Else T2 = Mid(Tmp, p2, p - p2) 'Extract record
    GetField = T2                           'Return the string
    
End Function

'---- Retrieve Field number 'n' from delimited record string 'Tmp'.
' Delimiter is passed to function. If Delimiter is null then TAB will be used
' Note: There MAY be multiple TABs between fields!
Public Function GetDField(ByVal Tmp As String, Delim As String, n As Integer) As String
    Dim C As Integer, p As Integer, p2 As Integer, T2 As String
    
    T2 = ""
    If Delim = "" Then Delim = Chr(9)       'Use TAB delimiter if not specified
    p2 = 1: C = 1                           'String starts at position 1 and is Field#1
    
    Do
        p = InStr(p2, Tmp, Delim)           'Look for the TAB
        If p = 0 Then Exit Do               'None, then exit
        If p > p2 Then
            If C = n Then Exit Do           'We found the last record (no TAB after it)
            p2 = p + 1: C = C + 1           'Move the start, increment the Field#
        Else
            p2 = p + 1                      'if p=p2 then we found two delimiters together, so we increment pointer but not field#
        End If
    Loop
    
    If p = 0 Then T2 = Mid(Tmp, p2) Else T2 = Mid(Tmp, p2, p - p2) 'Extract record
    
    GetDField = T2                          'Return the string

End Function

'---- Return C64 Colour
Public Function C64Colour(ByVal n As Integer) As Long
    
    Select Case n
        Case 0: C64Colour = RGB(0, 0, 0)
        Case 1: C64Colour = RGB(255, 255, 255)
        Case 2: C64Colour = RGB(255, 0, 0)
        Case 3: C64Colour = RGB(0, 255, 255)
        Case 4: C64Colour = RGB(255, 0, 255)
        Case 5: C64Colour = RGB(0, 255, 0)
        Case 6: C64Colour = RGB(0, 0, 255)
        Case 7: C64Colour = RGB(255, 255, 0)
        Case 8: C64Colour = RGB(255, 102, 0)
        Case 9: C64Colour = RGB(170, 68, 0)
        Case 10: C64Colour = RGB(255, 119, 119)
        Case 11: C64Colour = RGB(85, 85, 85)
        Case 12: C64Colour = RGB(136, 136, 136)
        Case 13: C64Colour = RGB(153, 255, 153)
        Case 14: C64Colour = RGB(153, 153, 255)
        Case 15: C64Colour = RGB(187, 187, 187)
        Case Else: C64Colour = 0
    End Select

End Function
