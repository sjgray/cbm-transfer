Attribute VB_Name = "modCommon"
' CBM-Transfer - Copyright (C) 2007-2021 Steve J. Gray
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
' FileNameBase..... Returns the Filename without path or extension
' PathOnly......... Returns the PATH when supplied either a valid directory path, or a filename
' PathUp........... Returns the parent PATH
' AddSlash......... Adds a slash to the end of a Path if it does not have one
' NoSlash.......... Removes trailing slash
' KillFile......... Checks if a file exists, then deletes it
' KillTemp......... Deletes files in the TEMP directory
'
' SupportedImg..... Checks if known Disk Image type (ie: D64,D80, etc)
' SupportedExt..... Checks if Extension is a valid CBM file (PRG,SEQ,ROM,BIN,or NO EXTENSION)
'
' DOSName.......... Returns DOS Name (FILENAME.EXT) from CBM Directory Listing
' DOSExt........... Returns CBM file extension (PRG,SEQ,USR,REL,DEL,CBM)
'
' CBMName.......... Returns CBM Name (FILENAME,P) from CBM Directory Listing
' CBMExt........... Converts DOSExt to CBM File type (PRG=,p  SEQ=,s  USR=,u  REL=,r  DEL=,d CBM=,c)
' CBMType.......... Extracts the CBM Type (",p" , ",s" etc ) from CBM filename ("filename,p")
'
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
' MyBin............ Convert byte to 8-bit binary
' MyTrim........... Removes leading and trailing spaces
' MyRGB............ Returns LONG RGB value given R,G,B 2-digit hex strings
' GetBrowseDir..... Prompt for new directory using "Browse for Folder" popup
' GetLoadAddress... Reads CBM Load Address from specified file
' ViceEXE.......... Returns VICE executable from list
' GetMachine....... Finds target machine using Commodore file Load Address
' DriveModel....... Converts Commodore DOS Disk ID to model number string
' ViewFile......... Opens app associated with the file type (ie: Notepad for TXT files)
' BatchName........ Makes filename given template. Uses *, % and ^ as substitution characters
' GetNamedField.... Retrieve Named Field from <CR> delimited Record
' GetNamedV........ Retrieve Named Field as above, and convert to numeric value
' GetField......... Retrieve Field from Comma-Delimited Record
' GetDField........ Retrieve Field from Delimited Record with specified Delimiter (default=TAB)
' GetVNameU........ Return Variable Name in UPPERCASE, from string in format: "variable=value"
' GetVstr.......... Return Value from string as above
' GetCharWidth..... Return Width based on character width index
' C64Colour........ Return the RGB value for specified C64 colour
' Pad.............. Pad a string to sepecified length
' Warning.......... Display a numbered warning
' SetAllPaths...... Sets All Paths to Utilities (Calls MakeUPath)
' MakeUPath........ Returns combined path and commandline exe string
' LoadBuffer....... Load file contents to Buffer string
' ASCIItoScreen.... Convert standard ASCII to Screen Code

'---- Convert ASCII to SCREEN CODE
Public Function ASCIItoScreen(ByVal V As Integer) As Integer
            
    Select Case V                                                       'Convert ASCII to screen code
        Case 64: V = 0                                                  '@
        Case 65 To 90 'V = V - 64                                       'A..Z - no change
        Case 91 - 95: V = V - 64                                        '[\]^_
        Case 96 To 126: V = V - 96                                      'a..z
    End Select
    
    ASCIItoScreen = V                                                   'Return the screen code

End Function

'---- MessageBox popup with default title
Public Sub MyMsg(ByVal Tmp As String)
    
    MsgBox Tmp, , MsgTitle

End Sub

'---- Set All Paths
' Uses the current cbmxfer.exe path as "current" path and then
' sets Utility paths accordingly (CBM-Transfer directory or User-Specified directory)
Public Sub SetAllPaths()

    CMDSTR = "command "
    
    CBMOpen = MakeUPath(0, "")              'PATH to OpenCBM    - Utilities for X-cables and Zoom Floppy
    CBMCtrl = MakeUPath(0, "cbmctrl")       'CBMCTRL.EXE        - Part of OpenCBM to talk to real drives via x-cable
    CBMCopy = MakeUPath(0, "cbmcopy")       'CBMCOPY.EXE        - Part of OpenCBM to copy files
    CBMFormat = MakeUPath(0, "cbmforng")    'CBMFORNG.EXE       - Part of OpenCBM to Fast-Format 1541 drives
    CBMC1541 = MakeUPath(1, "c1541")        'C1541.EXE          - Part of VICE to work with Disk Images
    CBMVICE = MakeUPath(1, "")              'PATH TO VICE       - To start selected Emulator
    CBMNib = MakeUPath(2, "")               'PATH TO NIBTOOLS   - Read and Write Protected Disks
    CBMLink = MakeUPath(3, "cbmlink")       'CBMLINK.EXE        - Talk to CBM Drives via CBM Computer
    CBMAcme = MakeUPath(4, "acme")          'ACME.EXE           - A 6502-Family Assembler
    CBMMD5 = MakeUPath(4, "md5")            'MD5.EXE            - Calculates 'unique' ID's for binary files
    
End Sub

'---- Make Utility Path from Config and EXE name
' Adds path to utility string. If path is empty then uses CBM-Transfer directory
' Adds .EXE to end of command if not included.
' Set EXEStr="" to just return the path (for utilites with multiple EXE)

Public Function MakeUPath(ByVal N As Integer, EXEStr As String) As String
    Dim Tmp As String

    If EXEStr <> "" Then
        Tmp = ".exe"
        If LCase(Right(EXEStr, 4)) <> Tmp Then EXEStr = EXEStr & Tmp    'make sure command has .exe at the end
    End If
    
    Tmp = UPath(N)                                                      'Get the user-specified path string
    If Tmp = "" Then Tmp = ExeDir                                       'If it is BLANK then use CBM-Transfer directory
    MakeUPath = AddSlash(Tmp) & EXEStr                                  'Return string with path/cmdname.exe

End Function

'---- Changes current path
Public Sub MyChDir(ByVal Path As String)
    On Local Error GoTo MCDErr
    ChDrive Left(Path, 1)
    ChDir Path
MCDErr:
End Sub

'---- Check if Directory exists. If so, returns TRUE
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
    Dim P As Integer
    
    If Right(Filename, 1) = Qu Then Filename = Left(Filename, Len(Filename) - 1)
    P = InStrRev(Filename, ".")
    If P > 0 Then FileExt = Mid(Filename, P + 1) Else FileExt = ""
End Function

'---- Return file Extension Uppercased
Public Function FileExtU(ByVal Filename As String) As String
    Dim P As Integer
    If Right(Filename, 1) = Qu Then Filename = Left(Filename, Len(Filename) - 1)
    P = InStrRev(Filename, ".")
    If P > 0 Then FileExtU = UCase(Mid(Filename, P + 1)) Else FileExtU = ""
    
End Function

'---- Return Path without filename
Public Function FilePath(ByVal Filename As String) As String
    Dim P As Integer
    
    P = InStrRev(Filename, "\")
    If P > 0 Then FilePath = Left(Filename, P - 1) Else FilePath = ""
    
End Function

'---- Return Filename without Extension (do not remove path if included)
Public Function FileBase(ByVal Filename As String) As String
    Dim P As Integer
    
    P = InStrRev(Filename, ".")
    If P > 0 Then FileBase = Left(Filename, P - 1) Else FileBase = Filename
    
End Function

'---- Return Filename without Path
Public Function FileNameOnly(ByVal Filename As String) As String
    Dim P As Integer
    
    P = InStrRev(Filename, "\")
    If P > 0 Then FileNameOnly = Mid(Filename, P + 1) Else FileNameOnly = Filename
    
End Function

'---- Return Filename without Path and Extension
Public Function FileNameBase(ByVal Filename As String) As String
    Dim P As Integer, Tmp As String
    
    Tmp = Filename
    P = InStrRev(Tmp, "\"): If P > 0 Then Tmp = Mid(Tmp, P + 1)         'Find the LAST "\"
    P = InStrRev(Tmp, "."): If P > 0 Then Tmp = Left(Tmp, P - 1)        'Find the LAST "."
    FileNameBase = Tmp
    
End Function


'---- Return PATH only
' Tmp must contain a path.
' - If Tmp is a DIRECTORY, then it is returned
' - If Tmp is a FILE, then it's path is returned
' Returns path without ending "\"
Public Function PathOnly(ByVal FileSpec As String) As String
    Dim P As Integer, Tmp As String
    
    P = InStr(FileSpec, ".")
    If P = 0 Then
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
    Dim P As Integer
    
    PathUp = Path
    If Len(Path) > 3 Then
        P = InStrRev(Path, "\", Len(Path) - 1)
        If P > 0 Then PathUp = Left(Path, P)
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
        Case "D64", "D71", "D80", "D81", "D82", "G64", "G71", "X64": SupportedImg = True
        Case "NIB", "NBZ": If WriteFlag = True Then SupportedImg = True Else SupportedImg = False
        Case "D1M", "D2M", "D4M": If WriteFlag = True Then SupportedImg = False Else SupportedImg = True
        Case Else: SupportedImg = False
    End Select
    
End Function

'---- Check if supported Image Extension
' Flag: False=reading, True=writing
Public Function SupportedExt(ByVal Ext As String) As Boolean
    
    Select Case UCase(Ext)
        Case "PRG", "SEQ", "USR", "ROM", "BIN", "": SupportedExt = True
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
    
    Tmp = UCase(Str)
    If Left(Tmp, 1) = "." Then Tmp = Mid(Tmp, 2)                'Remove "." if included and convert to uppercase
    
    Select Case Tmp
        Case "PRG", "P00", "P01": CBMExt = ",p"
        Case "SEQ", "S00", "S01": CBMExt = ",s"
        Case "USR", "U00", "U01": CBMExt = ",u"
        Case "REL", "R00", "R01": CBMExt = ",r"
        Case "CBM": CBMExt = ",c"
        Case "DEL": CBMExt = ",d"
        Case Else: CBMExt = ",p"                                'Default for files with no extension
    End Select
End Function

'---- Extract type from full CBM NAME
' ie: "filename,p" -> ",p"
'     "filename,s" -> ",s"
Public Function CBMType(ByVal Str As String) As String
    Dim P As Integer
    
    P = InStr(1, Str, ",")
    If P > 0 Then CBMType = Mid(Str, P) Else CBMType = ""
    
End Function

'---- Reverse Case of PETSCII Text (Mostly for original PET BASIC 1 text strings
Public Function Reverse(ByVal N As Integer) As Integer
    Select Case N
        Case 65 To 90: N = N + 32
        Case 97 To 122: N = N - 32
    End Select
    
    Reverse = N
End Function

'---- Validate PC filename - Check for invalid characters
' Any character outside the range SPACE to TILDA is invalid
' Any character in the BAD list is invalid
Public Function CheckPCFilename(ByVal Filename As String) As Boolean
    Dim Bad As String, Tmp As String, Flag As String
    Dim J As Integer
    
    CheckPCFilename = True                                              'Assume all okay
        
    If Left(Filename, 1) = " " Then CheckPCFilename = False: Exit Function
    
    Bad = "/\:*?<>|" & Qu                                               'Invalid characters
    Flag = True                                                         'Assume filename is valid
    
    For J = 1 To Len(Filename)
        Tmp = Mid(Filename, J, 1)
        If Tmp < " " Or Tmp > "~" Then Flag = False: Exit For            'Check outside range
        If InStr(1, Bad, Tmp) > 0 Then Flag = False: Exit For            'Check invalid characters
    Next J

    CheckPCFilename = Flag                                               'Return result
    
End Function

'---- Fix CBM Filename to be PC friendly according to Option MODE - updated feb 1/2011
' This will automatically fix filenames, or it will prompt for manual entry
Public Function MakePCName(ByVal Filename As String) As String
    Dim J As Integer, EdFlag As Boolean, OldName As String
    
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
    Dim J As Integer, Tmp As String, Tmp2 As String, Bad As String, Flag As Boolean, CFlag As Boolean
    
    Tmp = "": Bad = "/\:*?<>|" & Qu                                         'String of Invallid characters
    Flag = False                                                            'Flag for spaces at beginning of filename
    
    For J = 1 To Len(Filename)
        Tmp2 = Mid(Filename, J, 1)                                          'Get one character of filename
        CFlag = True                                                        'Assume it is valid
        If Tmp2 <> " " Or Flag = True Then
            If Tmp2 < " " Or Tmp2 > "~" Then CFlag = False                  'Invalidate outside normal alpha-numeric range
            If InStr(1, Bad, Tmp2) > 0 Then CFlag = False                   'Invalidate specific characters
            If CFlag = True Then Tmp = Tmp & Tmp2 Else Tmp = Tmp & RStr     'If character valid then add it, otherwise use replacement
            Flag = True                                                     'First non-space will set Flag=true
        End If
    Next J

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

'---- Convert fixed-length HEX string to a decimal value
Function MyDec(ByVal H As String) As Long
    On Local Error Resume Next
    MyDec = CLng(Hx & H)
End Function

'---- Convert decimal value to fixed-length HEX value with leading zeros
' D= Number of digits. -D adds a $ to the front
Function MyHex(ByVal N As Single, D As Integer) As String
    Dim Tmp As String
    
    If D < 0 Then Tmp = "$"
    MyHex = Tmp & Right("00000000" & Hex(N), Abs(D))
End Function

'---- Convert decimal value to fixed-length HEX value with leading zeros
Function MyBin(ByVal D As Integer) As String
    Dim Tmp As String, i As Integer
    
    If D > 255 Then MyBin = "--------": Exit Function
    
    For i = 7 To 0 Step -1
        If (D And (2 ^ i)) > 0 Then Tmp = Tmp & "1" Else Tmp = Tmp & "0"
    Next i
    
    MyBin = Tmp
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
'Index: 0=none, 1=ask me, 2 to 11=Specific Emulator
Public Function ViceEXE(ByVal Index As Integer) As String
    
    Select Case Index
        Case 2: ViceEXE = "x64"
        Case 3: ViceEXE = "x64sc"
        Case 4: ViceEXE = "x64dtv"
        Case 5: ViceEXE = "x128"
        Case 6: ViceEXE = "xvic"
        Case 7: ViceEXE = "xscpu64"
        Case 8: ViceEXE = "xcbm2"
        Case 9: ViceEXE = "xcbm5x0"
        Case 10: ViceEXE = "xplus4"
        Case 11: ViceEXE = "xpet"
        Case Else: ViceEXE = ""
    End Select
    
End Function

'---- Converts a Commodore file Load Address to associated computer family or model
Public Function GetMachine(ByVal LA As Long) As Integer
    Dim N As Integer
    
    Select Case LA
        Case 2049:       N = 2  'C64
        '                n = 3  'C64sc as of VICE 2.3
        '                n = 4  'C64DTV
        Case 7169:       N = 5  'C128 Basic 7 [Also? 16385 - C128 mode++]
        Case 4097, 4609: N = 6  'Vic20
        '                N = 7  'SuperCPU
        Case 3:          N = 8  'CBM2
        'Case 3:         N = 9  'CBM2 P500 as of VICE 2.4
        Case 8193:       N = 10 'C16/Plus4 (Also 4097 which conflicts with VIC-20)  [Also? 8193 - Plus/4-C16++]
        Case 1024, 1025: N = 11 'PET (1025 conflict with VIC-20 +3K)
        Case 12289:      N = 0  'CLCD (future)
        Case Else:       N = 0  'Unknown
    End Select
    
    GetMachine = N
End Function

'---- Converts Commodore DOS Disk ID to model number string
Public Function DriveModel(ByVal Tmp As String) As String
    Select Case UCase(Right(Tmp, 2))
        Case "2A": DriveModel = "1540/1541/1570/1571"
        Case "2C": DriveModel = "8050/8250/SFD"
        Case "3D", "1D": DriveModel = "1581"
        Case Else: DriveModel = Qu & Tmp & Qu & " is an unknown ID!"
    End Select
End Function

'---- Opens the specified File with associated application (ie: notepad for TXT files)
Public Sub ViewFile(ByVal Filename As String)
    Dim hWnd As Long
    
    ShellExecute hWnd, "open", Filename, vbNullString, ExeDir, 1
End Sub

'---- Makes filename given template. Uses *, % and ^ as substitution characters
Public Function BatchName(ByVal Num As Integer, ByVal Side As Integer, FStr As String) As String
    Dim P As Integer, P2 As Integer, L As Integer
    
    BatchName = FStr
    L = Len(FStr): P = InStr(1, FStr, "#"):    If P = 0 Then Exit Function
    
    P2 = 1
    Do While P2 < L
        If Mid(FStr, P + P2, 1) <> "#" Then Exit Do
        P2 = P2 + 1
    Loop
    
    Mid(BatchName, P, P2) = Right("000000" & Format(Num), P2)
    P = InStr(1, BatchName, "*"): If P > 0 Then Mid(BatchName, P, 1) = Format(Side)
    P = InStr(1, BatchName, "%"): If P > 0 Then Mid(BatchName, P, 1) = Chr(96 + Side)
    P = InStr(1, BatchName, "^"): If P > 0 Then Mid(BatchName, P, 1) = Chr(64 + Side)
End Function

'---- Retrieve Named Field 'FS' from string
'String contains multiple <CR> delimited lines (could be an entire text file)
'Note: string to search must end with <CR>!
Public Function GetNamedField(ByVal Tmp As String, FS As String) As String
    Dim P As Integer, PP As Integer, L As Integer, Tmp2 As String
        
    L = Len(FS)                 'Length of Field String
    P = InStr(1, Tmp, FS)       'Look for the string
    Tmp2 = ""
    
    If P > 0 Then
        P2 = InStr(P + L, Tmp, Cr) 'Now look for carriage return
        If P2 > 0 Then Tmp2 = Mid(Tmp, P + L, P2 - P - L)
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
Public Function GetField(ByVal Tmp As String, N As Integer) As String
    Dim C As Integer, P As Integer, P2 As Integer, Comma As String, T2 As String
    
    Comma = ",": P2 = 1: C = 1

    Do
        P = InStr(P2, Tmp, Comma)           'Look for the Comma
        If P = 0 Then Exit Do               'None, then exit
        If P > 0 And C = N Then Exit Do     'We found the last record (no comma after it)
        P2 = P + 1: C = C + 1               'Move the start, count the comma
    Loop
    
    If P = 0 Then T2 = Mid(Tmp, P2) Else T2 = Mid(Tmp, P2, P - P2) 'Extract record
    GetField = T2                           'Return the string
    
End Function

'---- Retrieve Field number 'n' from delimited record string 'Tmp'.
' Delimiter is passed to function. If Delimiter is null then TAB will be used
' Note: There MAY be multiple TABs between fields!
Public Function GetDField(ByVal Tmp As String, Delim As String, N As Integer) As String
    Dim C As Integer, P As Integer, P2 As Integer, T2 As String
    
    T2 = ""
    If Delim = "" Then Delim = Chr(9)       'Use TAB delimiter if not specified
    P2 = 1: C = 1                           'String starts at position 1 and is Field#1
    
    Do
        P = InStr(P2, Tmp, Delim)           'Look for the TAB
        If P = 0 Then Exit Do               'None, then exit
        If P > P2 Then
            If C = N Then Exit Do           'We found the last record (no TAB after it)
            P2 = P + 1: C = C + 1           'Move the start, increment the Field#
        Else
            P2 = P + 1                      'if p=p2 then we found two delimiters together, so we increment pointer but not field#
        End If
    Loop
    
    If P = 0 Then T2 = Mid(Tmp, P2) Else T2 = Mid(Tmp, P2, P - P2) 'Extract record
    
    GetDField = T2                          'Return the string

End Function

'--- Return Variable Name
' Str is in format: variable=string
' This returns the variable name in uppercase. Leading and Ending spaces are trimmed.
Public Function GetVNameU(ByVal Str As String) As String
  Dim P As Integer, Tmp As String
  
  GetVNameU = ""
  P = InStr(Str, "="): If P > 1 Then GetVNameU = UCase(MyTrim(Left(Str, P - 1)))
End Function

'--- Return Variable Name
' Str is in format: variable=value
' This returns the value string. Leading and Ending spaces are trimmed.
Public Function GetVstr(ByVal Str As String) As String
  Dim P As Integer, Tmp As String
  
  GetVstr = ""
  P = InStr(Str, "="): If P > 1 Then GetVstr = MyTrim(Mid(Str, P + 1))
End Function


'---- Return C64 Colour
Public Function C64Colour(ByVal N As Integer) As Long
    
    Select Case N
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

'---- Return RGB value from R,G,B 2-digit strings
Public Function MyRGB(ByVal R As String, G As String, B As String) As Long
    Dim RD As Integer, GD As Integer, BD As Integer
    
    RD = MyDec(R)
    GD = MyDec(G)
    BD = MyDec(B)
    MyRGB = RGB(RD, GD, BD)
    
End Function

'---- Pad a string to a specified length. Warning!: Will truncate string if longer than pad length!
Public Function Pad(ByVal S1 As String, L As Integer) As String
    
    Pad = Left(S1 & String(L, " "), L)
    
End Function

'---- Display Warning Message by number
Public Sub Warning(ByVal N As Integer, ByVal Filename As String)
    Dim W As String, Tmp As String, Tmp2 As String
    
    W = "Warning!"
    Tmp = "": Tmp2 = ""
    
    If Filename <> "" Then
        Tmp = "File: " & Filename & Cr & Cr
        'If InStr(1, Filename, ".") > 0 Then
        '    Tmp2 = Cr & Cr & "This file contains a '.' which MAY be a placeholder" & Cr & _
        '    "for a PETSCII character that can not be represented!" & Cr & _
        '    "Such files cannot be transferred!"
        'End If
    End If
        
    Select Case N
        Case 1:
            MsgBox Tmp & "The file could not be extracted from the image." & Tmp2, vbExclamation, W
            
        Case 3:
            MsgBox Tmp & "Could not read source file!" & Tmp2, vbExclamation, W
        Case 4:
            MsgBox Tmp & "Could not extract file from image!" & Tmp2, vbExclamation, W
        Case 5:
            MsgBox Tmp & "Could not transfer file from source drive!" & Tmp2, vbExclamation, W
    End Select
    
End Sub

'---- Load Contents of File into string buffer
' Max 32768 bytes
Public Function LoadBuffer(ByRef Buf As String, ByVal SrcFile As String) As Integer
    Dim FIO As Integer, Tmp As String, VLen As Integer

    LoadBuffer = 0                                                              'Default return value=0
    If SrcFile = "" Then Exit Function
    If Exists(SrcFile) = False Then Exit Function
    
    FIO = FreeFile
    Open SrcFile For Binary As FIO
        VLen = intLOF(FIO)
        If VLen > 32760 Then VLen = 32760                                      'Set to Max
                
        Buf = Input(VLen, FIO)                                                 'Read contents to buffer
        
    Close FIO
    
    LoadBuffer = VLen                                                           'Return length loaded
End Function

