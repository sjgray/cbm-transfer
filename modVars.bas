Attribute VB_Name = "modVars"
' CBM-Transfer - Copyright (C) 2007-2017 Steve J. Gray
' ====================================================
'
' modVars - Module with Global Variables and Data Types
'
Option Explicit

Public Cr As String, LF As String, Qu As String  'Carriage Return, Linefeed and Quote
Public Nu As String, Hx As String                'Null character. Hex prefix for conversion
Public MsgTitle As String                        'Title for MsgBox popup windows

'--- Directories
Public ExeDir           As String
Public CurDir           As String
Public LocalDir(1)      As String

'--- Mode and Program Variables
Public SrcMode          As Integer      'Selected TAB on LEFT
Public DstMode          As Integer      'Selected TAB on RIGHT
Public TEMPFILE1 As String, TEMPFILE2 As String, TEMPFILE3 As String
Public LogFile          As String
Public DDFile(1)        As String
Public PathFile         As String
Public PathHistory      As Boolean
Public LastCMDError     As String
Public KillFlag         As Boolean      'Global Kill flag to abort processes (experimental)

'--- Global Variables for Passing data
Public Response         As String       'String Returned from PROMPT dialog form
Public PickedColour     As Long         'Colour Value
Public Layout           As Integer
Public Layout2          As Integer
Public MenuNum          As Integer      'The index for the dropdown menu's list 0=left, 1=right

'=== CONFIG OPTIONS
'--- General Options
Public AutoRefreshDir   As Boolean
Public PreviewCheck     As Boolean
Public ConfirmD64       As Boolean
Public P00Flag          As Boolean
Public IgnoreD          As Boolean
Public CheckEXE         As Boolean
Public LogAll           As Boolean
Public StartDAD         As Boolean
Public IgnoreBadID      As Boolean

'--- Path Options
Public UseLP            As Boolean

'--- X-Cable Options
Public NoWarpString     As String
Public TransferString   As String
Public DriveNum         As Integer      'XCable DriveNum

'--- CBMLink Options
Public CBMUnit          As Integer
Public CBMDrive         As Integer      'CBMLink Drive and Unit Number
Public LinkCStr         As String       'Link Connection string 'example: -c serial 19200,com1 -d 8

'--- Vice Options
Public UseVice          As Boolean
Public VicePath         As String

'--- NIB Options
Public UseNIB           As Boolean
Public UseNBZ           As Boolean
Public NIBstr           As String
Public CreateNIB        As Boolean
Public CreateG64        As Boolean
Public CreateD64        As Boolean
Public WriteD64         As Boolean
Public UseNibCustom     As Boolean

'--- Filename Options
Public FNMode           As Integer
Public FNEdit           As Boolean
Public FNChr            As String

'--- Batch Imaging Options
Public UseBatch         As Boolean
Public BatchMode        As Integer
Public BatchFilename    As String
Public Batch2Sided      As Boolean
Public DiskNum          As Integer
Public DiskSide         As Integer

'--- Font Options
Public UseCBMFont       As Boolean

'--- Disk Image Parameters
Type DskImg
    FileSize    As Long         'Length of Disk Image File
    Desc        As String * 40  'Description (drives that use this format)
    SectSize    As Integer      'Sector Size (usually 256)
    SectMin     As Integer      'Min sectors/track
    SectMax     As Integer      'Max sectors/track
    SectMap     As String * 80  'Number of sectors per track
    HeaderT     As Integer      'Header Track
    HeaderS     As Integer      'Header Sector
    HeaderPos   As Integer      'Header Position
    DirT        As Integer      'Directory Track
    DirS        As Integer      'Directory Sector
    DirSize     As Integer      'Max directory sectors
    BAMT        As Integer      'BAM Track
    BAMS        As Integer      'BAM Sector
    BAMPos      As Integer      'Start byte for BAM
    BAMSize     As Integer      'Max BAM sectors
    MaxTrack    As Integer      'Max Track#
    MaxFiles    As Integer      'Max File Entries
    MaxSize     As Integer      'Max File Size for entire disk not including Err map
    MaxErr      As Integer      'Max Error Map Size
End Type

'---- BAM Entries (2 different formats)
Type BAM4Type                   '-- DOS TYPE A (1541,4040 etc) - First BAM at position 4
    TotFree     As String * 1   'Total Blocks free in track
    Map         As String * 3   'Allocation bits for track - 3*8=24 sectors max (0=USED,1=FREE)
End Type

Type BAM5Type                   '-- DOS TYPE C (8050, 8250)
    TotFree     As String * 1   'Total Blocks free in track
    Map         As String * 4   'Allocation bits for track - 4*8=32 sectors max (0=USED,1=FREE)
End Type

'---- Disk Headers (two different formats)
Type Header1Type                '-- DOS TYPE A (4040, 1541)
    FName       As String * 16  'Header Name padded with shift-space (160)
    ID          As String * 2   'Disk ID
    Unused      As String * 1   'Unused
    DOSVer      As String * 2   'DOS Version "2a"
End Type

Type Header2Type                '-- DOS TYPE C (8050, 8250)
    DOSVer2     As String * 1   'DOS Version "c"
    Unused      As String * 3   'Unused
    FName       As String * 16  'Header Name padded with shift-space (160)
    Unused2     As String * 2   'Unused
    ID          As String * 2   'Disk ID
    Unused3     As String * 1   'Unused
    DOSVer      As String * 2   'DOS Version "2c"
End Type

Type Header3Type                '-- DOS TYPE D (1581)
    DOSVer2     As String * 1   'DOS Version "d"
    Unused      As String * 1   'Unused
    FName       As String * 16  'Header Name padded with shift-space (160)
    Unused2     As String * 2   'Unused
    ID          As String * 2   'Disk ID
    Unused3     As String * 1   'Unused
    DOSVer      As String * 2   'DOS Version "3d"
End Type

'---- Directory Entry Structure
Type DirEntryType
    LinkT       As String * 1   'Link to next directory Track  (first entry of sector only, otherwise 0)
    LinkS       As String * 1   'Link to next directory Sector (first entry of sector only, otherwise 0)
    FType       As String * 1   'File Type (DEL,SEQ,PRG,USR,REL)
    FirstT      As String * 1   'Link to file TRACK
    FirstS      As String * 1   'Link to file SECTOR
    FName       As String * 16  'Filename padded with Shift-space ($60)
    RelSSTrk    As String * 1   'Relative File Side-Sector TRACK link
    RelSSSect   As String * 1   'Relative File Side-Sector SECTOR link
    RelLen      As String * 1   'Relative File Length
    Unused      As String * 6   'Unused
    FSizeLO     As String * 1   'File Size LO
    FSizeHI     As String * 1   'File Size HI
End Type

