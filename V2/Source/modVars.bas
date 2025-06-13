Attribute VB_Name = "modVars"
' CBM-Transfer - Copyright (C) 2007-2021 Steve J. Gray
' ====================================================
'
' modVars - Module with Global Variables and Data Types
'
Option Explicit

Public Cr As String, LF As String, Qu As String                         'Carriage Return, Linefeed and Quote
Public Nu As String, Hx As String                                       'Null character. Hex prefix for conversion
Public MsgTitle         As String                                       'Title for MsgBox popup windows

Public Theme            As Integer                                      'Colour Theme

Public ThemeFG          As Long                                         'Window   Foreground Colour
Public ThemeBG          As Long                                         'Window   Background Colour
Public ThemeTitleFG     As Long                                         'Title    Foreground Colour
Public ThemeTitleBG     As Long                                         'Title    Background Colour
Public ThemeMenuFG      As Long                                         'Menu     Foreground Colour
Public ThemeMenuBG      As Long                                         'Mwnu     Background Colour
Public ThemeListFG      As Long                                         'Listing  Foreground Colour
Public ThemeListBG      As Long                                         'Listing  Background Colour
Public ThemeFrFG        As Long                                         'Frame    Foreground Colour
Public ThemeFrBG        As Long                                         'Frame    Background Colour
Public ThemeFr2FG       As Long                                         'Frame2   Foreground Colour
Public ThemeFr2BG       As Long                                         'Frame2   Background Colour

'--- Global Stuff

Public DiskName(3)     As String                                        'Storage for Disk Names
Public DiskID(3)       As String                                        'Storage for Disk IDs

'--- Commandline strings and paths

Public CBMOpen          As String                                       'Path to OpenCBM
Public CBMCtrl          As String                                       'String for CBMCTRL command string name
Public CBMCopy          As String                                       'String for CBMCOPY command string name
Public CBMFormat        As String                                       'String for FORMAT command (cbmforng)
Public CBMC1541         As String                                       'String for C1541 command string name
Public CBMVICE          As String                                       'Path to VICE
Public CBMNib           As String                                       'Path to NIBTOOLS
Public CBMLink          As String                                       'String for CBMLINK command string name
Public CBMAcme          As String                                       'String for ACME command string name
Public CBMMD5           As String                                       'String for MD5 command string name

Public CMDSTR           As String                                       'String for COMMAND parameter string

'--- Directories

Public ExeDir           As String                                       'Path to Executable
Public CurDir           As String                                       'Current Directory
Public LocalDir(1)      As String                                       'Local PC Directories
Public ThemeDir         As String                                       'Theme Directory
Public INIFile          As String

'--- Mode and Program Variables

Public SrcMode          As Integer                                      'Selected TAB on LEFT
Public DstMode          As Integer                                      'Selected TAB on RIGHT
Public LogFile          As String                                       'Log File path
Public CatalogFile      As String                                       'Catalog file
Public HistoryFile      As String                                       'Path History file
Public AddPathFlag      As Boolean                                      'Setting to auto Add Path to History
Public LastCMDError     As String
Public KillFlag         As Boolean                                      'Global Kill flag to abort processes (experimental)

Public DDFile(1)        As String

Public TEMPFILE1 As String, TEMPFILE2 As String, TEMPFILE3 As String

'--- Global Variables for Passing data

Public Response         As String                                       'String Returned from PROMPT dialog form
Public PickedColour     As Long                                         'Colour Value
Public Layout           As Integer                                      'GUI Layout
Public Layout2          As Integer                                      'GUI Layout
Public MenuNum          As Integer                                      'The index for the dropdown menu's list 0=left, 1=right
Public MenuForm         As Integer                                      'The target Form for the menu

'=== CONFIG OPTIONS

'--- General Options

Public AutoRefreshDir   As Boolean
Public PreviewCheck     As Boolean
Public ConfirmD64       As Boolean
Public P00Flag          As Boolean
Public IgnoreD          As Boolean
Public LogAll           As Boolean
Public StartDAD         As Boolean
Public IgnoreBadID      As Boolean

'--- Path Options

Public UseLP            As Boolean

'--- X-Cable Options

Public NoWarpString     As String
Public TransferString   As String
Public DriveNum         As Integer                                      'XCable DriveNum
Public UseFirstDrive    As Boolean                                      'XCable Use first drive found on first scan

'--- CBMLink Options

Public LinkUnit         As Integer
Public LinkDrive        As Integer                                      'CBMLink Drive and Unit Number
Public LinkCStr         As String                                       'Link Connection string 'example: -c serial 19200,com1 -d 8

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
Public NIBPrompt        As Boolean

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
Public CBMFontName      As String
Public CBMFontSize      As Integer

'--- Utility Paths

Public UPath(4)         As String                                       'Utility PATH strings

'---- General Common Variables

Public Pow(7)           As Integer                                      'binary powers array
