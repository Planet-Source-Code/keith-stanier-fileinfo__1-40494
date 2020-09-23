VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Information"
   ClientHeight    =   6540
   ClientLeft      =   1800
   ClientTop       =   1530
   ClientWidth     =   9510
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9510
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   75
      ScaleHeight     =   2955
      ScaleWidth      =   6675
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   75
      Width           =   6735
      Begin VB.Label lblValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   10
         Left            =   1650
         TabIndex        =   42
         Top             =   2420
         Width           =   4920
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Comments:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   0
         TabIndex        =   41
         Top             =   2420
         Width           =   840
      End
      Begin VB.Label lblValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   40
         Top             =   0
         Width           =   4920
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "File Name:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   840
      End
      Begin VB.Label lblValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   38
         Top             =   240
         Width           =   4920
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "File Date/Size:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   37
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Company Name:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   36
         Top             =   480
         Width           =   1290
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "File Description:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "File Version:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   34
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Internal Name:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   33
         Top             =   1200
         Width           =   1140
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Legal Copyright:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   32
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Original File Name:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   31
         Top             =   1680
         Width           =   1440
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Product Name:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   0
         TabIndex        =   30
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Product Version:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   0
         TabIndex        =   29
         Top             =   2175
         Width           =   1290
      End
      Begin VB.Label lblValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   28
         Top             =   480
         Width           =   4920
      End
      Begin VB.Label lblValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   1680
         TabIndex        =   27
         Top             =   720
         Width           =   4920
      End
      Begin VB.Label lblValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   1680
         TabIndex        =   26
         Top             =   960
         Width           =   4920
      End
      Begin VB.Label lblValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   1680
         TabIndex        =   25
         Top             =   1200
         Width           =   4920
      End
      Begin VB.Label lblValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   1680
         TabIndex        =   24
         Top             =   1440
         Width           =   4920
      End
      Begin VB.Label lblValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   1680
         TabIndex        =   23
         Top             =   1680
         Width           =   4920
      End
      Begin VB.Label lblValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   1680
         TabIndex        =   22
         Top             =   1920
         Width           =   4920
      End
      Begin VB.Label lblValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   1680
         TabIndex        =   21
         Top             =   2175
         Width           =   4920
      End
   End
   Begin VB.Frame fraFileAttributes 
      BackColor       =   &H80000004&
      Caption         =   "Attributes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   150
      TabIndex        =   12
      Top             =   3150
      Width           =   4950
      Begin VB.CommandButton cmdProperties 
         Caption         =   "File Properties"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3225
         TabIndex        =   19
         Top             =   225
         Width           =   1515
      End
      Begin VB.CommandButton cmdSetFileAttr 
         Caption         =   "Set Attributes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3225
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   750
         Width           =   1515
      End
      Begin VB.CheckBox chkNormal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   420
         TabIndex        =   17
         Top             =   300
         Width           =   1245
      End
      Begin VB.CheckBox chkArchive 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Archive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   420
         TabIndex        =   16
         Top             =   570
         Width           =   1245
      End
      Begin VB.CheckBox chkSystem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2025
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   600
         Width           =   1035
      End
      Begin VB.CheckBox chkHidden 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hidden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2025
         TabIndex        =   14
         Top             =   300
         Width           =   1035
      End
      Begin VB.CheckBox chkReadOnly 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Read Only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   420
         TabIndex        =   13
         Top             =   860
         Width           =   1245
      End
   End
   Begin VB.Frame fraFileTimeStamp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Date && Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   150
      TabIndex        =   3
      Top             =   4500
      Width           =   4950
      Begin VB.CommandButton cmdChangeLastModifiedTime 
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3825
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Change Last Modified Time"
         Top             =   675
         Width           =   795
      End
      Begin VB.CommandButton cmdChangeCreatedTime 
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3825
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Change Created Time"
         Top             =   300
         Width           =   795
      End
      Begin VB.TextBox txtLastAccessed 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1875
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1050
         Width           =   1215
      End
      Begin VB.TextBox txtLastModified 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1875
         MaxLength       =   16
         TabIndex        =   5
         Top             =   675
         Width           =   1740
      End
      Begin VB.TextBox txtCreated 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1875
         MaxLength       =   16
         TabIndex        =   4
         Top             =   300
         Width           =   1740
      End
      Begin VB.Label lblLastAccessed 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Last Accessed:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   375
         TabIndex        =   11
         Top             =   1125
         Width           =   1425
      End
      Begin VB.Label lblLastModified 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Last Modified:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   390
         TabIndex        =   10
         Top             =   675
         Width           =   1305
      End
      Begin VB.Label lblCreated 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Created:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   420
         TabIndex        =   9
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Hidden          =   -1  'True
      Left            =   5250
      System          =   -1  'True
      TabIndex        =   2
      Top             =   3225
      Width           =   4065
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   6825
      TabIndex        =   1
      Top             =   600
      Width           =   2490
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   6825
      TabIndex        =   0
      Top             =   150
      Width           =   2490
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FName As String
Dim mResult As Long
Dim File As String
Dim RetVal As Long
Dim Char As String
Dim lSizeof As Long
Dim Verbuf() As Byte
Dim Fressize As Long
Dim mFileAttr As Long

' OpenFile() Flags
Const OF_READ = &H0
Const OF_WRITE = &H1
Const OF_READWRITE = &H2
Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40
Const OF_PARSE = &H100
Const OF_DELETE = &H200
Const OF_VERIFY = &H400
Const OF_CANCEL = &H800
Const OF_CREATE = &H1000
Const OF_PROMPT = &H2000
Const OF_EXIST = &H4000
Const OF_REOPEN = &H8000

Const OFS_MAXPATHNAME = 128

Const GENERIC_ALL = &H10000000
Const GENERIC_EXECUTE = &H20000000
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000
Const OPEN_EXISTING = 3

Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_SYSTEM = &H4          ' X - Cannot be set
Const FILE_ATTRIBUTE_DIRECTORY = &H10      ' X - CreateDirectory to set
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_TEMPORARY = &H100
Const FILE_ATTRIBUTE_COMPRESSED = &H800    ' X -

Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400

Const LOCALE_SYSTEM_DEFAULT = &H800
Const LOCALE_USER_DEFAULT = &H400

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

' OpenFile() Structure
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName As String * 64
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName As String * 64
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileSpec As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal mHandle As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal mHandle As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileSpec As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetBinaryType Lib "kernel32" (ByVal szFileName As String, fType As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Private Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As Long, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
Private Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As Long, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function GetLocaleInfo& Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long)
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Sub GetSystemTimeAdjustment Lib "kernel32" (lpTimeAdjustment As Long, lpTimeIncrement As Long, lpTimeAdjustmentDisabled As Long)
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Byte, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Private Declare Sub agCopyData Lib "apigid32.dll" (source As Any, dest As Any, ByVal nCount&)

Private Sub cmdProperties_Click()

If File1.filename = "" Then
    MsgBox "No file selected yet", 48, "File Information Error"
    Exit Sub
End If
RetVal = ShowFileProperties(File, Me.hwnd)
If RetVal <= 32 Then MsgBox "Error", 48, "File Information Error"

End Sub

Private Sub Dir1_Change()

File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()

Dir1.Path = Drive1.Drive

End Sub

Private Sub File1_Click()

Dim Fressize As Long
Dim Freshnd As Long

File = Dir1.Path
If Right(File, 1) <> "\" Then
    File = File & "\"
End If
File = File & File1.List(File1.ListIndex)
Fressize = GetFileVersionInfoSize(File, Freshnd)
If Fressize > 64000 Then Fressize = 64000
ReDim Verbuf(Fressize + 1)
RetVal = GetFileVersionInfo(File, Freshnd, Fressize, Verbuf(0))
If RetVal = 0 Then ReDim Verbuf(1)
FName = File1.filename
GetFileVersionData
DispFileInfo

End Sub

Private Sub File1_DblClick()

cmdProperties_Click

End Sub

Private Sub Form_Load()

If App.PrevInstance = True Then
    MsgBox "This Application is already running!", 48, "Application Running"
    End
End If
CentreMe Me

End Sub

Private Sub cmdSetFileAttr_Click()

On Error GoTo ErrHandler
Dim attr As Long

If File = "" Then
    MsgBox "No file specified yet", 48, "File Information Error"
    Exit Sub
ElseIf IsFileThere(File) = False Then
    MsgBox "File specification not found", 48, "File Information Error"
    Exit Sub
End If

If chkReadOnly.Value = Checked Then
    attr = FILE_ATTRIBUTE_READONLY
End If
If chkArchive.Value = Checked Then
    attr = attr + FILE_ATTRIBUTE_ARCHIVE
End If
If chkSystem.Value = Checked Then
    attr = attr + FILE_ATTRIBUTE_SYSTEM
End If
If chkHidden.Value = Checked Then
    attr = attr + FILE_ATTRIBUTE_HIDDEN
End If
If chkNormal.Value = Checked Then
    attr = attr + FILE_ATTRIBUTE_NORMAL
End If

SetFileAttributes File, attr
File1_Click

Exit Sub
ErrHandler:
MsgBox Err.Description, 48, "File Information Error"

End Sub

Private Sub cmdChangeCreatedTime_Click()

EffectChangeTimeStamp 1, txtCreated

End Sub

Private Sub cmdChangeLastModifiedTime_Click()

EffectChangeTimeStamp 2, txtLastModified

End Sub

Private Sub mnuAbout_Click()

frmAbout.Show 1

End Sub

Private Sub mnuExit_Click()

End

End Sub

Private Sub txtCreated_GotFocus()

Char = txtCreated.Text

End Sub

Private Sub txtCreated_KeyPress(KeyAscii As Integer)

If Not (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9" Or Chr(KeyAscii) = "/" Or Chr(KeyAscii) = ":") Then
    KeyAscii = 0
End If

End Sub

Private Sub txtLastAccessed_GotFocus()

Char = txtLastAccessed.Text

End Sub

Private Sub txtLastAccessed_KeyPress(KeyAscii As Integer)

If Not (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9" Or Chr(KeyAscii) = "/" Or Chr(KeyAscii) = ":") Then
    KeyAscii = 0
End If

End Sub

Private Sub txtLastModified_GotFocus()

Char = txtLastModified.Text

End Sub

Private Sub txtLastModified_KeyPress(KeyAscii As Integer)

If Not (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9" Or Chr(KeyAscii) = "/" Or Chr(KeyAscii) = ":") Then
    KeyAscii = 0
End If

End Sub

Sub CentreMe(F1 As Form)

Dim L1 As Long
Dim L2 As Long

L1 = (Screen.Width - F1.Width) / 2
L2 = (Screen.Height - F1.Height) / 2
F1.Move L1, L2

End Sub

Private Sub DispFileInfo()

On Error Resume Next
    
If File = "" Then
    MsgBox "File not selected", 48, "File Information Error"
    Exit Sub
ElseIf IsFileThere(File) = False Then
    MsgBox "File specification not found", 48, "File Information Error"
    Exit Sub
End If

Dim mHandle As Long
Dim OpenFileStruct As OFSTRUCT
Dim STime As SYSTEMTIME
Dim LTime As FILETIME
Dim mCreationTime As FILETIME
Dim mLastAccessTime As FILETIME
Dim mLastModifyTime As FILETIME
Dim mCreatedStamp As String
Dim mAccessedStamp As String
Dim mModifiedStamp As String
Dim mFileSize As Long

mFileSize = FileLen(File)
'lblFileSize.Caption = "File Size: " & Format(mFileSize, "###,###,### bytes")
'lblFileSize.Caption = "File Size: " & FileLen(File) & " bytes"

mHandle = OpenFile(File, OpenFileStruct, OF_READ Or OF_SHARE_DENY_NONE)

'------------------------------------------
mResult = GetFileTime(mHandle, mCreationTime, mLastAccessTime, mLastModifyTime)
Call CloseHandle(mHandle)

' Convert the 64-bit file time to system time format
mResult = FileTimeToSystemTime(mCreationTime, STime)
' Format it for display
If STime.wYear <= 1980 Then
    mCreatedStamp = "[Unknown]"
Else
    mCreatedStamp = Format$(STime.wDay, "00") & "/" & Format$(STime.wMonth, "00") & "/" & Format$(STime.wYear, "0000") & " " & Format$(STime.wHour, "00") & ":" & Format$(STime.wMinute, "00") & ":" & Format$(STime.wSecond, "00")
End If

mResult = FileTimeToSystemTime(mLastModifyTime, STime)
mModifiedStamp = Format$(STime.wDay, "00") & "/" & Format$(STime.wMonth, "00") & "/" & Format$(STime.wYear, "0000") & " " & Format$(STime.wHour, "00") & ":" & Format$(STime.wMinute, "00") & ":" & Format$(STime.wSecond, "00")

mResult = FileTimeToSystemTime(mLastAccessTime, STime)
mAccessedStamp = Format$(STime.wDay, "00") & "/" & Format$(STime.wMonth, "00") & "/" & Format$(STime.wYear, "0000") & " " & Format$(STime.wHour, "00") & ":" & Format$(STime.wMinute, "00") & ":" & Format$(STime.wSecond, "00")

'------------------------------------------
mFileAttr = GetFileAttributes(File)

If mFileAttr And FILE_ATTRIBUTE_READONLY Then
    chkReadOnly.Value = vbChecked
Else
    chkReadOnly.Value = vbUnchecked
End If
If mFileAttr And FILE_ATTRIBUTE_ARCHIVE Then
    chkArchive.Value = vbChecked
Else
    chkArchive.Value = vbUnchecked
End If
If mFileAttr And FILE_ATTRIBUTE_SYSTEM Then
    chkSystem.Value = vbChecked
Else
    chkSystem.Value = vbUnchecked
End If
If mFileAttr And FILE_ATTRIBUTE_HIDDEN Then
    chkHidden.Value = vbChecked
Else
    chkHidden.Value = vbUnchecked
End If
If mFileAttr And FILE_ATTRIBUTE_NORMAL Then
    chkNormal.Value = vbChecked
Else
    chkNormal.Value = vbUnchecked
End If
    
txtCreated.Text = Format$(mCreatedStamp, "dd/mm/yyyy HH:MM")
txtLastModified.Text = Format$(mModifiedStamp, "dd/mm/yyyy HH:MM")
txtLastAccessed.Text = Format$(mAccessedStamp, "dd/mm/yyyy")

End Sub

Private Sub EffectChangeTimeStamp(inIndex As Integer, Cont As TextBox)

On Error GoTo ErrHandler

If File = "" Then
    MsgBox "No file specified yet", 48, "File Information Error"
    Exit Sub
ElseIf IsFileThere(File) = False Then
    MsgBox "File specification not found", 48, "File Information Error"
    Exit Sub
End If
If mFileAttr And FILE_ATTRIBUTE_SYSTEM Then
    MsgBox "Remove the System Attribute before changing its date or time!", 48, "Attribute Error"
    Exit Sub
End If
If mFileAttr And FILE_ATTRIBUTE_READONLY Then
    MsgBox "Remove the Read Only Attribute before changing its date or time!", 48, "Attribute Error"
    Exit Sub
End If
If mFileAttr And FILE_ATTRIBUTE_HIDDEN Then
    MsgBox "Remove the Hidden Attribute before changing its date or time!", 48, "Attribute Error"
    Exit Sub
End If

Dim mHandle As Long
Dim OpenFileStruct As OFSTRUCT
Dim STime As SYSTEMTIME
Dim LTime As FILETIME
Dim mCreationTime As FILETIME
Dim mLastAccessTime As FILETIME
Dim mLastModifyTime As FILETIME
Dim tmp As Date
Dim Y As Integer
Dim M As Integer
Dim D As Integer
Dim HH As Integer
Dim MM As Integer
Dim SS As Integer
Dim mTimeStamp As Variant
Dim mfilespec As String
    
Select Case inIndex
    Case 1
        If txtCreated.Text = "" Then
            MsgBox "No time stamp entered for Creation Time yet", 48, "File Information Error"
            Exit Sub
        End If
        tmp = CDate(txtCreated.Text)
    Case 2
        If txtCreated.Text = "" Then
            MsgBox "No time stamp entered for Last Modified Time yet", 48, "File Information Error"
            Exit Sub
        End If
        tmp = CDate(txtLastModified.Text)
    Case 3
        If txtCreated.Text = "" Then
            MsgBox "No time stamp entered for Last Accessed Time yet", 48, "File Information Error"
            Exit Sub
        End If
        tmp = CDate(txtLastAccessed.Text)
End Select

Y = Year(tmp)
M = Month(tmp)
D = Day(tmp)
HH = Hour(tmp)
MM = Minute(tmp)
SS = 0

mTimeStamp = DateSerial(Y, M, D) + TimeSerial(HH, MM, SS)

With STime
    .wYear = Year(mTimeStamp)
    .wMonth = Month(mTimeStamp)
    .wDay = Day(mTimeStamp)
    .wDayOfWeek = WeekDay(mTimeStamp) - 1
    .wHour = Hour(mTimeStamp)
    .wSecond = Second(mTimeStamp)
    .wMilliseconds = 0
End With

' Convert system time format to 64-bit local time
mResult = SystemTimeToFileTime(STime, LTime)
    
' Convert local file time to file time based on UTC.
Select Case inIndex
    Case 1
        mResult = LocalFileTimeToFileTime(LTime, mCreationTime)
    Case 2
        mResult = LocalFileTimeToFileTime(LTime, mLastModifyTime)
    Case 3
        mResult = LocalFileTimeToFileTime(LTime, mLastAccessTime)
End Select

mHandle = OpenFile(File, OpenFileStruct, OF_WRITE Or OF_SHARE_EXCLUSIVE)

' Effect changing file time stamp
SetFileTime mHandle, mCreationTime, mLastAccessTime, mLastModifyTime
    
CloseHandle mHandle

Exit Sub
ErrHandler:
If Err = 13 Then
    MsgBox "Incorrect Date/Time Format!", 48, "File Information Error"
    CloseHandle mHandle
    Cont.Text = Char
    Exit Sub
End If

End Sub

Private Sub GetFileVersionData()

On Error GoTo GetFileVersionData_Error

Dim sInfo As String
Dim sMsg As String
Dim lResult As Long
Dim iDelim As Integer
Dim N As Integer
Dim lHandle As Long

For N = 2 To 10
    lblValue(N).Caption = ""
Next
lblValue(0).Caption = FName
If FileLen(File) = 0 Then
    lblValue(1).Caption = Format$(FileDateTime(File), "DD/MM/YY HH:MM:SS") & Space(10) & "0 bytes"
Else
    lblValue(1).Caption = Format$(FileDateTime(File), "DD/MM/YY HH:MM:SS") & Space(10) & Format$(FileLen(File), "###,###,###") & " bytes"
End If
lHandle = 0
'how big is the Version Info block?
lSizeof = GetFileVersionInfoSize(File, lHandle)
If lSizeof > 0 Then
    sInfo = String$(lSizeof, 0)
    lResult = GetFileVersionInfo(ByVal File, 0&, ByVal lSizeof, ByVal sInfo)
    If lResult Then
        iDelim = InStr(sInfo, "CompanyName")
        If iDelim > 0 Then
            iDelim = iDelim + 12
            lblValue(2).Caption = Mid$(sInfo, iDelim)
        End If

        iDelim = InStr(sInfo, "FileDescription")
        If iDelim > 0 Then
            iDelim = iDelim + 16
            lblValue(3).Caption = Mid$(sInfo, iDelim)
        End If

        iDelim = InStr(sInfo, "FileVersion")
        If iDelim > 0 Then
            iDelim = iDelim + 12
            lblValue(4).Caption = Mid$(sInfo, iDelim)
        End If

        iDelim = InStr(sInfo, "InternalName")
        If iDelim > 0 Then
            iDelim = iDelim + 16
            lblValue(5).Caption = Mid$(sInfo, iDelim)
        End If

        iDelim = InStr(sInfo, "LegalCopyright")
        If iDelim > 0 Then
            iDelim = iDelim + 16
            lblValue(6).Caption = Mid$(sInfo, iDelim)
        End If

        iDelim = InStr(sInfo, "OriginalFilename")
        If iDelim > 0 Then
            iDelim = iDelim + 20
            lblValue(7).Caption = Mid$(sInfo, iDelim)
        End If

        iDelim = InStr(sInfo, "ProductName")
        If iDelim > 0 Then
            iDelim = iDelim + 12
            lblValue(8).Caption = Mid$(sInfo, iDelim)
        End If

        iDelim = InStr(sInfo, "ProductVersion")
        If iDelim > 0 Then
            iDelim = iDelim + 16
            lblValue(9).Caption = Mid$(sInfo, iDelim)
        End If
    Else
        GoTo invalid_file_info_error
    End If
Else
    GoTo invalid_file_info_error
End If

lblValue(10).Caption = GetInfoString("Comments")

GetFileVersionData_Exit:

Exit Sub
GetFileVersionData_Error:
MsgBox "Error " & Format$(Err) & ": " & Error$ & " in GetFileVersionData"
Resume GetFileVersionData_Exit

invalid_file_info_error:

lblValue(3).Caption = "No information available."
GoTo GetFileVersionData_Exit

End Sub

Private Function GetInfoString$(stringtoget$)

Dim Tbuf As String
Dim Nullpos As Integer
Dim Xlatelang As Integer
Dim Xlatecode As Integer
Dim Numentries As Integer
Dim Fiiaddr As Long
Dim Xlatestring As String
Dim Xlateval As Long
Dim Fiilen As Long
Dim Di As Long
Dim X As Integer
    
Di = VerQueryValue(Verbuf(0), "\VarFileInfo\Translation", Fiiaddr, Fiilen)
If (Di <> 0) Then ' Translation table exists
    Numentries = Fiilen / 4
    Xlateval = 0
    For X = 1 To Numentries
        ' Copy the 4 byte tranlation entry for the first
        agCopyData ByVal Fiiaddr, Xlatelang, 2
        agCopyData ByVal (Fiiaddr + 2), Xlatecode, 2
        ' Exit if U.S. English was found
        If Xlatelang = &H409 Then Exit For
        Fiiaddr = Fiiaddr& + 4
    Next
Else
    ' No translation table - Assume standard ASCII
    Xlatelang% = &H409
    Xlatecode% = 0
End If
    
Xlatestring = Hex$(Xlatecode)
' Make sure hex string is 4 chars long
While Len(Xlatestring) < 4
    Xlatestring = "0" + Xlatestring
Wend
Xlatestring$ = Hex$(Xlatelang) + Xlatestring
' Make sure hex string is 8 chars long
While Len(Xlatestring) < 8
    Xlatestring = "0" + Xlatestring
Wend

Di = VerQueryValue(Verbuf(0), "\StringFileInfo\" + Xlatestring + "\" + stringtoget$, Fiiaddr, Fiilen)
If Di = 0 Then
    GetInfoString = "Unavailable"
    Exit Function
End If

Tbuf = String(Fiilen + 1, Chr(0))

' Copy the fixed file info into the structure
agCopyData ByVal Fiiaddr, ByVal Tbuf, Fiilen

Nullpos = InStr(Tbuf, Chr(0))
If (Nullpos > 1) Then
    GetInfoString = Left(Tbuf, Nullpos - 1)
Else
    GetInfoString = "None"
End If

End Function

Function IsFileThere(inFileSpec As String) As Boolean

On Error Resume Next
Dim FileNum As Integer

FileNum = FreeFile()
Open inFileSpec For Input As #FileNum
If Err Then
    IsFileThere = False
Else
    Close #FileNum
    IsFileThere = True
End If

End Function

Function ShowFileProperties(filename As String, OwnerhWnd As Long) As Long

Dim SEI As SHELLEXECUTEINFO

With SEI
    .cbSize = Len(SEI)
    .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
    .hwnd = OwnerhWnd
    .lpVerb = "properties"
    .lpFile = filename
    .lpParameters = vbNullChar
    .lpDirectory = vbNullChar
    .nShow = 0
    .hInstApp = 0
    .lpIDList = 0
End With

ShellExecuteEX SEI
ShowFileProperties = SEI.hInstApp

End Function

