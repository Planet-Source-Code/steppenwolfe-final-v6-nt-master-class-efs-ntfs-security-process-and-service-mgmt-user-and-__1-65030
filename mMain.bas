Attribute VB_Name = "mMain"
Option Explicit

Private Const VER_PLATFORM_WIN32s               As Integer = 0
Public Const VER_PLATFORM_WIN32_WINDOWS         As Integer = 1
Public Const VER_PLATFORM_WIN32_NT              As Integer = 2
Public Const ICC_USEREX_CLASSES                 As Long = &H200

Type OSVersionInfo
    dwOSVersionInfoSize                         As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion                                As String * 128
End Type

Public Type tagInitCommonControlsEx
   lngSize                                      As Long
   lngICC                                       As Long
End Type

Private Type SHITEMID
    cb                                          As Long
    abID                                        As Byte
End Type

Private Type ITEMIDLIST
    mkid                                        As SHITEMID
End Type

Private Type BrowseInfo
    hwndOwner                                    As Long
    pIDLRoot                                     As Long
    pszDisplayName                               As Long
    lpszTitle                                    As Long
    ulFlags                                      As Long
    lpfnCallback                                 As Long
    lParam                                       As Long
    iImage                                       As Long
End Type

Public Enum eDirectories
    CSIDL_DESKTOPDIRECTORY = 0
    CSIDL_START_PROGRAMS = 2
    CSIDL_MYDOCUMENTS = 5
    CSIDL_FAVORITES = 6
    CSIDL_STARTUP = 7
    CSIDL_RECENT = 8
    CSIDL_SENDTO = 9
    CSIDL_START_MENU = 11
    CSIDL_MYMUSIC = 13
    CSIDL_MYVIDEO = 14
    CSIDL_DESKTOP = 16
    CSIDL_NETHOOD = 19
    CSIDL_FONTS = 20
    CSIDL_TEMPLATES = 21
    CSIDL_COMMON_STARTMENU = 22
    CSIDL_COMMON_PROGRAMS = 23
    CSIDL_COMMON_STARTUP = 24
    CSIDL_COMMON_DESKTOP = 25
    CSIDL_APPDATA = 26
    CSIDL_PRINTHOOD = 27
    CSIDL_SETTINGS_APPDATA = 28
    CSIDL_COMMON_FAVORITES = 31
    CSIDL_INTERNET_CACHE = 32
    CSIDL_COOKIES = 33
    CSIDL_HISTORY = 34
    CSIDL_COMMON_APPDATA = 35
    CSIDL_WINDOWS = 36
    CSIDL_SYSTEM = 37
    CSIDL_PROGRAM_FILES = 38
    CSIDL_MYPICTURES = 39
    CSIDL_PROFILE = 40
    CSIDL_COMMON_SYSTEM = 42
    CSIDL_COMMON_FILES = 43
    CSIDL_COMMON_TEMPLATES = 45
    CSIDL_COMMON_DOCUMENTS = 46
    CSIDL_COMMON_MUSIC = 53
    CSIDL_COMMON_PICTURES = 54
    CSIDL_COMMON_VIDEO = 55
    CSIDL_RESOURCES = 56
    CSIDL_CD_BURN_AREA = 56
End Enum

Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, _
                                                            ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
                                                                  ByVal lpString2 As String) As Long

Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, _
                                                                   ByVal lpNewFileName As String, _
                                                                   ByVal bFailIfExists As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                                                               ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, _
                                                                               ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, _
                                                                               ByVal nShowCmd As Long) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWnd As Long, _
                                                                       ByVal csidl As Long, _
                                                                       ByRef ppidl As ITEMIDLIST) As Long

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVersionInfo) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lLongPath As String, _
                                                                                    ByVal lShortPath As String, _
                                                                                    ByVal lBuffer As Long) As Long

Public m_RegPath   As String
Public m_SysPath   As String
Public m_WinPath   As String

Public Sub Main()
'/* load and check

    Init_Controls
    Load frmMain
    '/* older os - disable features
    If OS_Check Then frmMain.Disable_Items
    '/* get paths
    m_RegPath = "Software\" + App.ProductName
    m_SysPath = Get_Folder(CSIDL_SYSTEM) + Chr$(92)
    m_WinPath = Get_Folder(CSIDL_WINDOWS) + Chr$(92)
    
    '/* show links
    If File_Exists(m_SysPath + "lusrmgr.msc") Then
        frmMain.cmdVerify(0).Visible = True
    End If
    If File_Exists(m_SysPath + "services.msc") Then
        frmMain.cmdVerify(1).Visible = True
    End If
    If File_Exists(m_SysPath + "taskmgr.exe") Then
        frmMain.cmdVerify(2).Visible = True
    End If
    '/* show
    frmMain.Show

End Sub

Public Sub Init_Controls()
'/* init common controls

On Error Resume Next

Dim iccex As tagInitCommonControlsEx

   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex

On Error GoTo 0

End Sub

Public Function OS_Check() As Boolean

Dim rOsVersionInfo  As OSVersionInfo
Dim sOpSys          As String

On Error GoTo Handler

    rOsVersionInfo.dwOSVersionInfoSize = Len(rOsVersionInfo)
    If GetVersionEx(rOsVersionInfo) = 0 Then
        sOpSys = "Can not be determined"
        GoTo Handler
    End If
    
    Select Case rOsVersionInfo.dwPlatformId
    Case 0
        sOpSys = "Win32s"
    
    Case 1
        If rOsVersionInfo.dwMajorVersion >= 5 Then
            sOpSys = "Windows ME"
            GoTo Handler
        ElseIf rOsVersionInfo.dwMajorVersion = 4 And rOsVersionInfo.dwMinorVersion > 0 Then
            sOpSys = "Windows 98"
            GoTo Handler
        Else
            sOpSys = "Windows 95"
            GoTo Handler
        End If
        
    Case 2
        If rOsVersionInfo.dwMajorVersion >= 5 Then
            If rOsVersionInfo.dwMinorVersion = 0 Then
                sOpSys = "Windows 2000"
            Else
                sOpSys = "Windows XP"
            End If
        Else
            sOpSys = "Windows NT"
        End If
        
    Case Else
    
    End Select

On Error GoTo 0
Exit Function

Handler:
    '/* older os - warn
    OS_Check = True
    With frmOSCheck
        .p_OSType = sOpSys
        .Show vbModal, frmMain
    End With

End Function

Public Function Get_Folder(SHFlag As eDirectories) As String
'/* get default folder locations

Dim lRes            As Long
Dim sPath           As String
Dim ItemIdL         As ITEMIDLIST

On Error GoTo Handler

    lRes = SHGetSpecialFolderLocation(100, SHFlag, ItemIdL)

    If lRes Then
        Get_Folder = vbNullString
    Else
        sPath = Space$(512)
        lRes = SHGetPathFromIDList(ByVal ItemIdL.mkid.cb, ByVal sPath)
        Get_Folder = Left(sPath, InStr(sPath, Chr$(0)) - 1)
    End If

Handler:

End Function

Public Sub Open_File(sPath As String)

    ShellExecute frmMain.hWnd, "open", sPath, "", "", 1

End Sub

Public Function Get_Path() As String
'/* get scan path dialog

Dim lpIDList    As Long
Dim szTitle     As String
Dim tBrowseInfo As BrowseInfo
Dim sBuffer     As String

On Error Resume Next

    szTitle = "Select a Directory to Scan: "

    With tBrowseInfo
        .hwndOwner = frmMain.hWnd
        .lpszTitle = lstrcat(szTitle, vbNullString)
        .ulFlags = 1 + 2
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If lpIDList Then
        sBuffer = Space$(260)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
        Get_Path = sBuffer + Chr$(92)
    End If

On Error GoTo 0

End Function

Public Function File_Exists(ByVal sDir As String) As Boolean
'/* test file

Dim lRes        As Long
Dim sPath       As String

    sPath = String$(260, 0)
    lRes = GetShortPathName(sDir, sPath, 259)
    File_Exists = lRes > 0

End Function
