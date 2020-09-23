VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEFS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************************
'*  NMC! V1.5    EFS Member Class                                                      *
'*                                                                                     *
'*  Created:     April 14, 2006                                                        *
'*  Updated:     April 29, 2006                                                        *
'*  Purpose:     NT System Security Master Class                                       *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.5.0                                                                 *
'*  Compile:     Native                                                                *
'*  Referenced:  Throughout Project                                                    *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*                                                                                     *
'***************************************************************************************

'/~ Encrypted File System ~/
'/~ EFS_Status              - file/folder state         [in -string | out -long]
'/~ EFS_Encrypt             - encrypt file/folder       [in -string | out -bool]
'/~ EFS_Decrypt             - decrypt file/folder       [in -string | out -bool]
'/~ EFS_Enable              - enable EFS                [in -string | out -bool]
'/~ EFS_Disable             - disable EFS               [in -string | out -bool]
'/~ Create_Directory        - create a directory        [in -string | out -bool]
'/~ File_Exists             - test for path/file        [in -string | out -bool]


Private Const FO_COPY                       As Long = &H2
Private Const FO_DELETE                     As Long = &H3
Private Const FOF_NOCONFIRMATION            As Long = &H10
Private Const FOF_SILENT                    As Long = &H4

Private Const FILE_ENCRYPTABLE              As Long = &H0
Private Const FILE_IS_ENCRYPTED             As Long = &H1
Private Const FILE_SYSTEM_ATTR              As Long = &H2
Private Const FILE_ROOT_DIR                 As Long = &H3
Private Const FILE_SYSTEM_DIR               As Long = &H4
Private Const FILE_UNKNOWN                  As Long = &H5
Private Const FILE_SYSTEM_NOT_SUPPORT       As Long = &H6
Private Const FILE_READ_ONLY                As Long = &H8
Private Const FILE_DIR_DISALLOWED           As Long = &H9

Private Const LOGON32_PROVIDER_DEFAULT      As Long = &H0
Private Const LOGON32_LOGON_INTERACTIVE     As Long = &H2

Private Type SHFILEOPSTRUCT
    hWnd                                    As Long
    wFunc                                   As Long
    pFrom                                   As String
    pTo                                     As String
    fFlags                                  As Integer
    fAborted                                As Long
    hNameMaps                               As Long
    sProgress                               As String
End Type

Private Declare Function EncryptFile Lib "advapi32.dll" Alias "EncryptFileA" (ByVal lpFileName As String) As Long

Private Declare Function DecryptFile Lib "advapi32.dll" Alias "DecryptFileA" (ByVal lpFileName As String, _
                                                                              ByVal dwReserved As Long) As Long

Private Declare Function FileEncryptionStatus Lib "advapi32.dll" Alias "FileEncryptionStatusA" (ByVal lpFileName As String, _
                                                                                                lpStatus As Long) As Long

Private Declare Function EncryptionDisable Lib "advapi32.dll" (ByVal DirPath As String, _
                                                               ByVal Disable As Long) As Long

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (ByRef lpFileOp As SHFILEOPSTRUCT) As Long

Private Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" (ByVal lpszUsername As String, _
                                                                          ByVal lpszDomain As String, _
                                                                          ByVal lpszPassword As String, _
                                                                          ByVal dwLogonType As Long, _
                                                                          ByVal dwLogonProvider As Long, _
                                                                          phToken As Long) As Long

Private Declare Function ImpersonateLoggedOnUser Lib "advapi32" (ByVal hToken As Long) As Long

Private Declare Function RevertToSelf Lib "advapi32.dll" () As Long

Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

Private Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, _
                                                                                      nSize As Long) As Long

Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
                                                                                                ByVal lpVolumeNameBuffer As String, _
                                                                                                ByVal nVolumeNameSize As Long, _
                                                                                                lpVolumeSerialNumber As Long, _
                                                                                                lpMaximumComponentLength As Long, _
                                                                                                lpFileSystemFlags As Long, _
                                                                                                ByVal lpFileSystemNameBuffer As String, _
                                                                                                ByVal nFileSystemNameSize As Long) As Long

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
                                                                                                ByVal lpBuffer As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Public Event eNComplete(ByVal sTask As String)
Public Event eNErrorCond(ByVal sRoutine As String, ByVal sError As String)

Private m_lToken    As Long

'* Name           : EFS_Status
'* Purpose        : return EFS status for path
'* Inputs         : path
'* Outputs        : long
'*********************************************
Public Function EFS_Status(ByVal sPath As String) As Long

Dim lStatus         As Long
Dim lReturn         As Long
Dim sMessage        As String

On Error GoTo Handler

    '/* query status
    lReturn = FileEncryptionStatus(sPath, lStatus)
    If lReturn = 0 Then GoTo Handler
    '/* interpret
    Select Case lStatus
    Case FILE_ENCRYPTABLE
        sMessage = "The Object can be Encrypted."
    Case FILE_IS_ENCRYPTED
        sMessage = "The Object is Encrypted."
    Case FILE_SYSTEM_ATTR
        sMessage = "Aborted: The Object is a System File."
    Case FILE_ROOT_DIR
        sMessage = "Aborted: The Object is on a Root Directory."
    Case FILE_SYSTEM_DIR
        sMessage = "Aborted: The Object is on a System Directory."
    Case FILE_UNKNOWN
        sMessage = "Aborted: The Object is a System File or Directory."
    Case FILE_SYSTEM_NOT_SUPPORT
        sMessage = "Aborted: This File System does Not Support Encryption."
    Case FILE_READ_ONLY
        sMessage = "Aborted: The File is Read Only."
    Case FILE_DIR_DISALLOWED
        sMessage = "Reserved."
    Case Else
    End Select
    '/* success
    RaiseEvent eNComplete(sMessage)
    EFS_Status = lReturn
    On Error GoTo 0

Exit Function

Handler:
    RaiseEvent eNErrorCond("EFS_Status", CStr(Err.LastDllError))

End Function

'* Name           : EFS_Encrypt
'* Purpose        : encrypt a file or folder
'* Inputs         : path
'* Outputs        : boolean
'*********************************************
Public Function EFS_Encrypt(ByVal sPath As String) As Boolean

Dim lReturn         As Long

On Error GoTo Handler

    lReturn = EncryptFile(sPath)
    Select Case lReturn
    '/* denied
    Case 6010
        RaiseEvent eNComplete("Encrypted is Disabled for this Object.")
    '/* success
    Case 1
        RaiseEvent eNComplete("The Object has been Encrypted.")
        EFS_Encrypt = True
    '/* unknown
    Case Else
        GoTo Handler
    End Select
    On Error GoTo 0

Exit Function

Handler:
    RaiseEvent eNErrorCond("EFS_Encrypt", CStr(Err.LastDllError))

End Function

'* Name           : EFS_Decrypt
'* Purpose        : decrypt a file or folder
'* Inputs         : path
'* Outputs        : boolean
'*********************************************
Public Function EFS_Decrypt(ByVal sPath As String) As Boolean

Dim lReturn         As Long

On Error GoTo Handler

    lReturn = DecryptFile(sPath, 0&)
    Select Case lReturn
    '/* denied
    Case 6010
        RaiseEvent eNComplete("Encrypted is Disabled for this Object.")
    '/* success
    Case 1
        RaiseEvent eNComplete("The Object has been Decrypted.")
        EFS_Decrypt = True
    '/* unknown
    Case Else
        GoTo Handler
    End Select
    On Error GoTo 0

Exit Function

Handler:
    RaiseEvent eNErrorCond("EFS_Decrypt", CStr(Err.LastDllError))

End Function

'* Name           : EFS_Enable
'* Purpose        : enable EFS on the path
'* Inputs         : path
'* Outputs        : boolean
'*********************************************
Public Function EFS_Enable(ByVal sPath As String) As Boolean

Dim lReturn         As Long
Dim bPath           As Byte

On Error GoTo Handler

    If InStr(1, sPath, ".") > 0 Then GoTo Handler
    bPath = sPath + vbNullChar
    lReturn = EncryptionDisable(bPath, 0)
    If lReturn > 0 Then GoTo Handler
    RaiseEvent eNComplete("Encryption for " + sPath + " has been enabled.")
    EFS_Enable = True

    On Error GoTo 0

Exit Function

Handler:
    RaiseEvent eNErrorCond("EFS_Enable", CStr(Err.LastDllError))

End Function

'* Name           : EFS_Disable
'* Purpose        : disable EFS on the path
'* Inputs         : path
'* Outputs        : boolean
'*********************************************
Public Function EFS_Disable(ByVal sPath As String) As Boolean

Dim lReturn         As Long
Dim bPath           As Byte

On Error GoTo Handler

    If InStr(1, sPath, ".") > 0 Then GoTo Handler
    bPath = sPath + vbNullChar
    lReturn = EncryptionDisable(bPath, 1)
    If lReturn > 0 Then GoTo Handler
    RaiseEvent eNComplete("Encryption for " + sPath + " has been disabled.")
    EFS_Disable = True
    
    On Error GoTo 0

Exit Function

Handler:
    RaiseEvent eNErrorCond("EFS_Disable", CStr(Err.LastDllError))

End Function


Public Function Local_Name() As String

Dim sName       As String

    sName = Space(255)
    GetComputerName sName, 255
    Local_Name = "\\" + Left$(sName, InStr(sName, vbNullChar) - 1)

End Function

'* Name           : Create_Directory
'* Purpose        : create a new directory
'* Inputs         : path
'* Outputs        : boolean
'*********************************************
Public Function Create_Directory(ByVal sPath As String) As Boolean
'/* create a new directory

Dim lResult     As Long

    lResult = MakeSureDirectoryPathExists(sPath)
    If lResult = 0 Then Create_Directory = True

End Function

'* Name           : User_Impersonate
'* Purpose        : process impersonate user
'* Inputs         : user, pass, domain
'* Outputs        : boolean
'*********************************************
Public Function User_Impersonate(ByVal sUser As String, _
                                 ByVal sPass As String, _
                                 ByVal sDomain As String) As Boolean

Dim lProfile    As Long
Dim lHandle     As Long

    If LogonUser(sUser, sDomain, sPass, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, m_lToken) <> 0 Then
        If ImpersonateLoggedOnUser(m_lToken) <> 0 Then
            User_Impersonate = True
        End If
    End If

End Function

'* Name           : User_Revert
'* Purpose        : restore user token
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Function User_Revert() As Boolean

Dim lReturn     As Long

    CloseHandle m_lToken
    lReturn = RevertToSelf
    User_Revert = (Not lReturn = 0)

End Function

'* Name           : NTFS_Drive
'* Purpose        : test ntfs status
'* Inputs         : path
'* Outputs        : boolean
'*********************************************
Public Function NTFS_Drive(ByVal sPath As String) As Boolean

Dim lFlags      As Long
Dim lMaxLen     As Long
Dim lSerial     As Long
Dim sName       As String * 256
Dim sType       As String * 256

    If Len(sPath) > 3 Then sPath = Left$(sPath, 3)
    GetVolumeInformation sPath, sName, Len(sName), lSerial, lMaxLen, lFlags, sType, Len(sType)
    If lFlags And &H40000 Then NTFS_Drive = True

End Function

'* Name           : NTFS_Check
'* Purpose        : return list of ntfs drives
'* Inputs         : none
'* Outputs        : collection
'*********************************************
Public Function NTFS_Check() As Collection

Dim sDrives     As String
Dim lBuffer     As Long
Dim cTemp       As Collection
Dim aDrives()   As String
Dim lCount      As Long

On Error Resume Next

    Set cTemp = New Collection
    '//get the buffer size
    lBuffer = GetLogicalDriveStrings(0, sDrives)
    '//set string len
    sDrives = String$(lBuffer, 0)
    '//get the drive list
    GetLogicalDriveStrings lBuffer, sDrives
    '/* split
    sDrives = Left$(sDrives, Len(sDrives) - 2)
    aDrives = Split(sDrives, Chr$(0))
    For lCount = 0 To UBound(aDrives)
        '/* skip floppy drive
        If Not LCase$(aDrives(lCount)) = "a:\" Then
            If NTFS_Drive(aDrives(lCount)) Then
                cTemp.Add aDrives(lCount)
            End If
        End If
    Next lCount
    '/* success
    If cTemp.Count > 0 Then
        Set NTFS_Check = cTemp
    End If

On Error GoTo 0

End Function
