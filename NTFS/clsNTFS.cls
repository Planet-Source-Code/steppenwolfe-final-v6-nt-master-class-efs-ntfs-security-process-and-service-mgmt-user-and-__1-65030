VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNTFS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************************
'*  NMC! V1.5    NTFS Member Class                                                     *
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

'/~ NTFS File System ~/
'/~ NTFS_Drive              - test drive for ntfs       [in -string(1) | out -bool]
'/~ NTFS_Check              - test all drives for ntfs  [out -col]
'/~ NTFS_Folder             - modify object security    [in -string/enum | out -bool]
'/~ NTFS_Recursive          - recurse permissions set   [in -string | out -bool]

'/~ NTFS Registry ~/
'/~ NTFS_Key                - modify key security       [in -string/enum | out -bool]


'/* memory constants
Private Const LMEM_FIXED                                As Long = &H0
Private Const LMEM_ZEROINIT                             As Long = &H40
Private Const lPtr                                      As Double = (LMEM_FIXED + LMEM_ZEROINIT)
'/* access rights
Private Const GENERIC_READ                              As Long = &H80000000
Private Const DELETE                                    As Long = &H10000
Private Const READ_CONTROL                              As Long = &H20000
Private Const WRITE_DAC                                 As Long = &H40000
Private Const SYNCHRONIZE                               As Long = &H100000
Private Const STANDARD_RIGHTS_REQUIRED                  As Long = &HF0000
Private Const STANDARD_RIGHTS_READ                      As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE                     As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE                   As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_ALL                       As Long = &H1F0000
Private Const ACCESS_SYSTEM_SECURITY                    As Long = &H1000000
Private Const MAXIMUM_ALLOWED                           As Long = &H2000000
'/* descriptor flags
Private Const DACL_SECURITY_INFORMATION                 As Long = &H4
Private Const SECURITY_DESCRIPTOR_REVISION              As Integer = 1
Private Const SECURITY_DESCRIPTOR_MIN_LENGTH            As Integer = 20
Private Const ACL_REVISION                              As Integer = 2
Private Const MAXDWORD                                  As Long = &HFFFFFFFF
'/* inherit flags of an ace header
Private Const OBJECT_INHERIT_ACE                        As Long = &H1
Private Const CONTAINER_INHERIT_ACE                     As Long = &H2
Private Const NO_PROPAGATE_INHERIT_ACE                  As Long = &H4
Private Const INHERIT_ONLY_ACE                          As Long = &H8
Private Const INHERITED_ACE                             As Long = &H10
'/* security descriptor flags.
Private Const SE_DACL_AUTO_INHERIT_REQ                  As Long = &H100
Private Const SE_DACL_AUTO_INHERITED                    As Long = &H400
Private Const SE_DACL_PROTECTED                         As Long = &H1000
'/* ACE being added.
Private Const ACCESS_ALLOWED_ACE_TYPE                   As Long = &H0
Private Const ACCESS_DENIED_ACE_TYPE                    As Long = &H1
'/* well-known sid, users and groups
Private Const SECURITY_WORLD_SID_AUTHORITY              As Long = &H1
Private Const SECURITY_WORLD_RID                        As Long = &H0
Private Const INVALID_HANDLE_VALUE                      As Integer = -1
Private Const OPEN_EXISTING                             As Integer = 3
Private Const FILE_FLAG_BACKUP_SEMANTICS                As Long = &H2000000
'/* folder specific access rights
Private Const FILE_NO_ACCESS                            As Long = &H0
Private Const FILE_LIST_DIRECTORY                       As Long = &H1
Private Const FILE_ADD_FILE                             As Long = &H2
Private Const FILE_ADD_SUBDIRECTORY                     As Long = &H4
Private Const FILE_TRAVERSE                             As Long = &H20
Private Const FILE_DELETE_CHILD                         As Long = &H40
Private Const FILE_READ_DATA                            As Long = &H1
Private Const FILE_WRITE_DATA                           As Long = &H2
Private Const FILE_APPEND_DATA                          As Long = &H4
Private Const FILE_EXECUTE                              As Long = &H20
Private Const FILE_READ_EA                              As Long = &H8
Private Const FILE_WRITE_EA                             As Long = &H10
Private Const FILE_READ_ATTRIBUTES                      As Long = &H80
Private Const FILE_WRITE_ATTRIBUTES                     As Long = &H100

'/* generic access masks for files
Private Const FILE_ALL_ACCESS                           As Long = _
        STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or 511
Private Const FILE_GENERIC_READ                         As Long = _
        STANDARD_RIGHTS_READ Or FILE_READ_DATA Or _
        FILE_READ_ATTRIBUTES Or FILE_READ_EA Or SYNCHRONIZE
Private Const FILE_GENERIC_WRITE                        As Long = _
        STANDARD_RIGHTS_WRITE Or FILE_WRITE_DATA Or _
        FILE_WRITE_ATTRIBUTES Or FILE_WRITE_EA Or _
        FILE_APPEND_DATA Or SYNCHRONIZE
Private Const FILE_GENERIC_EXECUTE                      As Long = _
        STANDARD_RIGHTS_EXECUTE Or FILE_READ_ATTRIBUTES Or _
        FILE_EXECUTE Or SYNCHRONIZE

'/* registry access masks
Private Const KEY_QUERY_VALUE                           As Long = &H1
Private Const KEY_SET_VALUE                             As Long = &H2
Private Const KEY_CREATE_SUB_KEY                        As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS                    As Long = &H8
Private Const KEY_NOTIFY                                As Long = &H10
Private Const KEY_CREATE_LINK                           As Long = &H20
Private Const KEY_READ                                  As Double = _
        ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or _
        KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE                                 As Double = _
        ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or _
        KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE                               As Double = _
        ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS                            As Double = _
        ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or _
        KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or _
        KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or _
        KEY_CREATE_LINK) And (Not SYNCHRONIZE))

'/* file time struct
Private Type FT
    lLD                             As Long
    lHD                             As Long
End Type

'/* file find struct
Private Type WFD
    lFA                             As Long
    fCT                             As FT
    fLA                             As FT
    fLWT                            As FT
    lFSH                            As Long
    lFSL                            As Long
    lR0                             As Long
    lR1                             As Long
    cFN                             As String * 260
    cAL                             As String * 14
End Type

'/* acl data
Private Type ACL
    AclRevision                                         As Byte
    Sbz1                                                As Byte
    AclSize                                             As Integer
    AceCount                                            As Integer
    Sbz2                                                As Integer
End Type

'/* ace size data
Private Type ACL_SIZE_INFORMATION
    AceCount                                            As Long
    AclBytesInUse                                       As Long
    AclBytesFree                                        As Long
End Type

'/* header data
Private Type ACE_HEADER
    AceType                                             As Byte
    AceFlags                                            As Byte
    AceSize                                             As Integer
End Type

'/* ace struct
Private Type ACE
    Header                                              As ACE_HEADER
    Mask                                                As Long
    SidStart                                            As Long
End Type

'/* sid id
Private Type SID_IDENTIFIER_AUTHORITY
    Value(6)                                            As Byte
End Type

'/* account struct
Private Type ACCOUNT_PERM
    AccountName                                         As String
    AccessMask                                          As Long
    AceFlags                                            As Byte
    AceType                                             As Byte
    pSid                                                As Long
    SidPassedByCaller                                   As Boolean
End Type

'/* store data
Private Type MEM_DATA
    pSD                                                 As Long
    pAcl                                                As Long
End Type

'~*** Exposed Enum Structures ***~
'/~ ace access mask
Public Enum ACCESS_MASK
    Delete_Only = DELETE Or READ_CONTROL
    Execute_Only = STANDARD_RIGHTS_EXECUTE Or STANDARD_RIGHTS_READ
    Read_Only = STANDARD_RIGHTS_READ
    Read_Write = READ_CONTROL Or STANDARD_RIGHTS_WRITE
    Read_Write_Delete = STANDARD_RIGHTS_READ Or STANDARD_RIGHTS_WRITE Or DELETE
    Read_Write_Execute = STANDARD_RIGHTS_READ Or STANDARD_RIGHTS_WRITE Or STANDARD_RIGHTS_EXECUTE
    Read_Write_Execute_Delete = STANDARD_RIGHTS_READ Or STANDARD_RIGHTS_WRITE Or STANDARD_RIGHTS_EXECUTE Or DELETE
    Standard_Rights = STANDARD_RIGHTS_ALL
    Full_Control = MAXIMUM_ALLOWED
    System_Control = MAXIMUM_ALLOWED Or ACCESS_SYSTEM_SECURITY
End Enum

'/~ folder access masks
Public Enum FOLDER_PERMISSIONS
    '/* generic permissions structures
    Folder_Read = FILE_GENERIC_READ
    Folder_Read_Execute = FILE_GENERIC_READ Or FILE_GENERIC_EXECUTE
    Folder_Read_Write = FILE_GENERIC_READ Or FILE_GENERIC_WRITE
    Folder_Read_Write_List = FILE_GENERIC_READ Or FILE_GENERIC_WRITE Or FILE_LIST_DIRECTORY
    Folder_Read_Write_Execute = FILE_GENERIC_READ Or FILE_GENERIC_WRITE Or FILE_GENERIC_EXECUTE
    Folder_Read_Write_Execute_List = FILE_GENERIC_READ Or FILE_GENERIC_WRITE Or FILE_GENERIC_EXECUTE Or FILE_LIST_DIRECTORY
    Folder_Read_Execute_List = FILE_GENERIC_READ Or FILE_GENERIC_EXECUTE Or FILE_LIST_DIRECTORY
    Folder_Read_List = FILE_GENERIC_READ Or FILE_LIST_DIRECTORY
    '/* specific attributes
    Folder_List = FILE_LIST_DIRECTORY
    Folder_Delete = FILE_DELETE_CHILD
    Folder_Execute = FILE_GENERIC_EXECUTE
    Folder_Full_Control = FILE_ALL_ACCESS
    Folder_No_Access = FILE_NO_ACCESS
End Enum

'/~ inheritence properties
Public Enum Inheritence_Flags
    '/~ singular inheritence attributes first ~/
    '/* folder and and future subfolders inherit attributes
    Container_Inherit = CONTAINER_INHERIT_ACE
    '/* folder and future files inherit attributes
    Object_Inherit = OBJECT_INHERIT_ACE
    '/* just the direct children
    Non_Propogate = NO_PROPAGATE_INHERIT_ACE
    '/* ace applies not to this object, but to child objects
    Inherit_Only = INHERIT_ONLY_ACE
    '/* apply to parent and child
    Inherit_Ace = INHERITED_ACE
    '/~ compound inheritence structures ~/
    '/* folder, and future subfolders and files inherit attributes
    Object_Container_Inherit = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE
    '/* subfolders, files, children only, one level
    Child_Inherit_Level = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE Or NO_PROPAGATE_INHERIT_ACE
    '/* subfolders, files, children only
    Child_Container_Inherit = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE Or INHERIT_ONLY_ACE
    '/* subfolders, files, parent and children
    Family_Container_Inherit = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE Or INHERITED_ACE
End Enum

'/~ dacl type
Public Enum Access_Type
    '/* permissive dacl
    Access_Allowed
    '/* deny dacl
    Access_Denied
End Enum

'/~ access flags
Public Enum REGISTRY_ACCESS
    '/* unique rights
    Registry_Read = KEY_READ
    Registry_Write = KEY_WRITE
    Registry_Execute = KEY_EXECUTE
    Registry_Full_Control = KEY_ALL_ACCESS
    '/* compound rights
    Registry_Read_Write = KEY_READ Or KEY_WRITE
    Registry_Read_Execute = KEY_READ Or KEY_EXECUTE
    Registry_Read_Write_Execute = KEY_READ Or KEY_WRITE Or KEY_EXECUTE
End Enum

Private Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, _
                                                               ByVal SecurityInformation As Long, _
                                                               ByVal pSecurityDescriptor As Long) As Long

Private Declare Function LocalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, _
                                                        ByVal wBytes As Long) As Long

Private Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, _
                                                                         ByVal hpvSource As Long, _
                                                                         ByVal cbCopy As Long)

Private Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, _
                                                                          ByVal dwRevision As Long) As Long

Private Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (ByVal lpSystemName As Long, _
                                                                                          ByVal lpAccountName As String, _
                                                                                          ByVal Sid As Long, _
                                                                                          cbSid As Long, _
                                                                                          ByVal ReferencedDomainName As String, _
                                                                                          cbReferencedDomainName As Long, _
                                                                                          peUse As Long) As Long

Private Declare Function GetLengthSid Lib "advapi32.dll" (ByVal pSid As Long) As Long

Private Declare Function InitializeAcl Lib "advapi32.dll" (ByVal pAcl As Long, _
                                                           ByVal nAclLength As Long, _
                                                           ByVal dwAclRevision As Long) As Long

Private Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, _
                                                                       ByVal bDaclPresent As Long, _
                                                                       ByVal pDacl As Long, _
                                                                       ByVal bDaclDefaulted As Long) As Long

Private Declare Function GetAce Lib "advapi32.dll" (ByVal pAcl As Long, _
                                                    ByVal dwAceIndex As Long, _
                                                    pACE As Long) As Long

Private Declare Function GetSecurityDescriptorDacl Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, _
                                                                       lpbDaclPresent As Long, _
                                                                       pDacl As Long, _
                                                                       lpbDaclDefaulted As Long) As Long

Private Declare Function GetAclInformation Lib "advapi32.dll" (ByVal pAcl As Long, _
                                                               pAclInformation As Any, _
                                                               ByVal nAclInformationLength As Long, _
                                                               ByVal dwAclInformationClass As Long) As Long

Private Declare Function GetSecurityDescriptorControl Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, _
                                                                          pControl As Long, _
                                                                          lpdwRevision As Long) As Long

Private Declare Function SetSecurityDescriptorControl Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, _
                                                                          ByVal controlBitsOfInterest As Long, _
                                                                          ByVal controlBitsToSet As Long) As Long

Private Declare Function EqualSid Lib "advapi32.dll" (ByVal pSid1 As Long, _
                                                      ByVal pSid2 As Long) As Long

Private Declare Function AddAce Lib "advapi32.dll" (ByVal pAcl As Long, _
                                                    ByVal dwAceRevision As Long, _
                                                    ByVal dwStartingAceIndex As Long, _
                                                    ByVal pAceList As Long, _
                                                    ByVal nAceListLength As Long) As Long

Private Declare Function AllocateAndInitializeSid Lib "advapi32.dll" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, _
                                                                      ByVal nSubAuthorityCount As Byte, _
                                                                      ByVal nSubAuthority0 As Long, _
                                                                      ByVal nSubAuthority1 As Long, _
                                                                      ByVal nSubAuthority2 As Long, _
                                                                      ByVal nSubAuthority3 As Long, _
                                                                      ByVal nSubAuthority4 As Long, _
                                                                      ByVal nSubAuthority5 As Long, _
                                                                      ByVal nSubAuthority6 As Long, _
                                                                      ByVal nSubAuthority7 As Long, _
                                                                      lpPSid As Long) As Long

Private Declare Sub FreeSid Lib "advapi32.dll" (ByVal pSid As Long)

Private Declare Function GetKernelObjectSecurity Lib "advapi32.dll" (ByVal hObject As Long, _
                                                                     ByVal RequestedInformation As Long, _
                                                                     ByVal pSecurityDescriptor As Long, _
                                                                     ByVal nLength As Long, _
                                                                     lpnLengthNeeded As Long) As Long

Private Declare Function SetKernelObjectSecurity Lib "advapi32.dll" (ByVal hObject As Long, _
                                                                     ByVal SecurityInformation As Long, _
                                                                     ByVal pSecurityDescriptor As Long) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, _
                                                                            ByVal dwDesiredAccess As Long, _
                                                                            ByVal dwShareMode As Long, _
                                                                            ByVal lpSecurityAttributes As Long, _
                                                                            ByVal dwCreationDisposition As Long, _
                                                                            ByVal dwFlagsAndAttributes As Long, _
                                                                            ByVal hTemplateFile As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, _
                                                               ByVal SecurityInformation As Long, _
                                                               ByVal pSecurityDescriptor As Long, _
                                                               lpcbSecurityDescriptor As Long) As Long

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

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lFN As String, _
                                                                              lFD As WFD) As Long

Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal lHD As Long, _
                                                                            lFD As WFD) As Long

Private Declare Function FindClose Lib "kernel32" (ByVal lHD As Long) As Long


'>  Borrowed User Management
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Const LOGON_WITH_PROFILE                        As Long = &H1
Private Const CREATE_DEFAULT_ERROR_MODE                 As Long = &H4000000
Private Const CREATE_NEW_CONSOLE                        As Long = &H10
Private Const CREATE_NEW_PROCESS_GROUP                  As Long = &H200
Private Const DEFAULT_LOGON                             As Double = _
    CREATE_DEFAULT_ERROR_MODE Or CREATE_NEW_CONSOLE Or _
    CREATE_NEW_PROCESS_GROUP

Private Type PROCESS_INFORMATION
    hProcess                                            As Long
    hThread                                             As Long
    dwProcessId                                         As Long
    dwThreadId                                          As Long
End Type

Private Type STARTUPINFO
    cb                                                  As Long
    lpReserved                                          As Long
    lpDesktop                                           As Long
    lpTitle                                             As Long
    dwX                                                 As Long
    dwY                                                 As Long
    dwXSize                                             As Long
    dwYSize                                             As Long
    dwXCountChars                                       As Long
    dwYCountChars                                       As Long
    dwFillAttribute                                     As Long
    dwFlags                                             As Long
    wShowWindow                                         As Integer
    cbReserved2                                         As Integer
    lpReserved2                                         As Byte
    hStdInput                                           As Long
    hStdOutput                                          As Long
    hStdError                                           As Long
End Type

Private Declare Function CreateProcessWithLogonW Lib "advapi32" (ByVal lpUserName As Long, _
                                                                 ByVal lpDomain As Long, _
                                                                 ByVal lpPassword As Long, _
                                                                 ByVal dwLogonFlags As Long, _
                                                                 ByVal lpApplicationName As Long, _
                                                                 ByVal lpCommandLine As Long, _
                                                                 ByVal dwCreationFlags As Long, _
                                                                 ByVal lpEnvironment As Long, _
                                                                 ByVal lpCurrentDirectory As Long, _
                                                                 lpStartupInfo As STARTUPINFO, _
                                                                 lpProcessInfo As PROCESS_INFORMATION) As Long

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, _
                                                                              nSize As Long) As Long

Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

Public Event eNComplete(ByVal sTask As String)
Public Event eNErrorCond(ByVal sRoutine As String, ByVal sError As String)

'/* recurse
Private m_Next                  As String
Private c_Storage               As Collection
Private W32                     As WFD

'/* enum access
Public Function Return_Access(eAccess As ACCESS_MASK) As Long
    Return_Access = eAccess
End Function

Public Function Return_Folder(eFolder As FOLDER_PERMISSIONS) As Long
    Return_Folder = eFolder
End Function

Public Function Return_Inherit(eInherit As Inheritence_Flags) As Long
    Return_Inherit = eInherit
End Function

Public Function Return_Type(eType As Access_Type) As Long
    Return_Type = eType
End Function

Public Function Return_RegAccess(eRegAccess As REGISTRY_ACCESS) As Long
    Return_RegAccess = eRegAccess
End Function

'>  Public Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'* Name           : NTFS_Folder
'* Purpose        : modify ntfs permissions
'* Inputs         : path, account, perm, inherit, access
'* Outputs        : boolean
'*********************************************
Public Function NTFS_Folder(ByVal sFolder As String, _
                            ByVal sAccount As String, _
                            ByVal eFile As FOLDER_PERMISSIONS, _
                            ByVal eFlags As Inheritence_Flags, _
                            ByVal eType As Access_Type) As Boolean

'/* set folder security (gate)

Dim lFile           As Long
Dim lReturn         As Long
Dim tAccount        As ACCOUNT_PERM
Dim tAuthority      As SID_IDENTIFIER_AUTHORITY

    '/* default struct
    With tAccount
        .AccountName = ""
        .AccessMask = GENERIC_READ
        .AceFlags = CONTAINER_INHERIT_ACE Or OBJECT_INHERIT_ACE
        .AceType = ACCESS_ALLOWED_ACE_TYPE
    End With
    
    '/* test access
    tAuthority.Value(5) = SECURITY_WORLD_SID_AUTHORITY
    If AllocateAndInitializeSid(tAuthority, 1, SECURITY_WORLD_RID, _
        0&, 0&, 0&, 0&, 0&, 0&, 0&, tAccount.pSid) = 0 Then GoTo Handler
    
    '/* set user account params
    With tAccount
        .SidPassedByCaller = True
        .AccountName = sAccount
        .AccessMask = eFile
        .AceFlags = eFlags
        .AceType = eType
        .pSid = 0
        .SidPassedByCaller = False
    End With
    
    '/* get handle to object
    lFile = CreateFile(sFolder, READ_CONTROL Or WRITE_DAC, 0&, 0&, _
        OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
    If Not lFile = 0 Then
        '/* success
        If Kernal_Permissions(lFile, tAccount) Then
            RaiseEvent eNComplete("The Object: " + sFolder + " permissions were changed successfully.")
            NTFS_Folder = True
        End If
    End If

    '/* cleanup
    If Not lFile = 0 Then CloseHandle lFile
    With tAccount
        If .pSid <> 0 And .SidPassedByCaller Then
            FreeSid (.pSid)
            .pSid = 0
        End If
    End With
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("NTFS_Folder", CStr(Err.LastDllError))
    '/* cleanup
    If Not lFile = 0 Then CloseHandle lFile
    With tAccount
        If .pSid <> 0 And .SidPassedByCaller Then
            FreeSid (.pSid)
            .pSid = 0
        End If
    End With

End Function

'* Name           : NTFS_Recursive
'* Purpose        : ntfs permissions for all children
'* Inputs         : path, account, perm, inherit, access
'* Outputs        : boolean
'*********************************************
Public Function NTFS_Recursive(ByVal sFolder As String, _
                               ByVal sAccount As String, _
                               ByVal eFile As FOLDER_PERMISSIONS, _
                               ByVal eFlags As Inheritence_Flags, _
                               ByVal eType As Access_Type, _
                               Optional ByVal bChildren As Boolean = False) As Boolean

Dim vItem               As Variant

On Error GoTo Handler

    '/* cycle through each item
    For Each vItem In Children_Return(sFolder, bChildren)
        '/* length test
        If Len(vItem) > 0 Then
            '/* apply permissions set
            If Not NTFS_Folder(CStr(vItem), sAccount, eFile, eFlags, eType) Then
                GoTo Handler
            End If
        End If
    Next vItem
    
    '/* success
    RaiseEvent eNComplete("The Object: " + sFolder + " permissions were changed successfully.")
    NTFS_Recursive = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("NTFS_Recursive", CStr(Err.LastDllError))

End Function


'>  Support Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Function Build_ACE(ByVal lAclId As Long, _
                           ByVal bType As Byte, _
                           ByVal bFlags As Byte, _
                           ByVal lMask As Long, _
                           ByVal lPointer As Long) As Boolean
'/* build an ace entry

Dim lReturn         As Long
Dim lAceLen         As Long
Dim lSidLen         As Long
Dim lAce            As Long
Dim tTempAce        As ACE

On Error GoTo Handler

    '/* get len
    lSidLen = GetLengthSid(lPointer)
    lAceLen = Len(tTempAce) + lSidLen - 4
    '/* allocate space
    lAce = LocalAlloc(lPtr, lAceLen)
    If lAce = 0 Then GoTo Handler

    '/* ace struct
    With tTempAce
        .Header.AceType = bType
        .Header.AceFlags = bFlags
        .Header.AceSize = lAceLen
        .Mask = lMask
    End With
    
    '/* copy to struct
    CopyMemory ByVal lAce, VarPtr(tTempAce), LenB(tTempAce)
    CopyMemory ByVal lAce + 8, lPointer, lSidLen
    
    '/* add to acl
    lReturn = AddAce(lAclId, ACL_REVISION, MAXDWORD, lAce, lAceLen)
    If Not lReturn = 0 Then Build_ACE = True
    LocalFree lAce
    
Handler:

End Function

Private Function Forge_Descriptor(ByVal lOldSD As Long, _
                                  tAccount As ACCOUNT_PERM, _
                                  tInfo As MEM_DATA) As Boolean

'/* reconstruct security descriptor

Dim lAceIndex           As Long
Dim lControlBits        As Long
Dim lControlSet         As Long
Dim lCount              As Long
Dim lDefault            As Long
Dim lDomain             As Long
Dim lFlag               As Long
Dim lNewACL             As Long
Dim lLength             As Long
Dim lPAcl               As Long
Dim lPresent            As Long
Dim lPsd                As Long
Dim lPointer            As Long
Dim lReturn             As Long
Dim lRevision           As Long
Dim lSidLen             As Long
Dim lTotal              As Long
Dim lUse                As Long
Dim sDomain             As String
Dim tTempACL            As ACL
Dim tempAce             As ACE
Dim tSize               As ACL_SIZE_INFORMATION

On Error GoTo Handler

    With tInfo
        .pAcl = 0
        .pSD = 0
    End With
    
    '/* get size
    lPsd = LocalAlloc(lPointer, SECURITY_DESCRIPTOR_MIN_LENGTH)
    If lPsd = 0 Then GoTo Handler

    tInfo.pSD = lPsd
    '/* init descriptor
    lReturn = InitializeSecurityDescriptor(lPsd, SECURITY_DESCRIPTOR_REVISION)
    If lReturn = 0 Then GoTo Handler

    '/* check for existing sd
    If lOldSD Then
        lReturn = GetSecurityDescriptorDacl(lOldSD, lPresent, lPAcl, lDefault)
        If lReturn = 0 Then GoTo Handler
        '/* extract dacl
        If (lPresent <> 0 And lPAcl <> 0) Then
            lReturn = GetAclInformation(lPAcl, tSize, Len(tSize), 2&)
            If lReturn = 0 Then GoTo Handler
            lTotal = tSize.AclBytesInUse
        Else
            lTotal = Len(tTempACL)
        End If
    Else
        lTotal = Len(tTempACL)
    End If
    
    '/* allocate sid
    sDomain = Space$(256)
    With tAccount
        '/* get callers sid
        If .pSid = 0 Then
            lDomain = 256
            '/* get size
            lReturn = LookupAccountName(0, .AccountName, 0, lSidLen, sDomain, lDomain, lUse)
            .pSid = LocalAlloc(lPointer, lSidLen)
            If .pSid = 0 Then GoTo Handler
            '/* get the sid
            lReturn = LookupAccountName(0, .AccountName, .pSid, lSidLen, sDomain, lDomain, lUse)
            If lReturn = 0 Then GoTo Handler
        End If
    End With
    
    '/* ace buffer
    lLength = Len(tempAce) + GetLengthSid(tAccount.pSid) - 4
    lTotal = lTotal + lLength
    lNewACL = LocalAlloc(lPointer, lTotal)
    If lNewACL = 0 Then GoTo Handler
    tInfo.pAcl = lNewACL
    
    '/* init acl
    lReturn = InitializeAcl(lNewACL, lTotal, ACL_REVISION)
    If lReturn = 0 Then GoTo Handler
    With tAccount
        '/* build dacl in sequence
        If .AceType = ACCESS_DENIED_ACE_TYPE Then
            lReturn = Build_ACE(lNewACL, .AceType, .AceFlags, .AccessMask, .pSid)
            If lReturn = 0 Then GoTo Handler
            lAceIndex = lAceIndex + 1
        End If
    End With
    
    '/* copy non-inherited ace
    If (lPresent <> 0 And lPAcl <> 0 And tSize.AceCount > 0) Then
        '/* combine old and new ACE entries
        For lCount = 0 To (tSize.AceCount - 1)
            '/* next ACE
            GetAce lPAcl, lCount, lPointer
            CopyMemory tempAce, lPointer, LenB(tempAce)
            '/* exit on inherited ace
            If ((tempAce.Header.AceFlags And INHERITED_ACE) = INHERITED_ACE) Then Exit For
            '/* check ace value
            If Not (Is_Equal(tAccount, lPointer + 8)) Then
                '/* add ace
                AddAce lNewACL, ACL_REVISION, MAXDWORD, lPointer, tempAce.Header.AceSize
                lAceIndex = lAceIndex + 1
            End If
        Next lCount
    End If
    
    '/* add explicit permit
    With tAccount
        If tAccount.AceType = ACCESS_ALLOWED_ACE_TYPE Then
            Build_ACE lNewACL, .AceType, .AceFlags, .AccessMask, .pSid
            lAceIndex = lAceIndex + 1
        End If
    End With
    
    '/* enties with inheritence flag
    If (lPresent <> 0 And lPAcl <> 0 And tSize.AceCount > 0) Then
        For lCount = lCount To (tSize.AceCount - 1)
            GetAce lPAcl, lCount, lPointer
            CopyMemory tempAce, lPointer, LenB(tempAce)
            AddAce lNewACL, ACL_REVISION, MAXDWORD, lPointer, tempAce.Header.AceSize
            lAceIndex = lAceIndex + 1
        Next lCount
    End If
    
    '/* descriptor flags
    If lOldSD <> 0 Then
        lReturn = GetSecurityDescriptorControl(lOldSD, lFlag, lRevision)
        If lReturn <> 0 Then
            If ((lFlag And SE_DACL_AUTO_INHERITED) = SE_DACL_AUTO_INHERITED) Then
                lControlBits = SE_DACL_AUTO_INHERIT_REQ Or SE_DACL_AUTO_INHERITED
                lControlSet = lControlBits
            ElseIf ((lFlag And SE_DACL_PROTECTED) = SE_DACL_PROTECTED) Then
                lControlBits = SE_DACL_PROTECTED
                lControlSet = lControlBits
            End If
            If lControlSet <> 0 Then
                SetSecurityDescriptorControl lPsd, lControlBits, lControlSet
            End If
        End If
    End If
    
    '/* add dacl
    lReturn = SetSecurityDescriptorDacl(lPsd, 1, lNewACL, 0)
    If Not lReturn = 0 Then Forge_Descriptor = True
    
    '/* free SIDs
    With tAccount
        If .pSid <> 0 And Not (.SidPassedByCaller) Then
            LocalFree (.pSid)
            .pSid = 0
        End If
    End With

On Error GoTo 0
Exit Function

Handler:
    With tInfo
        If .pSD <> 0 Then LocalFree .pSD
        If .pAcl <> 0 Then LocalFree .pAcl
        .pSD = 0
        .pAcl = 0
    End With

End Function

Private Function Kernal_Permissions(ByVal lObject As Long, _
                                    tAccount As ACCOUNT_PERM) As Boolean

'/* apply object security

Dim lLength         As Long
Dim lOldSD          As Long
Dim lReturn         As Long
Dim tInfo           As MEM_DATA

On Error GoTo Handler

    With tInfo
        .pAcl = 0
        .pSD = 0
    End With
    
    '/* allocate space
    GetKernelObjectSecurity lObject, DACL_SECURITY_INFORMATION, 0&, lLength, lLength
    If lLength = 0 Then GoTo Handler
    lOldSD = LocalAlloc(lPtr, lLength)
    If lOldSD = 0 Then GoTo Handler

    '/* get the structure
    lReturn = GetKernelObjectSecurity(lObject, DACL_SECURITY_INFORMATION, lOldSD, lLength, lLength)
    If lReturn = 0 Then GoTo Handler
    
    '/* initialize the descriptor
    If Forge_Descriptor(lOldSD, tAccount, tInfo) Then
        '/* load the structure
        lReturn = SetKernelObjectSecurity(lObject, DACL_SECURITY_INFORMATION, tInfo.pSD)
        '/* success
        If Not lReturn = 0 Then Kernal_Permissions = True
    End If

On Error GoTo 0

Handler:
    '/* cleanup
    If lOldSD <> 0 Then LocalFree lOldSD
    lOldSD = 0
    With tInfo
        If Not .pSD = 0 Then LocalFree .pSD
        If Not .pAcl = 0 Then LocalFree .pAcl
        .pSD = 0
        .pAcl = 0
    End With

End Function

Private Function Is_Equal(Account As ACCOUNT_PERM, _
                          lSid As Long) As Boolean

    Is_Equal = (EqualSid(Account.pSid, lSid))

End Function

Private Function Children_Return(ByVal sPath As String, _
                                 ByVal bChildren As Boolean) As Collection

'/* recurse and return children

On Error GoTo Handler

    Set c_Storage = New Collection
    '/* get paths
    Children_Recurse sPath, bChildren
    '/* valid return
    If c_Storage.Count > 1 Then
        Set Children_Return = c_Storage
        '/* add root
        Children_Return.Add sPath
        Set c_Storage = Nothing
    Else
        '/* root only
        Set Children_Return = New Collection
        Children_Return.Add sPath
    End If

On Error GoTo 0

Handler:

End Function

Private Sub Children_Recurse(ByVal sPath As String, _
                             ByVal bChildren As Boolean)

'/* dump folders and [children] into a collection

Dim bValue              As Boolean
Dim lHandle             As Long

On Error Resume Next

    lHandle = FindFirstFile(sPath + "*.*", W32)
    bValue = True
    Do While bValue
        With W32
            m_Next = Left$(.cFN, InStr(1, .cFN, Chr$(0)) - 1)
        End With
        If Not LenB(m_Next) > 0 Then GoTo Skip
        If AscW(m_Next) = 46 Then GoTo Skip
        Select Case True
        '/* add folders
        Case W32.lFA And &H10
            c_Storage.Add sPath + m_Next + Chr$(92)
            Children_Recurse sPath + m_Next + Chr$(92), bChildren
        '/* add files
        Case Else
            If bChildren Then
                c_Storage.Add sPath + m_Next
            End If
        End Select
Skip:
        bValue = FindNextFile(lHandle, W32)
    Loop
    FindClose lHandle

On Error GoTo 0

End Sub





'* Name           : NTFS_Key
'* Purpose        : modify key permissions
'* Inputs         : hkey, path, account, mask, access, inheritence
'* Outputs        : boolean
'*********************************************
Public Function NTFS_Key(ByVal lHKey As Long, _
                         ByVal sKey As String, _
                         ByVal sAccount As String, _
                         ByVal eMask As REGISTRY_ACCESS, _
                         ByVal eType As Access_Type, _
                         ByVal eFlags As Inheritence_Flags) As Boolean

'/* set key permissions (gate)

Dim lKey            As Long
Dim lReturn         As Long
Dim tAccount        As ACCOUNT_PERM
Dim tAuthority      As SID_IDENTIFIER_AUTHORITY

On Error GoTo Handler

    '/* default account
    With tAccount
        .AccountName = ""
        .AccessMask = GENERIC_READ
        .AceFlags = CONTAINER_INHERIT_ACE
        .AceType = ACCESS_ALLOWED_ACE_TYPE
    End With
    
    '/* test access
    tAuthority.Value(5) = SECURITY_WORLD_SID_AUTHORITY
    lReturn = AllocateAndInitializeSid(tAuthority, 1, SECURITY_WORLD_RID, _
        0&, 0&, 0&, 0&, 0&, 0&, 0&, tAccount.pSid)
    If lReturn = 0 Then GoTo Handler
    
    '/* set up account
    With tAccount
        .AccountName = sAccount
        .AccessMask = eMask
        .AceFlags = eFlags
        .AceType = eType
        .pSid = 0
        .SidPassedByCaller = False
    End With
    
    '/* open the key
    lReturn = RegOpenKeyEx(lHKey, sKey, 0&, READ_CONTROL Or WRITE_DAC, lKey)
    '/* apply change to key
    If lReturn = 0 Then
        If Key_Permissions(lKey, tAccount) Then
            RaiseEvent eNComplete("The Key: " + sKey + " permissions were changed successfully.")
            NTFS_Key = True
        End If
    End If
    
    '/* cleanup
    RegCloseKey lKey
    With tAccount
        If .pSid <> 0 And .SidPassedByCaller Then
            FreeSid .pSid
            .pSid = 0
        End If
    End With

Exit Function
On Error GoTo 0

Handler:
    RaiseEvent eNErrorCond("NTFS_Key", CStr(Err.LastDllError))

End Function

Private Function Key_Permissions(ByVal lHKey As Long, _
                                 tAccount As ACCOUNT_PERM) As Boolean

'/* apply key permissions

Dim lLength         As Long
Dim lOldSD          As Long
Dim lReturn         As Long
Dim tInfo           As MEM_DATA

On Error GoTo Handler

    With tInfo
        .pAcl = 0
        .pSD = 0
    End With
    
    '/* allocate space
    lReturn = RegGetKeySecurity(lHKey, DACL_SECURITY_INFORMATION, 0&, lLength)
    If lLength = 0 Then GoTo Handler
    lOldSD = LocalAlloc(lPtr, lLength)
    If lOldSD = 0 Then GoTo Handler

    '/* load structure
    lReturn = RegGetKeySecurity(lHKey, DACL_SECURITY_INFORMATION, lOldSD, lLength)
    If lReturn = 0 Then
        '/* create dacl
        If Forge_Descriptor(lOldSD, tAccount, tInfo) Then
            '/* set descriptor
            lReturn = RegSetKeySecurity(lHKey, DACL_SECURITY_INFORMATION, tInfo.pSD)
            If lReturn = 0 Then
                '/* success
                Key_Permissions = True
            End If
        End If
    End If

On Error GoTo 0

Handler:
    '/* cleanup
    If Not lOldSD = 0 Then LocalFree lOldSD
    lOldSD = 0
    With tInfo
        If Not .pSD = 0 Then LocalFree tInfo.pSD
        If Not .pAcl = 0 Then LocalFree tInfo.pAcl
        .pSD = 0
        .pAcl = 0
    End With

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

'>  User Management Function
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'* Name           : User_RunAs
'* Purpose        : start process as alternate user
'* Inputs         : user, pass, domain, command
'* Outputs        : boolean
'*********************************************
Private Function User_RunAs(ByVal sUserName As String, _
                            ByVal sPassword As String, _
                            ByVal sDomain As String, _
                            ByVal sCommand As String) As Boolean

Dim lReturn         As Long
Dim sApplication    As String
Dim sDirectory      As String
Dim tPInfo          As PROCESS_INFORMATION
Dim tStart          As STARTUPINFO

On Error GoTo Handler

    '/* default struct
    sApplication = ""
    sDirectory = ""
    With tStart
        .cb = LenB(tStart)
        .dwFlags = 0&
    End With
    
    '/* launch process
    lReturn = CreateProcessWithLogonW(StrPtr(sUserName), StrPtr(sDomain), _
        StrPtr(sPassword), &H1, StrPtr(sApplication), StrPtr(sCommand), _
        DEFAULT_LOGON, ByVal 0&, StrPtr(sDirectory), tStart, tPInfo)
        '/* success
        If lReturn = 0 Then
            RaiseEvent eNComplete("The Application: " + sCommand + " launched successfully.")
            User_RunAs = True
        End If
        
    '/* cleanup
    With tPInfo
        CloseHandle .hThread
        CloseHandle .hProcess
    End With

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("User_RunAs", CStr(Err.LastDllError))

End Function

'* Name           : User_Name
'* Purpose        : get logged in user name
'* Inputs         : none
'* Outputs        : string
'*********************************************
Public Function User_Name() As String

Dim sName   As String

    sName = String(255, Chr$(0))
    GetUserName sName, 255
    User_Name = Left$(sName, InStr(1, sName, Chr$(0)) - 1)
    
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

