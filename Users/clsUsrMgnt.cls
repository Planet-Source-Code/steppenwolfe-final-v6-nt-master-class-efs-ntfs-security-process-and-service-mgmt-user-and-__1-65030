VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsrMgnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************************
'*  NMC! V1.5    Service Managment Member Class                                        *
'*                                                                                     *
'*  Created:     April 14, 2006                                                        *
'*  Updated:     April 30, 2006                                                        *
'*  Purpose:     NT System Security Master Class                                       *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.5.0                                                                 *
'*  Compile:     Native                                                                *
'*  Referenced:  Throughout Project                                                    *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*                                                                                     *
'***************************************************************************************

'/~ User and Group Management ~/
'/~ User_Create             - create a new user         [in -string (6) | out -bool]
'/~ User_Delete             - delete a user account     [in -string (2) | out -bool]
'/~ User_Exist              - test a user account       [in -string (2) | out -bool]
'/~ User_Data               - get user profile data     [in -string (2) | out -bool]
'/~ User_LoadProfile        - load the users key        [in string (4) | out -bool]
'/~ User_UnloadProfile      - unload a key              [in string | out -bool]
'/~ Users_List              - list local|global users   [in -string (3) | out -col]
'/~ User_Impersonate        - impersonate a user        [in -string (3) | out -bool]
'/~ User_Revert             - revert to native token    [out -bool]
'/~ User_RunAs              - launch process as user    [in -string (4) | out -bool]
'/~ User_Name               - logged in user            [out -string]
'/~ User_SetData            - apply account changes     [in -string (2) + enum | out -bool]
'/~ User_Password           - change user password      [in -string (4) | out -bool]
'/~ Group_Create            - create a user group       [in -string (3) | out -bool]
'/~ Group_Add               - add user to group         [in -string (3) | out -bool]
'/~ Group_Remove            - remove user from group    [in -string (3) | out -bool]
'/~ Group_Delete            - delete a group            [in -string (3) | out -bool]
'/~ Groups_List             - list built-in groups      [in -none | out -col]
'/~ Computer_Name           - local computer name       [in -none | out -string]
'/~ Get_Domain              - get primary dc name       [in -string | out -string]


'/* netusersetinfo return params
Private Const USER_NAME_PARMNUM                         As Long = &H1
Private Const USER_PASSWORD_PARMNUM                     As Long = &H3
Private Const USER_PASSWORD_AGE_PARMNUM                 As Long = &H4
Private Const USER_PRIV_PARMNUM                         As Long = &H5
Private Const USER_HOME_DIR_PARMNUM                     As Long = &H6
Private Const USER_COMMENT_PARMNUM                      As Long = &H7
Private Const USER_FLAGS_PARMNUM                        As Long = &H8
Private Const USER_SCRIPT_PATH_PARMNUM                  As Long = &H9
Private Const USER_AUTH_FLAGS_PARMNUM                   As Long = &HA
Private Const USER_FULL_NAME_PARMNUM                    As Long = &HB
Private Const USER_USR_COMMENT_PARMNUM                  As Long = &HC
Private Const USER_PARMS_PARMNUM                        As Long = &HD
Private Const USER_WORKSTATIONS_PARMNUM                 As Long = &HE
Private Const USER_LAST_LOGON_PARMNUM                   As Long = &HF
Private Const USER_LAST_LOGOFF_PARMNUM                  As Long = &H10
Private Const USER_ACCT_EXPIRES_PARMNUM                 As Long = &H11
Private Const USER_MAX_STORAGE_PARMNUM                  As Long = &H12
Private Const USER_UNITS_PER_WEEK_PARMNUM               As Long = &H13
Private Const USER_LOGON_HOURS_PARMNUM                  As Long = &H14
Private Const USER_PAD_PW_COUNT_PARMNUM                 As Long = &H15
Private Const USER_NUM_LOGONS_PARMNUM                   As Long = &H16
Private Const USER_LOGON_SERVER_PARMNUM                 As Long = &H17
Private Const USER_COUNTRY_CODE_PARMNUM                 As Long = &H18
Private Const USER_CODE_PAGE_PARMNUM                    As Long = &H19
Private Const USER_PRIMARY_GROUP_PARMNUM                As Long = &H33
Private Const USER_PROFILE                              As Long = &H34
Private Const USER_PROFILE_PARMNUM                      As Long = &H34
Private Const USER_HOME_DIR_DRIVE_PARMNUM               As Long = &H35
'/* user params
Private Const USER_PRIV_MASK                            As Long = &H3
Private Const USER_PRIV_GUEST                           As Long = &H0
Private Const USER_PRIV_USER                            As Long = &H1
Private Const USER_PRIV_ADMIN                           As Long = &H2
Private Const UF_SCRIPT                                 As Long = &H1
Private Const UF_ACCOUNTDISABLE                         As Long = &H2
Private Const UF_HOMEDIR_REQUIRED                       As Long = &H8
Private Const UF_LOCKOUT                                As Long = &H10
Private Const UF_PASSWD_NOTREQD                         As Long = &H20
Private Const UF_PASSWD_CANT_CHANGE                     As Long = &H40
Private Const UF_NORMAL_ACCOUNT                         As Long = &H200
Private Const TOKEN_QUERY                               As Long = &H8&
Private Const TOKEN_ADJUST_PRIVILEGES                   As Long = &H20&
Private Const SE_PRIVILEGE_ENABLED                      As Long = &H2
Private Const REG_FORCE_RESTORE                         As Long = 8&
Private Const READ_CONTROL                              As Long = &H20000
Private Const SYNCHRONIZE                               As Long = &H100000
Private Const STANDARD_RIGHTS_READ                      As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE                     As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_ALL                       As Long = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL                       As Long = &HFFFF
Private Const KEY_QUERY_VALUE                           As Long = &H1
Private Const KEY_SET_VALUE                             As Long = &H2
Private Const KEY_CREATE_SUB_KEY                        As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS                    As Long = &H8
Private Const KEY_NOTIFY                                As Long = &H10
Private Const KEY_CREATE_LINK                           As Long = &H20
'/* loadprofile constants
Private Const LOGON32_LOGON_INTERACTIVE                 As Integer = 2
Private Const LOGON32_PROVIDER_DEFAULT                  As Integer = 0
'/* check credentials constants
Private Const TOKEN_READ                                As Long = &H20008
Private Const SECURITY_BUILTIN_DOMAIN_RID               As Long = &H20&
Private Const DOMAIN_ALIAS_RID_ADMINS                   As Long = &H220&
Private Const SECURITY_NT_AUTHORITY                     As Long = &H5
Private Const TOKEN_GROUPS                              As Long = 2
Private Const FILTER_NORMAL_ACCOUNT                     As Long = &H2
Private Const LOGON_WITH_PROFILE                        As Long = &H1
Private Const CREATE_DEFAULT_ERROR_MODE                 As Long = &H4000000
Private Const CREATE_NEW_CONSOLE                        As Long = &H10
Private Const CREATE_NEW_PROCESS_GROUP                  As Long = &H200
Private Const PI_NOUI                                   As Long = &H1
Private Const PI_APPLYPOLICY                            As Long = &H2

Private Const DEFAULT_LOGON                             As Double = _
    CREATE_DEFAULT_ERROR_MODE Or CREATE_NEW_CONSOLE Or _
    CREATE_NEW_PROCESS_GROUP

'/* sid id
Private Type SID_IDENTIFIER_AUTHORITY
    Value(6)                                            As Byte
End Type

Private Type PROFILE_INFORMATION
    dwSize                                              As Long
    dwFlags                                             As Long
    sUserName                                           As String
    lpProfilePath                                       As Long
    sDefaultPath                                        As String
    lpServerName                                        As Long
    lpPolicyPath                                        As Long
    hProfile                                            As Long
End Type

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

'/* user structure
Private Type USER_ATTRIBUTES
    ptrname                                             As Long
    ptrPassword                                         As Long
    dwPasswordAge                                       As Long
    dwPriv                                              As Long
    ptrHomeDir                                          As Long
    ptrcomment                                          As Long
    dwFlags                                             As Long
    ptrScriptHomeDir                                    As Long
End Type

Private Type GROUP_INFO
    tGroup                                              As String
    tComment                                            As String
End Type

Private Type LOCALGROUP_MEMBER
    UNamePtr                                            As Long
End Type

Private Type SID_AND_ATTRIBUTES
    Sid                                                 As Long
    Attributes                                          As Long
End Type

Private Type TOKEN_GROUPS
    GroupCount                                          As Long
    Groups(500)                                         As SID_AND_ATTRIBUTES
End Type

Private Type LOCALGROUP_MEMBERS_INFO_2
    lgrmi2_sid As Long
    lgrmi2_sidusage As Long
    lgrmi2_domainandname As Long
End Type

Private Type USER_INFO_3
    usri3_name                                          As Long
    usri3_password                                      As Long
    usri3_password_age                                  As Long
    usri3_priv                                          As Long
    usri3_home_dir                                      As Long
    usri3_comment                                       As Long
    usri3_flags                                         As Long
    usri3_script_path                                   As Long
    usri3_auth_flags                                    As Long
    usri3_full_name                                     As Long
    usri3_usr_comment                                   As Long
    usri3_parms                                         As Long
    usri3_workstations                                  As Long
    usri3_last_logon                                    As Long
    usri3_last_logoff                                   As Long
    usri3_acct_expires                                  As Long
    usri3_max_storage                                   As Long
    usri3_units_per_week                                As Long
    usri3_logon_hours                                   As Byte
    usri3_bad_pw_count                                  As Long
    usri3_num_logons                                    As Long
    usri3_logon_server                                  As String
    usri3_country_code                                  As Long
    usri3_code_page                                     As Long
    usri3_user_id                                       As Long
    usri3_primary_group_id                              As Long
    usri3_profile                                       As Long
    usri3_home_dir_drive                                As Long
    usri3_password_expired                              As Long
End Type

Private Type USER_INFO_1003
    usri1003_password As Long
End Type

Private Type USER_INFO_1008
    usri1008_flags As Long
End Type

Public Enum eAccountType
    Temp_Duplicate = &H100
    Normal_Account = &H200
    InterDomain_Trust = &H800
    WorkStation_Trust = &H1000
    Server_Trust = &H2000
End Enum

Public Enum eAccountFlags
    User_Script = &H1
    Account_Disabled = &H2
    Require_HomeDir = &H8
    Account_Lockout = &H10
    No_Password = &H20
    Password_NoChange = &H40
End Enum

Public Enum eAccountAccess
    User_Guest = 0
    User_Privileged = 1
    User_Admin = 2
End Enum

Public Enum eUserFlags
    User_ScriptReqd = UF_SCRIPT
    User_Disabled = UF_ACCOUNTDISABLE
    User_HomeReqd = UF_HOMEDIR_REQUIRED
    User_Lockout = UF_LOCKOUT
    User_PassNotRequired = UF_PASSWD_NOTREQD
    User_PassNoChange = UF_PASSWD_CANT_CHANGE
    User_Normal = UF_NORMAL_ACCOUNT
End Enum

'NetUserChangePassword (
'    IN  LPCWSTR   domainname OPTIONAL,
'    IN  LPCWSTR   username OPTIONAL,
'    IN  LPCWSTR   oldpassword,
'    IN  LPCWSTR   newpassword
'    );

'NetUserSetInfo (
'    IN  LPCWSTR    servername OPTIONAL,
'    IN  LPCWSTR    username,
'    IN  DWORD     level,
'    IN  LPBYTE    buf,
'    OUT LPDWORD   parm_err OPTIONAL
'    );
Private Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, _
                                                                            ByVal lpSubKey As String, _
                                                                            ByVal lpFile As String) As Long
                                                                            
Private Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String) As Long

Private Declare Function LoadUserProfile Lib "userenv.dll" Alias "LoadUserProfileA" (ByVal hToken As Long, _
                                                                                     ByRef lpProfileInfo As PROFILE_INFORMATION) As Boolean
                                                                                     
Private Declare Function UnloadUserProfile Lib "userenv.dll" (ByVal hToken As Long, _
                                                              ByVal hProfile As Long) As Long



Private Declare Function NetUserSetInfo Lib "Netapi32" (servername As Byte, _
                                                        username As Byte, _
                                                        ByVal level As Long, _
                                                        bufptr As Any, _
                                                        parmerr As Long) As Long

Private Declare Function NetUserChangePassword Lib "Netapi32" (domainname As Any, _
                                                               username As Any, _
                                                               oldpassword As Byte, _
                                                               newpassword As Byte) As Long

Private Declare Function NetGetDCName Lib "NETAPI32.dll" (servername As Long, _
                                                          domainname As Byte, _
                                                          bufptr As Long) As Long

Private Declare Function NetUserGetInfo Lib "Netapi32" (lpServer As Any, _
                                                        username As Byte, _
                                                        ByVal level As Long, _
                                                        lpBuffer As Long) As Long

Private Declare Function NetLocalGroupDel Lib "NETAPI32.dll" (servername As Byte, _
                                                              groupname As Byte) As Long

Private Declare Function NetGroupAdd1 Lib "NETAPI32.dll" Alias "NetGroupAdd" (servername As Byte, _
                                                                              ByVal level As Long, _
                                                                              Buffer As GROUP_INFO, _
                                                                              ParmError As Long) As Long
    
Private Declare Function NetUserEnum Lib "NETAPI32.dll" (servername As Byte, _
                                                         ByVal level As Long, _
                                                         ByVal lFilter As Long, _
                                                         Buffer As Long, _
                                                         ByVal prefmaxlen As Long, _
                                                         entriesread As Long, _
                                                         totalentries As Long, _
                                                         ResumeHandle As Long) As Long

Private Declare Function NetGroupGetUsers Lib "NETAPI32.dll" (servername As Byte, _
                                                              groupname As Byte, _
                                                              ByVal level As Long, _
                                                              Buffer As Long, _
                                                              ByVal prefmaxlen As Long, _
                                                              entriesread As Long, _
                                                              totalentries As Long, _
                                                              ResumeHandle As Long) As Long

Private Declare Function NetLocalGroupGetMembers Lib "NETAPI32.dll" (servername As Byte, _
                                                                    localgroupname As Byte, _
                                                                    ByVal level As Long, _
                                                                    Buffer As Long, _
                                                                    ByVal prefmaxlen As Long, _
                                                                    entriesread As Long, _
                                                                    totalentries As Long, _
                                                                    ResumeHandle As Long) As Long

Private Declare Function NetGroupDel Lib "NETAPI32.dll" (servername As Byte, _
                                                         groupname As Byte) As Long
    
Private Declare Function NetApiBufferFree Lib "Netapi32" (ByVal Buffer As Long) As Long

Private Declare Function NetLocalGroupAdd Lib "NETAPI32.dll" (servername As Byte, _
                                                              ByVal level As Long, _
                                                              Buffer As GROUP_INFO, _
                                                              parm_err As Long) As Long

Private Declare Function NetLocalGroupAddMembers Lib "NETAPI32.dll" (servername As Byte, _
                                                                     groupname As Byte, _
                                                                     ByVal level As Long, _
                                                                     buf As LOCALGROUP_MEMBER, _
                                                                     ByVal totalentries As Long) As Long

Private Declare Function NetLocalGroupDelMembers Lib "NETAPI32.dll" (servername As Byte, _
                                                                     groupname As Byte, _
                                                                     ByVal level As Long, _
                                                                     buf As LOCALGROUP_MEMBER, _
                                                                     ByVal totalentries As Long) As Long

Private Declare Function NetAPIBufferAllocate Lib "NETAPI32.dll" Alias "NetApiBufferAllocate" (ByVal ByteCount As Long, _
                                                                                               Ptr As Long) As Long

Private Declare Function NetUserDel Lib "NETAPI32.dll" (servername As Byte, _
                                                        username As Byte) As Long

Private Declare Function NetUserAdd1 Lib "NETAPI32.dll" Alias "NetUserAdd" (servername As Byte, _
                                                                            ByVal level As Long, _
                                                                            Buffer As USER_ATTRIBUTES, _
                                                                            lParmError As Long) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSource As Any, _
                                                                     ByVal dwLength As Long)

Private Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codepage As Long, _
                                                             ByVal dwFlags As Long, _
                                                             lpWideCharStr As Any, _
                                                             ByVal cchWideChar As Long, _
                                                             lpMultiByteStr As Any, _
                                                             ByVal cchMultiByte As Long, _
                                                             ByVal lpDefaultChar As String, _
                                                             ByVal lpUsedDefaultChar As Long) As Long

Private Declare Function StrToPtr Lib "kernel32" Alias "lstrcpyW" (ByVal Ptr As Long, _
                                                                   Source As Byte) As Long

Private Declare Function GetComputerName Lib "Kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, _
                                                                                      nSize As Long) As Long

Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, _
                                                                                        ByVal Sid As Long, _
                                                                                        ByVal Name As String, _
                                                                                        cbName As Long, _
                                                                                        ByVal ReferencedDomainName As String, _
                                                                                        cbReferencedDomainName As Long, _
                                                                                        peUse As Long) As Long

Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, _
                                                              ByVal DesiredAccess As Long, _
                                                              TokenHandle As Long) As Long

Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, _
                                                                 ByVal TokenInformationClass As Long, _
                                                                 TokenInformation As Any, _
                                                                 ByVal TokenInformationLength As Long, _
                                                                 ReturnLength As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" (ByVal lpszUsername As String, _
                                                                          ByVal lpszDomain As String, _
                                                                          ByVal lpszPassword As String, _
                                                                          ByVal dwLogonType As Long, _
                                                                          ByVal dwLogonProvider As Long, _
                                                                          phToken As Long) As Long

Private Declare Function ImpersonateLoggedOnUser Lib "advapi32" (ByVal hToken As Long) As Long

Private Declare Function RevertToSelf Lib "advapi32.dll" () As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal Length As Long)

Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, _
                                                                  ByVal Ptr As Long) As Long

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

Private Declare Sub lstrcpyW Lib "kernel32" (dest As Any, ByVal src As Any)

Private Declare Sub FreeSid Lib "advapi32.dll" (ByVal pSid As Long)


Public Event eNComplete(ByVal sTask As String)
Public Event eNErrorCond(ByVal sRoutine As String, ByVal sError As String)

Private m_lToken        As Long


'>  User Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


'***************************************************************************************
'*                                  USERS AND GROUPS
'***************************************************************************************

'>  User Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


'* Name           : User_Create
'* Purpose        : create a new user
'* Inputs         : machine name, user, password, directory, comment, script
'* Outputs        : boolean
'*********************************************
Public Function User_Create(ByVal sMachine As String, _
                            ByVal sUserName As String, _
                            ByVal sPassword As String, _
                            ByVal sDirectory As String, _
                            Optional ByVal sComment As String, _
                            Optional ByVal sScript As String) As Boolean

Dim bMachine()          As Byte
Dim bUserName()         As Byte
Dim bPassword()         As Byte
Dim bDirectory()        As Byte
Dim bComment()          As Byte
Dim bScript()           As Byte
Dim lReturn             As Long
Dim lParm               As Long
Dim lName               As Long
Dim lPass               As Long
Dim lDir                As Long
Dim lCmt                As Long
Dim lScript             As Long
Dim UserStruct          As USER_ATTRIBUTES

On Error GoTo Handler

    '/* add to byte arrays
    bMachine = sMachine + vbNullChar
    bUserName = sUserName + vbNullChar
    bPassword = sPassword + vbNullChar
    bDirectory = sDirectory + vbNullChar
    bComment = sComment + vbNullChar
    bScript = sScript + vbNullChar
    
    '/* create buffers
    NetAPIBufferAllocate UBound(bUserName) + 1, lName
    NetAPIBufferAllocate UBound(bPassword) + 1, lPass
    NetAPIBufferAllocate UBound(bDirectory) + 1, lDir
    NetAPIBufferAllocate UBound(bComment) + 1, lCmt
    NetAPIBufferAllocate UBound(bScript) + 1, lScript
    
    '/* get pointer and pass to user structure
    StrToPtr lName, bUserName(0)
    StrToPtr lPass, bPassword(0)
    StrToPtr lDir, bDirectory(0)
    StrToPtr lCmt, bComment(0)
    StrToPtr lScript, bScript(0)
    
    '/* fill structure
    With UserStruct
        .ptrname = lName
        .ptrPassword = lPass
        .dwPasswordAge = 3
        .dwPriv = USER_PRIV_USER
        .ptrHomeDir = lDir
        .ptrcomment = lCmt
        .dwFlags = UF_NORMAL_ACCOUNT Or UF_SCRIPT Or UF_PASSWD_CANT_CHANGE
        .ptrScriptHomeDir = lScript
    End With
    
    '/* call to add user
    lReturn = NetUserAdd1(bMachine(0), 1, UserStruct, lParm)
    
    Select Case lReturn
    Case 0
        '/* success
        RaiseEvent eNComplete("The User account was Created successfully")
        User_Create = True
    Case 2224
        RaiseEvent eNErrorCond("User_Create", "User already exists. Error #2224")
        '/* user exists
    Case Else
        RaiseEvent eNErrorCond("User_Create", "An unknown error has occured. Error# " + CStr(lReturn))
    End Select
    
    '/* free the buffers
    NetApiBufferFree lName
    NetApiBufferFree lPass
    NetApiBufferFree lDir
    NetApiBufferFree lCmt
    NetApiBufferFree lScript
    Exit Function
    
Handler:
    RaiseEvent eNErrorCond("User_Create", "An unknown error has occured. Error# " + CStr(Err.LastDllError))
    On Error GoTo 0

End Function

'* Name           : User_Delete
'* Purpose        : delete a user account
'* Inputs         : machine name, user name
'* Outputs        : boolean
'*********************************************
Public Function User_Delete(ByVal sMachine As String, _
                            ByVal sUserName As String) As Boolean

Dim bMachine()          As Byte
Dim bUserName()         As Byte
Dim lReturn             As Long
Dim bServer             As String

On Error GoTo Handler

    '/* prep vars
    bUserName = sUserName + Chr$(0)
    bMachine = sMachine + Chr$(0)
    '/* call delete user
    lReturn = NetUserDel(bMachine(0), bUserName(0))
    '/* success
    If lReturn = 0 Then
        RaiseEvent eNComplete("The User account was Deleted successfully")
         User_Delete = True
    Else
        GoTo Handler
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("User_Delete", "An unknown error has occured. Error# " + CStr(lReturn))

End Function

'* Name           : User_Exist
'* Purpose        : test a users existence
'* Inputs         : machine name, user name
'* Outputs        : boolean
'*********************************************
Public Function User_Exist(ByVal sMachine As String, _
                           ByVal sUserName As String) As Boolean

Dim bMachine()          As Byte
Dim bUserName()         As Byte
Dim lBuffer             As Long
Dim lReturn             As Long

On Error GoTo Handler

    bUserName = sUserName + Chr$(0)
    bMachine = sMachine + Chr$(0)
    '/* test for the account
    lReturn = NetUserGetInfo(bMachine(0), bUserName(0), 3, lBuffer)
    '/* success
    If lReturn = 0 Then
        RaiseEvent eNComplete("The User account: " + sUserName + " Exists.")
        User_Exist = True
    Else
        RaiseEvent eNComplete("The User account: " + sUserName + " does Not Exist.")
    End If

On Error GoTo 0
Exit Function
    
Handler:
    RaiseEvent eNErrorCond("User_Exist", "An unknown error has occured. Error# " + CStr(Err.LastDllError))

End Function

'* Name           : User_Password
'* Purpose        : change the users password
'* Inputs         : machine name, user name, old password, new password
'* Outputs        : boolean
'*********************************************
Public Function User_Password(ByVal sMachine As String, _
                              ByVal sUserName As String, _
                              ByVal sOldPass As String, _
                              ByVal sNewPass As String) As Boolean

Dim bMachine()          As Byte
Dim bUserName()         As Byte
Dim bOldPass()          As Byte
Dim bNewPass()          As Byte
Dim lReturn             As Long
Dim tUserInfo           As USER_INFO_1003

On Error GoTo Handler

    bMachine = sMachine + Chr$(0)
    bUserName = sUserName + Chr$(0)
    bNewPass = sNewPass + Chr$(0)

    If sOldPass = "" Then
        tUserInfo.usri1003_password = StrPtr(bNewPass)
        lReturn = NetUserSetInfo(bMachine(0), bUserName(0), 1003, tUserInfo, 0&)
    Else
        bOldPass = sOldPass + Chr$(0)
        lReturn = NetUserChangePassword(bMachine(0), bUserName(0), bOldPass(0), bNewPass(0))
    End If

    '/* success
    If lReturn = 0 Then
        RaiseEvent eNComplete("The User account: " + sUserName + " password has been changed.")
        User_Password = True
    ElseIf lReturn = 2221 Then
        RaiseEvent eNComplete("The User account: " + sUserName + " does not exist.")
    End If

On Error GoTo 0
Exit Function
    
Handler:
    RaiseEvent eNErrorCond("User_Password", "An unknown error has occured. Error# " + CStr(Err.LastDllError))

End Function

'* Name           : User_SetData
'* Purpose        : set user account flags
'* Inputs         : machine name, user name, flags
'* Outputs        : boolean
'*********************************************
Public Function User_SetData(ByVal sMachine As String, _
                             ByVal sUserName As String, _
                             eFlags As eUserFlags) As Boolean

Dim bServer()           As Byte
Dim bUserName()         As Byte
Dim lParam              As Long
Dim lReturn             As Long
Dim tFlags              As USER_INFO_1008

On Error GoTo Handler

    bUserName = sUserName & Chr$(0)
    bServer = sMachine & Chr$(0)
    '/* script is required
    tFlags.usri1008_flags = UF_SCRIPT Or eFlags
    lReturn = NetUserSetInfo(bServer(0), bUserName(0), 1008, tFlags.usri1008_flags, lParam)
    '/* success
    If lReturn = 0 Then
        RaiseEvent eNComplete("The User account: " + sUserName + " has been changed.")
        User_SetData = True
    Else
        RaiseEvent eNComplete("The User account: " + sUserName + " could not be changed.")
    End If

    NetApiBufferFree tFlags.usri1008_flags
   
On Error GoTo 0
Exit Function
    
Handler:
    RaiseEvent eNErrorCond("User_SetData", "An unknown error has occured. Error# " + CStr(lParam))

End Function

'* Name           : User_Data
'* Purpose        : get user info
'* Inputs         : machine name, user name
'* Outputs        : boolean
'*********************************************
Public Function User_Data(ByVal sMachine As String, _
                          ByVal sUserName As String) As Collection

Dim bServer()           As Byte
Dim bUserName()         As Byte
Dim lBuffer             As Long
Dim lReturn             As Long
Dim cTemp               As Collection
Dim tUData              As USER_INFO_3

On Error GoTo Handler

    '/* prep out variables
    Set cTemp = New Collection
    bServer = sMachine + Chr$(0)
    bUserName = sUserName + Chr$(0)
    '/* fill structure
    lReturn = NetUserGetInfo(bServer(0), bUserName(0), 3, lBuffer)
    If lReturn = 0 Then
        '/* put results into collection
        MoveMemory tUData, ByVal lBuffer, Len(tUData)
        With tUData
            cTemp.Add "Account Expires in: " + CStr(.usri3_acct_expires)
            cTemp.Add "Bad Password Count: " + CStr(.usri3_bad_pw_count)
            cTemp.Add "Country Code: " + CStr(.usri3_country_code)
            cTemp.Add "Number of Logons: " + CStr(.usri3_num_logons)
            cTemp.Add "Privilege Level: " + CStr(.usri3_priv)
        End With
        '/* release buffer
        NetApiBufferFree ByVal lBuffer
        '/* success
        Set User_Data = cTemp
        RaiseEvent eNComplete("User data for: " + sUserName + " has been recovered.")
    Else
        GoTo Handler
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("User_Data", "An unknown error has occured. Error# " + CStr(Err.LastDllError))

End Function

'* Name           : Users_List
'* Purpose        : list the local user accounts
'* Inputs         : machine name, group name, local
'* Outputs        : collection
'*********************************************
Public Function Users_List(ByVal sMachine As String, _
                           ByVal sGroup As String, _
                           ByVal bLocal As Boolean) As Collection

Dim bATemp(99)          As Byte
Dim bGroup()            As Byte
Dim bMachine()          As Byte
Dim lBuffer             As Long
Dim lCount              As Long
Dim lHandle             As Long
Dim lPointer            As Long
Dim lReceived           As Long
Dim lReturn             As Long
Dim lSize               As Long
Dim lTotal              As Long
Dim lUPoint             As Long
Dim sUserName           As String
Dim cTemp               As Collection
Dim tUData              As LOCALGROUP_MEMBERS_INFO_2

On Error GoTo Handler

    '/* prep containers
    Set cTemp = New Collection
    bMachine = sMachine + Chr$(0)
    bGroup = sGroup + Chr$(0)
    lBuffer = 255
    lHandle = 0
    
    '/* get all members
    If sGroup = vbNullString Then
        lReturn = NetUserEnum(bMachine(0), 0, FILTER_NORMAL_ACCOUNT, _
        lPointer, lBuffer, lReceived, lTotal, lHandle)
    Else
        '/* query local group
        If bLocal Then
            lReturn = NetLocalGroupGetMembers(bMachine(0), bGroup(0), _
            2, lPointer, lBuffer, lReceived, lTotal, lHandle)
        '/* query global group
        Else
            lReturn = NetGroupGetUsers(bMachine(0), bGroup(0), 0, _
            lPointer, lBuffer, lReceived, lTotal, lHandle)
        End If
    End If
    '/* failure
    If Not lReturn = 0 Then GoTo Handler
    
    '/* copy to struct
    lSize = Len(tUData)
    For lCount = 1 To lReceived
        Erase bATemp
        If sGroup <> vbNullString Then
            CopyMemory tUData, ByVal lPointer + (lSize * (lCount - 1)), lSize
            PtrToStr bATemp(0), tUData.lgrmi2_domainandname
        Else
            CopyMemory lUPoint, ByVal lPointer + (4 * (lCount - 1)), 4
            PtrToStr bATemp(0), lUPoint
        End If
        '/* add to temp collection
        sUserName = bATemp
        cTemp.Add sUserName
    Next lCount
    '/* success
    If cTemp.Count > 0 Then
        RaiseEvent eNComplete("User list has been recovered.")
        Set Users_List = cTemp
    End If
    '/* cleanup
    NetApiBufferFree lPointer

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Users_List", "An unknown error has occured. Error# " + CStr(Err.LastDllError))

End Function

'* Name           : User_Impersonate
'* Purpose        : process impersonate user
'* Inputs         : user, pass, domain
'* Outputs        : boolean
'*********************************************
Public Function User_Impersonate(ByVal sUser As String, _
                                 ByVal sPass As String, _
                                 ByVal sDomain As String) As Boolean

Dim lHandle             As Long
Dim lProfile            As Long

On Error GoTo Handler

    '/* user logon
    If LogonUser(sUser, sDomain, sPass, LOGON32_LOGON_INTERACTIVE, _
        LOGON32_PROVIDER_DEFAULT, m_lToken) <> 0 Then
        '/* test account token
        If ImpersonateLoggedOnUser(m_lToken) <> 0 Then
            User_Impersonate = True
            RaiseEvent eNComplete("Impersonation of: " + sUser + " was successful.")
        Else
            GoTo Handler
        End If
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("User_Impersonate", "An unknown error has occured. Error# " + CStr(Err.LastDllError))
    
End Function

'* Name           : User_Revert
'* Purpose        : restore user token
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Function User_Revert() As Boolean

Dim lReturn             As Long

    '/* release token
    CloseHandle m_lToken
    '/* revert
    lReturn = RevertToSelf
    '/* success
    If Not lReturn = 0 Then
        User_Revert = True
        RaiseEvent eNComplete("Reversion to native token was successful.")
    Else
        GoTo Handler
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("User_Revert", "An unknown error has occured. Error# " + CStr(Err.LastDllError))
    
End Function

'* Name           : User_LoadHive
'* Purpose        : load a key into the registry
'* Inputs         : root key, key name, key path
'* Outputs        : boolean
'*********************************************
Public Function User_LoadHive(ByVal lRootKey As Long, _
                              ByVal sKey As String, _
                              ByVal sPath As String) As Boolean

Dim lReturn             As Long

On Error GoTo Handler

    '/* load the key
    lReturn = RegLoadKey(lRootKey, sKey, sPath)
    
    '/* success
    If lReturn = 0 Then
        User_LoadHive = True
        RaiseEvent eNComplete("The Key " + sKey + " was loaded successfully.")
    Else
        GoTo Handler
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("User_LoadHive", "An unknown error has occured. Error# " + CStr(Err.LastDllError))

End Function

'* Name           : User_UnloadHive
'* Purpose        : unload key from registry
'* Inputs         : root key, key name
'* Outputs        : boolean
'*********************************************
Public Function User_UnloadHive(ByVal lRootKey As Long, _
                                ByVal sKey As String) As Boolean

Dim lReturn             As Long

On Error GoTo Handler

    '/* unload the key
    lReturn = RegUnLoadKey(lRootKey, sKey)
    
    '/* success
    If lReturn = 0 Then
        User_UnloadHive = True
        RaiseEvent eNComplete("The Key " + sKey + " was unloaded successfully.")
    Else
        GoTo Handler
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("User_LoadHive", "An unknown error has occured. Error# " + CStr(Err.LastDllError))

End Function

'* Name           : User_LoadProfile
'* Purpose        : load a users profile
'* Inputs         : user name, user token, profile, path
'* Outputs        : boolean
'*********************************************
Public Function User_LoadProfile(ByVal sUser As String, _
                                 ByVal lToken As Long, _
                                 ByVal lProfile As Long, _
                                 ByVal sPath As String) As Boolean

Dim lReturn             As Long
Dim PI                  As PROFILE_INFORMATION

On Error GoTo Handler

    '/* profile structure
    With PI
        .dwSize = Len(PI)
        .dwFlags = PI_NOUI Or PI_APPLYPOLICY
        .sUserName = sUser
        .lpProfilePath = 0
        .sDefaultPath = 0
        .lpServerName = 0
        .lpPolicyPath = 0
    End With

    '/* load the profile
    lReturn = LoadUserProfile(lToken, PI)
    
    '/* success
    If lReturn = 0 Then
        User_LoadProfile = True
        RaiseEvent eNComplete("The Profile for user: " + sUser + " was loaded successfully.")
    Else
        GoTo Handler
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("User_LoadProfile", "An unknown error has occured. Error# " + CStr(Err.LastDllError))

End Function

'* Name           : User_UnloadProfile
'* Purpose        : unload a users profile
'* Inputs         : token, profile
'* Outputs        : boolean
'*********************************************
Public Function User_UnloadProfile(ByVal lToken As Long, _
                                   ByVal lProfile As Long) As Boolean

Dim lReturn             As Long

On Error GoTo Handler

    '/* unload the profile
    lReturn = UnloadUserProfile(lToken, lProfile)
    
    '/* success
    If lReturn = 0 Then
        User_UnloadProfile = True
        RaiseEvent eNComplete("The Profile was unloaded successfully.")
    Else
        GoTo Handler
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("User_UnloadProfile", "An unknown error has occured. Error# " + CStr(Err.LastDllError))

End Function

'* Name           : User_RunAs
'* Purpose        : start process as alternate user
'* Inputs         : user, pass, domain, command
'* Outputs        : boolean
'*********************************************
Public Function User_RunAs(ByVal sUserName As String, _
                           ByVal sPassword As String, _
                           ByVal sDomain As String, _
                           ByVal sCommand As String) As Boolean

Dim lReturn             As Long
Dim sApplication        As String
Dim sDirectory          As String
Dim tPInfo              As PROCESS_INFORMATION
Dim tStart              As STARTUPINFO

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

Dim sName               As String

    '/* prep var
    sName = String(255, Chr$(0))
    '/* call for name
    GetUserName sName, 255
    '/* trim nulls
    User_Name = Left$(sName, InStr(1, sName, Chr$(0)) - 1)
    
End Function


'>  Group Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'* Name           : Get_Domain
'* Purpose        : get primary dc name
'* Inputs         : domain name
'* Outputs        : string
'*********************************************
Public Function Get_Domain(ByVal sDomainName As String) As String

Dim bDomainName()       As Byte
Dim lPointer            As Long
Dim lReturn             As Long
Dim sTemp               As String

    bDomainName = sDomainName + Chr$(0)
    ' Lookup the Primary Domain Controller
    lReturn = NetGetDCName(0&, bDomainName(0), lPointer)
    '/* success
    If lReturn = 0 Then
        RaiseEvent eNComplete("Query completed successfully.")
    Else
        RaiseEvent eNComplete("The query found no results.")
    End If
    lstrcpyW bDomainName(0), lPointer
    NetApiBufferFree lPointer
    sTemp = bDomainName()
    Get_Domain = Left(sTemp, InStr(sTemp, Chr(0)) - 1)
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Get_Domain", CStr(Err.LastDllError))

End Function

'* Name           : Group_Create
'* Purpose        : create a new group
'* Inputs         : machine name, user name, comment
'* Outputs        : boolean
'*********************************************
Public Function Group_Create(ByVal sMachine As String, _
                             ByVal sGroupName As String, _
                             ByVal sComment As String, _
                             Optional ByVal bRemote As Boolean) As Boolean

Dim bMachine()          As Byte
Dim lReturn             As Long
Dim tGInfo              As GROUP_INFO

On Error GoTo Handler

    '/* fill our structure
    With tGInfo
        .tGroup = StrConv(sGroupName, vbUnicode)
        .tComment = StrConv(sComment, vbUnicode)
    End With
    bMachine = sMachine + Chr$(0)
    '/* local group
    If Not bRemote Then
        lReturn = NetLocalGroupAdd(bMachine(0), 1, tGInfo, 0)
    '/* remote
    Else
        lReturn = NetGroupAdd1(bMachine(0), 1, tGInfo, 0)
    End If
    '/* success
    If lReturn = 0 Then
        RaiseEvent eNComplete("The Group: " + sGroupName + " has been created.")
        Group_Create = True
    '/* group exists
    ElseIf lReturn = 1379 Then
        RaiseEvent eNComplete("The Group: " + sGroupName + " already exists.")
    Else
        GoTo Handler
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Group_Create", "An unknown error has occured. Error# " + CStr(Err.LastDllError))
    
End Function

'* Name           : Group_Add
'* Purpose        : add a user to a group
'* Inputs         : machine name, user name, group name
'* Outputs        : boolean
'*********************************************
Public Function Group_Add(ByVal sMachine As String, _
                          ByVal sUserName As String, _
                          ByVal sGroup As String) As Boolean

Dim bGroup()            As Byte
Dim bMachine()          As Byte
Dim bUser()             As Byte
Dim lReturn             As Long
Dim lUser               As Long
Dim ULocal              As LOCALGROUP_MEMBER

On Error GoTo Handler

    '/* format and assign to byte arrays
    bMachine = sMachine + Chr$(0)
    bGroup = sGroup + Chr$(0)
    bUser = Chr$(92) + sUserName + Chr$(0)
    '/* create buffer
    NetAPIBufferAllocate UBound(bUser) + 1, lUser
    '/* get pointer and fill struct
    StrToPtr lUser, bUser(0)
    ULocal.UNamePtr = lUser
    '/* assign the user
    lReturn = NetLocalGroupAddMembers(bMachine(0), bGroup(0), 3, ULocal, 1)
    '/* success
    If lReturn = 0 Then
        RaiseEvent eNComplete("The User: " + sUserName + " has been added to " + sGroup)
        Group_Add = True
    Else
        GoTo Handler
    End If
    '/* free the buffer
    NetApiBufferFree lUser

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Group_Add", "An unknown error has occured. Error# " + CStr(Err.LastDllError))

End Function

'* Name           : Group_Remove
'* Purpose        : remove a user to a group
'* Inputs         : machine name, user name, group name
'* Outputs        : boolean
'*********************************************
Public Function Group_Remove(ByVal sMachine As String, _
                             ByVal sUserName As String, _
                             ByVal sGroup As String) As Boolean


Dim bGroup()            As Byte
Dim bMachine()          As Byte
Dim bUserName()         As Byte
Dim lReturn             As Long
Dim lUser               As Long
Dim tLocal              As LOCALGROUP_MEMBER

On Error GoTo Handler

    '/* create byte arrays
    bMachine = sMachine + Chr$(0)
    bGroup = sGroup + Chr$(0)
    bUserName = Chr$(92) + sUserName + Chr$(0)
    '/* get buffer
    NetAPIBufferAllocate UBound(bUserName) + 1, lUser
    '/* create pointer
    StrToPtr lUser, bUserName(0)
    tLocal.UNamePtr = lUser
    '/* remove user from group
    lReturn = NetLocalGroupDelMembers(bMachine(0), bGroup(0), 3, tLocal, 1)
    '/* success
    If lReturn = 0 Then
        RaiseEvent eNComplete("The User: " + sUserName + " was removed from " + sGroup)
        Group_Remove = True
    Else
        GoTo Handler
    End If
    '/* release buffer
    NetApiBufferFree lUser

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Group_Remove", "An unknown error has occured. Error# " + CStr(Err.LastDllError))

End Function

'* Name           : Group_Delete
'* Purpose        : delete a group
'* Inputs         : machine name, group name
'* Outputs        : boolean
'*********************************************
Public Function Group_Delete(ByVal sMachine As String, _
                             ByVal sGroup As String, _
                             Optional ByVal bRemote As Boolean) As Boolean

Dim bGroup()            As Byte
Dim bMachine()          As Byte
Dim lReturn             As Long

On Error GoTo Handler

    bGroup = sGroup + Chr$(0)
    bMachine = sMachine + Chr$(0)
    '/* local group
    If Not bRemote Then
        '/* delete group
        lReturn = NetLocalGroupDel(bMachine(0), bGroup(0))
    '/* server
    Else
        lReturn = NetGroupDel(bMachine(0), bGroup(0))
    End If
    '/* success
    If lReturn = 0 Then
        RaiseEvent eNComplete("The Group: " + sGroup + " has been removed.")
        Group_Delete = True
    Else
        GoTo Handler
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Group_Delete", "An unknown error has occured. Error# " + CStr(Err.LastDllError))

End Function

'* Name           : Groups_List
'* Purpose        : list the groups with local accounts
'* Inputs         : none
'* Outputs        : collection
'*********************************************
Public Function Groups_List() As Collection

Dim lAcctLen            As Long
Dim lBuffer             As Long
Dim lCount              As Long
Dim lDomainLen          As Long
Dim lLength             As Long
Dim lProcessID          As Long
Dim lSid                As Long
Dim lToken              As Long
Dim lUse                As Long
Dim sAcctName1          As String
Dim sAcctName2          As String
Dim sDomain             As String
Dim cTemp               As Collection
Dim SIA                 As SID_IDENTIFIER_AUTHORITY
Dim TG                  As TOKEN_GROUPS

On Error GoTo Handler

    Set cTemp = New Collection
    lProcessID = GetCurrentProcess()
    If lProcessID = 0 Then GoTo Handler
    '/* open current process token
    If Not OpenProcessToken(lProcessID, TOKEN_READ, lToken) = 1 Then GoTo Handler
    '/* get our buffer
    If GetTokenInformation(lToken, TOKEN_GROUPS, TG, lLength, lBuffer) = 0 Then
        lLength = lBuffer
        '/* fill the groups struct
        If GetTokenInformation(lToken, TOKEN_GROUPS, TG, lLength, lBuffer) = 1 Then
            SIA.Value(5) = SECURITY_NT_AUTHORITY
            '/* extract the sids
            If AllocateAndInitializeSid(SIA, 2, SECURITY_BUILTIN_DOMAIN_RID, _
                DOMAIN_ALIAS_RID_ADMINS, 0&, 0&, 0&, 0&, 0&, 0&, lSid) = 1 Then
                sAcctName1 = Space$(255)
                sDomain = Space$(255)
                lAcctLen = 255
                lDomainLen = 255
                '/* lookup sid authority
                If LookupAccountSid(vbNullString, lSid, sAcctName1, _
                    lAcctLen, sDomain, lDomainLen, lUse) = 1 Then
                    '/* add to collection
                    For lCount = 0 To TG.GroupCount - 1
                        sAcctName2 = Space$(255)
                        sDomain = Space$(255)
                        lAcctLen = 255
                        lDomainLen = 255
                        LookupAccountSid vbNullString, TG.Groups(lCount).Sid, _
                            sAcctName2, lAcctLen, sDomain, lDomainLen, lUse
                        cTemp.Add Trim$(sAcctName2) + Chr$(31) + CStr(TG.Groups(lCount).Sid) + _
                            Chr$(31) + CStr(lUse)
                    Next lCount
                End If
            End If
        End If
    End If
    
    '/* cleanup
    FreeSid ByVal lSid
    CloseHandle lToken
    CloseHandle lProcessID
    '/* success
    If cTemp.Count > 0 Then
        Set Groups_List = cTemp
        RaiseEvent eNComplete("User Groups enumerated Succesfully.")
    Else
        GoTo Handler
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Group_Remove", "An unknown error has occured. Error# " + CStr(Err.LastDllError))

End Function

'* Name           : Computer_Name
'* Purpose        : get the local computer name
'* Inputs         : none
'* Outputs        : string
'*********************************************
Public Function Computer_Name() As String

Dim sName               As String

    '/* prep var
    sName = Space$(255)
    '/* get name
    GetComputerName sName, 255
    '/* format and return
    Computer_Name = "\\" + Left$(sName, InStr(sName, Chr$(0)) - 1)

End Function


Private Function GetStrFromBufferA(sz As String) As String
'/* string from pointer

    If InStr(sz, Chr$(0)) Then
        GetStrFromBufferA = Left$(sz, InStr(sz, Chr$(0)) - 1)
    Else
        GetStrFromBufferA = sz
    End If

End Function


Private Function GetStrFromPtrW(lpszW As Long) As String
'/* sfp wide

Dim sRtn                As String

    sRtn = String$(lstrlenW(ByVal lpszW) * 2, 0)
    WideCharToMultiByte 0, 0, ByVal lpszW, -1, ByVal sRtn, Len(sRtn), 0, 0
    GetStrFromPtrW = GetStrFromBufferA(sRtn)

End Function


