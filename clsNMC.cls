VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************************
'*  NMC! V1.6    NT Master Class v1.6.1                                                *
'*                                                                                     *
'*  Created:     April 14, 2006                                                        *
'*  Updated:     June 02, 2006                                                         *
'*  Purpose:     NT System Security Master Class                                       *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.6.1                                                                 *
'*  Compile:     Native                                                                *
'*  Referenced:  Throughout Project                                                    *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*                                                                                     *
'***************************************************************************************


' ~*** Notes ***~

'/~ Some things to know..
'/~ First off, some errors are not actually errors, but are caused by invalid input. For example,
'/~ you can not delete a user that does not exist, or create a group that already exists, these
'/~ sorts of things are user input errors, and have to be handled or ignored by your application
'/~ handlers. Some services can not be paused/restarted, that is an attribute of the service design.
'/~ You must have at least one NTFS partition in order to test the NTFS security routines, if NTFS
'/~ is not detected, these controls are disabled. Some things will work in 98/ME, some, [of course],
'/~ will not. Most of the api for service and process management is backwards compatable.
'/~ You may use this code in any way you like, as long as you agree that no warranty or
'/~ liability of any kind is expressed or implied. If you wish you can add a mention to the
'/~ about page of your project/app, ex. Credit to: John Underhill of NSPowertools.com.

'/~ April 17, 2006 - V2
'/~ Gave the code a once over and fixed a couple of things, added some options to NTFS recurse
'/~ and group routines.

'/~ April 22, 2006 - V3
'/~ Added functions Service_Remove and Service_Add. Updated User and group functions, global groups
'/~ can now be added, and extended user properties like default permissions set were added.

'/~ April 27, 2006 - V4
'/~ Fixed a bug in User_Create, that caused the example to work improperly. Added three new routines,
'/~ User_SetData (sets account properties), User_Password (change account password), and Get_Domain
'/~ (retrieve the primary dc name).

'/~ April 29, 2006 - V5  ~Final~ (!60 Functions!)
'/~ Excepting if bugs are found, this is the final rev.
'/~ Added 10 routines to Process management. Process functions are all mirrored to psapi equivilents
'/~ for NT Server compatability. Example was expanded to demonstrate split methods.
'/~ Module_EnumG2 (module enum using psapi for 2k server), Process_EnumG2 (process enum using psapi),
'/~ Return_ProcessID and  Return_ProcessIDG2 (return process id from process name - psapi and kernal32 versions)
'/~ Process_ExistsG2 (test for the existence of a process using psapi)
'/~ Process_GetClass (get current process class), Process_SetClass (set process priority),
'/~ Thread_Terminate (kill a thread), Thread_GetPriority (get thread priority),
'/~ Thread_SetPriority (set new thread priority).

'/~ Added 4 new routines to user management:
'/~ User_LoadProfile (generate a new user profile), User_UnloadProfile (unload user profile)
'/~ User_LoadHive (load users profile hive into the registry), User_UnloadHive (unload the hive).

'/~ Updated the sample program with the "Extended Example" demo, to demonstrate usage of some of
'/~ the new features.

'/~ April 30, 2006
'/~ One small fix to NetLocalGroupDel and NetGroupDel declarations, and updated documentation.

'/ June 02, 2006 - V6
'/~ Added Service_Desc function to change a services description in the mmc. Proofed the Service_Add and
'/~ Service_Remove functions and updated them. Some small formatting changes.

'/ June 03, 2006
'/~ As per Lite's suggestion, added remote capabilities to all the service routines, (simply specify
'/~ UNC path to machine in optional 'sMachine' param).

'/~ Cheers
'/~ John - steppenwolfe_2000@yahoo.com


' ~*** Exposed Functions ***~

'/~ User and Group Management ~/
'/~ User_Create             - create a new user         [in -string (6) | out -bool]
'/~ User_Delete             - delete a user account     [in -string (2) | out -bool]
'/~ User_Exist              - test a user account       [in -string (2) | out -bool]
'/~ User_Data               - get user profile data     [in -string (2) | out -bool]
'/~ User_LoadProfile        - create user profile       [in string (4) | out -bool]
'/~ User_UnloadProfile      - unload profile            [in string | out -bool]
'/~ User_LoadHive           - load user key             [in string | out -bool]
'/~ User_UnloadHive         - unload user key           [in string | out -bool]
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

'/~ Service Management ~/
'/~ Service_Start           - start a service           [in -string | out -bool]
'/~ Service_Stop            - stop a service            [in -string | out -bool]
'/~ Service_Pause           - pause a service           [in -string | out -bool]
'/~ Service_Resume          - resume a service          [in -string | out -bool]
'/~ Service_Query           - query service state       [in -string | out -long]
'/~ Service_Change          - change startup attr       [in -string + enum | out -bool]
'/~ Service_Enumerate       - list running services     [out -col]
'/~ Service_Add             - add a new service         [in -string (6) | out -bool]
'/~ Service_Remove          - delete a service          [in -string | out -bool]
'/~ Service_Desc            - change svc description    [in -string (2) | out -bool]

'/~ Process Management ~/
'/~ Process_Enumerate       - list processes (kernal32) [out -col]
'/~ Process_EnumG2          - list processes (psapi)    [out -col]
'/~ Process_Exists          - test existence (kernal32) [in -string | out -bool]
'/~ Process_ExistsG2        - test existence (psapi)    [in -string | out -bool]
'/~ Process_Terminate       - terminate a process       [in -string | out -bool]
'/~ Process_GetClass        - get process class         [in -long | out -long]
'/~ Process_SetClass        - change process priority   [in - string + enum | out -bool]
'/~ Return_ProcessID        - return prc id (kernal32)  [in -string | out -long]
'/~ Return_ProcessIDG2      - return prc id (psapi)     [in -string | out -long]
'/~ Thread_Enumerate        - list a process threads    [in -string | out -bool]
'/~ Thread_Suspend          - suspend a thread          [in -long + string | out -bool]
'/~ Thread_Resume           - resume a thread           [in -long + string | out -bool]
'/~ Thread_GetPriority      - get thread priority       [in -long | out -long]
'/~ Thread_SetPriority      - set thread priority       [in -long + enum | -out bool]
'/~ Thread_Terminate        - kill a thread             [in -long | out -bool]
'/~ Module_Enumerate        - list a process modules    [in -string | out -col]
'/~ Module_EnumG2           - server enum (psapi)       [in -string | out -col]

'/~ Encrypted File System ~/
'/~ EFS_Status              - file/folder state         [in -string | out -long]
'/~ EFS_Encrypt             - encrypt file/folder       [in -string | out -bool]
'/~ EFS_Decrypt             - decrypt file/folder       [in -string | out -bool]
'/~ EFS_Enable              - enable EFS                [in -string | out -bool]
'/~ EFS_Disable             - disable EFS               [in -string | out -bool]
'/~ Create_Directory        - create a directory        [in -string | out -bool]
'/~ File_Exists             - test for path/file        [in -string | out -bool]

'/~ NTFS File System ~/
'/~ NTFS_Drive              - test drive for ntfs       [in -string(1) | out -bool]
'/~ NTFS_Check              - test all drives for ntfs  [out -col]
'/~ NTFS_Folder             - modify object security    [in -string/enum | out -bool]
'/~ NTFS_Recursive          - recurse permissions set   [in -string | out -bool]

'/~ NTFS Registry ~/
'/~ NTFS_Key                - modify key security       [in -string/enum | out -bool]


' ~*** Events ***~

'/~ eNComplete              - task completed successfully
'/~ eNErrorCond             - error condition



'>  Start User/Group Dependencies
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
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
    lgrmi2_sid                                          As Long
    lgrmi2_sidusage                                     As Long
    lgrmi2_domainandname                                As Long
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
    usri1003_password                                   As Long
End Type

Private Type USER_INFO_1008
    usri1008_flags                                      As Long
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

'/~ example
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

Private Declare Function NetUserAdd1 Lib "NETAPI32.dll" Alias "NetUserAdd" (servername As Byte, _
                                                                            ByVal level As Long, _
                                                                            Buffer As USER_ATTRIBUTES, _
                                                                            lParmError As Long) As Long

Private Declare Function NetUserDel Lib "NETAPI32.dll" (servername As Byte, _
                                                        username As Byte) As Long

Private Declare Function NetUserEnum Lib "NETAPI32.dll" (servername As Byte, _
                                                         ByVal level As Long, _
                                                         ByVal lFilter As Long, _
                                                         Buffer As Long, _
                                                         ByVal prefmaxlen As Long, _
                                                         entriesread As Long, _
                                                         totalentries As Long, _
                                                         ResumeHandle As Long) As Long

Private Declare Function NetGroupAdd1 Lib "NETAPI32.dll" Alias "NetGroupAdd" (servername As Byte, _
                                                                              ByVal level As Long, _
                                                                              Buffer As GROUP_INFO, _
                                                                              ParmError As Long) As Long

Private Declare Function NetLocalGroupAdd Lib "NETAPI32.dll" (servername As Byte, _
                                                              ByVal level As Long, _
                                                              Buffer As GROUP_INFO, _
                                                              parm_err As Long) As Long

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

Private Declare Function NetLocalGroupAddMembers Lib "NETAPI32.dll" (servername As Byte, _
                                                                     groupname As Byte, _
                                                                     ByVal level As Long, _
                                                                     buf As LOCALGROUP_MEMBER, _
                                                                     ByVal totalentries As Long) As Long

Private Declare Function NetLocalGroupDel Lib "NETAPI32.dll" (servername As Byte, _
                                                              groupname As Byte) As Long

Private Declare Function NetGroupDel Lib "NETAPI32.dll" (servername As Byte, _
                                                         groupname As Byte) As Long

Private Declare Function NetLocalGroupDelMembers Lib "NETAPI32.dll" (servername As Byte, _
                                                                     groupname As Byte, _
                                                                     ByVal level As Long, _
                                                                     buf As LOCALGROUP_MEMBER, _
                                                                     ByVal totalentries As Long) As Long

Private Declare Function NetAPIBufferAllocate Lib "NETAPI32.dll" Alias "NetApiBufferAllocate" (ByVal ByteCount As Long, _
                                                                                               Ptr As Long) As Long

Private Declare Function NetApiBufferFree Lib "Netapi32" (ByVal Buffer As Long) As Long

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

Private Declare Sub lstrcpyW Lib "kernel32" (dest As Any, ByVal src As Any)

Private Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, _
                                                                            ByVal lpSubKey As String, _
                                                                            ByVal lpFile As String) As Long
                                                                            
Private Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String) As Long

Private Declare Function LoadUserProfile Lib "userenv.dll" Alias "LoadUserProfileA" (ByVal hToken As Long, _
                                                                                     ByRef lpProfileInfo As PROFILE_INFORMATION) As Boolean
                                                                                     
Private Declare Function UnloadUserProfile Lib "userenv.dll" (ByVal hToken As Long, _
                                                              ByVal hProfile As Long) As Long

'>  Start Services Dependencies
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Const ERROR_MORE_DATA                           As Integer = 234
Private Const SC_MANAGER_ENUMERATE_SERVICE              As Long = &H4
Private Const SERVICE_STATE_ALL                         As Long = &H3
Private Const SERVICE_KERNEL_DRIVER                     As Long = &H1
Private Const SERVICE_FILE_SYSTEM_DRIVER                As Long = &H2
Private Const SERVICE_ADAPTER                           As Long = &H4
Private Const SERVICE_RECOGNIZER_DRIVER                 As Long = &H8
Private Const SERVICE_WIN32_OWN_PROCESS                 As Long = &H10
Private Const SERVICE_WIN32_SHARE_PROCESS               As Long = &H20
Private Const SERVICE_INTERACTIVE_PROCESS               As Long = &H100
Private Const SERVICE_STOPPED                           As Long = &H1
Private Const SERVICE_START_PENDING                     As Long = &H2
Private Const SERVICE_STOP_PENDING                      As Long = &H3
Private Const SERVICE_RUNNING                           As Long = &H4
Private Const SERVICE_CONTINUE_PENDING                  As Long = &H5
Private Const SERVICE_PAUSE_PENDING                     As Long = &H6
Private Const SERVICE_PAUSED                            As Long = &H7
Private Const SC_MANAGER_CONNECT                        As Long = &H1
Private Const SC_MANAGER_CREATE_SERVICE                 As Long = &H2
Private Const SERVICE_BOOT_START                        As Long = &H0
Private Const SERVICE_SYSTEM_START                      As Long = &H1
Private Const SERVICE_AUTO_START                        As Long = &H2
Private Const SERVICE_DEMAND_START                      As Long = &H3
Private Const SERVICE_DISABLED                          As Long = &H4
Private Const SERVICE_ERROR_NORMAL                      As Long = &H1
Private Const SERVICE_NO_CHANGE                         As Long = &HFFFFFFFF
Private Const SERVICE_QUERY_CONFIG                      As Long = &H1
Private Const SERVICE_CHANGE_CONFIG                     As Long = &H2
Private Const SERVICE_QUERY_STATUS                      As Long = &H4
Private Const SERVICE_ENUMERATE_DEPENDENTS              As Long = &H8
Private Const SERVICE_PAUSE_CONTINUE                    As Long = &H40
Private Const SERVICE_INTERROGATE                       As Long = &H80
Private Const SERVICE_USER_DEFINED_CONTROL              As Long = &H100
Private Const SERVICE_CSTART                            As Long = &H10
Private Const SERVICE_CSTOP                             As Long = &H20
Private Const SERVICE_CONTROL_STOP                      As Long = &H1
Private Const SERVICE_CONTROL_PAUSE                     As Long = &H2
Private Const SERVICE_CONTROL_CONTINUE                  As Long = &H3
Private Const SERVICE_CONTROL_INTERROGATE               As Long = &H4
Private Const SERVICE_CONTROL_SHUTDOWN                  As Long = &H5
Private Const SERVICE_CONTROL_PARAMCHANGE               As Long = &H6
Private Const SERVICE_CONTROL_NETBINDADD                As Long = &H7
Private Const SERVICE_CONTROL_NETBINDREMOVE             As Long = &H8
Private Const SERVICE_CONTROL_NETBINDENABLE             As Long = &H9
Private Const SERVICE_CONTROL_NETBINDDISABLE            As Long = &HA
Private Const STANDARD_RIGHTS_REQUIRED                  As Long = &HF0000

Private Const SERVICE_ALL_ACCESS                        As Double = _
        (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or _
        SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or _
        SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_CSTART Or _
        SERVICE_CSTOP Or SERVICE_PAUSE_CONTINUE Or _
        SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)
Private Const SERVICE_WIN32                             As Long = _
        SERVICE_WIN32_OWN_PROCESS Or SERVICE_WIN32_SHARE_PROCESS
Private Const SERVICE_DRIVER                            As Long = _
        SERVICE_KERNEL_DRIVER Or SERVICE_FILE_SYSTEM_DRIVER Or _
        SERVICE_RECOGNIZER_DRIVER
Private Const SERVICE_TYPE_ALL                          As Long = _
        SERVICE_WIN32 Or SERVICE_ADAPTER Or SERVICE_DRIVER Or _
        SERVICE_INTERACTIVE_PROCESS

'/* service status
Private Type SERVICE_STATUS
    dwServiceType                                       As Long
    dwCurrentState                                      As Long
    dwControlsAccepted                                  As Long
    dwWin32ExitCode                                     As Long
    dwServiceSpecificExitCode                           As Long
    dwCheckPoint                                        As Long
    dwWaitHint                                          As Long
End Type

Private Type ENUM_SERVICE_STATUS
   lpServiceName                                        As Long
   lpDisplayName                                        As Long
   ServiceStatus                                        As SERVICE_STATUS
End Type

Private Type SERVICE_DESCRIPTION
    lpDescription                                       As String
End Type

'/* description
Private Enum InfoLevel
    SERVICE_CONFIG_DESCRIPTION = 1&
    SERVICE_CONFIG_FAILURE_ACTIONS = 2&
End Enum

'/* startup type
Public Enum eServiceStartType
    START_BOOT = &H0
    START_SYSTEM = &H1
    START_AUTO = &H2
    START_DEMAND = &H3
    START_DISABLED = &H4
End Enum

'/* service api
Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal strMachineName As String, _
                                                                                  ByVal strDBName As String, _
                                                                                  ByVal lAccessReq As Long) As Long

Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, _
                                                                              ByVal strServiceName As String, _
                                                                              ByVal lAccessReq As Long) As Long

Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, _
                                                                                ByVal lNumServiceArgs As Long, _
                                                                                ByVal strArgs As String) As Boolean

Private Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, _
                                                            ByVal lControlCode As Long, _
                                                            lpServiceStatus As SERVICE_STATUS) As Boolean

Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hHandle As Long) As Boolean

Private Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, _
                                                                lpServiceStatus As SERVICE_STATUS) As Boolean

Private Declare Function ChangeServiceConfig Lib "advapi32.dll" Alias "ChangeServiceConfigA" (ByVal hService As Long, _
                                                                                              ByVal dwServiceType As Long, _
                                                                                              ByVal dwStartType As eServiceStartType, _
                                                                                              ByVal dwErrorControl As Long, _
                                                                                              ByVal lpBinaryPathName As String, _
                                                                                              ByVal lpLoadOrderGroup As String, _
                                                                                              ByVal lpdwTagId As Long, _
                                                                                              ByVal lpDependencies As String, _
                                                                                              ByVal lpServiceStartName As String, _
                                                                                              ByVal lpPassword As String, _
                                                                                              ByVal lpDisplayName As String) As Boolean

Private Declare Function EnumServicesStatus Lib "advapi32" Alias "EnumServicesStatusA" (ByVal hSCManager As Long, _
                                                                                        ByVal dwServiceType As Long, _
                                                                                        ByVal dwServiceState As Long, _
                                                                                        lpServices As Any, _
                                                                                        ByVal cbBufSize As Long, _
                                                                                        pcbBytesNeeded As Long, _
                                                                                        lpServicesReturned As Long, _
                                                                                        lpResumeHandle As Long) As Long

Private Declare Function CreateService Lib "advapi32" Alias "CreateServiceA" (ByVal hSCManager As Long, _
                                                                              ByVal lpServiceName As String, _
                                                                              ByVal lpDisplayName As String, _
                                                                              ByVal dwDesiredAccess As Long, _
                                                                              ByVal dwServiceType As Long, _
                                                                              ByVal dwStartType As Long, _
                                                                              ByVal dwErrorControl As Long, _
                                                                              ByVal lpBinaryPathName As String, _
                                                                              ByVal lpLoadOrderGroup As String, _
                                                                              ByVal lpdwTagId As String, _
                                                                              ByVal lpDependencies As String, _
                                                                              ByVal lp As String, _
                                                                              ByVal lpPassword As String) As Long

Private Declare Function DeleteService Lib "advapi32" (ByVal hService As Long) As Long

Private Declare Function ChangeServiceConfig2 Lib "advapi32.dll" Alias "ChangeServiceConfig2W" (ByVal hService As Long, _
                                                                                                ByVal dwInfoLevel As InfoLevel, _
                                                                                                lpInfo As Any) As Boolean

Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, _
                                                  ByVal Ptr As Long) As Long

Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long


'>  Start Process Dependencies
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Const PROCESS_VM_READ                           As Integer = 16
Private Const PROCESS_SET_INFORMATION                   As Long = &H200
Private Const PROCESS_QUERY_INFORMATION                 As Long = &H400
Private Const MAX_PATH                                  As Long = 260
Private Const TH32CS_SNAPTHREAD                         As Long = &H4
Private Const TH32CS_SNAPHEAPLIST                       As Long = &H1
Private Const TH32CS_SNAPPROCESS                        As Long = &H2
Private Const TH32CS_SNAPMODULE                         As Long = &H8
Private Const PROCESS_CTERMINATE                        As Long = &H1
Private Const THREAD_SUSPEND_RESUME                     As Long = &H2
Private Const THREAD_SET_INFORMATION                    As Long = &H20
Private Const THREAD_QUERY_INFORMATION                  As Long = &H40
Private Const TH32CS_SNAPALL                            As Double = _
        (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or _
        TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)

Private Type PROCESSENTRY32
    dwSize                                              As Long
    cntUsage                                            As Long
    th32ProcessID                                       As Long
    th32DefaultHeapID                                   As Long
    th32ModuleID                                        As Long
    cntThreads                                          As Long
    th32ParentProcessID                                 As Long
    pcPriClassBase                                      As Long
    dwFlags                                             As Long
    szexeFile                                           As String * MAX_PATH
End Type

Private Type MODULEENTRY32
    dwSize                                              As Long
    th32ModuleID                                        As Long
    th32ProcessID                                       As Long
    GlblcntUsage                                        As Long
    ProccntUsage                                        As Long
    modBaseAddr                                         As Long
    modBaseSize                                         As Long
    hModule                                             As Long
    szModule                                            As String * 256
    szExePath                                           As String * MAX_PATH
End Type

Private Type THREADENTRY32
    dwSize                                              As Long
    cntUsage                                            As Long
    th32ThreadID                                        As Long
    th32OwnerProcessID                                  As Long
    tpBasePri                                           As Long
    tpDeltaPri                                          As Long
    dwFlags                                             As Long
End Type

Public Enum eThreadPriority
    Thread_Idle = -15
    Thread_LowRT = 15
    Thread_Minimum = -2
    Thread_Normal = 0
    Thread_Maximum = 2
End Enum
    
Public Enum eProcessClass
    Process_Idle = &H40
    Process_Normal = &H20
    Process_Above = &H8000&
    Process_High = &H80
    Process_RealTime = &H100
End Enum

'/* psapi g2
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, _
                                                        ByVal lBuffer As Long, _
                                                        ByRef lAllocate As Long) As Long

Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal lProcess As Long, _
                                                               ByVal hModule As Long, _
                                                               ByVal ModuleName As String, _
                                                               ByVal lSize As Long) As Long

Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, _
                                                                                           ByVal hModule As Long, _
                                                                                           ByVal ModuleName As String, _
                                                                                           ByVal nSize As Long) As Long

Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal lProcess As Long, _
                                                             ByRef lphModule As Long, _
                                                             ByVal lBuffer As Long, _
                                                             ByRef lAllocate As Long) As Long

'/* process api
Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, _
                                                         ByVal dwExitCode As Long) As Long

Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long

Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, _
                                                           ByVal nPriority As Long) As Long

Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long

Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, _
                                                          ByVal dwPriorityClass As Long) As Long

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, _
                                                                  ByVal lProcessID As Long) As Long

Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, _
                                                        uProcess As PROCESSENTRY32) As Long

Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, _
                                                       uProcess As PROCESSENTRY32) As Long

Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, _
                                                          ByVal uExitCode As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                                                     ByVal blnheritHandle As Long, _
                                                     ByVal dwAppProcessId As Long) As Long

Private Declare Function Module32First Lib "kernel32" (ByVal hSnapShot As Long, _
                                                       uProcess As MODULEENTRY32) As Long

Private Declare Function Module32Next Lib "Kernel32.dll" (ByVal hSnapShot As Long, _
                                                          ByRef lpme As MODULEENTRY32) As Boolean

Private Declare Function ResumeThread Lib "Kernel32.dll" (ByVal hThread As Long) As Long

Private Declare Function SuspendThread Lib "Kernel32.dll" (ByVal hThread As Long) As Long

Private Declare Function GetExitCodeProcess Lib "Kernel32.dll" (ByVal hProcess As Long, _
                                                                ByRef lpExitCode As Long) As Boolean

Private Declare Function Thread32First Lib "Kernel32.dll" (ByVal hSnapShot As Long, _
                                                           ByRef lpte As THREADENTRY32) As Boolean

Private Declare Function Thread32Next Lib "Kernel32.dll" (ByVal hSnapShot As Long, _
                                                          ByRef lpte As THREADENTRY32) As Boolean

Private Declare Function OpenThread Lib "Kernel32.dll" (ByVal dwDesiredAccess As Long, _
                                                        ByVal bInheritHandle As Boolean, _
                                                        ByVal dwThreadId As Long) As Long


'>  Start NTFS Dependencies
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'/* memory constants
Private Const LMEM_FIXED                                As Long = &H0
Private Const LMEM_ZEROINIT                             As Long = &H40
Private Const lPtr                                      As Double = (LMEM_FIXED + LMEM_ZEROINIT)
'/* access rights
Private Const GENERIC_READ                              As Long = &H80000000
Private Const DELETE                                    As Long = &H10000
Private Const WRITE_DAC                                 As Long = &H40000
Private Const STANDARD_RIGHTS_EXECUTE                   As Long = READ_CONTROL
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
    lLD                                                 As Long
    lHD                                                 As Long
End Type

'/* file find struct
Private Type WFD
    lFA                                                 As Long
    fCT                                                 As FT
    fLA                                                 As FT
    fLWT                                                As FT
    lFSH                                                As Long
    lFSL                                                As Long
    lR0                                                 As Long
    lR1                                                 As Long
    cFN                                                 As String * 260
    cAL                                                 As String * 14
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
Public Enum eAccessMask
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
Public Enum eFolderPermissions
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
Public Enum eInheritenceFlags
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
Public Enum eAccessType
    '/* permissive dacl
    Access_Allowed
    '/* deny dacl
    Access_Denied
End Enum

'/~ access flags
Public Enum eRegistryAccess
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

Private Declare Function LocalAlloc Lib "Kernel32.dll" (ByVal wFlags As Long, _
                                                        ByVal wBytes As Long) As Long

Private Declare Function LocalFree Lib "Kernel32.dll" (ByVal hMem As Long) As Long

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

Private Declare Function CreateFile Lib "Kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, _
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

Private Declare Function GetVolumeInformation Lib "Kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
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


'>  Start EFS Dependencies
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Const FO_COPY                                   As Long = &H2
Private Const FO_DELETE                                 As Long = &H3
Private Const FOF_NOCONFIRMATION                        As Long = &H10
Private Const FOF_SILENT                                As Long = &H4

Private Const FILE_ENCRYPTABLE                          As Long = &H0
Private Const FILE_IS_ENCRYPTED                         As Long = &H1
Private Const FILE_SYSTEM_ATTR                          As Long = &H2
Private Const FILE_ROOT_DIR                             As Long = &H3
Private Const FILE_SYSTEM_DIR                           As Long = &H4
Private Const FILE_UNKNOWN                              As Long = &H5
Private Const FILE_SYSTEM_NOT_SUPPORT                   As Long = &H6
Private Const FILE_READ_ONLY                            As Long = &H8
Private Const FILE_DIR_DISALLOWED                       As Long = &H9

Private Type SHFILEOPSTRUCT
    hWnd                                                As Long
    wFunc                                               As Long
    pFrom                                               As String
    pTo                                                 As String
    fFlags                                              As Integer
    fAborted                                            As Long
    hNameMaps                                           As Long
    sProgress                                           As String
End Type

Private Declare Function EncryptFile Lib "advapi32.dll" Alias "EncryptFileA" (ByVal lpFileName As String) As Long

Private Declare Function DecryptFile Lib "advapi32.dll" Alias "DecryptFileA" (ByVal lpFileName As String, _
                                                                              ByVal dwReserved As Long) As Long

Private Declare Function FileEncryptionStatus Lib "advapi32.dll" Alias "FileEncryptionStatusA" (ByVal lpFileName As String, _
                                                                                                lpStatus As Long) As Long

Private Declare Function EncryptionDisable Lib "advapi32.dll" (ByVal DirPath As String, _
                                                               ByVal Disable As Long) As Long

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (ByRef lpFileOp As SHFILEOPSTRUCT) As Long

Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lLongPath As String, _
                                                                                    ByVal lShortPath As String, _
                                                                                    ByVal lBuffer As Long) As Long

'>  Global Events
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Public Event eNComplete(ByVal sTask As String)
Public Event eNErrorCond(ByVal sRoutine As String, ByVal sError As String)

'>  Local Variables
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private m_lToken        As Long
Private m_Next          As String
Private c_Storage       As Collection
Private W32             As WFD


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
    bMachine = sMachine + Chr$(0)
    bUserName = sUserName + Chr$(0)
    bPassword = sPassword + Chr$(0)
    bDirectory = sDirectory + Chr$(0)
    bComment = sComment + Chr$(0)
    bScript = sScript + Chr$(0)
    
    '/* create buffers
    NetAPIBufferAllocate UBound(bUserName), lName
    NetAPIBufferAllocate UBound(bPassword), lPass
    NetAPIBufferAllocate UBound(bDirectory), lDir
    NetAPIBufferAllocate UBound(bComment), lCmt
    NetAPIBufferAllocate UBound(bScript), lScript
    
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
    '/* success
    Case 0
        User_Create = True
        RaiseEvent eNComplete("The User account was Created successfully")
    '/* user exists
    Case 2224
        RaiseEvent eNErrorCond("User_Create", "User already exists. Error #2224")
    '/* error
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
        User_Delete = True
        RaiseEvent eNComplete("The User account was Deleted successfully")
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
        User_Exist = True
        RaiseEvent eNComplete("The User account: " + sUserName + " Exists.")
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
        User_Password = True
        RaiseEvent eNComplete("The User account: " + sUserName + " password has been changed.")
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
        User_SetData = True
        RaiseEvent eNComplete("The User account: " + sUserName + " has been changed.")
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
        If Not sGroup = vbNullString Then
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
        Set Users_List = cTemp
        RaiseEvent eNComplete("User list has been recovered.")
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
    If Not LogonUser(sUser, sDomain, sPass, LOGON32_LOGON_INTERACTIVE, _
        LOGON32_PROVIDER_DEFAULT, m_lToken) = 0 Then
        '/* test account token
        If Not ImpersonateLoggedOnUser(m_lToken) = 0 Then
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
    lReturn = LoadUserProfile(m_lToken, PI)
    
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
    sApplication = vbNullString
    sDirectory = vbNullString
    tStart.cb = LenB(tStart)
    tStart.dwFlags = 0&
    
    '/* launch process
    lReturn = CreateProcessWithLogonW(StrPtr(sUserName), StrPtr(sDomain), _
        StrPtr(sPassword), &H1, StrPtr(sApplication), StrPtr(sCommand), _
        DEFAULT_LOGON, ByVal 0&, StrPtr(sDirectory), tStart, tPInfo)
        '/* success
        If Not lReturn = 0 Then
            User_RunAs = True
            RaiseEvent eNComplete("The Application: " + sCommand + " launched successfully.")
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
        Group_Create = True
        RaiseEvent eNComplete("The Group: " + sGroupName + " has been created.")
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
        Group_Add = True
        RaiseEvent eNComplete("The User: " + sUserName + " has been added to " + sGroup)
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
        Group_Remove = True
        RaiseEvent eNComplete("The User: " + sUserName + " was removed from " + sGroup)
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
        Group_Delete = True
        RaiseEvent eNComplete("The Group: " + sGroup + " has been removed.")
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


'***************************************************************************************
'*                                    SERVICE MANAGEMENT
'***************************************************************************************

'* Name           : Service_Start
'* Purpose        : start a service
'* Inputs         : service name
'* Outputs        : boolean
'*********************************************
Public Function Service_Start(sService As String, _
                              Optional ByVal sMachine As String) As Boolean

Dim lHManager           As Long
Dim lHService           As Long

On Error GoTo Handler

    If Len(sMachine) = 0 Then
        sMachine = vbNullString
    End If
    '/* get handle to service manager
    lHManager = OpenSCManager(sMachine, vbNullString, SC_MANAGER_CONNECT)
    '/* get service handle
    lHService = OpenService(lHManager, sService, SERVICE_ALL_ACCESS)
    '/* start service
    Service_Start = StartService(lHService, 0&, 0&)
    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager
    '/* notify and exit
    If Service_Start Then
        RaiseEvent eNComplete("Service has Started")
    Else
        RaiseEvent eNComplete("Service could Not be Started")
    End If
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Service_Start", CStr(Err.LastDllError))
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

'* Name           : Service_Stop
'* Purpose        : stop a service
'* Inputs         : service name
'* Outputs        : boolean
'*********************************************
Public Function Service_Stop(sService As String, _
                             Optional ByVal sMachine As String) As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim tStatus             As SERVICE_STATUS

On Error GoTo Handler

    If Len(sMachine) = 0 Then
        sMachine = vbNullString
    End If
    '/* get app and service handles
    lHManager = OpenSCManager(sMachine, vbNullString, SC_MANAGER_CONNECT)
    lHService = OpenService(lHManager, sService, SERVICE_ALL_ACCESS)
    '/* stop service
    Service_Stop = ControlService(lHService, SERVICE_CONTROL_STOP, tStatus)
    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager
    '/* notify and exit
    RaiseEvent eNComplete("Service has Stopped")
    
On Error GoTo 0
Exit Function
    
Handler:
    RaiseEvent eNErrorCond("Service_Stop", CStr(Err.LastDllError))
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

'* Name           : Service_Pause
'* Purpose        : pause a service
'* Inputs         : service name
'* Outputs        : boolean
'*********************************************
Public Function Service_Pause(sService As String, _
                              Optional ByVal sMachine As String) As Boolean

Dim Resume_Service      As Boolean
Dim lHManager           As Long
Dim lHService           As Long
Dim tStatus             As SERVICE_STATUS

On Error GoTo Handler

    If Len(sMachine) = 0 Then
        sMachine = vbNullString
    End If
    '/* get app and service handles
    lHManager = OpenSCManager(sMachine, vbNullString, SC_MANAGER_CONNECT)
    lHService = OpenService(lHManager, sService, SERVICE_ALL_ACCESS)
    '/* pause service
    Service_Pause = ControlService(lHService, SERVICE_CONTROL_PAUSE, tStatus)
    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager
    '/* notify and exit
    RaiseEvent eNComplete("Service is Paused")
    
On Error GoTo 0
Exit Function
    
Handler:
    RaiseEvent eNErrorCond("Service_Pause", CStr(Err.LastDllError))
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

'* Name           : Service_Resume
'* Purpose        : restart a service
'* Inputs         : service name
'* Outputs        : boolean
'*********************************************
Public Function Service_Resume(sService As String, _
                               Optional ByVal sMachine As String) As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim tStatus             As SERVICE_STATUS

On Error GoTo Handler

    If Len(sMachine) = 0 Then
        sMachine = vbNullString
    End If
    '/* get app and service handles
    lHManager = OpenSCManager(sMachine, vbNullString, SC_MANAGER_CONNECT)
    lHService = OpenService(lHManager, sService, SERVICE_ALL_ACCESS)
    '/* stop service
    Service_Resume = ControlService(lHService, SERVICE_CONTROL_CONTINUE, tStatus)
    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager
    '/* notify and exit
    RaiseEvent eNComplete("Service has Resumed")
    
On Error GoTo 0
Exit Function
    
Handler:
    RaiseEvent eNErrorCond("Service_Resume", CStr(Err.LastDllError))
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

'* Name           : Service_Add
'* Purpose        : install a service
'* Inputs         : name, display name, path, account, password, domain
'* Outputs        : boolean
'*********************************************
Public Function Service_Add(ByVal sServiceName As String, _
                            ByVal sDisplayName As String, _
                            ByVal sServicePath As String, _
                            ByVal sAccountName As String, _
                            ByVal sPassword As String, _
                            Optional ByVal sDomain As String) As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim lResult             As Long
Dim tStatus             As SERVICE_STATUS

On Error GoTo Handler

    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CREATE_SERVICE)
    If Not Len(sDomain) = 0 Then
        sAccountName = sDomain + Chr$(92) + sAccountName
    End If

    lResult = CreateService(lHManager, sServiceName, sDisplayName, SERVICE_ALL_ACCESS, _
        SERVICE_WIN32_OWN_PROCESS, SERVICE_DEMAND_START, SERVICE_ERROR_NORMAL, _
        sServicePath, vbNullString, vbNullString, vbNullString, sAccountName, sPassword)
    
    '/* success
    If Not lResult = 0 Then
        Service_Add = True
        RaiseEvent eNComplete("Service has been added")
    Else
        GoTo Handler
    End If
    
    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager
    
On Error GoTo 0
Exit Function
    
Handler:
    RaiseEvent eNErrorCond("Service_Add", CStr(Err.LastDllError))
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

'* Name           : Service_Remove
'* Purpose        : remove a service
'* Inputs         : service name
'* Outputs        : boolean
'*********************************************
Public Function Service_Remove(sService As String, _
                               Optional ByVal sMachine As String) As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim lResult             As Long
Dim tStatus             As SERVICE_STATUS

On Error GoTo Handler

    If Len(sMachine) = 0 Then
        sMachine = vbNullString
    End If
    
    '/* get app and service handles
    lHManager = OpenSCManager(sMachine, vbNullString, SC_MANAGER_CONNECT)
    lHService = OpenService(lHManager, sService, SERVICE_ALL_ACCESS)

    '/* stop the service
    ControlService lHService, SERVICE_CONTROL_STOP, tStatus
    '/* delete service
    lResult = DeleteService(lHService)
    
    '/* success
    If Not lResult = 0 Then
        Service_Remove = True
        RaiseEvent eNComplete("Service has been removed")
    Else
        GoTo Handler
    End If
    
    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager
    
On Error GoTo 0
Exit Function
    
Handler:
    RaiseEvent eNErrorCond("Service_Remove", CStr(Err.LastDllError))
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

'* Name           : Service_Desc
'* Purpose        : change service description
'* Inputs         : service name, description
'* Outputs        : boolean
'*********************************************
Public Function Service_Desc(ByVal sService As String, _
                             ByVal sDescription As String, _
                             Optional ByVal sMachine As String) As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim lResult             As Long
Dim tDesc               As SERVICE_DESCRIPTION

On Error GoTo Handler

    If Len(sMachine) = 0 Then
        sMachine = vbNullString
    End If
    
    '/* get app and service handles
    lHManager = OpenSCManager(sMachine, vbNullString, SC_MANAGER_CONNECT)
    lHService = OpenService(lHManager, sService, SERVICE_ALL_ACCESS)
    sDescription = StrConv(sDescription, vbUnicode) + Chr$(0)
    tDesc.lpDescription = sDescription
    lResult = ChangeServiceConfig2(lHService, SERVICE_CONFIG_DESCRIPTION, tDesc)
    
    '/* success
    If Not lResult = 0 Then
        Service_Desc = True
        RaiseEvent eNComplete("Service description has been changed")
    Else
        GoTo Handler
    End If

    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager
    
On Error GoTo 0
Exit Function
    
Handler:
    RaiseEvent eNErrorCond("Service_Desc", CStr(Err.LastDllError))
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

'* Name           : Service_Query
'* Purpose        : get service running state
'* Inputs         : service name
'* Outputs        : long
'*********************************************
Public Function Service_Query(sService As String, _
                              Optional ByVal sMachine As String) As Long

Dim lHManager           As Long
Dim lHService           As Long
Dim sState              As String
Dim tStatus             As SERVICE_STATUS

On Error GoTo Handler

    If Len(sMachine) = 0 Then
        sMachine = vbNullString
    End If
    
    '/* get app and service handles
    lHManager = OpenSCManager(sMachine, vbNullString, SC_MANAGER_CONNECT)
    lHService = OpenService(lHManager, sService, SERVICE_QUERY_STATUS)
    '/* query service status
    QueryServiceStatus lHService, tStatus
    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager
    
    '/* return service state
    '/* 0 - not exist, 1 - stopped, 2 - paused, 3 - waiting, 4 - running
    '/* 5 - shutting down 6 - param change 7 - binding
    Service_Query = tStatus.dwCurrentState
    
    '/* return status event
    Select Case tStatus.dwCurrentState
    Case 0:         sState = "Service does not exist"
    Case 1:         sState = "Service is Stopped"
    Case 2:         sState = "Service is Paused"
    Case 3:         sState = "Service is Continuing"
    Case 4:         sState = "Service is Running"
    Case 5:         sState = "Service is Shutting Down"
    Case 6:         sState = "Service Paramater Change"
    Case 7:         sState = "Service is Binding"
    Case Else:      sState = "Service is in State Change"
    End Select
    
    '/* notify and exit
    RaiseEvent eNComplete(sState)
    
On Error GoTo 0
Exit Function
    
Handler:
    RaiseEvent eNErrorCond("Service_Query", CStr(Err.LastDllError))
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

'* Name           : Service_Change
'* Purpose        : change startup type
'* Inputs         : service name, start_type
'* Outputs        : long
'*********************************************
Public Function Service_Change(ByVal sService As String, _
                               ByRef svcStartType As eServiceStartType, _
                               Optional ByVal sMachine As String) As Boolean

Dim lHManager           As Long
Dim lHService           As Long

On Error GoTo Handler

    If Len(sMachine) = 0 Then
        sMachine = vbNullString
    End If
    '/* get app and service handles
    lHManager = OpenSCManager(sMachine, vbNullString, SC_MANAGER_CONNECT)
    lHService = OpenService(lHManager, sService, SERVICE_CHANGE_CONFIG)
    '/* change startup type
    Service_Change = ChangeServiceConfig(lHService, SERVICE_NO_CHANGE, _
    svcStartType, SERVICE_NO_CHANGE, vbNullString, vbNullString, 0&, _
    vbNullString, vbNullString, vbNullString, vbNullString)
    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager
    '/* notify and exit
    RaiseEvent eNComplete("Service startup type has been changed")
    
On Error GoTo 0
Exit Function
    
Handler:
    RaiseEvent eNErrorCond("Service_Change", CStr(Err.LastDllError))
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

'* Name           : Service_Enumerate
'* Purpose        : return list of services
'* Inputs         : none
'* Outputs        : collection
'*********************************************
Public Function Service_Enumerate(Optional ByVal sMachine As String) As Collection

Dim lAlloc              As Long
Dim lBuffer             As Long
Dim lCount              As Long
Dim lHandle             As Long
Dim lMngrHndl           As Long
Dim lResult             As Long
Dim lResume             As Long
Dim lReturn             As Long
Dim lStatus             As Long
Dim sSvcName            As String
Dim sDispName           As String
Dim sStatus             As String
Dim cTemp               As Collection
Dim eSPtr()             As ENUM_SERVICE_STATUS

On Error GoTo Handler

    If Len(sMachine) = 0 Then
        sMachine = vbNullString
    End If
    
    '/* open service manager
    lMngrHndl = OpenSCManager(sMachine, vbNullString, SC_MANAGER_ENUMERATE_SERVICE)
    If lMngrHndl = 0 Then GoTo Handler
    
    '/* get buffer size
    lResult = EnumServicesStatus(lMngrHndl, SERVICE_WIN32, SERVICE_STATE_ALL, _
    ByVal 0&, 0&, lReturn, lHandle, lResume)
    If Not lResult = 0 Then
        If Err.LastDllError = ERROR_MORE_DATA Then
            GoTo Handler
        End If
    End If
    
    '/* set up our array
    lAlloc = (lReturn \ 36) + 1
    ReDim eSPtr(0 To lAlloc) As ENUM_SERVICE_STATUS
    lBuffer = lAlloc * 36
    lResume = 0
    
    '/* pass data to buffer
    If EnumServicesStatus(lMngrHndl, SERVICE_WIN32, SERVICE_STATE_ALL, _
        eSPtr(0), lBuffer, lReturn, lHandle, lResume) = 0 Then
        GoTo Handler
    End If
    
    '/* write out results
    Set cTemp = New Collection
    For lCount = 0 To lHandle - 1
        With eSPtr(lCount)
            sDispName = GetStrFromPtrA(ByVal .lpDisplayName)
            sSvcName = GetStrFromPtrA(ByVal .lpServiceName)
            lStatus = .ServiceStatus.dwCurrentState
        End With
        
        Select Case lStatus
        Case SERVICE_STOPPED:               sStatus = "Stopped"
        Case SERVICE_START_PENDING:         sStatus = "Start Pending"
        Case SERVICE_STOP_PENDING:          sStatus = "Stop Pending"
        Case SERVICE_RUNNING:               sStatus = "Running"
        Case SERVICE_CONTINUE_PENDING:      sStatus = "Resume Pending"
        Case SERVICE_PAUSE_PENDING:         sStatus = "Pause Pending"
        Case SERVICE_PAUSED:                sStatus = "Paused"
        Case Else:                          sStatus = "Unknown"
        End Select
        cTemp.Add sDispName + Chr$(31) + sSvcName + Chr$(31) + sStatus
    Next lCount
    
    '/* cleanup
    CloseServiceHandle lMngrHndl
    
    '/* return collection
    If cTemp.Count > 0 Then
        Set Service_Enumerate = cTemp
        RaiseEvent eNComplete("Services have been enumerated.. Found: " + CStr(lCount) + " items")
    End If
    
On Error GoTo 0
Exit Function
    
Handler:
    RaiseEvent eNErrorCond("Service_Enumerate", CStr(Err.LastDllError))
    If Not lMngrHndl = 0 Then CloseServiceHandle lMngrHndl

End Function

Private Function GetStrFromPtrA(ByVal lpszA As Long) As String
'/* str fpm pnt
    GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0&)
    lstrcpyA ByVal GetStrFromPtrA, ByVal lpszA

End Function


'***************************************************************************************
'*                                  PROCESS MANAGEMENT
'***************************************************************************************

'>  Process Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'* Name           : Process_Enumerate
'* Purpose        : list running processes
'* Inputs         : none
'* Outputs        : collection
'*********************************************
Public Function Process_Enumerate() As Collection

Dim lEntry              As Long
Dim lSnapShot           As Long
Dim sItem               As String
Dim cTemp               As Collection
Dim tProc               As PROCESSENTRY32

On Error GoTo Handler

    Set cTemp = New Collection
    With tProc
        '/* create our snapshot
        lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
        .dwSize = Len(tProc)
        '/* first entry
        lEntry = Process32First(lSnapShot, tProc)
        '/* loop through entries
        Do While lEntry
            lEntry = Process32Next(lSnapShot, tProc)
            If lEntry = 0 Then Exit Do
            '/* name/id/thrdcnt/priority/parent
            sItem = Left$(.szexeFile, InStr(1, .szexeFile, Chr$(0)) - 1) + Chr$(31) + _
            CStr(.th32ProcessID) + Chr$(31) + _
            CStr(.cntThreads) + Chr$(31) + _
            CStr(.pcPriClassBase) + Chr$(31) + _
            CStr(.th32ParentProcessID)
            cTemp.Add sItem
        Loop
        CloseHandle lSnapShot
    End With
    
    '/* success
    If cTemp.Count > 0 Then
        Set Process_Enumerate = cTemp
        RaiseEvent eNComplete("Processes have been enumerated.. Found: " + CStr(cTemp.Count) + " items")
    End If
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Process_Enumerate", CStr(Err.LastDllError))
    If Not lSnapShot = 0 Then CloseHandle lSnapShot

End Function

'* Name           : Process_EnumG2
'* Purpose        : list running processes (psapi)
'* Inputs         : none
'* Outputs        : collection
'*********************************************
Public Function Process_EnumG2() As Collection

Dim lBuffer            As Long
Dim lAllocate          As Long
Dim lItems             As Long
Dim aProcessID()       As Long
Dim lDimension         As Long
Dim aModule(1 To 1024) As Long
Dim lReturn            As Long
Dim lSize              As Long
Dim lHandle            As Long
Dim lCount             As Long
Dim sProcess           As String
Dim cTemp              As Collection

On Error GoTo Handler

    Set cTemp = New Collection
    '/* size our array
    lBuffer = 8
    lAllocate = 96
    Do While lBuffer <= lAllocate
        lBuffer = lBuffer * 2
        ReDim aProcessID(lBuffer / 4) As Long
        lReturn = EnumProcesses(aProcessID(1), lBuffer, lAllocate)
    Loop
    
    '/* loop through return
    lItems = lAllocate / 4
    For lCount = 1 To lItems
        '/* open process
        lHandle = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, aProcessID(lCount))
        '/* invalid
        If lHandle = 0 Then GoTo skip
        
        '/* size module array
        lReturn = EnumProcessModules(lHandle, aModule(1), 1024, lDimension)
        lReturn = EnumProcessModules(lHandle, aModule(1), lDimension, lDimension)
        '/* get the file name
        If Not lReturn = 0 Then
            sProcess = Space$(MAX_PATH)
            lSize = 500
            'Get the module file name
            lReturn = GetModuleFileNameExA(lHandle, aModule(1), sProcess, lSize)
            sProcess = Left$(sProcess, lReturn)
            '/* add to collection name/id/path
            cTemp.Add Mid$(sProcess, InStrRev(sProcess, Chr$(92)) + 1) + Chr$(31) + _
                CStr(aProcessID(lCount)) + Chr$(31) + _
                sProcess
        End If
skip:
        '/* close handle
        CloseHandle lHandle
    Next lCount
    
    '/* success
    If cTemp.Count > 0 Then
        Set Process_EnumG2 = cTemp
        RaiseEvent eNComplete("Processes have been enumerated.. Found: " + CStr(cTemp.Count) + " items")
    End If

On Error GoTo 0
Exit Function

Handler:
    If Not lHandle = 0 Then CloseHandle lHandle
    RaiseEvent eNErrorCond("Process_EnumG2", CStr(Err.LastDllError))

End Function

'* Name           : Process_Exists
'* Purpose        : test process existence
'* Inputs         : process name, use loose naming
'* Outputs        : boolean
'*********************************************
Public Function Process_Exists(ByVal sProcess As String, _
                               Optional ByVal bLoose As Boolean) As Boolean

Dim lEntry              As Long
Dim lSnapShot           As Long
Dim sItem               As String
Dim tProc               As PROCESSENTRY32

On Error GoTo Handler

    With tProc
        '/* create our snapshot
        lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
        .dwSize = Len(tProc)
        '/* first entry
        lEntry = Process32First(lSnapShot, tProc)
        '/* loop through entries
        Do While lEntry
            lEntry = Process32Next(lSnapShot, tProc)
            If lEntry = 0 Then Exit Do
            '/* compare name against process list
            sItem = Left$(.szexeFile, InStr(1, .szexeFile, Chr$(0)) - 1)
            If Not bLoose Then
                If sProcess = sItem Then
                    '/* success
                    Process_Exists = True
                    RaiseEvent eNComplete("Process: " + sProcess + " Exists.")
                    Exit Do
                End If
            Else
                '/* success
                If LCase$(sProcess) = LCase$(sItem) Then
                    Process_Exists = True
                    RaiseEvent eNComplete("Process: " + sProcess + " Exists.")
                    Exit Do
                End If
            End If
        Loop
        CloseHandle lSnapShot
    End With
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Process_Exists", CStr(Err.LastDllError))
    If Not lSnapShot = 0 Then CloseHandle lSnapShot

End Function


'* Name           : Process_ExistsG2
'* Purpose        : test process existence (psapi)
'* Inputs         : process name, use loose naming
'* Outputs        : boolean
'*********************************************
Public Function Process_ExistsG2(ByVal sProcess As String, _
                                 Optional ByVal bLoose As Boolean) As Boolean
'/* return a process id from a process name

Dim lBuffer             As Long
Dim lAllocate           As Long
Dim lItems              As Long
Dim aProcessID()        As Long
Dim lDimension          As Long
Dim aModule(1 To 1024)  As Long
Dim lReturn             As Long
Dim lSize               As Long
Dim lHandle             As Long
Dim lCount              As Long
Dim sTemp               As String

On Error GoTo Handler

    '/* size our array
    lBuffer = 8
    lAllocate = 96
    Do While lBuffer <= lAllocate
        lBuffer = lBuffer * 2
        ReDim aProcessID(lBuffer / 4) As Long
        lReturn = EnumProcesses(aProcessID(1), lBuffer, lAllocate)
    Loop
    
    '/* loop through return
    lItems = lAllocate / 4
    For lCount = 1 To lItems
        '/* open process
        lHandle = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, aProcessID(lCount))
        '/* invalid
        If lHandle = 0 Then GoTo skip
        
        '/* size module array
        lReturn = EnumProcessModules(lHandle, aModule(1), 1024, lDimension)
        lReturn = EnumProcessModules(lHandle, aModule(1), lDimension, lDimension)
        '/* get the file name
        If Not lReturn = 0 Then
            sTemp = Space$(MAX_PATH)
            lSize = 500
            'Get the module file name
            lReturn = GetModuleFileNameExA(lHandle, aModule(1), sTemp, lSize)
            sTemp = Left$(sTemp, lReturn)
            '/* add to collection name/id/path
            sTemp = Mid$(sTemp, InStrRev(sTemp, Chr$(92)) + 1)
            If bLoose Then
                If LCase$(sTemp) = LCase$(sProcess) Then
                    Process_ExistsG2 = True
                    CloseHandle lHandle
                    Exit For
                End If
            Else
                '/* success
                If sTemp = sProcess Then
                    Process_ExistsG2 = True
                    CloseHandle lHandle
                    Exit For
                End If
            End If
        End If
skip:
        '/* close handle
        CloseHandle lHandle
    Next lCount

On Error GoTo 0
Exit Function

Handler:

End Function

'* Name           : Process_Terminate
'* Purpose        : terminate a process
'* Inputs         : process name
'* Outputs        : boolean
'*********************************************
Public Function Process_Terminate(ByVal lProcess As Long) As Boolean

Dim lExitCode           As Long
Dim lHandle             As Long
Dim lReturn             As Long

On Error GoTo Handler

    '/* get process handle
    lHandle = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_CTERMINATE, False, lProcess)
    If lHandle = 0 Then GoTo Handler
    '/* get exit code
    GetExitCodeProcess lHandle, lExitCode
    If lExitCode = 0 Then GoTo Handler
    '/* kill process
    lReturn = TerminateProcess(lHandle, lExitCode)
    If lReturn = 0 Then GoTo Handler
    
    '/* success
    CloseHandle lHandle
    Process_Terminate = True
    RaiseEvent eNComplete("Process ID: " + CStr(lProcess) + " has been terminated.")

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Process_Terminate", CStr(Err.LastDllError))
    If Not lHandle = 0 Then CloseHandle lHandle

End Function

'* Name           : Process_GetClass
'* Purpose        : return the process class
'* Inputs         : process name
'* Outputs        : long
'*********************************************
Public Function Process_GetClass(ByVal lProcess As Long) As Long
'/* return the process class level

Dim lReturn             As Long
Dim sResult             As String
Dim lHandle             As Long

On Error GoTo Handler

    '/* get process handle
    lHandle = OpenProcess(PROCESS_QUERY_INFORMATION, False, lProcess)
    If lHandle = 0 Then
        RaiseEvent eNErrorCond("Process_GetClass", "Could not aquire process handle.")
        Exit Function
    End If
    '/* return class
    lReturn = GetPriorityClass(lHandle)
    
    If Not lReturn = -1 Then
        Process_GetClass = lReturn
        Select Case lReturn
        '/* Idle
        Case &H40
            sResult = "Idle"
        '/* Normal
        Case &H20
            sResult = "Normal"
        Case &H8000&
            sResult = "Above Normal"
        '/* High
        Case &H80
            sResult = "High"
        '/* Priority
        Case &H100
            sResult = "Realtime"
        '/* unknown
        Case Else
            sResult = "Unknown"
            Debug.Print "&H" & Hex(lReturn)
        End Select
        RaiseEvent eNComplete("Process ID: " + CStr(lProcess) + " priority class is " + sResult + Chr(46))
    Else
        GoTo Handler
    End If
    
    '/* cleanup
    CloseHandle lHandle
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Process_GetClass", CStr(Err.LastDllError))
    If Not lHandle = 0 Then CloseHandle lHandle

End Function

'* Name           : Process_SetClass
'* Purpose        : change process class
'* Inputs         : process name, level
'* Outputs        : boolean
'*********************************************
Public Function Process_SetClass(ByVal lProcess As Long, _
                                 ByVal eLevel As eProcessClass) As Boolean

'/* alter process class for selected process

Dim lClass              As Long
Dim lHandle             As Long
Dim lReturn             As Long

On Error GoTo Handler

    lClass = eLevel
    '/* get process handle
    lHandle = OpenProcess(PROCESS_SET_INFORMATION, False, lProcess)
    If lHandle = 0 Then
        RaiseEvent eNErrorCond("Process_GetClass", "Could not aquire process handle.")
        Exit Function
    End If
    
    lReturn = SetPriorityClass(lHandle, lClass)
    
    '/* success
    If Not lReturn = 0 Then
        Process_SetClass = True
        RaiseEvent eNComplete("The Process ID: " + CStr(lProcess) + " priority has been changed.")
    Else
        GoTo Handler
    End If
    
    '/* cleanup
    CloseHandle lHandle
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Process_SetClass", CStr(Err.LastDllError))
    If Not lHandle = 0 Then CloseHandle lHandle

End Function


'>  Thread Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'* Name           : Thread_Enumerate
'* Purpose        : list all process threads
'* Inputs         : process name
'* Outputs        : collection
'*********************************************
Public Function Thread_Enumerate(ByVal lProcess As Long) As Collection

Dim lReturn             As Long
Dim lSnapShot           As Long
Dim sItem               As String
Dim cTemp               As Collection
Dim tThread             As THREADENTRY32

On Error GoTo Handler

    '/* set up variables
    Set cTemp = New Collection
    '/* create a snapshot
    lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, lProcess)
    '/* prepare structure
    tThread.dwSize = Len(tThread)
    lReturn = Thread32First(lSnapShot, tThread)
    
    '/* get first entry
    If Not lReturn = 0 Then
        With tThread
            If .th32OwnerProcessID = lProcess Then
                sItem = CStr(.th32ThreadID) + Chr$(31) + _
                CStr(.cntUsage) + Chr$(31) + _
                CStr(.dwFlags) + Chr$(31) + _
                CStr(.th32OwnerProcessID) + Chr$(31) + _
                CStr(.tpBasePri) + Chr$(31) + _
                CStr(.tpDeltaPri)
                cTemp.Add sItem
            End If
        End With
    Else
        GoTo Handler
    End If
    
    '/* loop through entries
    '/* id/usage/flags/owner/base/delta
    Do While lReturn
        lReturn = Thread32Next(lSnapShot, tThread)
        If lReturn = 0 Then Exit Do
        With tThread
            If .th32OwnerProcessID = lProcess Then
                sItem = CStr(.th32ThreadID) + Chr$(31) + _
                CStr(.cntUsage) + Chr$(31) + _
                CStr(.dwFlags) + Chr$(31) + _
                CStr(.th32OwnerProcessID) + Chr$(31) + _
                CStr(.tpBasePri) + Chr$(31) + _
                CStr(.tpDeltaPri)
                cTemp.Add sItem
            End If
        End With
    Loop

    '/* success
    If cTemp.Count > 0 Then
        RaiseEvent eNComplete("Process ID: " + CStr(lProcess) + " threads have been enumerated.")
        Set Thread_Enumerate = cTemp
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Thread_Enumerate", CStr(Err.LastDllError))
    If Not lSnapShot = 0 Then CloseHandle lSnapShot

End Function

'* Name           : Thread_Suspend
'* Purpose        : suspend thread operations
'* Inputs         : thread id, thread name
'* Outputs        : boolean
'*********************************************
Public Function Thread_Suspend(ByVal lThread As Long, _
                               ByVal sThread As String) As Boolean

Dim lHandle             As Long
Dim lReturn             As Long

On Error GoTo Handler

    '/* get thread handle
    lHandle = OpenThread(THREAD_SUSPEND_RESUME, False, lThread)
    If lHandle = 0 Then GoTo Handler
    '/* suspend
    lReturn = SuspendThread(lHandle)
    
    '/* success
    If lReturn = 0 Then
        Thread_Suspend = True
        RaiseEvent eNComplete("Thread: " + sThread + " has been suspended.")
    Else
        GoTo Handler
    End If
    
    '/* cleanup
    CloseHandle lHandle
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Thread_Suspend", CStr(Err.LastDllError))
    If Not lHandle = 0 Then CloseHandle lHandle

End Function

'* Name           : Thread_Resume
'* Purpose        : resume thread operations
'* Inputs         : thread id, thread name
'* Outputs        : boolean
'*********************************************
Public Function Thread_Resume(ByVal lThread As Long, _
                              ByVal sThread As String) As Boolean

Dim lHandle             As Long
Dim lReturn             As Long

On Error GoTo Handler

    '/* get thread handle
    lHandle = OpenThread(THREAD_SUSPEND_RESUME, False, lThread)
    If lHandle = 0 Then GoTo Handler
    '/* resume
    lReturn = ResumeThread(lHandle)
    
    '/* success
    If lReturn = 1 Then
        Thread_Resume = True
        RaiseEvent eNComplete("Thread: " + sThread + " has Resumed.")
    Else
        GoTo Handler
    End If
    
    '/* cleanup
    CloseHandle lHandle

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Thread_Resume", CStr(Err.LastDllError))
    If Not lHandle = 0 Then CloseHandle lHandle

End Function

'* Name           : Thread_Terminate
'* Purpose        : terminate an active thread
'* Inputs         : thread id
'* Outputs        : boolean
'*********************************************
Public Function Thread_Terminate(ByVal lThread As Long) As Boolean
'/* terminate a thread

Dim lReturn             As Long

On Error GoTo Handler

    If lThread = 0 Then GoTo Handler
    '/* kill the thread
    lReturn = TerminateThread(lThread, 0&)
    
    '/* success
    If lReturn = 0 Then
        Thread_Terminate = True
        RaiseEvent eNComplete("Thread: " + CStr(lThread) + " has been Terminated.")
    Else
        GoTo Handler
    End If
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Thread_Terminate", CStr(Err.LastDllError))

End Function

'* Name           : Thread_GetPriority
'* Purpose        : return the current thread priority
'* Inputs         : thread id
'* Outputs        : long
'*********************************************
Public Function Thread_GetPriority(ByVal lThread As Long) As Long
'/* return the thread priority level

Dim lHandle             As Long
Dim lReturn             As Long
Dim sResult             As String

On Error GoTo Handler

    If lThread = 0 Then GoTo Handler
    '/* get thread handle
    lHandle = OpenThread(THREAD_QUERY_INFORMATION, False, lThread)
    If lHandle = 0 Then GoTo Handler
    '/* call for thread priority
    lReturn = GetThreadPriority(lHandle)
    '/* translate return value
    If Not lReturn = -1 Then
        Thread_GetPriority = lReturn
        Select Case lReturn
        '/* Idle
        Case -15
            sResult = "Idle"
        '/* LowRT
        Case 15
            sResult = "LowRT"
        '/* Minimum
        Case -2
            sResult = "Minimum"
        '/* Normal
        Case 0
            sResult = "Normal"
        '/* Maximum
        Case 2
            sResult = "Maximum"
        End Select
        RaiseEvent eNComplete("Thread: " + CStr(lThread) + " is " + sResult + Chr(46))
    Else
        GoTo Handler
    End If

    '/* cleanup
    CloseHandle lHandle
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Thread_GetPriority", CStr(Err.LastDllError))
    If Not lHandle = 0 Then CloseHandle lHandle

End Function

'* Name           : Thread_SetPriority
'* Purpose        : modify thread priority level
'* Inputs         : thread id, priority
'* Outputs        : boolean
'*********************************************
Public Function Thread_SetPriority(ByVal lThread As Long, _
                                   ByVal eLevel As eThreadPriority) As Boolean

'/* alter thread priority

Dim lPriority           As Long
Dim lReturn             As Long
Dim lHandle             As Long

On Error GoTo Handler
    
    If lThread = 0 Then GoTo Handler
    '/* get thread handle
    lHandle = OpenThread(THREAD_SET_INFORMATION, False, lThread)
    If lHandle = 0 Then GoTo Handler
    lPriority = eLevel
    lReturn = SetThreadPriority(lHandle, lPriority)
    
    '/* success
    If Not lReturn = 0 Then
        Thread_SetPriority = True
        RaiseEvent eNComplete("Thread: " + CStr(lThread) + " has priority has been changed.")
    Else
        GoTo Handler
    End If
    
    '/* cleanup
    CloseHandle lHandle
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Thread_SetPriority", CStr(Err.LastDllError))
    If Not lHandle = 0 Then CloseHandle lHandle

End Function

'* Name           : Module_Enumerate
'* Purpose        : list all process modules
'* Inputs         : process name
'* Outputs        : collection
'*********************************************
Public Function Module_Enumerate(ByVal lProcess As Long) As Collection

Dim lReturn             As Long
Dim lSnapShot           As Long
Dim sItem               As String
Dim cTemp               As Collection
Dim tModule             As MODULEENTRY32

On Error GoTo Handler

    '/* set up variables
    Set cTemp = New Collection
    '/* create a snapshot
    lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, lProcess)
    If lSnapShot = -1 Then
        RaiseEvent eNComplete("This Process has No Modules associated with it.")
        Exit Function
    End If
    '/* prepare struct
    tModule.dwSize = Len(tModule)
    '/* fill struct
    lReturn = Module32First(lSnapShot, tModule)
    
    '/* get first entry
    If Not lReturn = 0 Then
        With tModule
            sItem = .szModule + Chr$(31) + _
            .szExePath + Chr$(31) + _
            CStr(.GlblcntUsage) + Chr$(31) + _
            CStr(.hModule) + Chr$(31) + _
            CStr(.modBaseAddr) + Chr$(31) + _
            CStr(.ProccntUsage) + Chr$(31) + _
            CStr(.th32ProcessID) + Chr$(31)
        End With
        cTemp.Add sItem
    Else
        GoTo Handler
    End If
    
    '/* loop through entries
    '/* name/exe path/id32/usage/modid/address/procuse/prcid
    Do While lReturn
        lReturn = Module32Next(lSnapShot, tModule)
        If lReturn = 0 Then Exit Do
        With tModule
            sItem = CStr(.szModule) + Chr$(31) + _
            CStr(.szExePath) + Chr$(31) + _
            CStr(.GlblcntUsage) + Chr$(31) + _
            CStr(.hModule) + Chr$(31) + _
            CStr(.modBaseAddr) + Chr$(31) + _
            CStr(.ProccntUsage) + Chr$(31) + _
            CStr(.th32ProcessID) + Chr$(31)
        End With
        cTemp.Add sItem
    Loop
    
    '/* success
    If cTemp.Count > 0 Then
        RaiseEvent eNComplete("Process ID: " + CStr(lProcess) + " modules have been enumerated.")
        Set Module_Enumerate = cTemp
    End If
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Module_Enumerate", CStr(Err.LastDllError))
    If Not lSnapShot = 0 Then CloseHandle lSnapShot

End Function

'* Name           : Module_EnumG2
'* Purpose        : list all process modules using psapi
'* Inputs         : process name
'* Outputs        : collection
'*********************************************
Public Function Module_EnumG2(ByVal lProcess As Long) As Collection

Dim aModule(1 To 1024)  As Long
Dim lBuffer             As Long
Dim lCount              As Long
Dim lHandle             As Long
Dim lReturn             As Long
Dim lSize               As Long
Dim sModule             As String
Dim cTemp               As Collection

On Error GoTo Handler

    Set cTemp = New Collection
    lSize = 500
    '/* get the process handle
    lHandle = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcess)
    If lHandle = 0 Then GoTo Handler
    '/* size the buffer
    lReturn = EnumProcessModules(lHandle, aModule(1), 1024, lBuffer)
    '/* get the modules
    lReturn = EnumProcessModules(lHandle, aModule(1), lBuffer, lBuffer)
    sModule = Space$(MAX_PATH)
    
    '/* get the first file name
    lReturn = GetModuleFileNameExA(lHandle, aModule(1), sModule, lSize)
    '/* no modules
    If lReturn = 0 Then
        RaiseEvent eNComplete("This Process has No Modules associated with it.")
        CloseHandle lHandle
        Exit Function
    End If
    
    Do
        lCount = lCount + 1
        sModule = Space$(MAX_PATH)
        lSize = 500
        lReturn = GetModuleFileNameExA(lHandle, aModule(lCount), sModule, lSize)
        sModule = Left$(sModule, lReturn)
        '/* add to collection name/path/handle/parent id
        If LenB(Trim$(sModule)) Then
            cTemp.Add Mid$(sModule, InStrRev(sModule, Chr$(92)) + 1) + Chr$(31) + _
                sModule + Chr$(31) + _
                CStr(aModule(lCount)) + Chr$(31) + _
                CStr(lProcess)
        End If
    Loop While aModule(lCount) > 0

    '/* cleanup
    CloseHandle lHandle
    
    '/* success
    If cTemp.Count > 0 Then
        Set Module_EnumG2 = cTemp
        RaiseEvent eNComplete("Process ID: " + CStr(lProcess) + " modules have been enumerated.")
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Module_EnumG2", CStr(Err.LastDllError))
    If Not lHandle = 0 Then CloseHandle lHandle
    
End Function

'* Name           : Return_ProcessID
'* Purpose        : return process id from name (kernal32)
'* Inputs         : process name
'* Outputs        : long
'*********************************************
Public Function Return_ProcessID(ByVal sProcess As String) As Long
'/* return a process id from a process name

Dim lEntry              As Long
Dim lSnapShot           As Long
Dim sItem               As String
Dim tProc               As PROCESSENTRY32

On Error GoTo Handler

    With tProc
        '/* create our snapshot
        lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
        tProc.dwSize = Len(tProc)
        '/* first entry
        lEntry = Process32First(lSnapShot, tProc)
        '/* loop through entries
        Do While lEntry
            lEntry = Process32Next(lSnapShot, tProc)
            If lEntry = 0 Then Exit Do
            '/* compare name against process list
            sItem = Left$(.szexeFile, InStr(1, .szexeFile, Chr$(0)) - 1)
            If LCase$(sProcess) = LCase$(sItem) Then
                Return_ProcessID = .th32ProcessID
                Exit Do
            End If
        Loop
        CloseHandle lSnapShot
    End With
    
On Error GoTo 0
Exit Function

Handler:
    If Not lSnapShot = 0 Then CloseHandle lSnapShot
    Return_ProcessID = 0

End Function

'* Name           : Return_ProcessIDG2
'* Purpose        : return process id from name (psapi)
'* Inputs         : process name
'* Outputs        : long
'*********************************************
Public Function Return_ProcessIDG2(ByVal sProcess As String) As Long
'/* return a process id from a process name

Dim lBuffer             As Long
Dim lAllocate           As Long
Dim lItems              As Long
Dim aProcessID()        As Long
Dim lDimension          As Long
Dim aModule(1 To 1024)  As Long
Dim lReturn             As Long
Dim lSize               As Long
Dim lHandle             As Long
Dim lCount              As Long
Dim sTemp               As String

On Error GoTo Handler

    '/* size our array
    lBuffer = 8
    lAllocate = 96
    Do While lBuffer <= lAllocate
        lBuffer = lBuffer * 2
        ReDim aProcessID(lBuffer / 4) As Long
        lReturn = EnumProcesses(aProcessID(1), lBuffer, lAllocate)
    Loop
    
    '/* loop through return
    lItems = lAllocate / 4
    For lCount = 1 To lItems
        '/* open process
        lHandle = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, aProcessID(lCount))
        '/* invalid
        If lHandle = 0 Then GoTo skip
        
        '/* size module array
        lReturn = EnumProcessModules(lHandle, aModule(1), 1024, lDimension)
        lReturn = EnumProcessModules(lHandle, aModule(1), lDimension, lDimension)
        '/* get the file name
        If Not lReturn = 0 Then
            sTemp = Space$(MAX_PATH)
            lSize = 500
            'Get the module file name
            lReturn = GetModuleFileNameExA(lHandle, aModule(1), sTemp, lSize)
            sTemp = Left$(sTemp, lReturn)
            '/* add to collection name/id/path
            sTemp = Mid$(sTemp, InStrRev(sTemp, Chr$(92)) + 1)
            If LCase$(sTemp) = LCase$(sProcess) Then
                Return_ProcessIDG2 = aProcessID(lCount)
                CloseHandle lHandle
                Exit For
            End If
        End If
skip:
        '/* close handle
        CloseHandle lHandle
    Next lCount

On Error GoTo 0
Exit Function

Handler:
    Return_ProcessIDG2 = 0

End Function


'***************************************************************************************
'*                                 ENCRYPTED FILE SYSTEM
'***************************************************************************************

'* Name           : EFS_Status
'* Purpose        : return EFS status for path
'* Inputs         : path
'* Outputs        : long
'*********************************************
Public Function EFS_Status(ByVal sPath As String) As Long

Dim lReturn             As Long
Dim lStatus             As Long
Dim sMessage            As String

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

Dim lReturn             As Long

On Error GoTo Handler

    lReturn = EncryptFile(sPath)
    Select Case lReturn
    '/* denied
    Case 6010
        RaiseEvent eNComplete("Encrypted is Disabled for this Object.")
    '/* success
    Case 1
        EFS_Encrypt = True
        RaiseEvent eNComplete("The Object has been Encrypted.")
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

Dim lReturn             As Long

On Error GoTo Handler

    lReturn = DecryptFile(sPath, 0&)
    Select Case lReturn
    '/* denied
    Case 6010
        RaiseEvent eNComplete("Encryption is Disabled for this Object.")
    '/* success
    Case 1
        EFS_Decrypt = True
        RaiseEvent eNComplete("The Object has been Decrypted.")
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

Dim bPath               As Byte
Dim lReturn             As Long

On Error GoTo Handler

    '/* test for valid path
    If InStr(1, sPath, ".") > 0 Then GoTo Handler
    bPath = sPath + Chr$(0)
    '/* reenable encryption
    lReturn = EncryptionDisable(bPath, 0)
    If Not lReturn = 0 Then GoTo Handler
    
    '/* success
    EFS_Enable = True
    RaiseEvent eNComplete("Encryption for " + sPath + " has been enabled.")

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

Dim bPath               As Byte
Dim lReturn             As Long

On Error GoTo Handler

    '/* test path
    If InStr(1, sPath, ".") > 0 Then GoTo Handler
    bPath = sPath + Chr$(0)
    '/* disable encrytion
    lReturn = EncryptionDisable(bPath, 1)
    If Not lReturn = 0 Then GoTo Handler
    
    '/* success
    EFS_Disable = True
    RaiseEvent eNComplete("Encryption for " + sPath + " has been disabled.")
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("EFS_Disable", CStr(Err.LastDllError))

End Function

'* Name           : Create_Directory
'* Purpose        : create a new directory
'* Inputs         : path
'* Outputs        : boolean
'*********************************************
Public Function Create_Directory(ByVal sPath As String) As Boolean
'/* create a new directory

Dim lResult             As Long

    lResult = MakeSureDirectoryPathExists(sPath)
    If lResult = 0 Then Create_Directory = True

End Function

'* Name           : File_Exists
'* Purpose        : test for file/path
'* Inputs         : path
'* Outputs        : boolean
'*********************************************
Public Function File_Exists(ByVal sDir As String) As Boolean
'/* test file

Dim lRes                As Long
Dim sPath               As String

    '/* prep var
    sPath = String$(255, 0)
    '/* get short path
    lRes = GetShortPathName(sDir, sPath, 254)
    '/* success
    File_Exists = lRes > 0

End Function


'***************************************************************************************
'*                                 FILE SYSTEM SECURITY
'***************************************************************************************

'* Name           : NTFS_Drive
'* Purpose        : test ntfs status
'* Inputs         : path
'* Outputs        : boolean
'*********************************************
Public Function NTFS_Drive(ByVal sPath As String) As Boolean

Dim lFlags              As Long
Dim lMaxLen             As Long
Dim lSerial             As Long
Dim sName               As String * 256
Dim sType               As String * 256

    '/* test and shorten string
    If Len(sPath) > 3 Then sPath = Left$(sPath, 3)
    '/* get volume flags
    GetVolumeInformation sPath, sName, Len(sName), lSerial, lMaxLen, lFlags, sType, Len(sType)
    '/* success
    If lFlags And &H40000 Then NTFS_Drive = True

End Function

'* Name           : NTFS_Check
'* Purpose        : return list of ntfs drives
'* Inputs         : none
'* Outputs        : collection
'*********************************************
Public Function NTFS_Check() As Collection

Dim lBuffer             As Long
Dim lCount              As Long
Dim sDrives             As String
Dim aDrives()           As String
Dim cTemp               As Collection

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

'* Name           : NTFS_Folder
'* Purpose        : modify ntfs permissions
'* Inputs         : path, account, perm, inherit, access
'* Outputs        : boolean
'*********************************************
Public Function NTFS_Folder(ByVal sFolder As String, _
                            ByVal sAccount As String, _
                            ByVal eFile As eFolderPermissions, _
                            ByVal eFlags As eInheritenceFlags, _
                            ByVal eType As eAccessType) As Boolean

'/* set folder security (gate)

Dim lFile               As Long
Dim lReturn             As Long
Dim tAccount            As ACCOUNT_PERM
Dim tAuthority          As SID_IDENTIFIER_AUTHORITY

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
            NTFS_Folder = True
            RaiseEvent eNComplete("The Object: " + sFolder + " permissions were changed successfully.")
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
                               ByVal eFile As eFolderPermissions, _
                               ByVal eFlags As eInheritenceFlags, _
                               ByVal eType As eAccessType, _
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
    NTFS_Recursive = True
    RaiseEvent eNComplete("The Object: " + sFolder + " permissions were changed successfully.")

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("NTFS_Recursive", CStr(Err.LastDllError))

End Function


'>  NTFS Support Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Function Build_ACE(ByVal lAclId As Long, _
                           ByVal bType As Byte, _
                           ByVal bFlags As Byte, _
                           ByVal lMask As Long, _
                           ByVal lPointer As Long) As Boolean
'/* build an ace entry

Dim lReturn             As Long
Dim lAceLen             As Long
Dim lSidLen             As Long
Dim lAce                As Long
Dim tTempAce            As ACE

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
    CopyMemory ByVal lAce, ByVal VarPtr(tTempAce), LenB(tTempAce)
    CopyMemory ByVal lAce + 8, ByVal lPointer, lSidLen
    
    '/* add to acl
    lReturn = AddAce(lAclId, ACL_REVISION, MAXDWORD, lAce, lAceLen)
    If Not lReturn = 0 Then Build_ACE = True
    LocalFree lAce

Handler:
On Error GoTo 0

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
            '/* next ace
            GetAce lPAcl, lCount, lPointer
            CopyMemory tempAce, ByVal lPointer, LenB(tempAce)
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
            CopyMemory tempAce, ByVal lPointer, LenB(tempAce)
            AddAce lNewACL, ACL_REVISION, MAXDWORD, lPointer, tempAce.Header.AceSize
            lAceIndex = lAceIndex + 1
        Next lCount
    End If
    
    '/* descriptor flags
    If lOldSD <> 0 Then
        lReturn = GetSecurityDescriptorControl(lOldSD, lFlag, lRevision)
        If Not lReturn = 0 Then
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
        If Not .pSD = 0 Then LocalFree .pSD
        If Not .pAcl = 0 Then LocalFree .pAcl
        .pSD = 0
        .pAcl = 0
    End With

End Function

Private Function Kernal_Permissions(ByVal lObject As Long, _
                                    tAccount As ACCOUNT_PERM) As Boolean

'/* apply object security

Dim lLength             As Long
Dim lOldSD              As Long
Dim lReturn             As Long
Dim tInfo               As MEM_DATA

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
        If Not LenB(m_Next) > 0 Then GoTo skip
        If AscW(m_Next) = 46 Then GoTo skip
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
skip:
        bValue = FindNextFile(lHandle, W32)
    Loop
    FindClose lHandle

On Error GoTo 0

End Sub


'***************************************************************************************
'*                                  REGSITRY SECURITY
'***************************************************************************************

'* Name           : NTFS_Key
'* Purpose        : modify key permissions
'* Inputs         : hkey, path, account, mask, access, inheritence
'* Outputs        : boolean
'*********************************************
Public Function NTFS_Key(ByVal lHKey As Long, _
                         ByVal sKey As String, _
                         ByVal sAccount As String, _
                         ByVal eMask As eRegistryAccess, _
                         ByVal eType As eAccessType, _
                         ByVal eFlags As eInheritenceFlags) As Boolean

'/* set key permissions (gate)

Dim lKey                As Long
Dim lReturn             As Long
Dim tAccount            As ACCOUNT_PERM
Dim tAuthority          As SID_IDENTIFIER_AUTHORITY

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
            NTFS_Key = True
            RaiseEvent eNComplete("The Key: " + sKey + " permissions were changed successfully.")
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

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("NTFS_Key", CStr(Err.LastDllError))

End Function

Private Function Key_Permissions(ByVal lHKey As Long, _
                                 tAccount As ACCOUNT_PERM) As Boolean

'/* apply key permissions

Dim lLength             As Long
Dim lOldSD              As Long
Dim lReturn             As Long
Dim tInfo               As MEM_DATA

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


'>  Enum Access
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Function Return_Access(eAccess As eAccessMask) As Long
'/* access mask
    Return_Access = eAccess
End Function

Public Function Return_Folder(eFolder As eFolderPermissions) As Long
'/* permissions mask
    Return_Folder = eFolder
End Function

Public Function Return_Inherit(eInherit As eInheritenceFlags) As Long
'/* inheritence flags
    Return_Inherit = eInherit
End Function

Public Function Return_Type(eType As eAccessType) As Long
'/* access type
    Return_Type = eType
End Function

Public Function Return_RegAccess(eRegAccess As eRegistryAccess) As Long
'/* registry access mask
    Return_RegAccess = eRegAccess
End Function

Public Function Return_UserFlags(eFlags As eUserFlags) As Long
'/* user account flags
    Return_UserFlags = eFlags
End Function
