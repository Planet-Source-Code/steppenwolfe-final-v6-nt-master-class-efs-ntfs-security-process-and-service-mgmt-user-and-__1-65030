VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************************
'*  NMC! V1.6    Service Managment Member Class                                        *
'*                                                                                     *
'*  Created:     April 14, 2006                                                        *
'*  Updated:     June 02, 2006                                                        *
'*  Purpose:     NT System Security Master Class                                       *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.6.0                                                                 *
'*  Compile:     Native                                                                *
'*  Referenced:  Throughout Project                                                    *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*                                                                                     *
'***************************************************************************************

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


'/* service constants
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
   lpServiceName As Long
   lpDisplayName As Long
   ServiceStatus As SERVICE_STATUS
End Type

Private Type SERVICE_DESCRIPTION
    lpDescription                               As String
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

Public Event eNErrorCond(ByVal sRoutine As String, ByVal sError As String)
Public Event eNComplete(ByVal sTask As String)

'* Name           : Service_Start
'* Purpose        : start a service
'* Inputs         : service name
'* Outputs        : boolean
'*********************************************
Public Function Service_Start(sService As String) As Boolean

Dim lHManager           As Long
Dim lHService           As Long

On Error GoTo Handler

    '/* get handle to service manager
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
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
Public Function Service_Stop(sService As String) As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim tStatus             As SERVICE_STATUS

On Error GoTo Handler

    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
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
Public Function Service_Pause(sService As String) As Boolean

Dim Resume_Service      As Boolean
Dim lHManager           As Long
Dim lHService           As Long
Dim tStatus             As SERVICE_STATUS

On Error GoTo Handler

    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
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
Public Function Service_Resume(sService As String) As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim tStatus             As SERVICE_STATUS

On Error GoTo Handler

    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
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
    If Len(sDomain) = 0 Then sDomain = Chr$(46)
    'sAccountName = sDomain + Chr$(92) + sAccountName

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
Public Function Service_Remove(sService As String) As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim lResult             As Long
Dim tStatus             As SERVICE_STATUS

On Error GoTo Handler

    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
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
                             ByVal sDescription As String) As Boolean

Dim bDescription        As Byte
Dim lHManager           As Long
Dim lHService           As Long
Dim lResult             As Long
Dim tDesc               As SERVICE_DESCRIPTION

On Error GoTo Handler

    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    lHService = OpenService(lHManager, sService, SERVICE_ALL_ACCESS)
    bDescription = sDescription + Chr$(0)
    tDesc.lpDescription = bDescription(0)
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
Public Function Service_Query(sService As String) As Long

Dim lHManager           As Long
Dim lHService           As Long
Dim sState              As String
Dim tStatus             As SERVICE_STATUS

On Error GoTo Handler

    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
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
                               ByRef svcStartType As eServiceStartType) As Boolean

Dim lHManager           As Long
Dim lHService           As Long

On Error GoTo Handler

    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
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
Public Function Service_Enumerate() As Collection

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

    '/* open service manager
    lMngrHndl = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_ENUMERATE_SERVICE)
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


