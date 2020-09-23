Attribute VB_Name = "modServiceCode"
'Constants*************************************************************************
'Service Control Manager object specific access types
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SC_MANAGER_CONNECT = &H1
Public Const SC_MANAGER_CREATE_SERVICE = &H2
Public Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Public Const SC_MANAGER_LOCK = &H8
Public Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
Public Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
Public Const SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)

Public Const SERVICES_ACTIVE_DATABASE = "ServicesActive"
Public Const SERVICE_ACTIVE = &H1
Public Const SERVICE_INACTIVE = &H2
Public Const SERVICE_STATE_ALL = (SERVICE_ACTIVE Or SERVICE_INACTIVE)
Public Const SERVICE_WIN32_OWN_PROCESS = &H10
Public Const SERVICE_WIN32_SHARE_PROCESS = &H20
Public Const SERVICE_WIN32 = SERVICE_WIN32_OWN_PROCESS + SERVICE_WIN32_SHARE_PROCESS
Public Const SERVICE_KERNEL_DRIVER = &H1
Public Const SERVICE_FILE_SYSTEM_DRIVER = &H2
Public Const SERVICE_RECOGNIZER_DRIVER = &H8
Public Const SERVICE_ADAPTER = &H4
Public Const SERVICE_INTERACTIVE_PROCESS = &H100
Public Const SERVICE_DRIVER = SERVICE_KERNEL_DRIVER Or SERVICE_FILE_SYSTEM_DRIVER Or SERVICE_RECOGNIZER_DRIVER

'Service object specific access types
Public Const SERVICE_QUERY_CONFIG = &H1
Public Const SERVICE_CHANGE_CONFIG = &H2
Public Const SERVICE_QUERY_STATUS = &H4
Public Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Public Const SERVICE_START = &H10
Public Const SERVICE_STOP = &H20
Public Const SERVICE_PAUSE_CONTINUE = &H40
Public Const SERVICE_INTERROGATE = &H80
Public Const SERVICE_USER_DEFINED_CONTROL = &H100
Public Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)

'Error Constants
Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_INVALID_HANDLE = 6&
Private Const ERROR_INVALID_PARAMETER = 87
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_DATABASE_DOES_NOT_EXIST = 1065&
Private Const ERROR_INSUFFICIENT_BUFFER = 122
Public Const ERROR_SERVICE_SPECIFIC_ERROR = 1066&

'Status Constants
Public Const SERVICE_STOPPED = &H1
Public Const SERVICE_START_PENDING = &H2
Public Const SERVICE_STOP_PENDING = &H3
Public Const SERVICE_RUNNING = &H4
Public Const SERVICE_CONTINUE_PENDING = &H5
Public Const SERVICE_PAUSE_PENDING = &H6
Public Const SERVICE_PAUSED = &H7

'Accepted Controls Constants
Public Const SERVICE_ACCEPT_STOP = &H1
Public Const SERVICE_ACCEPT_PAUSE_CONTINUE = &H2
Public Const SERVICE_ACCEPT_SHUTDOWN = &H4

'Control Constants
Public Const SERVICE_CONTROL_STOP = &H1
Public Const SERVICE_CONTROL_PAUSE = &H2
Public Const SERVICE_CONTROL_CONTINUE = &H3
Public Const SERVICE_ERROR_IGNORE = &H0
Public Const SERVICE_ERROR_NORMAL = &H1
Public Const SERVICE_ERROR_SEVERE = &H2
Public Const SERVICE_ERROR_CRITICAL = &H3

'StartType Constants
Public Const SERVICE_BOOT_START = &H0
Public Const SERVICE_SYSTEM_START = &H1
Public Const SERVICE_AUTO_START = &H2
Public Const SERVICE_DEMAND_START = &H3
Public Const SERVICE_DISABLED = &H4

Public Const SC_GROUP_IDENTIFIER = "+"

'Type Declarations*****************************************************************
Public Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
End Type

Private Type ENUM_SERVICE_STATUS
    lpServiceName As Long
    lpDisplayName As Long
    ServiceStatus As SERVICE_STATUS
End Type

Type QUERY_SERVICE_CONFIG
    dwServiceType As Long
    dwStartType As Long
    dwErrorControl As Long
    lpBinaryPathName As Long
    lpLoadOrderGroup As Long
    dwTagId As Long
    lpDependencies As Long
    lpServiceStartName As Long
    lpDisplayName As Long
End Type

Public Type ListOfServices
    bInit As Boolean
    lCount As Long
    lLastErr As Long
    sErrMessage As String
    List() As ENUM_SERVICE_STATUS
End Type

'API Declarations******************************************************************
Private Declare Function EnumServicesStatus Lib "advapi32.dll" Alias "EnumServicesStatusA" (ByVal hSCManager As Long, ByVal dwServiceType As Long, ByVal dwServiceState As Long, lpServices As Any, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long, lpResumeHandle As Long) As Long
Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Private Declare Function QueryServiceConfig Lib "advapi32.dll" Alias "QueryServiceConfigA" (ByVal hService As Long, lpServiceConfig As Byte, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Private Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function EnumDependentServices Lib "advapi32.dll" Alias "EnumDependentServicesA" (ByVal hService As Long, ByVal dwServiceState As Long, lpServices As Any, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long) As Long
Private Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Public Declare Function lStringCopy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lStringLength Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

Public colDepServices As New Collection
Public colServices As New Collection
Public Function EnumerateServices(sMachineName As String, lServiceType As Long, lType As Long) As Boolean

    Dim lSCMHandle As Long
    Dim ServiceStatusBuffer() As ENUM_SERVICE_STATUS
    Dim lBytesNeeded As Long
    Dim lServicesReturned As Long
    Dim lServiceStatusInfoBuffer As Long
    Dim lResumeHandle As Long
    Dim lStructsNeeded As Long
    Dim ServiceList As ListOfServices
    Dim CurrentService As CServices
    Dim ServiceConfigInfo As QUERY_SERVICE_CONFIG
    Dim sErrorMessage As String
    Dim lResult As Long
    
    'sMachineName is the machine to enumerate the services for.  This needs to be
    'NULL terminated or NULL (if empty).  The local machine is default.
    If sMachineName = "" Then
        sMachineName = vbNullString
    Else
        sMachineName = sMachineName & Chr(0)
    End If
    
    'Open connection to Service Control Manager.
    lSCMHandle = OpenSCManager(sMachineName, vbNullString, SC_MANAGER_ENUMERATE_SERVICE)
    
    'If an error results exit sub
    If lSCMHandle = 0 Then
        sErrorMessage = GetErrorMessage(ServiceList.lLastErr, "EnumerateServices", "OpenSCManager")
        MsgBox "The following error occurred (" & Err.LastDllError & "):" & vbCrLf & sErrorMessage, vbOKOnly + vbCritical, "Error enumerating services"
        EnumerateServices = False
        Exit Function
    End If
    
    'Set the pointer to the first entry
    lResumeHandle = 0
    
    'Call the function with the buffer set to 0 this will generate a ERROR_MORE_DATA
    'return, but it also return the size of the buffer we will need
    lResult = EnumServicesStatus(lSCMHandle, lType, lServiceType, ByVal &H0, &H0, lBytesNeeded, lServicesReturned, lResumeHandle)
    If Not Err.LastDllError = ERROR_MORE_DATA Then
        'If an error other than ERROR_MORE_DATA exit sub
        MsgBox "The following error occurred (" & Err.LastDllError & ")", vbOKOnly + vbCritical, "Error enumerating services"
        CloseServiceHandle (lSCMHandle)
        EnumerateServices = False
        Exit Function
    End If
    
    'Calculate the number of structures needed then redimension the array
    'according to the new values.  We then need to set the new buffer size.
    lStructsNeeded = lBytesNeeded / Len(ServiceStatusBuffer(0)) + 1
    ReDim ServiceStatusBuffer(lStructsNeeded - 1)
    lServiceStatusInfoBuffer = lStructsNeeded * Len(ServiceStatusBuffer(0))
    
    'Set the pointer to the first entry
    lResumeHandle = 0
    
    'Call the function again with the appropriate buffer
    lResult = EnumServicesStatus(lSCMHandle, lType, lServiceType, ServiceStatusBuffer(0), lServiceStatusInfoBuffer, lBytesNeeded, lServicesReturned, lResumeHandle)
    If lResult = 0 Then
        MsgBox "The following error occurred (" & Err.LastDllError & ")", vbOKOnly + vbCritical, "Error enumerating services"
        CloseServiceHandle (lSCMHandle)
        EnumerateServices = False
        Exit Function
    End If
    
    Dim i As Long
    Set colServices = Nothing
    For i = 0 To lServicesReturned - 1
        'Set a new temporary object based on the class
        Set CurrentService = New CServices
        
        CurrentService.ServiceName = LPSTRtoSTRING(ServiceStatusBuffer(i).lpServiceName)
        CurrentService.DisplayName = LPSTRtoSTRING(ServiceStatusBuffer(i).lpDisplayName)
        CurrentService.ServiceType = ServiceStatusBuffer(i).ServiceStatus.dwServiceType
        CurrentService.CurrentStatus = ServiceStatusBuffer(i).ServiceStatus.dwCurrentState
        CurrentService.ControlsAccepted = ServiceStatusBuffer(i).ServiceStatus.dwControlsAccepted
        CurrentService.Win32ExitCode = ServiceStatusBuffer(i).ServiceStatus.dwWin32ExitCode
        CurrentService.ServiceSpecificExitCode = ServiceStatusBuffer(i).ServiceStatus.dwServiceSpecificExitCode
        CurrentService.CheckPoint = ServiceStatusBuffer(i).ServiceStatus.dwCheckPoint
        CurrentService.WaitHint = ServiceStatusBuffer(i).ServiceStatus.dwWaitHint
        
        'We will also retrieve the configuration information for this service
        'at this time.  This will save time later and give more information easily
        'accessible since it will all be stored in the same class object.
        ServiceConfigInfo = GetServiceConfig(sMachineName, CurrentService.ServiceName)
        CurrentService.StartType = ServiceConfigInfo.dwStartType
        CurrentService.ErrorControl = ServiceConfigInfo.dwErrorControl
        CurrentService.PathName = LPSTRtoSTRING(ServiceConfigInfo.lpBinaryPathName)
        CurrentService.LoadOrderGroup = LPSTRtoSTRING(ServiceConfigInfo.lpLoadOrderGroup)
        CurrentService.TagID = ServiceConfigInfo.dwTagId
        CurrentService.Dependencies = LPSTRtoSTRING(ServiceConfigInfo.lpDependencies)
        CurrentService.StartName = LPSTRtoSTRING(ServiceConfigInfo.lpServiceStartName)
        
        'Add this service to our collection
        colServices.Add CurrentService, UCase(CurrentService.ServiceName)
        
        'Reset Service
        Set CurrentService = Nothing
    Next
    
    'Close our connection to the Service Control Manager
    CloseServiceHandle (lSCMHandle)
    EnumerateServices = True
  
End Function
Public Function GetServiceConfig(sMachineName As String, sServiceName As String) As QUERY_SERVICE_CONFIG

    Dim abServiceConfigInfo() As Byte
    Dim ServiceConfigInfo As QUERY_SERVICE_CONFIG
    Dim lBytesNeeded As Long
    Dim hSManager As Long
    Dim hService As Long
    Dim lResult As Long
    Dim service As CServices
    
    hSManager = OpenSCManager(sMachineName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, sServiceName, SERVICE_QUERY_CONFIG)
        If hService <> 0 Then
            ReDim abServiceConfigInfo(0) As Byte
            lResult = QueryServiceConfig(hService, abServiceConfigInfo(0), 0&, lBytesNeeded)
            
            If lResult = 0 And lBytesNeeded = 0 Then
                'No dependencies
                CloseServiceHandle hService
                CloseServiceHandle hSManager
                'GetDependentInfo = False
                Exit Function
            Else
                'If an error other than ERROR_INSUFFICIENT_BUFFER exit sub
                If Not Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
                    MsgBox "LastDLLError = " & CStr(Err.LastDllError)
                    CloseServiceHandle hService
                    CloseServiceHandle hSManager
                    Exit Function
                End If
            End If
            
            ReDim abServiceConfigInfo(lBytesNeeded) As Byte
            
            'Call the function again with the appropriate buffer
            lResult = QueryServiceConfig(hService, abServiceConfigInfo(0), lBytesNeeded, lBytesNeeded)
            
            If lResult = 0 Then
                MsgBox "QueryServicConfig failed. LastDllError = " & CStr(Err.LastDllError)
                CloseServiceHandle hService
                CloseServiceHandle hSManager
                Exit Function
            End If
            
            'Copy the information from the ByteArray to our structure
            CopyMemory ServiceConfigInfo, abServiceConfigInfo(0), Len(ServiceConfigInfo)
            
            GetServiceConfig = ServiceConfigInfo
            
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If

End Function
Public Function GetDependentInfo(sMachineName As String, ServiceName As String) As Boolean
    
    Dim lpEnumServiceStatus() As ENUM_SERVICE_STATUS
    Dim lServiceStatusInfoBuffer As Long
    Dim lBytesNeeded As Long
    Dim lServicesReturned As Long
    Dim lStructsNeeded As Long
    Dim hSManager As Long
    Dim hService As Long
    Dim lResult As Long
    Dim service As CServices
    
    'This function is like the EnumerateServices function, but it enumerates
    'all the dependent services for a given service.  We are simply gathering
    'each dependent service (only the name) and storing them in a collection
    'to be used when stopping a service.
    hSManager = OpenSCManager(sMachineName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ENUMERATE_DEPENDENTS)
        If hService <> 0 Then
            lResult = EnumDependentServices(hService, SERVICE_STATE_ALL, ByVal &H0, &H0, lBytesNeeded, lServicesReturned)
            
            If lResult = 1 And lBytesNeeded = 0 Then
                'No dependencies
                CloseServiceHandle hService
                CloseServiceHandle hSManager
                GetDependentInfo = False
                Exit Function
            Else
                'If an error other than ERROR_MORE_DATA exit sub
                If Not Err.LastDllError = ERROR_MORE_DATA Then
                    MsgBox "LastDLLError = " & CStr(Err.LastDllError)
                    CloseServiceHandle hService
                    CloseServiceHandle hSManager
                    GetDependentInfo = False
                    Exit Function
                End If
            End If
            
            'Calculate the number of structures needed then redimension the array
            'according to the new values.  We then need to set the new buffer size.
            lStructsNeeded = lBytesNeeded / Len(lpEnumServiceStatus(0)) + 1
            ReDim lpEnumServiceStatus(lStructsNeeded - 1)
            lServiceStatusInfoBuffer = lStructsNeeded * Len(lpEnumServiceStatus(0))
            
            'Call the function again with the appropriate buffer
            lResult = EnumDependentServices(hService, SERVICE_STATE_ALL, lpEnumServiceStatus(0), lServiceStatusInfoBuffer, lBytesNeeded, lServicesReturned)
            If lResult = 0 Then
                MsgBox "EnumDependentServices failed. LastDllError = " & CStr(Err.LastDllError)
                CloseServiceHandle hService
                CloseServiceHandle hSManager
                GetDependentInfo = False
                Exit Function
            End If
    
            Dim i As Long
            Set colDepServices = Nothing
            For i = 0 To lServicesReturned - 1
                'Set a new temporary object based on the class
                Set service = New CServices
                'Resolve the Service Names
                service.ServiceName = LPSTRtoSTRING(lpEnumServiceStatus(i).lpServiceName)
                'Add this service to our collection
                colDepServices.Add service, UCase(service.ServiceName)
                'Reset Service
                Set service = Nothing
            Next
            
            GetDependentInfo = True
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If
    
End Function
Public Function ServiceStatus(ComputerName As String, ServiceName As String) As SERVICE_STATUS
    
    Dim ServiceStat As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim hServiceStatus As Long

    'ServiceStatus = ""
    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            hServiceStatus = QueryServiceStatus(hService, ServiceStat)
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
        
        ServiceStatus = ServiceStat
    End If
    
End Function
Public Sub ServiceStart(ComputerName As String, ServiceName As String)
    
    Dim ServiceStatus As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim lReturn As Long

    DoEvents
    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            lReturn = StartService(hService, 0, 0)
            If lReturn = 0 Then
                MsgBox Err.LastDllError
            End If
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If
    
End Sub
Public Sub ServiceStop(ComputerName As String, ServiceName As String)
    
    Dim ServiceStatus As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim res As Long

    DoEvents

    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            res = ControlService(hService, SERVICE_CONTROL_STOP, ServiceStatus)
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If
    
End Sub
Public Sub ServicePause(ComputerName As String, ServiceName As String)
    
    Dim ServiceStatus As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim res As Long

    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            res = ControlService(hService, SERVICE_CONTROL_PAUSE, ServiceStatus)
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If
    
End Sub
Public Sub ServiceContinue(ComputerName As String, ServiceName As String)
    
    Dim ServiceStatus As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim res As Long

    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            res = ControlService(hService, SERVICE_CONTROL_CONTINUE, ServiceStatus)
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If
    
End Sub
Private Function GetErrorMessage(lErrorNumber As Long, sProcedureName As String, sUsedFunction As String) As String

    'This function returns a formated message depending on the passed
    'error number.
    Dim sMessage As String
    
    Select Case lErrorNumber
        Case ERROR_ACCESS_DENIED            '5
            sMessage = "Access denied."
        Case ERROR_INVALID_HANDLE           '6
            sMessage = "Invalid handle was passed to the function."
        Case ERROR_INVALID_PARAMETER        '87
            sMessage = "An invalid parameter was passed."
        Case ERROR_DATABASE_DOES_NOT_EXIST  '1065
            sMessage = "SCM Database does not exist."
        Case Else
            sMessage = "An untrapped error occurred (" & lErrorNumber & ")."
    End Select
    
    GetErrorMessage = sMessage & vbCrLf & "Calling Procedure:  " & sProcedureName & vbCrLf & "API Function:  " & sUsedFunction
    
End Function
Public Function LPSTRtoSTRING(ByVal lStringPointer As Long) As String

    'This function converts a string pointer to a string
    Dim lLength As Long
    
    'Get number of characters in string
    lLength = lStringLength(ByVal lStringPointer) * 2
    
    'Initialize string so we have something to copy the string into
    LPSTRtoSTRING = String(lLength, 0)
    
    'Copy the string
    CopyMemory ByVal StrPtr(LPSTRtoSTRING), ByVal lStringPointer, lLength
    
    'Convert to Unicode from ASCII
    LPSTRtoSTRING = TrimStr(StrConv(LPSTRtoSTRING, vbUnicode))
    
End Function
Public Function TrimStr(sName As String) As String

    'Finds a null then trims the string
    Dim iNullLocal As Integer
    iNullLocal = InStr(sName, vbNullChar)
    If iNullLocal > 0 Then TrimStr = Left(sName, iNullLocal - 1) Else TrimStr = sName

End Function
