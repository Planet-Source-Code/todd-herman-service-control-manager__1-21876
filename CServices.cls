VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'This Class is basically a glorified type.  It's used to hold information about
'a specific service.

' CLASS DESCRIPTION:
' sServiceName=The name of the service in the service control manager database.
' sDisplayName=The name used by user interface programs to identify the service.
' lServiceType=This flag indicates the type of service.  The possible flags are:
'      SERVICE_WIN32_OWN_PROCESS
'      SERVICE_WIN32_SHARE_PROCESS
'      SERVICE_KERNEL_DRIVER
'      SERVICE_FILE_SYSTEM_DRIVER
'      SERVICE_INTERACTIVE_PROCESS
' lCurrentStatus=Indicates the current state of the service.  Possible states are:
'      SERVICE_STOPPED
'      SERVICE_START_PENDING
'      SERVICE_STOP_PENDING
'      SERVICE_RUNNING
'      SERVICE_CONTINUE_PENDING
'      SERVICE_PAUSE_PENDING
'      SERVICE_PAUSED
' lControlsAccepted=Specifies the control codes that the service will accept
'      and process.  The possible control codes are:
'      SERVICE_ACCEPT_STOP
'      SERVICE_ACCEPT_PAUSE_CONTINUE
'      SERVICE_ACCEPT_SHUTDOWN
'      SERVICE_ACCEPT_PARAMCHANGE           (Windows 2000 only)
'      SERVICE_ACCEPT_NETBINDCHANGE         (Windows 2000 only)
'      SERVICE_ACCEPT_HARDWAREPROFILECHANGE (Windows 2000 only)
'      SERVICE_ACCEPT_POWEEVENT             (Windows 2000 only)
' lWin32ExitCode=Specifies an Win32 error code of an error that occurred when the
'      service was starting or stopping.
' lServiceSpecificExitCode=If an error occurs when the service is starting or
'      stopping that is specific to the service itself, then the error code is
'      recorded in this variable.
' lCheckPoint=Specifies a value that the service increments periodically to report
'      its progress during lengthy start, stop, pause, or continue operations.  This
'      value is not valid and should be zero when the service does not have a start,
'      stop, pause, or continue operation pending.
' lWaitHint=Specifies an estimate of the amount of time, in milliseconds, that the
'      service expects a pending start, stop, pause, or continue operation to take
'      before the service makes iits next call to its SetServiceStatus function
'      with either an incremented lCheckPoint value or a change in lCurrentStatus.
'      If the amount of time specified passes, and lCheckPoint has not been
'      incremented, or lCurrentStatus has not changed, the service control manager
'      or service control program can assume that an error has occurred.
' lStartType=Specifies when to start the service.  Possible flags are:
'      SERVICE_BOOT_START
'      SERVICE_SYSTEM_START
'      SERVICE_AUTO_START
'      SERVICE_DEMAND_START
'      SERVICE_DISABLED
' lErrorControl=Specifies the severity of the error if this service fails to start
'      during startup, and determines the action taken by the startup program if
'      failure occurs.  The possible flags:
'      SERVICE_ERROR_IGNORE
'      SERVICE_ERROR_NORMAL
'      SERVICE_ERROR_SEVERE
'      SERVICE_ERROR_CRITICAL
' sPathName=Path pointing to the service file.
' sLoadOrderGroup=Name of the load ordering group which this service is a member.
' lTagID=Specifies the unique tag value for this service in the group specified
'      by the sLoadOrderGroup parameter.
' sDependencies=Null-separated list of names of services or load ordering groups
'      that must start before this service.  A group name will be prefixed by
'      the character defined in SC_GROUP_IDENTIFIER.
' sServiceStartName=This is that account name (Domain\User) which the service
'      process will be logged on as when it runs.

'Service Status Properties
Private sDisplayName As String
Private sServiceName As String
Private lServiceType As Long
Private lCurrentStatus As Long
Private lControlsAccepted As Long
Private lWin32ExitCode As Long
Private lServiceSpecificExitCode As Long
Private lCheckPoint As Long
Private lWaitHint As Long

'Service Config Properties
Private lStartType As Long
Private lErrorControl As Long
Private sPathName As String
Private sLoadOrderGroup As String
Private lTagID As Long
Private sDependencies As String
Private sServiceStartName As String
Public Property Let WaitHint(ByVal lNewData As Long)
    lWaitHint = lNewData
End Property
Public Property Get WaitHint() As Long
    WaitHint = lWaitHint
End Property
Public Property Let CheckPoint(ByVal lNewData As Long)
    lCheckPoint = lNewData
End Property
Public Property Get CheckPoint() As Long
    CheckPoint = lCheckPoint
End Property
Public Property Let ServiceSpecificExitCode(ByVal lNewData As Long)
    lServiceSpecificExitCode = lNewData
End Property
Public Property Get ServiceSpecificExitCode() As Long
    ServiceSpecificExitCode = lServiceSpecificExitCode
End Property
Public Property Let Win32ExitCode(ByVal lNewData As Long)
    lWin32ExitCode = lNewData
End Property
Public Property Get Win32ExitCode() As Long
    Win32ExitCode = lWin32ExitCode
End Property
Public Property Let CurrentStatus(ByVal lNewData As Long)
    lCurrentStatus = lNewData
End Property
Public Property Get CurrentStatus() As Long
    CurrentStatus = lCurrentStatus
End Property
Public Property Let ControlsAccepted(ByVal lNewData As Long)
    lControlsAccepted = lNewData
End Property
Public Property Get ControlsAccepted() As Long
    ControlsAccepted = lControlsAccepted
End Property
Public Property Let ServiceName(ByVal sNewData As String)
    sServiceName = sNewData
End Property
Public Property Get ServiceName() As String
    ServiceName = sServiceName
End Property
Public Property Let DisplayName(ByVal sNewData As String)
    sDisplayName = sNewData
End Property
Public Property Get DisplayName() As String
    DisplayName = sDisplayName
End Property
Public Property Get ServiceType() As Long
    ServiceType = lServiceType
End Property
Public Property Let ServiceType(ByVal lNewValue As Long)
    lServiceType = lNewValue
End Property
Public Property Get StartType() As Long
    StartType = lStartType
End Property
Public Property Let StartType(ByVal lNewValue As Long)
    lStartType = lNewValue
End Property
Public Property Get ErrorControl() As Long
    ErrorControl = lErrorControl
End Property
Public Property Let ErrorControl(ByVal lNewValue As Long)
    lErrorControl = lNewValue
End Property
Public Property Get PathName() As String
    PathName = sPathName
End Property
Public Property Let PathName(ByVal sNewValue As String)
    sPathName = sNewValue
End Property
Public Property Get LoadOrderGroup() As String
    LoadOrderGroup = sLoadOrderGroup
End Property
Public Property Let LoadOrderGroup(ByVal sNewValue As String)
    sLoadOrderGroup = sNewValue
End Property
Public Property Get TagID() As Long
    TagID = lTagID
End Property
Public Property Let TagID(ByVal lNewValue As Long)
    lTagID = lNewValue
End Property
Public Property Get Dependencies() As String
    Dependencies = sDependencies
End Property
Public Property Let Dependencies(ByVal sNewValue As String)
    sDependencies = sNewValue
End Property
Public Property Get StartName() As String
    StartName = sServiceStartName
End Property
Public Property Let StartName(ByVal sNewValue As String)
    sServiceStartName = sNewValue
End Property
Public Property Get Running() As Boolean
    If (lCurrentStatus And SERVICE_RUNNING) = SERVICE_RUNNING Then
        Running = True
    Else
        Running = False
    End If
End Property
Public Property Get StartPending() As Boolean
    If (lCurrentStatus And SERVICE_START_PENDING) = SERVICE_START_PENDING Then
        StartPending = True
    Else
        StartPending = False
    End If
End Property
Public Property Get ContinuePending() As Boolean
    If (lCurrentStatus And SERVICE_CONTINUE_PENDING) = SERVICE_CONTINUE_PENDING Then
        ContinuePending = True
    Else
        ContinuePending = False
    End If
End Property
Public Property Get Paused() As Boolean
    If (lCurrentStatus And SERVICE_PAUSED) = SERVICE_PAUSED Then
        Paused = True
    Else
        Paused = False
    End If
End Property
Public Property Get PausePausing() As Boolean
    If (lCurrentStatus And SERVICE_PAUSE_PENDING) = SERVICE_PAUSE_PENDING Then
        PausePausing = True
    Else
        PausePausing = False
    End If
End Property
Public Property Get Stopped() As Boolean
    If (lCurrentStatus And SERVICE_STOPPED) = SERVICE_STOPPED Then
        Stopped = True
    Else
        Stopped = False
    End If
End Property
Public Property Get StopPending() As Boolean
    If (lCurrentStatus And SERVICE_STOP_PENDING) = SERVICE_STOP_PENDING Then
        StopPending = True
    Else
        StopPending = False
    End If
End Property
Public Property Get AcceptPause() As Boolean
    If (lControlsAccepted And SERVICE_ACCEPT_PAUSE_CONTINUE) = SERVICE_ACCEPT_PAUSE_CONTINUE Then
        AcceptPause = True
    Else
        AcceptPause = False
    End If
End Property
Public Property Get AcceptShutDown() As Boolean
    If (lControlsAccepted And SERVICE_ACCEPT_SHUTDOWN) = SERVICE_ACCEPT_SHUTDOWN Then
        AcceptShutDown = True
    Else
        AcceptShutDown = False
    End If
End Property
Public Property Get AcceptStop() As Boolean
    If (lControlsAccepted And SERVICE_ACCEPT_STOP) = SERVICE_ACCEPT_STOP Then
        AcceptStop = True
    Else
        AcceptStop = False
    End If
End Property
