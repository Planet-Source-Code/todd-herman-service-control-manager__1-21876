VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSCMsgBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1005
   ClientLeft      =   3750
   ClientTop       =   8205
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   5655
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4050
      Top             =   480
   End
   Begin VB.Timer tmrAnimation 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4560
      Top             =   465
   End
   Begin MSComctlLib.ImageList imglstTimerAni 
      Left            =   5055
      Top             =   375
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCMsgBox.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCMsgBox.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCMsgBox.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCMsgBox.frx":094E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCMsgBox.frx":0C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCMsgBox.frx":0F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCMsgBox.frx":129C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCMsgBox.frx":15B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCMsgBox.frx":18D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCMsgBox.frx":1BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCMsgBox.frx":1F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCMsgBox.frx":221E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgTimerIcons 
      Height          =   480
      Left            =   90
      Top             =   285
      Width           =   480
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   660
      TabIndex        =   0
      Top             =   180
      Width           =   4875
   End
End
Attribute VB_Name = "frmSCMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iCurrentImage As Integer
Dim lTimerCounter As Long
Dim lCheckPoint As Long
Private Sub Form_Load()
    
    Call AlwaysOnTop(Me.hwnd)
        
    tmrMain.Enabled = True
    tmrMain.Interval = SERVICE_CONTROL_WAITTIME
    
    tmrAnimation.Enabled = False
    tmrAnimation.Interval = 500
    lblMessage.Caption = SERVICE_CONTROL_MSGBOX_MESSAGE
    Me.Caption = SERVICE_CONTROL_MSGBOX_TITLE
    
    imgTimerIcons.Picture = imglstTimerAni.ListImages(1).Picture
    iCurrentImage = 1
    lTimerCounter = 0
    lCheckPoint = 0
    tmrAnimation.Enabled = True
    
End Sub
Private Sub tmrAnimation_Timer()

    If iCurrentImage < 12 Then
        iCurrentImage = iCurrentImage + 1
    Else
        iCurrentImage = 1
    End If
        
    imgTimerIcons.Picture = imglstTimerAni.ListImages(iCurrentImage).Picture
    
End Sub
Private Sub tmrMain_Timer()

    Dim CurrentStatus As SERVICE_STATUS
    Dim liCurrentItem As ListItem
    
    lTimerCounter = lTimerCounter + 1
    
    CurrentStatus = ServiceStatus(SERVICE_CONTROL_MACHINENAME, SERVICE_CONTROL_SERVICENAME)
    
    If CurrentStatus.dwCurrentState = SERVICE_CONTROL_STATUS_WANTED Then
        tmrMain.Enabled = False
        
        Set liCurrentItem = frmMain.lvServices.ListItems(UCase(SERVICE_CONTROL_SERVICENAME))
        Select Case CurrentStatus.dwCurrentState
            Case SERVICE_STOPPED
                liCurrentItem.SubItems(1) = ""
                liCurrentItem.ForeColor = vbRed
            Case SERVICE_RUNNING
                liCurrentItem.SubItems(1) = "Started"
                liCurrentItem.ForeColor = vbBlack
            Case SERVICE_PAUSED
                liCurrentItem.SubItems(1) = "Paused"
                liCurrentItem.ForeColor = vbGreen
        End Select
        
        If CurrentStatus.dwWaitHint > 0 Then
            MsgBox CurrentStatus.dwWaitHint
        End If
        
        colServices(liCurrentItem.Key).CurrentStatus = CurrentStatus.dwCurrentState
        colServices(liCurrentItem.Key).ControlsAccepted = CurrentStatus.dwControlsAccepted
        
        Set liCurrentItem = Nothing
        Unload Me
    Else
        If lCheckPoint > 0 And lCheckPoint = CurrentStatus.dwCheckPoint Then
            'check point failed to increment... this could be a potential problem
            MsgBox "I was unable to perform the selected action on the service."
            tmrMain.Enabled = False
            Unload Me
        End If
        
        If CurrentStatus.dwWaitHint > 0 Then
            'Wait hint gives us a hint at how long the service needs.  Change the
            'timer interval to this time.
            tmrMain.Interval = CurrentStatus.dwWaitHint
            lCheckPoint = CurrentStatus.dwCheckPoint
        End If
        
        If CurrentStatus.dwWin32ExitCode > 0 Or CurrentStatus.dwServiceSpecificExitCode > 0 Then
            'An error occurred while performing an action on the service
            If CurrentStatus.dwWin32ExitCode = ERROR_SERVICE_SPECIFIC_ERROR Then
                'The error is service specific and will be specified by dwServiceSpecificExitCode
                MsgBox CurrentStatus.dwServiceSpecificExitCode, vbOKOnly + vbCritical, "Service Specific Error"
            Else
                MsgBox CurrentStatus.dwWin32ExitCode, vbOKOnly + vbCritical, "Error performing an action on the service"
            End If
        End If
        
        Debug.Print "WAIT HINT:               " & CStr(CurrentStatus.dwWaitHint)
        Debug.Print "Check Point:             " & CStr(CurrentStatus.dwCheckPoint)
        Debug.Print "Error Code:              " & CStr(CurrentStatus.dwWin32ExitCode)
        Debug.Print "Service Specific Error:  " & CStr(CurrentStatus.dwServiceSpecificExitCode)
        Debug.Print "Current State:           " & CStr(CurrentStatus.dwCurrentState)
        
        'Microsoft documentation states:  If the amount of time specified by dwWaitHint
        'passes, and dwCheckPoint has not been incremented, or dwCurrentState has not
        'changed, the service control manager or service control program can assume
        'that an error has occurred.
    End If
    
    Set liCurrentItem = Nothing
    
End Sub
