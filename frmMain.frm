VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Services"
   ClientHeight    =   3990
   ClientLeft      =   2250
   ClientTop       =   4095
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7635
   Begin MSComctlLib.ListView lvServices 
      Height          =   3795
      Left            =   0
      TabIndex        =   6
      Top             =   195
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6694
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin Threed.SSCommand cmdClose 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   450
      Left            =   6060
      TabIndex        =   0
      Top             =   0
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   794
      _StockProps     =   78
      Caption         =   "Close"
   End
   Begin Threed.SSCommand cmdStart 
      Height          =   450
      Left            =   6060
      TabIndex        =   1
      Top             =   705
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   794
      _StockProps     =   78
      Caption         =   "&Start"
   End
   Begin Threed.SSCommand cmdStop 
      Height          =   450
      Left            =   6060
      TabIndex        =   2
      Top             =   1425
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   794
      _StockProps     =   78
      Caption         =   "S&top"
   End
   Begin Threed.SSCommand cmdPause 
      Height          =   450
      Left            =   6060
      TabIndex        =   3
      Top             =   2130
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   794
      _StockProps     =   78
      Caption         =   "&Pause"
   End
   Begin Threed.SSCommand cmdContinue 
      Height          =   450
      Left            =   6060
      TabIndex        =   4
      Top             =   2835
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   794
      _StockProps     =   78
      Caption         =   "&Continue"
   End
   Begin Threed.SSCommand cmdStartup 
      Height          =   450
      Left            =   6060
      TabIndex        =   5
      Top             =   3540
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   794
      _StockProps     =   78
      Caption         =   "Sta&rtup"
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Startup"
      Height          =   270
      Left            =   4575
      TabIndex        =   9
      Top             =   15
      Width           =   1140
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Status"
      Height          =   270
      Left            =   3150
      TabIndex        =   8
      Top             =   15
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   270
      Left            =   15
      TabIndex        =   7
      Top             =   15
      Width           =   3135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public colServices As New Collection

Public liPreviousItem As ListItem
Public bDoOnce As Boolean
Private Sub Form_Load()

    Call SetupServiceListView
    Call GetServiceList
    Call StoreServicesInListView
    bDoOnce = True
    
End Sub
Private Sub Form_Activate()

    If bDoOnce Then
        'Select the first item in the list view
        Set lvServices.SelectedItem = lvServices.ListItems(1)
        Call lvServices_ItemClick(lvServices.SelectedItem)
        Set liPreviousItem = lvServices.SelectedItem
        lvServices.SetFocus
        bDoOnce = False
    End If
    
End Sub
Private Sub cmdStart_Click()

    Dim liCurrentItem As ListItem
    Dim sDependencies As String
    Dim sSrvcName As String
    Dim sServiceList() As String
    Dim iCount As Integer
    Dim i As Integer
    
    Set liCurrentItem = lvServices.SelectedItem
    
    iCount = 0
    
    'This can get a little tricky.  We are going to attempt to duplicate a feature
    'found in Microsoft's Service Manager.  We are going to travel up the
    'dependency tree, starting each service we come to.
    
    'Check if the service is dependent on another service, don't worry about groups
    'at the moment.  For those of you that don't know, certain services can be
    'part of a defined group (sLoadOrderGroup in our class).  A service might very
    'well depend on an entire group of services to load first.  In this version
    'I am not dealing with groups at all.
    sDependencies = colServices(liCurrentItem.Key).Dependencies
    sSrvcName = colServices(liCurrentItem.Key).ServiceName
    Do While Not sDependencies = "" And InStr(1, sDependencies, SC_GROUP_IDENTIFIER) = 0
        'Lets store the current item in the array as the very first item
        If iCount = 0 Then
            ReDim sServiceList(iCount) As String
            sServiceList(iCount) = sSrvcName
        End If
        
        iCount = iCount + 1
        ReDim Preserve sServiceList(iCount) As String
        sServiceList(iCount) = sDependencies
        sDependencies = colServices(sDependencies).Dependencies
    Loop
    
    If iCount > 0 Then
        'Now we need to start each service going backwards through the array
        For i = iCount To 0 Step -1
            Call StartServiceByName(UCase(sServiceList(i)))
        Next i
    Else
        Call StartServiceByName(UCase(sSrvcName))
    End If
    
    Set liCurrentItem = Nothing
    
End Sub
Private Sub cmdStop_Click()
    
    Dim liCurrentItem As ListItem
    Dim lServiceStatus As Long
    Dim sHoldText As String
    Dim tStartTime As Date
    Dim bResult As Boolean
    Dim service As CServices
    
    Dim sMessage As String
    Dim iCount As Integer
    Dim lResult As Long
    
    Set liCurrentItem = lvServices.SelectedItem
    
    'Check if the service in question has any dependent services
    bResult = GetDependentInfo("", colServices(liCurrentItem.Key).ServiceName)
    If bResult Then
        'Check if the user is sure they want to stop all services
        sMessage = colServices(liCurrentItem.Key).ServiceName & " has the following dependent service(s) that also need to be stopped:  " & vbCrLf & vbCrLf
        
        iCount = 0
        For Each service In colDepServices
            iCount = iCount + 1
            If iCount = colDepServices.Count Then
                'We are on the last one
                sMessage = sMessage & service.DisplayName & vbCrLf
            Else
                sMessage = sMessage & service.DisplayName & vbCrLf
            End If
        Next
        
        sMessage = sMessage & vbCrLf & "Are you sure you wish to stop all the above services?"
        lResult = MsgBox(sMessage, vbYesNo + vbQuestion, "Stopping")
        If lResult = vbYes Then
            'Go ahead and stop services
                
            'The colDepServices should be filled with dependent service info now.
            For Each service In colDepServices
                'Is the service running?
                If colServices(UCase(service.ServiceName)).CurrentStatus = SERVICE_RUNNING Then
                    'The service is running, lets stop it
                    Call StopServiceByName(UCase(service.ServiceName))
                End If
            Next
            
            'Now stop the main service
            Call StopServiceByName(UCase(colServices(liCurrentItem.Key).ServiceName))
        Else
            'Don't stop services.. exit
            Exit Sub
        End If
    Else
        'Check if the user is sure they want to stop all services
        sMessage = "Are you sure you want to stop the " & colServices(liCurrentItem.Key).ServiceName & " service?"
        
        lResult = MsgBox(sMessage, vbYesNo + vbQuestion, "Services")
        If lResult = vbYes Then
            Call StopServiceByName(UCase(colServices(liCurrentItem.Key).ServiceName))
        End If
    End If
    
    Set liCurrentItem = Nothing
    
End Sub
Private Sub cmdPause_Click()

    Dim liCurrentItem As ListItem
    Dim sMessage As String
    Dim lResult As Long
    
    Set liCurrentItem = lvServices.SelectedItem
    
    sMessage = "Are you sure you want to pause the " & colServices(liCurrentItem.Key).ServiceName & " service?"
    
    lResult = MsgBox(sMessage, vbYesNo, "Services")
    If lResult = vbYes Then
        Call PauseServiceByName(UCase(colServices(liCurrentItem.Key).ServiceName))
    End If
    
    Set liCurrentItem = Nothing

End Sub
Private Sub cmdContinue_Click()

    Dim liCurrentItem As ListItem
    
    Set liCurrentItem = lvServices.SelectedItem
    
    Call ContinueServiceByName(UCase(colServices(liCurrentItem.Key).ServiceName))

End Sub
Private Sub cmdStartup_Click()
    
    frmStartup.Show vbModal
    lvServices.SetFocus
    
End Sub
Private Sub lvServices_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call subSetupButtons(Item.Key)
End Sub
Private Sub lvServices_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then
        If Not lvServices.SelectedItem.Index = lvServices.ListItems.Count Then
            'If we are at the end of the list we should not move
            Call subSetupButtons(lvServices.ListItems(lvServices.SelectedItem.Index + 1).Key)
        End If
    End If
    
    If KeyCode = vbKeyUp Then
        If Not lvServices.SelectedItem.Index = 1 Then
            'If we are at the beginning of the list we should not move
            Call subSetupButtons(lvServices.ListItems(lvServices.SelectedItem.Index - 1).Key)
        End If
    End If
    
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub SetupServiceListView()

    'This routine configures the column headings for the Report view mode
    Dim lLVWidth As Long
    Dim LVHeader As ColumnHeader
    
    lLVWidth = lvServices.Width - 5 * Screen.TwipsPerPixelX
    
    Set LVHeader = lvServices.ColumnHeaders.Add(1, , "Name", lLVWidth * 0.55)
    Set LVHeader = lvServices.ColumnHeaders.Add(2, , "Status", lLVWidth * 0.2, vbCenter)
    Set LVHeader = lvServices.ColumnHeaders.Add(3, , "Startup", lLVWidth * 0.2, vbCenter)
    
    lvServices.Sorted = True
    lvServices.View = lvwReport
    lvServices.FullRowSelect = True
    lvServices.HideColumnHeaders = True
    lvServices.LabelEdit = lvwManual
    
    Set LVHeader = Nothing

End Sub
Private Sub GetServiceList()

    Dim bResult As Boolean
    
    bResult = EnumerateServices("", SERVICE_STATE_ALL, SERVICE_WIN32)
    If Not bResult Then
        'Something went wrong enumerating services
        End
    End If
    
End Sub
Private Sub StoreServicesInListView()

    Dim CurrentService As CServices
    Dim liCurrentItem As ListItem
    
    'Loop through the service collection and add each item to the list view
    For Each CurrentService In colServices
        'SubItems stores data that will be displayed when in Report view
        Set liCurrentItem = lvServices.ListItems.Add(, UCase(CurrentService.ServiceName), CurrentService.DisplayName)
        
        'The following two case statements fill in the subitem information and
        'change the forecolor of the service line depending on it's state.
        Select Case CurrentService.CurrentStatus
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
        
        Select Case CurrentService.StartType
            Case SERVICE_BOOT_START
                liCurrentItem.SubItems(2) = "Boot"
            Case SERVICE_SYSTEM_START
                liCurrentItem.SubItems(2) = "System"
            Case SERVICE_AUTO_START
                liCurrentItem.SubItems(2) = "Automatic"
            Case SERVICE_DEMAND_START
                liCurrentItem.SubItems(2) = "Manual"
            Case SERVICE_DISABLED
                liCurrentItem.SubItems(2) = "Disabled"
                liCurrentItem.ForeColor = vbGrayText
        End Select
    Next
    
End Sub
Private Sub subSetupButtons(sKey As String)

    'This sub routine sets up the buttons depending on the currently selected
    'servies status.
    cmdStart.Enabled = Not colServices(sKey).Running And Not CBool((colServices(sKey).StartType And SERVICE_DISABLED))
    cmdStop.Enabled = colServices(sKey).AcceptStop And colServices(sKey).Running
    cmdPause.Enabled = colServices(sKey).AcceptPause And colServices(sKey).Running And Not colServices(sKey).Paused
    cmdContinue.Enabled = colServices(sKey).Paused
    
End Sub
Private Sub StartServiceByName(sServiceName As String)
    
    Dim liCurrentItem As ListItem
    
    Set liCurrentItem = lvServices.ListItems(sServiceName)
    
    'Don't try and start the service if it's already started
    If Not colServices(sServiceName).Running Then
        'sHoldText = sbmain.Panels("MAIN").Text
        Me.MousePointer = vbHourglass
        
        'Start the selected service
        ServiceStart "", colServices(liCurrentItem.Key).ServiceName
        
        'Set global variables for Service Control dialog box
        SERVICE_CONTROL_MSGBOX_MESSAGE = "Attempting to start the " & colServices(liCurrentItem.Key).DisplayName & " service."
        SERVICE_CONTROL_WAITTIME = 3000
        SERVICE_CONTROL_MSGBOX_TITLE = "Service Control"
        SERVICE_CONTROL_MACHINENAME = ""
        SERVICE_CONTROL_SERVICENAME = colServices(liCurrentItem.Key).ServiceName
        SERVICE_CONTROL_STATUS_WANTED = SERVICE_RUNNING
        frmSCMsgBox.Show vbModal
        
        colServices(liCurrentItem.Key).CurrentStatus = SERVICE_RUNNING
        Call subSetupButtons(liCurrentItem.Key)
        Me.MousePointer = vbDefault
        
        lvServices.SetFocus
    End If
    Set liCurrentItem = Nothing

End Sub
Private Sub StopServiceByName(sServiceName As String)

    Dim liCurrentItem As ListItem

    Set liCurrentItem = lvServices.ListItems(sServiceName)
    
    'Don't try and start the service if it's already started
    If Not colServices(sServiceName).CurrentStatus = SERVICE_STOPPED Then
        Me.MousePointer = vbHourglass
        ServiceStop "", sServiceName
            
        'Set global variables for Service Control dialog box
        SERVICE_CONTROL_MSGBOX_MESSAGE = "Attempting to stop the " & colServices(liCurrentItem.Key).DisplayName & " service."
        SERVICE_CONTROL_WAITTIME = 3000
        SERVICE_CONTROL_MSGBOX_TITLE = "Service Control"
        SERVICE_CONTROL_MACHINENAME = ""
        SERVICE_CONTROL_SERVICENAME = sServiceName
        SERVICE_CONTROL_STATUS_WANTED = SERVICE_STOPPED
        frmSCMsgBox.Show vbModal
        
        'colServices(liCurrentItem.Key).CurrentStatus = SERVICE_RUNNING
        Call subSetupButtons(liCurrentItem.Key)
        Me.MousePointer = vbDefault
        
        lvServices.SetFocus
    End If
    Set liCurrentItem = Nothing
    
End Sub
Private Sub PauseServiceByName(sServiceName As String)

    Dim liCurrentItem As ListItem

    Set liCurrentItem = lvServices.ListItems(sServiceName)

    'The service must be running to pause it
    If colServices(sServiceName).CurrentStatus = SERVICE_RUNNING Then
        Me.MousePointer = vbHourglass
        ServicePause "", sServiceName
            
        'Set global variables for Service Control dialog box
        SERVICE_CONTROL_MSGBOX_MESSAGE = "Attempting to pause the " & colServices(liCurrentItem.Key).DisplayName & " service."
        SERVICE_CONTROL_WAITTIME = 3000
        SERVICE_CONTROL_MSGBOX_TITLE = "Service Control"
        SERVICE_CONTROL_MACHINENAME = ""
        SERVICE_CONTROL_SERVICENAME = sServiceName
        SERVICE_CONTROL_STATUS_WANTED = SERVICE_PAUSED
        frmSCMsgBox.Show vbModal
        
        Call subSetupButtons(liCurrentItem.Key)
        Me.MousePointer = vbDefault
        
        lvServices.SetFocus
    End If
    Set liCurrentItem = Nothing

End Sub
Private Sub ContinueServiceByName(sServiceName As String)

    Dim liCurrentItem As ListItem

    Set liCurrentItem = lvServices.ListItems(sServiceName)

    'The service must be paused to continue it
    If colServices(sServiceName).CurrentStatus = SERVICE_PAUSED Then
        Me.MousePointer = vbHourglass
        ServiceContinue "", sServiceName
            
        'Set global variables for Service Control dialog box
        SERVICE_CONTROL_MSGBOX_MESSAGE = "Attempting to continue the " & colServices(liCurrentItem.Key).DisplayName & " service."
        SERVICE_CONTROL_WAITTIME = 3000
        SERVICE_CONTROL_MSGBOX_TITLE = "Service Control"
        SERVICE_CONTROL_MACHINENAME = ""
        SERVICE_CONTROL_SERVICENAME = sServiceName
        SERVICE_CONTROL_STATUS_WANTED = SERVICE_RUNNING
        frmSCMsgBox.Show vbModal
        
        Call subSetupButtons(liCurrentItem.Key)
        Me.MousePointer = vbDefault
        
        lvServices.SetFocus
    End If
    Set liCurrentItem = Nothing

End Sub

