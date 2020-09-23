VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmStartup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service - Under construction still"
   ClientHeight    =   4650
   ClientLeft      =   6210
   ClientTop       =   6345
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   4230
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame2 
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   1995
      Width           =   3990
      _Version        =   65536
      _ExtentX        =   7038
      _ExtentY        =   4471
      _StockProps     =   14
      Caption         =   "Log On As:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSCommand cmdAccountLookup 
         Height          =   330
         Left            =   3585
         TabIndex        =   18
         Top             =   1005
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin VB.TextBox txtConfirmPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1410
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   1920
         Width           =   2460
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1410
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   1455
         Width           =   2460
      End
      Begin VB.TextBox txtAccountName 
         Height          =   315
         Left            =   1410
         TabIndex        =   13
         Top             =   1020
         Width           =   2175
      End
      Begin Threed.SSCheck chkInteractive 
         Height          =   300
         Left            =   345
         TabIndex        =   12
         Top             =   540
         Width           =   3210
         _Version        =   65536
         _ExtentX        =   5662
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "A&llow Service to Interact with Desktop"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optSystem 
         Height          =   300
         Left            =   105
         TabIndex        =   10
         Top             =   270
         Width           =   2085
         _Version        =   65536
         _ExtentX        =   3678
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&System Account"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption optAccount 
         Height          =   300
         Left            =   105
         TabIndex        =   11
         Top             =   1020
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&This Account:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Confirm Password:"
         Height          =   435
         Left            =   345
         TabIndex        =   15
         Top             =   1860
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   270
         Left            =   345
         TabIndex        =   14
         Top             =   1485
         Width           =   870
      End
   End
   Begin Threed.SSCommand cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   3015
      TabIndex        =   6
      Top             =   510
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "OK"
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1440
      Left            =   105
      TabIndex        =   0
      Top             =   420
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   2540
      _StockProps     =   14
      Caption         =   "Startup Type"
      Begin Threed.SSOption optAutomatic 
         Height          =   300
         Left            =   105
         TabIndex        =   1
         Top             =   300
         Width           =   2085
         _Version        =   65536
         _ExtentX        =   3678
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&Automatic"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optManual 
         Height          =   300
         Left            =   105
         TabIndex        =   2
         Top             =   645
         Width           =   2085
         _Version        =   65536
         _ExtentX        =   3678
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&Manual"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optDisabled 
         Height          =   300
         Left            =   105
         TabIndex        =   3
         Top             =   975
         Width           =   2085
         _Version        =   65536
         _ExtentX        =   3678
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&Disabled"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSCommand cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3015
      TabIndex        =   7
      Top             =   990
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Cancel"
   End
   Begin Threed.SSCommand cmdHelp 
      Height          =   375
      Left            =   3015
      TabIndex        =   8
      Top             =   1485
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Help"
      Enabled         =   0   'False
   End
   Begin VB.Label lblServiceName 
      Height          =   270
      Left            =   795
      TabIndex        =   5
      Top             =   60
      Width           =   3315
   End
   Begin VB.Label Label1 
      Caption         =   "Service:"
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   660
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub Form_Load()

    'All the information displayed depends on the currently selected item
    'in the list view.  We must determine what this item is so we can
    'lookup all appropriate information from the corresponding class object
    'stored in the collection.
    
    Dim CurrentItem As ListItem
    
    Set CurrentItem = frmMain.lvServices.SelectedItem
    
    lblServiceName.Caption = colServices(CurrentItem.Key).DisplayName
    
    If (colServices(CurrentItem.Key).ServiceType And SERVICE_WIN32_SHARE_PROCESS) Then
        'A share process service logs in a special way and the user can not
        'change how it logs in.  Therefore we need to disable a few things.
        optSystem.Enabled = False
        optAccount.Enabled = False
        
        cmdAccountLookup.Enabled = False
        
        txtAccountName.Enabled = False
        txtAccountName.BackColor = Me.BackColor
        txtPassword.Enabled = False
        txtPassword.BackColor = Me.BackColor
        txtConfirmPassword.Enabled = False
        txtConfirmPassword.BackColor = Me.BackColor
        
        Label2.Enabled = False
        Label3.Enabled = False
    End If
    
    Debug.Print "Service Name:  " & colServices(CurrentItem.Key).ServiceName
    Debug.Print "Display Name:  " & colServices(CurrentItem.Key).DisplayName
    Debug.Print "Service Type:  " & colServices(CurrentItem.Key).ServiceType
    Debug.Print "Controls Accepted:  " & colServices(CurrentItem.Key).ControlsAccepted
    Debug.Print "CurrentStatus:  " & colServices(CurrentItem.Key).CurrentStatus
    Debug.Print "Start Type:  " & colServices(CurrentItem.Key).StartType
    Debug.Print "Error Control:  " & colServices(CurrentItem.Key).ErrorControl
    Debug.Print "Service Start Name:  " & colServices(CurrentItem.Key).StartName
    
End Sub
Private Sub Form_Activate()

    Dim CurrentItem As ListItem
    
    Set CurrentItem = frmMain.lvServices.SelectedItem

    Select Case colServices(CurrentItem.Key).StartType
        Case SERVICE_AUTO_START
            optAutomatic.Value = True
        Case SERVICE_DEMAND_START
            optManual.Value = True
        Case SERVICE_DISABLED
            optDisabled.Value = True
    End Select
    
    chkInteractive.Value = (colServices(CurrentItem.Key).ServiceType And SERVICE_INTERACTIVE_PROCESS)

End Sub
Private Sub cmdHelp_Click()
    MsgBox "This button is not currently activated", vbOKOnly + vbInformation, "DEMO VERSION"
End Sub
Private Sub cmdOK_Click()
    MsgBox "This button is not currently activated", vbOKOnly + vbInformation, "DEMO VERSION"
End Sub
Private Sub cmdAccountLookup_Click()
    MsgBox "This button is not currently activated", vbOKOnly + vbInformation, "DEMO VERSION"
End Sub
