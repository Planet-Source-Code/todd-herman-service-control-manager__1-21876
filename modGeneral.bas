Attribute VB_Name = "modGeneral"
Public SERVICE_CONTROL_MSGBOX_MESSAGE As String
Public SERVICE_CONTROL_WAITTIME As Long
Public SERVICE_CONTROL_MSGBOX_TITLE As String
Public SERVICE_CONTROL_SERVICENAME As String
Public SERVICE_CONTROL_MACHINENAME As String
Public SERVICE_CONTROL_STATUS_WANTED As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Sub AlwaysOnTop(lWindowHandle As Long)

    SetWindowPos lWindowHandle, -1, 0, 0, 0, 0, 1 Or 2
    
End Sub

