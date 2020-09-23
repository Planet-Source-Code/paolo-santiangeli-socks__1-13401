VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFormSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private WithEvents mlog As cLog
Attribute mlog.VB_VarHelpID = -1
Private WithEvents mSocketData As cSocketData
Attribute mSocketData.VB_VarHelpID = -1

Event sckLog(pLogEntry As String, hWnd As Long)
Event sckData(pCMD As String, pDATA As String, hWnd As Long)
Event sckBinaryData(pDATA As Variant, hWnd As Long)

Private mConnected As Boolean

Private mWaitForData As Boolean
Private mWaitedData As String
Private mCommandOK As Boolean

Private m_hWnd As Long
Private m_sckError As Boolean
Private m_lpPrevWndProc As Long
Property Get LastCommandOK() As Boolean

    LastCommandOK = mCommandOK

End Property

Public Function sckConnect(m_host, m_port, hWnd As Long) As Integer

    Dim temp, TimeOut As Variant
    
    'Progress = 0
    
    m_sckError = False
    
    'TimeOut = Timer + 60
    
    gLog.Log "Will timeout in 60 seconds", m_hWnd
    
    gRegistredForms(Hex(m_hWnd)).LastCommand = "Connecting"

    'make sure the port is closed!
    If gRegistredForms(Hex(m_hWnd)).SocketPointer <> 0 Then Call closesocket(gRegistredForms(Hex(m_hWnd)).SocketPointer)

    'let's connect!!!       host            port       handle
    temp = ConnectSock(m_host, m_port, 0, hWnd, True)
    
    gLog.Log "Connect socket return value " & temp, m_hWnd
    gLog.Log "Connecting to " & m_host & " at port " & m_port, m_hWnd
        
    'Winsock Error
    If temp = INVALID_SOCKET Then
        gLog.Log "Error -Invalid Socket", m_hWnd
        sckConnect = -1
        Exit Function
    End If
    
    sckConnect = 1

End Function
Function sckSendCommand(pServerCMD As String, pWaitFor As String)
    
    gLog.Log "CMD Sent: " & pServerCMD, m_hWnd
    gRegistredForms(Hex(m_hWnd)).LastCommand = pServerCMD
    
    mCommandOK = False
    mWaitForData = True
    mWaitedData = pWaitFor
    
    Call SendData(gRegistredForms(Hex(m_hWnd)).SocketPointer, pServerCMD & vbCrLf)
    
    While mWaitForData
        
        DoEvents
        
    Wend
    
    sckSendCommand = CMDAnswer

End Function
Function sckSendData(pServerCMD As String) As String

    gLog.Log "CMD: " & pServerCMD, m_hWnd
    gRegistredForms(Hex(m_hWnd)).LastCommand = pServerCMD
    
    Call SendData(gRegistredForms(Hex(m_hWnd)).SocketPointer, pServerCMD & vbCrLf)
    
    While DataInBuffer = False
        
        DoEvents
        
    Wend
    
    sckSendData = CMDAnswer
    
End Function
Public Sub sckHook(ByVal pHwnd As Long)
    Dim lpPrevWndProc As Long
    
    m_hWnd = pHwnd
    
    lpPrevWndProc = GetWindowLong(m_hWnd, GWL_WNDPROC)
    HelpPrevWndProc = lpPrevWndProc
    
    'Add Info to the global  Register Class
    gRegistredForms.Add lpPrevWndProc, "Handle-" & CStr(m_hWnd), m_hWnd, Hex(m_hWnd)
    
    lpPrevWndProc = SetWindowLong(m_hWnd, GWL_WNDPROC, AddressOf WindowProc)
    
End Sub
Public Sub UnHook()
    'if we dont un-subclass before we shutdown
    'the program, we get an illigal procedure error.
    'fun.
    Call SetWindowLong(m_hWnd, GWL_WNDPROC, lpPrevWndProc)
    
    gRegistredForms.Remove Hex(m_hWnd)
    
End Sub
Sub sckClose()

    Call closesocket(gRegistredForms(Hex(m_hWnd)).SocketPointer)

End Sub

Private Sub Class_Initialize()
    'ok, we have to start winsock, DUH!
    'Global from WinsockAPI
    
    Set mlog = gLog
    Set mSocketData = gSocketData
    
    Call StartWinsock("")

End Sub
Private Sub mlog_NewLogLine(pLogEntry As String, hWnd As Long)
        
    RaiseEvent sckLog(pLogEntry, hWnd)

End Sub

Private Sub mSocketData_BinaryData(ByVal pDATA As Variant, ByVal hWnd As Long)
    
    If mWaitForData = True Then
    
        If InStrB(1, pDATA, mWaitedData, vbBinaryCompare) Then
        
            mCommandOK = True
        
        End If
    
            
    mWaitForData = False
    
    End If
    
    RaiseEvent sckBinaryData(pDATA, hWnd)

End Sub

Private Sub mSocketData_SocketData(pCMD As String, pDATA As String, hWnd As Long)

    RaiseEvent sckData(pCMD, pDATA, hWnd)

End Sub
