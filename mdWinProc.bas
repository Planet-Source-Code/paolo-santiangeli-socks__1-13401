Attribute VB_Name = "mdWinProc"
Option Explicit

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'**********************************
Public DataInBuffer As Boolean
Public e_err As Variant
Public e_errstr As Variant
'**********************************

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim sckCurrentCommand As String

Dim x As Long
Dim wp As Integer
Dim temp As Variant
Dim ReadBuffer(1000) As Byte
Dim DataFlag As Boolean
Dim DataBuffer As Variant
Dim plpPrevWndProc As Long



'Debug.Print uMsg, wParam, lParam
    
    Select Case uMsg
        
        Case 1025:
            
            e_err = WSAGetAsyncError(lParam)
            e_errstr = GetWSAErrorString(e_err)
            
            If e_err <> 0 Then
            '********** Error *********
                gLog.Log "Error String returned -> " & e_err & " - " & e_errstr, hw
                gLog.Log "Terminating....", hw
               
                'Exit Function
            
            End If
            
            Select Case lParam
            
                Case FD_READ: 'lets check for data
                            
                        If gRegistredForms(Hex(hw)).LastCommand <> "" Then
                           'Command Pending
                           sckCurrentCommand = gRegistredForms(Hex(hw)).LastCommand
                           gRegistredForms(Hex(hw)).LastCommand = ""
                        
                        ElseIf gRegistredForms(Hex(hw)).LastCommand = "DATA" Then
                               DataFlag = True
                        End If
                        

                        x = recv(gRegistredForms(Hex(hw)).SocketPointer, ReadBuffer(0), 1000, 0) 'try to get some
                        
                        If x > 0 Then 'was there any?
                            
                            ReadFlag = False
                            
                            RecvBuffer = StrConv(ReadBuffer, vbUnicode) 'yep, lets change it to stuff we can understand
                            
                            gLog.Log RecvBuffer, hw
                            gSocketData.Log sckCurrentCommand, RecvBuffer, hw
                            
                            gSocketData.SetData RecvBuffer, hw
                            
                            'rtncode = Mid(RecvBuffer, 1, 3)
                            DataInBuffer = True
                        
                        End If
                
                Case FD_CONNECT: 'did we connect?
                        
                      gLog.Log "Connection Established... :" & lParam, hw
                      gRegistredForms(Hex(hw)).SocketPointer = wParam 'yep, we did! yayay

                Case FD_OOB:
                                                
                      'gSocketData.SetData(recv  , hw
                        
                Case FD_CLOSE: 'uh oh. they closed the connection
                      
                    Call closesocket(wp)   'so we need to close
                    gLog.Log "CLS-Connection Closed By Peer", hw
            
            End Select
    
    End Select
    
    plpPrevWndProc = gRegistredForms(Hex(hw)).PreviousWinProc
    'let the msg get through to the form
    
    WindowProc = CallWindowProc(plpPrevWndProc, hw, uMsg, wParam, lParam)

End Function
