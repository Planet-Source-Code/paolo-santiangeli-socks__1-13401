VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSocketData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event SocketData(pCMD As String, pDATA As String, hWnd As Long)
Event BinaryData(ByVal pDATA As Variant, ByVal hWnd As Long)

Public Sub Log(ByVal sCMD As String, ByVal sDATA As String, Optional pHwnd As Long = -1)
    
    RaiseEvent SocketData(sCMD, sDATA, pHwnd)

End Sub
Public Sub SetData(ByVal sDATA As Variant, Optional pHwnd As Long = -1)
    
    RaiseEvent BinaryData(sDATA, pHwnd)

End Sub

