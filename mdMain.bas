Attribute VB_Name = "mdMain"
Option Explicit
Global Const MAXSOCKETFRM = 3

Public gLog As New cLog
Public gSocketData As New cSocketData
Public gRegistredForms As New cFormsSockets

Public Const GWL_WNDPROC = -4
Public HelpPrevWndProc As Long

'Public CMDAnswer As String
'Public CMDExec As Boolean
'Public mySock As Long
Public RecvBuffer As String
Public TimeOut As Variant
Public ReadFlag As Boolean
