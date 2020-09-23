VERSION 5.00
Begin VB.Form frmPASV 
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPASV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents FormSocket As cFormSocket
Attribute FormSocket.VB_VarHelpID = -1
Private Sub Form_Load()

Set FormSocket = New cFormSocket
FormSocket.hWnd = Me.hWnd

FormSocket.Hook

End Sub
Function Retrieve(pHOST As String, pPort As String)
    Me.Show
    
    FormSocket.sckConnect pHOST, pPort, Me.hWnd
    FormSocket.sckSendData "MODE S" & vbCrLf
    
    Open App.Path & "\" & "tmpdata.tmp" For Binary As 1
    FormSocket.sckSendData "RETR /recvq/fax0014" & vbCrLf

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Close 1
FormSocket.sckClose
FormSocket.UnHook

End Sub

Private Sub FormSocket_sckBinaryData(pDATA As Variant, hWnd As Long)

Put #1, , pDATA

End Sub

