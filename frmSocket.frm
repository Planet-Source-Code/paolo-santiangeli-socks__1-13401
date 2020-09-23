VERSION 5.00
Begin VB.Form frmSocket 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Socket"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Log"
      Height          =   3255
      Left            =   3240
      TabIndex        =   19
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtLog 
         Height          =   2685
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   360
         Width           =   5655
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   3015
      Begin VB.CommandButton Command2 
         Caption         =   "SEND >>"
         Height          =   375
         Left            =   1680
         TabIndex        =   21
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtCMD 
         Height          =   285
         Left            =   600
         TabIndex        =   17
         Text            =   "USER anonymous"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "CMD:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   3015
      Begin VB.TextBox txtWaitFor 
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Wait For"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "CMD Status"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbStatus 
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.TextBox txtLastAnswer 
      Height          =   1395
      Left            =   0
      TabIndex        =   9
      Top             =   5280
      Width           =   10065
   End
   Begin VB.TextBox txtLastCommand 
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton Command1 
         Caption         =   "Connect"
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Text            =   "21"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtHost 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Text            =   "ftp.intel.com"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Host"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   4920
      Width           =   5415
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Last CMD"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   4560
      Width           =   1545
   End
End
Attribute VB_Name = "frmSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents FormSocket As cFormSocket
Attribute FormSocket.VB_VarHelpID = -1

Private Sub Command1_Click()
Dim aRes As Integer
    
    If Command1.Caption = "Connect" Then
                Me.Caption = "Server:" & txtServer & " Port:" & txtPort
            
                aRes = FormSocket.sckConnect(txtHost.Text, txtPort.Text, Me.hWnd)
                Command1.Caption = "Disconnect"
    Else
                    
                txtLog.Text = ""
                Command1.Caption = "Connect"
    End If


End Sub
Private Sub Command2_Click()
Dim a As Variant

    If txtWaitFor = "" Then
    a = FormSocket.sckSendData(txtCMD.Text)
    Else
    
    a = FormSocket.sckSendCommand(txtCMD.Text, txtWaitFor.Text)
    
        If FormSocket.LastCommandOK Then
           
           lbStatus.Caption = "Command OK"
        
        Else
           
           lbStatus.Caption = "Command Error"
        
        End If
        
    
    End If

End Sub


Private Sub Form_Load()
      
      Set FormSocket = New cFormSocket
      FormSocket.sckHook Me.hWnd

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
        FormSocket.sckClose
        FormSocket.UnHook

        Set FormSocket = Nothing

End Sub

Private Sub FormSocket_sckData(pCMD As String, pDATA As String, hWnd As Long)

If hWnd = Me.hWnd Then
    
    txtLastCommand = pCMD
    txtLastAnswer = pDATA

End If

End Sub

Private Sub FormSocket_sckLog(pLogEntry As String, hWnd As Long)
    
    If hWnd = Me.hWnd Or hWnd = -1 Then
      
      txtLog = txtLog & pLogEntry & vbCrLf
            
            If Left(pLogEntry, 3) = "CLS" Then
                Call Command1_Click
            End If
            
    End If
    
End Sub

