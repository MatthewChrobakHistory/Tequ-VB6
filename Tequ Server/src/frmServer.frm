VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   6720
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.ListBox lstindex 
      Height          =   1620
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
   Begin MSWinsockLib.Winsock socket 
      Index           =   0
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

frmPacketViewer.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Long

For i = 1 To Player_HighIndex
    Call SavePlayer(i, Trim$(Player(i).Name))
Next

For i = 1 To MAX_MAPS
    Call SaveMap(i)
Next

For i = 1 To MAX_PLAYERS
    Call Unload(socket(i))
Next
socket(0).Close
Running = False
End

End Sub

Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Call AcceptConnection(index, requestID)
End Sub

Private Sub Socket_Accept(index As Integer, SocketId As Integer)
    Call AcceptConnection(index, SocketId)
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)

    If IsConnected(index) Then
        Call IncomingData(index, bytesTotal)
    End If

End Sub

Private Sub Socket_Close(index As Integer)
    Call CloseSocket(index)
End Sub

