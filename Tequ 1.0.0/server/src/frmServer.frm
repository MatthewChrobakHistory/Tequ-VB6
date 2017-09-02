VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmServer 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   287
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   543
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   195
      Left            =   5280
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   5400
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ListBox lstIndex 
      Height          =   2010
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   4215
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6800
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblErrorNotification 
      BackStyle       =   0  'Transparent
      Caption         =   "An error just occured. Click here to view it."
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)

    Call DestroyServer

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
