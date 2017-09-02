VERSION 5.00
Begin VB.Form frmPacketViewer 
   Caption         =   "Form1"
   ClientHeight    =   8955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstSIndex 
      Height          =   8640
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.ListBox lstCIndex 
      Height          =   8640
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmPacketViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub RefreshList()
Dim i As Long
Dim ListSpot As Long

ListSpot = lstCIndex.ListIndex
lstCIndex.Clear

For i = 1 To ClientPackets.CMSG_COUNT - 1
    lstCIndex.AddItem (GetPacketName(True, i) & " received: x" & PacketsRecieved(i))
Next

lstCIndex = ListSpot

ListSpot = lstSIndex.ListIndex
lstSIndex.Clear
For i = 1 To ServerPackets.SMSG_COUNT - 1
    lstSIndex.AddItem (GetPacketName(False, i) & " sent: x" & PacketsSent(i))
Next
    
lstSIndex = ListSpot
    
End Sub

Public Function GetPacketName(ByVal ClientPackets As Boolean, ByVal Num As Long) As String

Select Case ClientPackets
    Case True
        Select Case Num
            Case CRequestLogin
                GetPacketName = "CRequestLogin"
            Case CCreatePlayer
                GetPacketName = "CCreatePlayer"
        End Select
    Case False
        Select Case Num
            Case SClientMsgBox
                GetPacketName = "SClientMsgBox"
            Case SPlayerData
                GetPacketName = "SClientData"
            Case SPlayerMyIndex
                GetPacketName = "SPlayerMyIndex"
            Case SEnterGame
                GetPacketName = "SEnterGame"
            Case SMapData
                GetPacketName = "SMapData"
        End Select
End Select
            
End Function


