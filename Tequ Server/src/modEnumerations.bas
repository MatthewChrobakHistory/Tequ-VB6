Attribute VB_Name = "modEnumerations"
Option Explicit

Public Enum ServerPackets
    SClientMsgBox = 1
    SPlayerData
    SPlayerMyIndex
    SEnterGame
    SMapData
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CRequestLogin = 1
    CCreatePlayer
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

Public HandleDataSub(CMSG_COUNT) As Long
