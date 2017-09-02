Attribute VB_Name = "modEnumerations"
Option Explicit

' The order of the packets must match with the server's packet enumeration

' Packets sent by server to client
Public Enum ServerPackets
    SEnterGame = 1
    SPlayerIndex
    SMsgBox
    SPlayerData
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

' Map Layers
Public Enum Layers
    Ground = 1
    Mask1
    Mask2
    Mask3
    MaskAnim
    Fringe1
    Fringe2
    Fringe3
    FringeAnim
    Layer_Count
End Enum

Public HandleDataSub(SMSG_COUNT) As Long
