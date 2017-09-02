Attribute VB_Name = "modGlobals"
Option Explicit

Public Running As Boolean
Public Player_HighIndex As Long

Public PacketsRecieved(1 To CMSG_COUNT - 1) As Long
Public PacketsSent(1 To SMSG_COUNT - 1) As Long
