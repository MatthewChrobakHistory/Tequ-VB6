Attribute VB_Name = "modEnumerations"
Option Explicit

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

' Vitals
Public Enum Vitals
    Health = 1
    Spirit
    Vital_Count
End Enum

' Stats
Public Enum Stats
    Attack = 1
    Strength
    Defense
    Agility
    Sagacity
    Stat_Count
End Enum

Public Enum Equipment
    Head = 1
    Body
    Legs
    Shield
    Weapon
    Equipment_Count
End Enum

Public Enum Skills
    Woodcutting = 1
    Mining
    Fishing
    Smithing
    Cooking
    Fletching
    Crafting
    PotionBrewing
    Skill_Count
End Enum


Public HandleDataSub(CMSG_COUNT) As Long

