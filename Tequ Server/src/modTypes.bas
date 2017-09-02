Attribute VB_Name = "modTypes"
Option Explicit

Public LoginPlayer As PlayerRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Map(1 To MAX_MAPS) As MapRec

Private Type EventRec
    Type As Long
    Data(1 To 5) As Long
    Text(1 To 5) As String * 50
End Type

Private Type LabelRec
    Left As Long
    Top As Long
    Width As Long
    Height As Long
    Caption As String * 50
    Visible As Boolean
    Event As EventRec
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Picture As Long
    Label(1 To 20) As LabelRec
End Type

Private Type PlayerRec
    Name As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    Map As Long
End Type

Private Type TempPlayerRec
    Buffer As clsBuffer
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    DoneLoadingData As Boolean
End Type
