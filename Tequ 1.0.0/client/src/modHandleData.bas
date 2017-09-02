Attribute VB_Name = "modHandleData"
Option Explicit

Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(SEnterGame) = GetAddress(AddressOf HandleEnterGame)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
End Sub

Sub HandleData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        DestroyGame
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.Length), 0, 0
    
End Sub

Public Sub HandleEnterGame(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MyIndex = Buffer.ReadLong
    Set Buffer = Nothing

    Call EnterGame

End Sub

Public Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim mapNum As Long
Dim X As Long, Y As Long, L As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    mapNum = Buffer.ReadLong
    Map(mapNum).Name = Buffer.ReadString

    Set Buffer = Nothing

End Sub

Public Sub HandlePlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim PlayerIndex As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PlayerIndex = Buffer.ReadLong
    Player(PlayerIndex).Name = Buffer.ReadString
    Player(PlayerIndex).Map = Buffer.ReadLong
    Player(PlayerIndex).X = Buffer.ReadLong
    Player(PlayerIndex).Y = Buffer.ReadLong
    Player(PlayerIndex).Sprite = Buffer.ReadLong
    Set Buffer = Nothing
    
End Sub


