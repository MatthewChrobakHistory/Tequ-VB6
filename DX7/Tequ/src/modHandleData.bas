Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(SClientMsgBox) = GetAddress(AddressOf HandleClientMsgBox)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMyIndex) = GetAddress(AddressOf HandlePlayerMyIndex)
    HandleDataSub(SEnterGame) = GetAddress(AddressOf HandleEnterGame)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
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

Public Sub HandleClientMsgBox(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Text As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Text = Buffer.ReadString()
    MsgBox Text
    Set Buffer = Nothing
End Sub

Public Sub HandlePlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim PlayerIndex As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PlayerIndex = Buffer.ReadLong
    With Player(PlayerIndex)
        .Name = Buffer.ReadString
        '.Password = buffer.ReadString         REMOVED FOR POSSIBLE SECURITY ISSUES
        .Map = Buffer.ReadLong
    End With
    Set Buffer = Nothing
End Sub

Public Sub HandlePlayerMyIndex(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MyIndex = Buffer.ReadLong
    Set Buffer = Nothing
End Sub

Public Sub HandleEnterGame(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Call EnterGame
    
End Sub

Public Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim MapNum As Long, i As Long, X As Long
Dim ToF As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MapNum = Buffer.ReadLong
    Map(MapNum).Name = Buffer.ReadString
    Map(MapNum).Picture = Buffer.ReadLong
    For i = 1 To 20
        ToF = Buffer.ReadByte
        If ToF = 1 Then
            Map(MapNum).Label(i).Visible = True
        Else
            Map(MapNum).Label(i).Visible = False
        End If
        Map(MapNum).Label(i).Width = Buffer.ReadLong
        Map(MapNum).Label(i).Height = Buffer.ReadLong
        Map(MapNum).Label(i).Left = Buffer.ReadLong
        Map(MapNum).Label(i).Top = Buffer.ReadLong
        Map(MapNum).Label(i).Caption = Buffer.ReadString
        Map(MapNum).Label(i).Event.Type = Buffer.ReadLong
        For X = 1 To 5
            Map(MapNum).Label(i).Event.Data(X) = Buffer.ReadLong
            Map(MapNum).Label(i).Event.Text(X) = Buffer.ReadString
        Next
    Next
    Set Buffer = Nothing
    If MapNum = 100 Then MsgBox "SUCESS"
End Sub
