Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CRequestLogin) = GetAddress(AddressOf HandleRequestLogin)
    HandleDataSub(CCreatePlayer) = GetAddress(AddressOf HandleCreatePlayer)
End Sub

Sub HandleData(ByVal index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long
        
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        Exit Sub
    End If
    
    If MsgType >= CMSG_COUNT Then
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), index, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Public Sub HandleRequestLogin(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Name As String, Password As String
Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    'lets check the login details. First, clear the data, and then load it.
    Name = Buffer.ReadString
    Password = Buffer.ReadString
    If FileExist(App.Path & "\data\players\" & Name & ".bin") = True Then
        Call LoadLDPlayer(Name)
        If Trim$(LDPlayer.Password) = Password Then
            Call LoadPlayer(index, Name)
            TempPlayer(index).InGame = True
            Call ClearLDPlayer
            Call SendGameDataTo(index)
            Call SendEnterGame(index)
        End If
    End If
    Set Buffer = Nothing
    
End Sub

Public Sub HandleCreatePlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Name As String, Password As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Name = Buffer.ReadString
    Password = Buffer.ReadString
    
    If FileExist(App.Path & "\data\players\" & Name & ".bin") = False Then
        Call MakePlayer(index, Name, Password)
        TempPlayer(index).InGame = True
        Call SendGameDataTo(index)
        Call SendEnterGame(index)
    End If
    Set Buffer = Nothing
    
End Sub

