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
    If FileExist(App.Path & "/data/players/" & Name & ".bin") = True Then
        Call LoadLoginPlayer(Trim$(Name))
        If Trim$(LoginPlayer.Password) = Trim$(Password) Then
            Call LoadPlayer(index, Name)
            Call SendPlayerMyIndex(index)
            Call SendGameData(index)
            Call SendEnterGame(index)
        Else
            Call SendClientMsgBox(index, "Login failed!")
        End If
    Else
        Call SendClientMsgBox(index, "Login failed!")
    End If
    Set Buffer = Nothing
    PacketsRecieved(CRequestLogin) = PacketsRecieved(CRequestLogin) + 1
End Sub

Public Sub HandleCreatePlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Name As String, Password As String
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Name = Trim$(Buffer.ReadString)
    Password = Trim$(Buffer.ReadString)
    
    If FileExist(App.Path & "/data/players/" & Name & ".bin") = True Then
        Call SendClientMsgBox(index, "Player already exists!")
        Exit Sub
    Else
        Call MakePlayer(index, Name, Password)
        Call SendPlayerMyIndex(index)
        Call SendGameData(index)
        Call SendEnterGame(index)
        For i = 1 To Player_HighIndex
            If i <> index Then Call SendPlayerData(i, index)
        Next
    End If
    
    PacketsRecieved(CCreatePlayer) = PacketsRecieved(CCreatePlayer) + 1
End Sub
