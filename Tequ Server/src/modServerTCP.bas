Attribute VB_Name = "modServerTCP"
Option Explicit

Public Sub SendGameData(ByVal index As Long)

TempPlayer(index).DoneLoadingData = False

End Sub

Sub SendDataTo(ByVal index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim TempData() As Byte

    If IsConnected(index) Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
              
        frmServer.socket(index).SendData Buffer.ToArray()
    End If
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex
        Call SendDataTo(i, Data)
    Next

End Sub

Function FindOpenPlayerSlot() As Long
    Dim i As Long
    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Function IsConnected(ByVal index As Long) As Boolean

    If frmServer.socket(index).State = sckConnected Then
        IsConnected = True
    End If

End Function

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.socket(i).Close
            frmServer.socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If

End Sub

Sub SocketConnected(ByVal index As Long)
Dim i As Long

    If index <> 0 Then
        ' re-set the high index
        Player_HighIndex = 0
        For i = MAX_PLAYERS To 1 Step -1
            If IsConnected(i) Then
                Player_HighIndex = i
                Exit For
            End If
        Next
        ' send the new highindex to all logged in players
        'SendHighIndex
    End If
    
    frmServer.lstindex.Clear
    If Player_HighIndex = 0 Then Exit Sub
    For i = 1 To Player_HighIndex
        frmServer.lstindex.AddItem i & ": " & frmServer.socket(i).RemoteHostIP
    Next
    
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long
Set TempPlayer(index).Buffer = New clsBuffer

        ' Check for data flooding
        If TempPlayer(index).DataBytes > 1000 Then
            Exit Sub
        End If
    
        ' Check for packet flooding
        If TempPlayer(index).DataPackets > 25 Then
            Exit Sub
        End If
            
    ' Check if elapsed time has passed
    TempPlayer(index).DataBytes = TempPlayer(index).DataBytes + DataLength
    If timeGetTime >= TempPlayer(index).DataTimer Then
        TempPlayer(index).DataTimer = timeGetTime + 1000
        TempPlayer(index).DataBytes = 0
        TempPlayer(index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.socket(index).GetData Buffer(), vbUnicode, DataLength
    TempPlayer(index).Buffer.WriteBytes Buffer()
    
    If TempPlayer(index).Buffer.Length >= 4 Then
        pLength = TempPlayer(index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(index).Buffer.Length - 4
        If pLength <= TempPlayer(index).Buffer.Length - 4 Then
            TempPlayer(index).DataPackets = TempPlayer(index).DataPackets + 1
            TempPlayer(index).Buffer.ReadLong
            HandleData index, TempPlayer(index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(index).Buffer.Length >= 4 Then
            pLength = TempPlayer(index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(index).Buffer.Trim
End Sub

Sub CloseSocket(ByVal index As Long)
Dim i As Long

    If index > 0 Then
        frmServer.socket(index).Close
        TempPlayer(index).InGame = False
    End If
    
    ' re-set the high index
    Player_HighIndex = 0
    For i = MAX_PLAYERS To 1 Step -1
        If IsConnected(i) Then
            Player_HighIndex = i
            Call SendPlayerMyIndex(i)
            Exit For
        End If
    Next
    
    frmServer.lstindex.Clear
    If Player_HighIndex = 0 Then Exit Sub
    For i = 1 To Player_HighIndex
        frmServer.lstindex.AddItem i & ": " & frmServer.socket(i).RemoteHostIP
    Next

End Sub
 
Sub SendClientMsgBox(ByVal index As Long, ByVal Text As String)
Dim Buffer As clsBuffer

Set Buffer = New clsBuffer
Buffer.WriteLong SClientMsgBox
Buffer.WriteString Text
SendDataTo index, Buffer.ToArray()
Set Buffer = Nothing

PacketsSent(SClientMsgBox) = PacketsSent(SClientMsgBox) + 1

End Sub

Sub SendPlayerData(ByVal SendIndex As Long, ByVal DatIndex As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong DatIndex
    Buffer.WriteString Trim$(Player(DatIndex).Name)
    Buffer.WriteLong Player(DatIndex).Map
    SendDataTo SendIndex, Buffer.ToArray()
    Set Buffer = Nothing
    PacketsSent(SPlayerData) = PacketsSent(SPlayerData) + 1
End Sub

Sub SendPlayerMyIndex(ByVal index As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerMyIndex
    Buffer.WriteLong index
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
    PacketsSent(SPlayerMyIndex) = PacketsSent(SPlayerMyIndex) + 1
End Sub

Sub SendEnterGame(ByVal index As Long)
Dim Buffer As clsBuffer
    
    TempPlayer(index).InGame = True
    Set Buffer = New clsBuffer
    Buffer.WriteLong SEnterGame
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
    PacketsSent(SEnterGame) = PacketsSent(SEnterGame) + 1
End Sub

Sub SendMapData(ByVal index As Long, ByVal MapNum As Long)
Dim Buffer As clsBuffer
Dim i As Long, x As Long
Dim ToF As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapData
    Buffer.WriteLong MapNum
    Buffer.WriteString Trim$(Map(MapNum).Name)
    Buffer.WriteLong Map(MapNum).Picture
    For i = 1 To 20
        If Map(index).Label(i).Visible = True Then
            ToF = 1
        Else
            ToF = 0
        End If
        Buffer.WriteByte ToF
        Buffer.WriteLong Map(index).Label(i).Width
        Buffer.WriteLong Map(index).Label(i).Height
        Buffer.WriteLong Map(index).Label(i).Left
        Buffer.WriteLong Map(index).Label(i).Top
        Buffer.WriteString Map(index).Label(i).Caption
        Buffer.WriteLong Map(index).Label(i).Event.Type
        For x = 1 To 5
            Buffer.WriteLong Map(index).Label(i).Event.Data(x)
            Buffer.WriteString Map(index).Label(i).Event.Text(x)
        Next
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
    PacketsSent(SMapData) = PacketsSent(SMapData) + 1
End Sub
