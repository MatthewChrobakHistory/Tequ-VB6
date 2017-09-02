Attribute VB_Name = "modServerTCP"
Option Explicit

Sub SendDataTo(ByVal index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim TempData() As Byte

    If IsConnected(index) Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
              
        frmServer.Socket(index).SendData Buffer.ToArray()
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

    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    End If

End Function

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If

End Sub

Sub SocketConnected(ByVal index As Long)

    
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
    frmServer.Socket(index).GetData Buffer(), vbUnicode, DataLength
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
        frmServer.Socket(index).Close
        TempPlayer(index).InGame = False
        Call ClearPlayer(index)
    End If

End Sub

Sub SendEnterGame(ByVal index As Long)
Dim Buffer As New clsBuffer

    TempPlayer(index).InGame = True
    Set Buffer = New clsBuffer
    Buffer.WriteLong SEnterGame
    Buffer.WriteLong index
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerIndex(ByVal index As Long)
Dim Buffer As clsBuffer

    'Set Buffer = New clsBuffer
    'Buffer.WriteBytes SPlayerIndex
    
End Sub

Sub SendPlayerData(ByVal index As Long, ByVal PlayerNum As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong PlayerNum
    Buffer.WriteString Trim$(Player(PlayerNum).Name)
    Buffer.WriteLong Player(PlayerNum).Map
    Buffer.WriteLong Player(PlayerNum).x
    Buffer.WriteLong Player(PlayerNum).y
    Buffer.WriteLong Player(PlayerNum).Sprite
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

