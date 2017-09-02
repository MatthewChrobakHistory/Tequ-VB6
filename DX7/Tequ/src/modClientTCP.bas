Attribute VB_Name = "modClientTCP"
Option Explicit

Private PlayerBuffer As clsBuffer

Public Sub IncomingData(ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long
Set PlayerBuffer = New clsBuffer

    frmMain.Socket.GetData buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes buffer()
    
    If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= PlayerBuffer.Length - 4
        If pLength <= PlayerBuffer.Length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop
    PlayerBuffer.Trim
    DoEvents
    
End Sub

Function IsConnected() As Boolean
    
    frmMenu.Label3.Caption = GetWinSockState(frmMain.Socket.State)
    
    If frmMain.Socket.State = sckConnected Then
        IsConnected = True
    End If

End Function

Function IsPlaying(ByVal Index As Long) As Boolean

        IsPlaying = True
    
End Function

Sub SendData(ByRef Data() As Byte)
Dim buffer As clsBuffer
    
    If IsConnected Then
        Set buffer = New clsBuffer
                
        buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        buffer.WriteBytes Data()
        frmMain.Socket.SendData buffer.ToArray()
    End If
    
End Sub

Public Sub SendRequestLogin(ByVal Username As String, ByVal Password As String)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestLogin
    buffer.WriteString Username
    buffer.WriteString Password
    SendData buffer.ToArray()
    Set buffer = Nothing
    
End Sub

Public Sub SendCreatePlayer(ByVal Username As String, ByVal Password As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CCreatePlayer
    buffer.WriteString Username
    buffer.WriteString Password
    SendData buffer.ToArray()
    Set buffer = Nothing
    
End Sub
