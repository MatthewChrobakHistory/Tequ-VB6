Attribute VB_Name = "modClientTCP"
Option Explicit

Private PlayerBuffer As clsBuffer

Public Sub IncomingData(ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long
Set PlayerBuffer = New clsBuffer

    frmMain.Socket.GetData Buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes Buffer()
    
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
    
    If frmMain.Socket.State = sckConnected Then
        IsConnected = True
    End If

End Function

Sub SendData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
    
    If IsConnected Then
        Set Buffer = New clsBuffer
                
        Buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        Buffer.WriteBytes Data()
        frmMain.Socket.SendData Buffer.ToArray()
    End If
    
End Sub

Public Sub SendRequestLogin(ByVal Username As String, ByVal Password As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestLogin
    Buffer.WriteString Username
    Buffer.WriteString Password
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

Public Sub SendCreatePlayer(ByVal Username As String, ByVal Password As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCreatePlayer
    Buffer.WriteString Username
    Buffer.WriteString Password
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub
