Attribute VB_Name = "modGeneral"
Option Explicit

Sub main()
Dim i As Long

    Call InitMessages
    
    Call LoadData

    ' Get the listening socket ready to go
    frmServer.socket(0).Close
    frmServer.socket(0).RemoteHost = frmServer.socket(0).LocalIP
    frmServer.socket(0).LocalPort = 7001
    
    For i = 1 To MAX_PLAYERS
        Load frmServer.socket(i)
    Next
    
    frmServer.socket(0).Listen
    frmServer.Show
    
    Running = True
    Call GameLoop

End Sub

Public Function FileExist(ByVal FileName As String) As Boolean
    If LenB(Dir$(FileName)) > 0 Then FileExist = True
End Function

Public Function GetWinSockState(ByVal State As Byte) As String

If State > 9 Or State < 0 Then
    GetWinSockState = " Null "
    Exit Function
End If

Select Case State
    Case sckClosed
        GetWinSockState = " Connection closed "
    Case sckOpen
        GetWinSockState = " Open "
    Case sckListening
        GetWinSockState = " Listening for incoming connections "
    Case sckConnectionPending
        GetWinSockState = " Connection pending "
    Case sckResolvingHost
        GetWinSockState = " Resolving remote host name "
    Case sckHostResolved
        GetWinSockState = " Remote host name successfully resolved "
    Case sckConnecting
        GetWinSockState = " Connecting to remote host "
    Case sckConnected
        GetWinSockState = " Connected to remote host "
    Case sckClosing
        GetWinSockState = " Connection is closing "
    Case sckError
        GetWinSockState = " Error occured "
End Select

End Function

Public Sub MakePlayer(ByVal index As Long, ByVal Name As String, ByVal Password As String)

Call ClearPlayer(index)
Player(index).Name = Name
Player(index).Password = Password
Player(index).Map = 1
Call SavePlayer(index, Name)

End Sub

Public Sub LoadData()
Dim i As Long

For i = 1 To MAX_MAPS
    Call LoadMap(i)
Next
End Sub
