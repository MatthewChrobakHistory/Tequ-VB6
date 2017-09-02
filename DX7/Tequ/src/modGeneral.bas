Attribute VB_Name = "modGeneral"
Option Explicit

Public Sub Main()

' show the menu
frmMenu.Show
Call InitMessages

End Sub

' Checks if a directory exists, if it doesn't, it makes it
Public Sub CheckDir(ByVal Path As String, ByVal Directory As String)
    If LCase$(Dir$(Path & Directory, vbDirectory)) <> Directory Then Call MkDir(Path & Directory)
End Sub

Public Function FileExist(ByVal FileName As String) As Boolean
    If LenB(Dir$(FileName)) > 0 Then FileExist = True
End Function

Public Sub DestroyGame()

If OnlineMode = True Then
    frmMain.Socket.Close
    Connecting = False
End If

LastMap = 0
Running = False
frmMenu.Visible = True
frmMenu.fraCharacterCreation.Visible = False
Unload frmMETools
Unload frmAdminPanel
Unload frmMapEditor
Unload frmMain

End Sub

Public Sub LoadGame()
Dim tmr3000 As Long
Dim Connected As Byte 'variable used in the loop
Dim Tick As Long

If OnlineMode = True Then
    If Connecting = False Then ' we're not connected
        frmMain.Socket.Close 'clear the socket, just in case
        frmMain.Socket.RemoteHost = "localhost" '"192.168.2.13"
        frmMain.Socket.RemotePort = 7001
        frmMain.Socket.Connect
        frmMain.label1.Caption = "Mode: Online"
        Connecting = True
    
    'CONNECTING TO THE SERVER LOOP
    
        tmr3000 = timeGetTime + 3000
            
        Do While Connected = 0
            Tick = timeGetTime
                
            If Tick > tmr3000 Then
                If frmMain.Socket.State <> 7 Then
                    frmMenu.Label3.Caption = GetWinSockState(frmMain.Socket.State)
                    DestroyGame
                    Connected = 2 'failed
                    frmMain.Socket.Close
                Else
                    Connected = 1 'sucess
                End If
            Else
                frmMenu.Caption = tmr3000 - Tick
                frmMenu.Label3.Caption = GetWinSockState(frmMain.Socket.State)
                If frmMain.Socket.State = 7 Then
                    Connected = 1
                End If
            End If
                        
            DoEvents
            Sleep 1
        Loop
                    
                    
        If frmMain.Socket.State <> 7 Then Exit Sub
                    
    'CONNECTING TO THE SERVER LOOP
    End If
    
    ' we should be connected now. Go ahead and load up everything
    
    LC = Login
    frmMenu.lblType.Caption = "Type: Login (click to switch)"
    frmMenu.fraCharacterCreation.Visible = True
    frmMenu.lblPassword.Visible = True
    'frmMenu.lblRetypePass.Visible = True
    frmMenu.txtPassword.Visible = True
    'frmMenu.txtRetypePass.Visible = True
    Exit Sub
Else
    ' we're not online, so just load up singleplayer
    MyIndex = 1
    Call LoadPlayer
    frmMain.label1.Caption = "Mode: Offline"
    
    If Player(MyIndex).Map = 0 Then ' check to see if the player doesn't exist.
        LC = Create
        frmMenu.lblType.Caption = "Type: Create (click to switch)"
        frmMenu.fraCharacterCreation.Visible = True
        frmMenu.lblPassword.Visible = False
        frmMenu.lblRetypePass.Visible = False
        frmMenu.txtPassword.Visible = False
        frmMenu.txtRetypePass.Visible = False
        Exit Sub
    Else
        Call EnterGame
    End If
End If

End Sub

Public Sub EnterGame()

    ' start the loop
    Running = True
    frmMenu.Hide
    Call LoadData
        
    frmMain.Show
    Call GameLoop

End Sub

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

Public Sub SetupMapLabels(ByVal MapNum As Long)
Dim i As Long

With frmMain
    For i = 1 To 20
        .Label(i).Left = Map(MapNum).Label(i).Left
        .Label(i).Top = Map(MapNum).Label(i).Top
        .Label(i).Caption = Trim$(Map(MapNum).Label(i).Caption)
        .Label(i).Height = Map(MapNum).Label(i).Height
        .Label(i).width = Map(MapNum).Label(i).width
        .Label(i).Visible = Map(MapNum).Label(i).Visible
        .Label(i).BackStyle = Transparent
    Next
End With

End Sub

Public Sub MakePlayer(ByVal Name As String)

Call ClearPlayer
Player(MyIndex).Name = Name
Player(MyIndex).Map = 1
Call SavePlayer

End Sub
