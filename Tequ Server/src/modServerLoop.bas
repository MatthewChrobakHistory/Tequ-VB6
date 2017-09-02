Attribute VB_Name = "modServerLoop"
Option Explicit

' Main timer
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' Halts thread of execution
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Public Sub GameLoop()
Dim Tick As Long
Dim tmr500 As Long, tmr1000 As Long
Dim i As Long, x As Long
Dim MapNum(1 To MAX_PLAYERS) As Long, PlayerNum(1 To MAX_PLAYERS) As Long

    For i = 1 To MAX_PLAYERS
        MapNum(i) = 1
        PlayerNum(i) = 1
    Next

Do While Running = True
    Tick = timeGetTime
    
    For i = 1 To Player_HighIndex
        If TempPlayer(i).DoneLoadingData = False Then
            If IsConnected(i) = True Then
                If MapNum(i) < MAX_MAPS Then
                    Call SendMapData(i, MapNum(i))
                    MapNum(i) = MapNum(i) + 1
                End If
            End If
                
            If PlayerNum(i) <= Player_HighIndex Then
                If IsConnected(PlayerNum(i)) = True Then
                    Call SendPlayerData(i, PlayerNum(i))
                    PlayerNum(i) = PlayerNum(i) + 1
                End If
            End If
        End If
            
        If PlayerNum(i) >= Player_HighIndex And MapNum(i) >= MAX_MAPS Then
            TempPlayer(i).DoneLoadingData = True
            MapNum(i) = 1
            PlayerNum(i) = 1
        End If
    Next
    
    If tmr500 < Tick Then
        'check for disconnections every half a second
        For i = 1 To Player_HighIndex
            If IsConnected(i) = False Then
                frmServer.socket(i).Close
                TempPlayer(i).InGame = False
            End If
        Next
        tmr500 = Tick + 500
    End If
        
    If tmr1000 < Tick Then
        If frmPacketViewer.Visible = True Then frmPacketViewer.RefreshList
        tmr1000 = Tick + 1000
    End If
        
    DoEvents
    Sleep 1
Loop
    
End Sub


