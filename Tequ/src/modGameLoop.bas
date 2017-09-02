Attribute VB_Name = "modGameLoop"
Option Explicit

' Main timer
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' Halts thread of execution
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Public Sub GameLoop()
Dim i As Long
Dim Tick As Long
Dim tmr1000 As Long
Dim tmr250 As Long

    Do While Running = True
        Tick = timeGetTime
        
        
        If tmr250 < Tick Then
        
            'render the map
            If LastMap <> Player(MyIndex).Map Then
                frmMain.picMap.Picture = Nothing
                If FileExist(App.Path & "/graphics/maps/" & Map(Player(MyIndex).Map).Picture & ".bmp") = True Then
                    frmMain.picMap.Picture = LoadPicture(App.Path & "/graphics/maps/" & Map(Player(MyIndex).Map).Picture & ".bmp")
                    LastMap = Player(MyIndex).Map
                End If
                Call SetupMapLabels(Player(MyIndex).Map)
            End If
            'render the map
            
            tmr250 = Tick + 250
        End If
        
        
        If tmr1000 < Tick Then
        
            frmMain.Caption = GetWinSockState(frmMain.Socket.State)
        
            'check to see if the player disconnected
            If IsConnected = False And OnlineMode = True Then
                DestroyGame
            End If
            'check to see if the player disconnected
            
            'EDITOR
            If InEditor = True Then
                With frmMETools
                    If LabelIndex = 0 Then
                        .txtHeight.Visible = False
                        .txtWidth.Visible = False
                        .txtLeft.Visible = False
                        .txtTop.Visible = False
                        .txtText.Visible = False
                        .cmdDelete.Visible = False
                    Else
                        .txtHeight.Visible = True
                        .txtWidth.Visible = True
                        .txtLeft.Visible = True
                        .txtTop.Visible = True
                        .txtText.Visible = True
                        .cmdDelete.Visible = True
                    End If
                End With
                
                frmMETools.lblIndex.Caption = "Index: " & LabelIndex
                
                If TransparentLabels = True Then
                    For i = 1 To 20
                        frmMapEditor.Label(i).BackStyle = Transparent
                    Next
                Else
                    For i = 1 To 20
                        frmMapEditor.Label(i).BackStyle = Opaque
                    Next
                End If
            End If
            'EDITOR
            
            tmr1000 = timeGetTime + 1000
        End If
        
        DoEvents
        Sleep 1
    Loop
    
End Sub

