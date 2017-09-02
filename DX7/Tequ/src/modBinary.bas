Attribute VB_Name = "modBinary"
Option Explicit

Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub LoadData()
Dim i As Long

If OnlineMode = True Then
    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next
    Exit Sub
Else

    Call LoadPlayer
    For i = 1 To MAX_MAPS
        Call LoadMap(i)
    Next
End If

End Sub

Public Sub SaveData()
Dim i As Long

Call SavePlayer
For i = 1 To MAX_MAPS
    Call SaveMap(i)
Next

End Sub

Sub ClearPlayer()
    Call ZeroMemory(ByVal VarPtr(Player(MyIndex)), LenB(Player(MyIndex)))
End Sub

Sub ClearMap(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Map(Index)), LenB(Map(Index)))
End Sub

Sub LoadPlayer()
Dim FileName As String
Dim F As Long

Call ClearPlayer

FileName = App.Path & "/data/player.bin"
F = FreeFile
Open FileName For Binary As #F
Get #F, , Player(MyIndex)
Close #F
End Sub

Sub LoadMap(ByVal Index As Long)
Dim FileName As String
Dim F As Long

Call ClearMap(Index)

FileName = App.Path & "/data/maps/map" & Index & ".bin"
F = FreeFile
Open FileName For Binary As #F
Get #F, , Map(Index)
Close #F
End Sub

Sub SavePlayer()
Dim FileName As String
Dim F As Long

FileName = App.Path & "/data/player.bin"
F = FreeFile
Open FileName For Binary As #F
Put #F, , Player(MyIndex)
Close #F
End Sub

Sub SaveMap(ByVal Index As Long)
Dim FileName As String
Dim F As Long

FileName = App.Path & "/data/maps/map" & Index & ".bin"
F = FreeFile
Open FileName For Binary As #F
Put #F, , Map(Index)
Close #F
End Sub
