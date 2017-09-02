Attribute VB_Name = "modBinary"
Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Sub ClearLoginPlayer()
    Call ZeroMemory(ByVal VarPtr(LoginPlayer), LenB(LoginPlayer))
End Sub

Sub LoadLoginPlayer(ByVal Name As String)
Dim FileName As String
Dim F As Long

Call ClearLoginPlayer

FileName = App.Path & "/data/players/" & Name & ".bin"
F = FreeFile
Open FileName For Binary As #F
Get #F, , LoginPlayer
Close #F
End Sub

Sub ClearPlayer(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Player(index)), LenB(Player(index)))
End Sub

Sub LoadPlayer(ByVal index As Long, ByVal Name As String)
Dim FileName As String
Dim F As Long

Call ClearPlayer(index)

FileName = App.Path & "/data/players/" & Name & ".bin"
F = FreeFile
Open FileName For Binary As #F
Get #F, , Player(index)
Close #F
End Sub

Sub SavePlayer(ByVal index As Long, ByVal Name As String)
Dim FileName As String
Dim F As Long

If TempPlayer(index).InGame = False Then Exit Sub

FileName = App.Path & "/data/players/" & Name & ".bin"
F = FreeFile
Open FileName For Binary As #F
Put #F, , Player(index)
Close #F
End Sub

Sub ClearMap(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Map(index)), LenB(Map(index)))
End Sub

Sub LoadMap(ByVal index As Long)
Dim FileName As String
Dim F As Long

Call ClearMap(index)

FileName = App.Path & "/data/maps/map" & index & ".bin"
F = FreeFile
Open FileName For Binary As #F
Get #F, , Map(index)
Close #F
End Sub

Sub SaveMap(ByVal index As Long)
Dim FileName As String
Dim F As Long

FileName = App.Path & "/data/maps/map" & index & ".bin"
F = FreeFile
Open FileName For Binary As #F
Put #F, , Map(index)
Close #F
End Sub
