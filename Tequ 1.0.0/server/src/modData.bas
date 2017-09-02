Attribute VB_Name = "modData"
Option Explicit

Public Sub LoadOptions()
Dim FileName As String

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    ' Get the filename
    FileName = App.Path & "\options.ini"
    
    ' If the file doesn't exist, save it and then it will continue as normal.
    If FileExist(FileName) = False Then
        SaveOptions
    End If
    
    Options.Debug = GetVar(FileName, "Options", "Debug")
    Options.Port = GetVar(FileName, "Options", "Port")

' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "LoadOptions", Err.Description)
Err.Clear
End Sub

Public Sub SaveOptions()
Dim FileName As String

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:
    
    FileName = App.Path & "\options.ini"
    Options.Debug = 0
    Options.Port = 7001
    Call PutVar(FileName, "Options", "Debug", Str(Options.Debug))
    Call PutVar(FileName, "Options", "Port", Str(Options.Port))

' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "SaveOptions", Err.Description)
Err.Clear
End Sub

Public Sub LoadData()
Dim i As Long
' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    Call LoadOptions
    
    For i = 1 To MAX_MAPS
        Call LoadMap(i)
        
    Next
    

' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "LoadData", Err.Description)
Err.Clear
End Sub

Public Sub SaveData()
Dim i As Long
' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    Call SaveOptions
    
    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next

' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "SaveData", Err.Description)
Err.Clear
End Sub

Public Sub SendGameDataTo(ByVal index As Long)
Dim i As Long
Dim L As Long

    For i = 1 To MAX_PLAYERS
        If TempPlayer(i).InGame = True Then
            Call SendPlayerData(index, i)
        End If
    Next

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

FileName = App.Path & "/data/players/" & Name & ".bin"
F = FreeFile
Open FileName For Binary As #F
Put #F, , Player(index)
Close #F
End Sub

Sub ClearLDPlayer()
    Call ZeroMemory(ByVal VarPtr(LDPlayer), LenB(LDPlayer))
End Sub

Sub LoadLDPlayer(ByVal Name As String)
Dim FileName As String
Dim F As Long

Call ClearLDPlayer

FileName = App.Path & "/data/players/" & Name & ".bin"
F = FreeFile
Open FileName For Binary As #F
Get #F, , LDPlayer
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
