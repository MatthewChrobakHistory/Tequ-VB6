Attribute VB_Name = "modData"
Option Explicit

Public Sub SaveData()
Dim i As Long

For i = 1 To MAX_MAPS
    Call SaveMap(i)
Next

End Sub

Public Sub LoadData()
Dim i As Long

For i = 1 To MAX_MAPS
    Call LoadMap(i)
Next

End Sub

Public Sub LoadOptions()
Dim FileName As String

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    ' Get the filename
    FileName = App.Path & "\options.ini"
    
    ' If the file doesn't exist, save it and then it will continue as normal.
    If FileExist(FileName) = False Then
        SaveOptions (True)
    End If
    
    Options.Debug = GetVar(FileName, "Options", "Debug")
    Options.IP = GetVar(FileName, "Options", "IP")
    Options.Port = GetVar(FileName, "Options", "Port")
    Options.InstallRuntimes = GetVar(FileName, "Options", "InstallRuntimes")
    Options.GameFont = GetVar(FileName, "Options", "GameFont")
    Options.Username = GetVar(FileName, "Options", "Username")
    Options.Password = GetVar(FileName, "Options", "Password")
    
' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "LoadOptions", Err.Description)
Err.Clear
End Sub

Public Sub SaveOptions(Optional ByVal NewFile As Boolean = False)
Dim FileName As String

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:
    
    FileName = App.Path & "\options.ini"
    
    If NewFile = True Then
        Options.Debug = 1
        Options.IP = "localhost"
        Options.Port = 7001
        Options.InstallRuntimes = True
        Options.GameFont = FontStyle(8) ' Tamoha
        Options.Username = vbNullString
        Options.Password = vbNullString
    End If
    
    Call PutVar(FileName, "Options", "Debug", Str(Options.Debug))
    Call PutVar(FileName, "Options", "IP", Trim$(Options.IP))
    Call PutVar(FileName, "Options", "Port", Str(Options.Port))
    Call PutVar(FileName, "Options", "InstallRuntimes", Str(Options.InstallRuntimes))
    Call PutVar(FileName, "Options", "GameFont", Trim$(Options.GameFont))
    Call PutVar(FileName, "Options", "Username", Trim$(Options.Username))
    Call PutVar(FileName, "Options", "Password", Trim$(Options.Password))

' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "SaveOptions", Err.Description)
Err.Clear
End Sub

Public Sub MakeAccount(ByVal Name As String)

    If FileExist(App.Path & "\data\players\" & Name & ".bin") = False Then
        MyIndex = 1
        With Player(MyIndex)
            .Name = Name
            .Map = 1
            .Sprite = 1
            .X = 1
            .Y = 1
        End With
        Call SavePlayer(Name)
        ' send them into the game
        Call InitGame(Options.OnlineMode)
    Else
        MsgBox "Player already exists!", vbCritical
        Exit Sub
    End If
    
End Sub

Sub ClearPlayer()
    Call ZeroMemory(ByVal VarPtr(Player(MyIndex)), LenB(Player(MyIndex)))
End Sub

Sub LoadPlayer(ByVal Name As String)
Dim FileName As String
Dim f As Long

Call ClearPlayer

FileName = App.Path & "/data/players/" & Name & ".bin"
f = FreeFile
Open FileName For Binary As #f
Get #f, , Player(MyIndex)
Close #f
End Sub

Sub SavePlayer(ByVal Name As String)
Dim FileName As String
Dim f As Long

FileName = App.Path & "/data/players/" & Name & ".bin"
f = FreeFile
Open FileName For Binary As #f
Put #f, , Player(MyIndex)
Close #f
End Sub

Sub ClearMap(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Map(Index)), LenB(Map(Index)))
End Sub

Sub LoadMap(ByVal Index As Long)
Dim FileName As String
Dim f As Long

Call ClearMap(Index)

If Options.OnlineMode = True Then
    FileName = App.Path & "/data/mapcache/map" & Index & ".bin"
Else
    FileName = App.Path & "/data/maps/map" & Index & ".bin"
End If

f = FreeFile
Open FileName For Binary As #f
Get #f, , Map(Index)
Close #f
End Sub

Sub SaveMap(ByVal Index As Long)
Dim FileName As String
Dim f As Long

If Options.OnlineMode = True Then
    FileName = App.Path & "/data/mapcache/map" & Index & ".bin"
Else
    FileName = App.Path & "/data/maps/map" & Index & ".bin"
End If

f = FreeFile
Open FileName For Binary As #f
Put #f, , Map(Index)
Close #f
End Sub

Sub SaveMapCache(ByVal Index As Long)
Dim FileName As String
Dim f As Long
Dim X As Long, Y As Long

    FileName = App.Path & "\data\mapcache\" & Index & ".map"

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Map.Name

    For X = 0 To MAX_MAP_X
        For Y = 0 To MAX_MAP_Y
            Put #f, , Map.Tile(X, Y)
        Next

        DoEvents
    Next

    Close #f
End Sub
