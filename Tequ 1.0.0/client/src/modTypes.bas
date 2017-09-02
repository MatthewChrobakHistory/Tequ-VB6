Attribute VB_Name = "modTypes"

Public Options As OptionRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Map(1 To MAX_MAPS) As MapRec

Private Type OptionRec
    Debug As Boolean
    OnlineMode As Boolean
    InGame As Boolean
    IP As String
    Port As String
    InstallRuntimes As Boolean
    GameFont As String
    Username As String
    Password As String
End Type

Private Type TempPlayerRec
    Buffer As clsBuffer
    InGame As Boolean
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
End Type

Private Type PlayerRec
    Name As String * NAME_LENGTH
    Sprite As Long
    Map As Long
    X As Long
    Y As Long
End Type

Private Type LayerRec
    tileset As Byte
    X As Byte
    Y As Byte
End Type

Private Type TileRec
    Layer(1 To Layers.Layer_Count) As LayerRec
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Tile(0 To MAX_MAP_X, 0 To MAX_MAP_Y) As TileRec
    MapPlayer(1 To MAX_PLAYERS) As Long
End Type
