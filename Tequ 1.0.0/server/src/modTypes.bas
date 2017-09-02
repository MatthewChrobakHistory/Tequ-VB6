Attribute VB_Name = "modTypes"

Public Options As OptionsRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public LDPlayer As PlayerRec 'loading data player
Public Map(1 To MAX_MAPS) As MapRec

Private Type OptionsRec
    Debug As Boolean
    IP As String
    Port As Long
End Type

Private Type TempPlayerRec
    Buffer As clsBuffer
    InGame As Boolean
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    DoneLoadingData As Boolean
End Type

Private Type PlayerRec
    Name As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    Sprite As Long
    Map As Long
    x As Long
    y As Long
End Type

Private Type LayerRec
    tileset As Byte
    x As Byte
    y As Byte
End Type

Private Type TileRec
    Layer(1 To Layers.Layer_Count) As LayerRec
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Tile(0 To MAX_MAP_X, 0 To MAX_MAP_Y) As TileRec
    MapPlayer(1 To MAX_PLAYERS) As Long
End Type
