Attribute VB_Name = "modGlobals"

' Error Report Handling
Public LastErrorDir As String

' Game Loop Variable
Public Running As Boolean

' Error Notification Timer
Public ErrorAlert As Boolean
Public AlertMsgWait As Integer

' Text variables
Public TexthDC As Long
Public GameFont As Long

Public SocketState As Byte
Public MyIndex As Long

Public FontStyle(1 To 10) As String
Public FontNumber As Byte

Public Layer As Byte
Public CurTileX As Byte
Public CurTileY As Byte
Public NumTileset As Byte
