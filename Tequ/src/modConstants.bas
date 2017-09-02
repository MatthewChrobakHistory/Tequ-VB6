Attribute VB_Name = "modConstants"
Option Explicit

' API Declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const NAME_LENGTH As Byte = 12


Public Const MAX_MAPS As Long = 100
Public Const MAX_PLAYERS As Long = 70

' MAP EDITOR LABEL VALUES
Public Const Transparent As Byte = 0
Public Const Opaque As Byte = 1


' TEXT COLORS
Public Const Blue As String = "&H8000000D&"
Public Const Red As String = "&H000000FF&"

' Login Details
Public LC As Byte
Public Const Login As Byte = 1
Public Const Create As Byte = 2
