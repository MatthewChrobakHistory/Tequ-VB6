Attribute VB_Name = "modConstants"

Public Const MAX_PLAYERS As Long = 10

Public Const MAX_MAPS As Long = 1
Public Const MAX_MAP_X As Byte = 16
Public Const MAX_MAP_Y As Byte = 12
Public Const NAME_LENGTH As Byte = 12

' text color constants
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15

' Font variables
Public Const FONT_SIZE As Byte = 14

' Menu Buttons
Public Const Button_SiPl As Byte = 1
Public Const Button_MuPl As Byte = 2
Public Const Button_Info As Byte = 3
Public Const Button_Exit As Byte = 4

' Socket Constants
Public Const State_Closed As Byte = 0
Public Const State_Connecting As Byte = 1
Public Const State_Connected As Byte = 2
