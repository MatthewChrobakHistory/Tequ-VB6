VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   8865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   591
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   784
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMyChat 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   720
      TabIndex        =   6
      Top             =   8415
      Width           =   6600
   End
   Begin VB.PictureBox picInventory 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3705
      Left            =   8115
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   4
      Top             =   4590
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   9360
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.ComboBox cmbText 
      Height          =   315
      ItemData        =   "frmMain.frx":155AEA
      Left            =   8040
      List            =   "frmMain.frx":155B0C
      TabIndex        =   2
      Text            =   "cmbText"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   150
      ScaleHeight     =   384
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   150
      Width           =   7200
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1800
      Left            =   180
      TabIndex        =   5
      Top             =   6630
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   3175
      _Version        =   393217
      BackColor       =   790032
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":155B86
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblErrorNotification 
      BackStyle       =   0  'Transparent
      Caption         =   "An error just occured. Click here to view it."
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8640
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbText_Click()

    If cmbText.ListIndex = 0 Then Exit Sub
    If cmbText.ListIndex > 10 Then Exit Sub
    Call SetFont(FontStyle(cmbText.ListIndex), FONT_SIZE)
    Options.GameFont = FontStyle(cmbText.ListIndex)

End Sub

Private Sub Command1_Click()
frmEditor_Map.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call DestroyGame

End Sub

Private Sub lblErrorNotification_Click()

    Call OpenLastErrorReport
    AlertMsgWait = 5000

End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MouseX As Single
Dim MouseY As Single
Dim ConvX As Byte
Dim ConvY As Byte

    MouseX = X / 32
    MouseY = Y / 32

    ' Exit out if out of parameters
    If MouseX < 0 Or MouseX > 480 Then Exit Sub
    If MouseY < 0 Or MouseY > 384 Then Exit Sub
    
    ConvX = MouseX
    ConvY = MouseY
    
    ' If the rounded number is bigger than the original number, we must have rounded up. Deduct one
    If ConvX - MouseX > 0 Then ConvX = ConvX - 1
    If ConvY - MouseY > 0 Then ConvY = ConvY - 1
    
    ConvX = ConvX + 1
    ConvY = ConvY + 1

    If frmEditor_Map.Visible = True Then ' We must be in the editor.
        Map(Player(MyIndex).Map).Tile(ConvX, ConvY).Layer(Layer).X = CurTileX
        Map(Player(MyIndex).Map).Tile(ConvX, ConvY).Layer(Layer).Y = CurTileY
        Map(Player(MyIndex).Map).Tile(ConvX, ConvY).Layer(Layer).Tileset = frmEditor_Map.scrlTileset.Value
    End If
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MouseX As Single
Dim MouseY As Single
Dim ConvX As Byte
Dim ConvY As Byte

If Button = vbLeftButton Or Button = vbRightButton Then

    MouseX = X / 32
    MouseY = Y / 32

    ' Exit out if out of parameters
    If MouseX < 0 Or MouseX > 480 / 32 Then Exit Sub
    If MouseY < 0 Or MouseY > 384 / 32 Then Exit Sub
    
    ConvX = MouseX
    ConvY = MouseY
    
    ' If the rounded number is bigger than the original number, we must have rounded up. Deduct one
    If ConvX - MouseX > 0 Then ConvX = ConvX - 1
    If ConvY - MouseY > 0 Then ConvY = ConvY - 1
    
    ConvX = ConvX + 1
    ConvY = ConvY + 1
    
    If ConvY > MAX_MAP_Y Or ConvX > MAX_MAP_X Then Exit Sub

    If frmEditor_Map.Visible = True Then ' We must be in the editor.
        Map(Player(MyIndex).Map).Tile(ConvX, ConvY).Layer(Layer).X = CurTileX
        Map(Player(MyIndex).Map).Tile(ConvX, ConvY).Layer(Layer).Y = CurTileY
    End If
End If

End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

End Sub

Private Sub txtChat_Click()

    Call SetFocusOnChat

End Sub

Public Sub SetFocusOnChat()

    On Error Resume Next ' prevent RTE5, no way to handle error
    
    If frmMain.txtMyChat.Visible = True Then
        frmMain.txtMyChat.SetFocus
    Else
        frmMain.txtMyChat.Visible = True
        frmMain.txtMyChat.SetFocus
        frmMain.txtMyChat.Visible = False
    End If
End Sub
