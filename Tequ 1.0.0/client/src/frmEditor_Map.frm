VERSION 5.00
Begin VB.Form frmEditor_Map 
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar scrlTileset 
      Height          =   255
      Left            =   120
      Min             =   1
      TabIndex        =   11
      Top             =   6720
      Value           =   1
      Width           =   6495
   End
   Begin VB.OptionButton optLayer 
      Caption         =   "Fringe - Anim"
      Height          =   255
      Index           =   9
      Left            =   6720
      TabIndex        =   10
      Top             =   3120
      Width           =   1695
   End
   Begin VB.OptionButton optLayer 
      Caption         =   "Fringe - 3"
      Height          =   255
      Index           =   8
      Left            =   6720
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin VB.OptionButton optLayer 
      Caption         =   "Fringe - 2"
      Height          =   255
      Index           =   7
      Left            =   6720
      TabIndex        =   8
      Top             =   2400
      Width           =   1695
   End
   Begin VB.OptionButton optLayer 
      Caption         =   "Fringe - 1"
      Height          =   255
      Index           =   6
      Left            =   6720
      TabIndex        =   7
      Top             =   2040
      Width           =   1695
   End
   Begin VB.OptionButton optLayer 
      Caption         =   "Mask - Anim"
      Height          =   255
      Index           =   5
      Left            =   6720
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
   Begin VB.OptionButton optLayer 
      Caption         =   "Mask - 3"
      Height          =   255
      Index           =   4
      Left            =   6720
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.OptionButton optLayer 
      Caption         =   "Mask - 2"
      Height          =   255
      Index           =   3
      Left            =   6720
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.OptionButton optLayer 
      Caption         =   "Mask - 1"
      Height          =   255
      Index           =   2
      Left            =   6720
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.OptionButton optLayer 
      Caption         =   "Ground"
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   2
      Top             =   240
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.PictureBox picTileset 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   120
      ScaleHeight     =   433
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   433
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.PictureBox picSel 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   6000
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   1
         Top             =   6000
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmEditor_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    scrlTileset.Max = NumTileset
    picTileset.Picture = Nothing
    picTileset.Picture = LoadPicture(App.Path & "\graphics\tilesets\" & scrlTileset.Value & ".bmp")
    Layer = Layers.Ground

End Sub

Private Sub optLayer_Click(Index As Integer)

    Layer = Index

End Sub

Private Sub picTileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TempX As Single, TempY As Single

    
    ' Convert the number, and make sure there's no rounding up or down.
    TempX = (X / 32)
    TempY = (Y / 32)
    
    
    CurTileX = TempX
    CurTileY = TempY
    If TempX - CurTileX < 0 Then CurTileX = CurTileX - 1
    If TempY - CurTileY < 0 Then CurTileY = CurTileY - 1
    
    picSel.Left = CurTileX * 32
    picSel.Top = CurTileY * 32
    
    frmEditor_Map.Caption = CurTileX & ":" & CurTileY

End Sub

Private Sub scrlTileset_Change()

    picTileset.Picture = Nothing
    picTileset.Picture = LoadPicture(App.Path & "\graphics\tilesets\" & scrlTileset.Value & ".bmp")

End Sub
