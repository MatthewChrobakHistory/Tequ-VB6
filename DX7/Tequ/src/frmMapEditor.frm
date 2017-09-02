VERSION 5.00
Begin VB.Form frmMapEditor 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   563
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   650
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picMap 
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   120
      ScaleHeight     =   545
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   633
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   20
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   19
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   18
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   17
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   16
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   15
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   14
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   13
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   12
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   11
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   10
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   9
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   8
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   7
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   6
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   5
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   4
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   3
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   2
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         Height          =   255
         Index           =   1
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMapEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click()

Call frmMETools.LoadEditorMap(InputBox("Enter the map you wish to load.", "Loading Map..."))

End Sub

Public Sub LoadMapEditor()

frmMapEditor.Show
frmMETools.Show
picMap.Picture = Nothing
InEditor = True

End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)

Source.Move X, Y

End Sub

Private Sub Form_Unload(Cancel As Integer)

If InEditor = True Then Cancel = 1

End Sub

Private Sub Label_DblClick(Index As Integer)

LabelIndex = Index
frmMETools.txtWidth.Text = Label(Index).Width
frmMETools.txtHeight.Text = Label(Index).Height
frmMETools.txtLeft.Text = Label(Index).Left
frmMETools.txtTop.Text = Label(Index).Top
frmMETools.txtText.Text = Trim$(Label(Index).Caption)

End Sub

Private Sub Label_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Label(Index).Drag vbBeginDrag

End Sub

Private Sub Label_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Label(Index).Drag vbEndDrag

End Sub
