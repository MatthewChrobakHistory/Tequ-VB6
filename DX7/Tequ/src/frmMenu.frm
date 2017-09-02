VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCharacterCreation 
      Height          =   5295
      Left            =   120
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   7815
      Begin VB.CommandButton cmdX 
         Caption         =   "x"
         Height          =   315
         Left            =   7440
         TabIndex        =   33
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play!"
         Height          =   375
         Left            =   2520
         TabIndex        =   30
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox txtRetypePass 
         Height          =   285
         Left            =   2400
         TabIndex        =   29
         Top             =   2640
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   2400
         TabIndex        =   27
         Top             =   2280
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   2400
         TabIndex        =   26
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label lblType 
         Caption         =   "Type: Create (click to switch)"
         Height          =   375
         Left            =   2400
         TabIndex        =   32
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label lblRetypePass 
         Caption         =   "Retype Pass:"
         Height          =   255
         Left            =   1320
         TabIndex        =   28
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   1320
         TabIndex        =   25
         Top             =   2280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Name:"
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   1920
         Width           =   1335
      End
   End
   Begin VB.Label Label5 
      Caption         =   "delete singleplayer account"
      Height          =   255
      Left            =   3840
      TabIndex        =   31
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label"
      Height          =   255
      Index           =   1
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   22
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
      TabIndex        =   21
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
      Index           =   4
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
      Index           =   5
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
      Index           =   6
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
      Index           =   7
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
      Index           =   8
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
      Index           =   9
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
      Index           =   10
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
      Index           =   11
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
      Index           =   12
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
      Index           =   13
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
      Index           =   14
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
      Index           =   15
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
      Index           =   16
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
      Index           =   17
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
      Index           =   18
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
      Index           =   19
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
      Index           =   20
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "Singleplayer"
      Height          =   1455
      Left            =   3840
      TabIndex        =   1
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Multiplayer"
      Height          =   1455
      Left            =   600
      TabIndex        =   0
      Top             =   2400
      Width           =   2895
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPlay_Click()

If Trim$(txtName.Text) = vbNullString Then Exit Sub

If OnlineMode = False Then
    Call MakePlayer(Trim$(txtName.Text))
    fraCharacterCreation.Visible = False
    Call EnterGame
Else
    If Trim$(txtPassword.Text) = vbNullString Then Exit Sub
    If LC = Create Then
        If Trim$(txtPassword.Text) <> Trim$(txtRetypePass.Text) Then
            MsgBox "Passwords don't match!"
            Exit Sub
        End If
        Call SendCreatePlayer(Trim$(txtName.Text), Trim$(txtPassword.Text))
    Else
        Call SendRequestLogin(Trim$(txtName.Text), Trim$(txtPassword.Text))
    End If
End If

End Sub

Private Sub cmdX_Click()

frmMain.Socket.Close
Connecting = False
Me.fraCharacterCreation.Visible = False

End Sub

Private Sub Form_Load()

txtName.MaxLength = NAME_LENGTH
txtPassword.MaxLength = NAME_LENGTH
txtRetypePass.MaxLength = NAME_LENGTH

End Sub

Private Sub Form_Unload(Cancel As Integer)

End

End Sub

Private Sub Label1_Click()

frmMain.Socket.Close
Connecting = False
OnlineMode = True
Call LoadGame

End Sub

Private Sub Label2_Click()

OnlineMode = False
Call LoadGame

End Sub

Private Sub Label5_Click()

MyIndex = 1
Call ClearPlayer
Call SavePlayer
MsgBox "Player deleted!"
MyIndex = 0

End Sub

Private Sub lblType_Click()

If LC = Login Then
    LC = Create
    cmdPlay.Caption = "Create!"
    lblType.Caption = "Type: Create (click to switch)"
    lblRetypePass.Visible = True
    txtRetypePass.Visible = True
    txtRetypePass.Text = vbNullString
ElseIf LC = Create And OnlineMode = True Then
    LC = Login
    cmdPlay.Caption = "Play!"
    lblType.Caption = "Type: Login (click to switch)"
    lblRetypePass.Visible = False
    txtRetypePass.Visible = False
End If

End Sub
