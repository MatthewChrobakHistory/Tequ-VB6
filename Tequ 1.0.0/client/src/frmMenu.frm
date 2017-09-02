VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMenu.frx":0000
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCreate 
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   555
      Picture         =   "frmMenu.frx":850C2
      ScaleHeight     =   243
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   442
      TabIndex        =   2
      Top             =   180
      Visible         =   0   'False
      Width           =   6630
      Begin VB.TextBox txtRetype 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   4
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label lblRetype 
         BackStyle       =   0  'Transparent
         Caption         =   "Retype:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblLogin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[ Login ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblCreate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[ Create ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   2640
         Width           =   855
      End
   End
   Begin VB.PictureBox picHomeScreen 
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   555
      Picture         =   "frmMenu.frx":D3D94
      ScaleHeight     =   3645
      ScaleWidth      =   6630
      TabIndex        =   0
      Top             =   180
      Width           =   6630
   End
   Begin VB.Image imgMenuButton 
      Height          =   435
      Index           =   4
      Left            =   5460
      Picture         =   "frmMenu.frx":122A66
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Image imgMenuButton 
      Height          =   435
      Index           =   3
      Left            =   3960
      Picture         =   "frmMenu.frx":124904
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Image imgMenuButton 
      Height          =   435
      Index           =   2
      Left            =   2460
      Picture         =   "frmMenu.frx":1267A2
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Image imgMenuButton 
      Height          =   435
      Index           =   1
      Left            =   960
      Picture         =   "frmMenu.frx":128640
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Label lblErrorNotification 
      BackStyle       =   0  'Transparent
      Caption         =   "An error just occured. Click here to view it."
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)

Call DestroyClient

End Sub

Private Sub imgMenuButton_Click(Index As Integer)

    Call MenuButton(Index)

End Sub

Private Sub lblCreate_Click()

    If Options.OnlineMode = True Then
        If Trim$(txtPass.Text) <> vbNullString Then
            If Trim$(txtPass.Text) = Trim$(txtRetype.Text) Then
                Call SendCreatePlayer(Trim$(txtUsername.Text), Trim$(txtPass.Text))
            Else
                'passwords didn't match
                MsgBox "Passwords don't match!", vbCritical
            End If
        End If
    ElseIf Options.OnlineMode = False Then
        If Len(Trim$(txtUsername.Text)) > 0 Then Call MakeAccount(Trim$(txtUsername.Text))
    End If

End Sub

Private Sub lblErrorNotification_Click()

Call OpenLastErrorReport
AlertMsgWait = 5000

End Sub

Private Sub MenuButton(ByVal Index As Integer)

    Select Case Index
        Case Button_SiPl
            Options.OnlineMode = False
            Call SetupScreen(False)
            txtUsername.Text = Options.Username
        Case Button_MuPl
            Options.OnlineMode = True
            Call SetupScreen(True)
            Call InitGame(Options.OnlineMode)
            txtUsername.Text = Options.Username
            txtPass.Text = Options.Password
            txtRetype.Text = Options.Password
        Case Button_Info
            
        Case Button_Exit
            End
    End Select
End Sub

Private Sub SetupScreen(ByVal Multiplayer As Boolean)
    
    frmMenu.picHomeScreen.Visible = False
    frmMenu.picCreate.Visible = True
    
    Select Case Multiplayer
        Case True
            txtPass.Visible = True
            txtRetype.Visible = True
            lblPassword.Visible = True
            lblRetype.Visible = True
        Case False
            txtPass.Visible = False
            txtRetype.Visible = False
            lblPassword.Visible = False
            lblRetype.Visible = False
    End Select
    
End Sub

Private Sub lblLogin_Click()

    If Options.OnlineMode = False Then
        If FileExist(App.Path & "\data\players\" & Trim$(txtUsername.Text) & ".bin") = True Then
            Call InitGame(False)
            Call LoadPlayer(Trim$(txtUsername.Text))
        Else
            MsgBox "Player does not exist!", vbCritical
            Exit Sub
        End If
    Else
        If Trim$(txtUsername.Text) <> vbNullString And Trim$(txtPass.Text) <> vbNullString Then
            If Trim$(txtRetype.Text) = Trim$(txtPass.Text) Then
                Call SendRequestLogin(Trim$(frmMenu.txtUsername.Text), Trim$(frmMenu.txtPass.Text))
            Else
                ' Passwords didn't match
                MsgBox "Passwords don't match!", vbCritical
            End If
        End If
    End If

End Sub

Private Sub picCreate_Click()
Call SendRequestLogin(Trim$(frmMenu.txtUsername.Text), Trim$(frmMenu.txtPass.Text))
End Sub

Private Sub txtPass_Change()

    Options.Password = Trim$(txtPass.Text)

End Sub

Private Sub txtUsername_Change()

    Options.Username = Trim$(txtUsername.Text)

End Sub
