VERSION 5.00
Begin VB.Form frmAdminPanel 
   Caption         =   "Admin Panel"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   1845
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   1845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMapEditor 
      Caption         =   "Map Editor"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmAdminPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMapEditor_Click()

Call frmMapEditor.LoadMapEditor
Unload Me

End Sub
