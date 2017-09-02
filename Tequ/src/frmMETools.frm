VERSION 5.00
Begin VB.Form frmMETools 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   2565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Label Data"
      Height          =   4335
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton cmdEvent 
         Caption         =   "Event"
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtText 
         Height          =   735
         Left            =   120
         MaxLength       =   50
         TabIndex        =   21
         Top             =   3120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtLeft 
         Height          =   285
         Left            =   720
         TabIndex        =   16
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTop 
         Height          =   285
         Left            =   720
         TabIndex        =   15
         Top             =   2160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Left:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Top:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Height:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblLabelFill 
         Alignment       =   2  'Center
         Caption         =   "Opaque"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label label1 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label lblIndex 
         Caption         =   "Index: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Map Data"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   2295
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         TabIndex        =   23
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.HScrollBar scrlPicture 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load"
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblPicture 
         Caption         =   "Picture: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmMETools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()

With frmMapEditor.Label(LabelIndex)
    .Caption = "Label"
    .Width = 57
    .Height = 17
    .Left = 0
    .Top = 0
    .Visible = False
End With

LabelIndex = 0

End Sub

Private Sub cmdLoad_Click()
Dim MapNum As String

MapNum = InputBox("Please enter the map you want to load.", "Loading Map...")
If IsNumeric(MapNum) = False Then Exit Sub

CurMap = MapNum
Call LoadEditorMap(MapNum)

End Sub

Public Sub LoadEditorMap(ByVal MapNum As Long)
Dim i As Long

With frmMapEditor
    For i = 1 To 20
        .Label(i).Caption = Map(MapNum).Label(i).Caption
        .Label(i).Left = Map(MapNum).Label(i).Left
        .Label(i).Top = Map(MapNum).Label(i).Top
        .Label(i).Width = Map(MapNum).Label(i).Width
        .Label(i).Height = Map(MapNum).Label(i).Height
        .Label(i).Visible = Map(MapNum).Label(i).Visible
    Next
    .picMap.Picture = Nothing
    If FileExist(App.Path & "\graphics\maps\" & Map(MapNum).Picture & ".bmp") = True Then .picMap.Picture = LoadPicture(App.Path & "\graphics\maps\" & Map(MapNum).Picture & ".bmp")
    Me.txtName.Text = Trim$(Map(MapNum).Name)
End With

MsgBox "Map loaded!"

End Sub

Private Sub cmdNew_Click()
Dim i As Long

If CurMap = 0 Then Exit Sub

For i = 1 To 20 'max labels
    If frmMapEditor.Label(i).Visible = False Then
        LabelIndex = i
        txtWidth.Text = frmMapEditor.Label(i).Width
        txtHeight.Text = frmMapEditor.Label(i).Height
        frmMapEditor.Label(i).Visible = True
        Exit For
    End If
Next


End Sub

Private Sub cmdSave_Click()
Dim i As Long

If CurMap = 0 Then Exit Sub

With frmMapEditor
    For i = 1 To 20
        Map(CurMap).Label(i).Caption = .Label(i).Caption
        Map(CurMap).Label(i).Left = .Label(i).Left
        Map(CurMap).Label(i).Top = .Label(i).Top
        Map(CurMap).Label(i).Width = .Label(i).Width
        Map(CurMap).Label(i).Height = .Label(i).Height
        Map(CurMap).Label(i).Visible = .Label(i).Visible
    Next
    Map(CurMap).Picture = Me.scrlPicture
    Map(CurMap).Name = Trim$(Me.txtName.Text)
End With

MsgBox "Map saved!"

End Sub

Private Sub Form_Unload(Cancel As Integer)

CurMap = 0
LabelIndex = 0
InEditor = False
Call Unload(frmMapEditor)

End Sub

Private Sub lblLabelFill_Click()

If TransparentLabels = False Then
    TransparentLabels = True
    lblLabelFill.BackStyle = Transparent
    lblLabelFill.Caption = "Transparent"
Else
    TransparentLabels = False
    lblLabelFill.BackStyle = Opaque
    lblLabelFill.Caption = "Opaque"
End If

End Sub

Private Sub scrlPicture_Change()

If CurMap = 0 Then Exit Sub

frmMapEditor.picMap.Picture = Nothing
If FileExist(App.Path & "/graphics/maps/" & scrlPicture.Value & ".bmp") = True Then frmMapEditor.picMap.Picture = LoadPicture(App.Path & "/graphics/maps/" & scrlPicture.Value & ".bmp")

Map(CurMap).Picture = scrlPicture

End Sub

Private Sub txtHeight_Change()

If IsNumeric(txtHeight.Text) = False Then txtHeight.Text = frmMapEditor.Label(LabelIndex).Height
If txtHeight.Text > 500 Or txtHeight.Text < 1 Then txtHeight.Text = frmMapEditor.Label(LabelIndex).Height

frmMapEditor.Label(LabelIndex).Height = txtHeight.Text
End Sub

Private Sub txtLeft_Change()

If IsNumeric(txtLeft.Text) = False Then txtLeft.Text = frmMapEditor.Label(LabelIndex).Left
If txtLeft.Text > frmMapEditor.picMap.Width Or txtLeft.Text < 0 Then txtLeft.Text = frmMapEditor.Label(LabelIndex).Left

frmMapEditor.Label(LabelIndex).Left = txtLeft.Text

End Sub

Private Sub txtText_Change()

frmMapEditor.Label(LabelIndex).Caption = txtText.Text

End Sub

Private Sub txtTop_Change()

If IsNumeric(txtTop.Text) = False Then txtTop.Text = frmMapEditor.Label(LabelIndex).Top
If txtTop.Text > frmMapEditor.picMap.Height Or txtTop.Text < 0 Then txtTop.Text = frmMapEditor.Label(LabelIndex).Top

frmMapEditor.Label(LabelIndex).Top = txtTop.Text

End Sub

Private Sub txtWidth_Change()

If IsNumeric(txtWidth.Text) = False Then txtWidth.Text = frmMapEditor.Label(LabelIndex).Width
If txtWidth.Text > 500 Or txtWidth.Text < 1 Then txtWidth.Text = frmMapEditor.Label(LabelIndex).Width

frmMapEditor.Label(LabelIndex).Width = txtWidth.Text
End Sub
