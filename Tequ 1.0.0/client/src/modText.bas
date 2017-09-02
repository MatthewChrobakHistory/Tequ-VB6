Attribute VB_Name = "modText"
Option Explicit

' Text declares
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Function GetFontNumber(ByVal Name As String) As Byte
    Select Case Name
        Case "Calibri"
            GetFontNumber = 1
        Case "Cambria"
            GetFontNumber = 2
        Case "Candara"
            GetFontNumber = 3
        Case "Courier New"
            GetFontNumber = 4
        Case "News Gothic"
            GetFontNumber = 5
        Case "Palantino Linotype"
            GetFontNumber = 6
        Case "Pescadero"
            GetFontNumber = 7
        Case "Tahoma"
            GetFontNumber = 8
        Case "Trajan Pro"
            GetFontNumber = 9
        Case "Trebuchet MS"
            GetFontNumber = 10
    End Select
            
End Function

Public Sub LoadFonts()

FontStyle(1) = "Calibri"
FontStyle(2) = "Cambria"
FontStyle(3) = "Candara"
FontStyle(4) = "Courier New"
FontStyle(5) = "News Gothic"
FontStyle(6) = "Palatino Linotype"
FontStyle(7) = "Pescadero"
FontStyle(8) = "Tahoma"
FontStyle(9) = "Trajan Pro"
FontStyle(10) = "Trebuchet MS"

End Sub

' Used to set a font for GDI text drawing
Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:
    
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
    frmMain.Font = Font
    frmMain.FontSize = Size - 5
    
' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "", Err.Description)
Err.Clear
End Sub

' GDI text drawing onto buffer
Public Sub DrawText(ByVal hDC As Long, ByVal X, ByVal Y, ByVal Text As String, color As Long)

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:
    
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, 0)
    Call TextOut(hDC, X + 1, Y + 1, Text, Len(Text))
    Call SetTextColor(hDC, color)
    Call TextOut(hDC, X, Y, Text, Len(Text))
    
' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "", Err.Description)
Err.Clear
End Sub

Public Sub DrawRandomText()
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim Text As String

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:
    
    ' Set the color
    color = QBColor(Cyan)
    ' Set the text you want to render

    Text = "LOL!"
    ' calc pos
    TextX = 64
    TextY = 64

    ' Draw name
    Call DrawText(TexthDC, TextX, TextY, Text, color)
    
' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "", Err.Description)
Err.Clear
End Sub
