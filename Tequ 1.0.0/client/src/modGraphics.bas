Attribute VB_Name = "modGraphics"
Option Explicit

' Main DirectX Object
Dim DD As DirectDraw7
' Clipper
Public DD_Clip As DirectDrawClipper

' Primary surface
Dim DDS_Primary As DirectDrawSurface7
Dim DDSD_Primary As DDSURFACEDESC2
' Backbuffer
Dim DDS_Backbuffer As DirectDrawSurface7
Dim DDSD_Backbuffer As DDSURFACEDESC2

Dim DDS_Tileset(1 To 255) As DirectDrawSurface7
Dim DDSD_Tileset(1 To 255) As DDSURFACEDESC2

Public DDSD_Temp As DDSURFACEDESC2

Private LoadedDX7 As Boolean

Public Sub Render_Graphics()
Dim X As Long
Dim Y As Long
Dim L As Long
Dim rec As DxVBLib.RECT
Dim rec_pos As DxVBLib.RECT
' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    If frmMain.WindowState = vbMinimized Then
        Exit Sub
    End If
    
    If LoadedDX7 = False Then
        Exit Sub
    End If
    
    If DD.TestCooperativeLevel <> DD_OK Then
        Call DestroyDirectX
        Call InitDirectX
        Exit Sub
    End If
    
    ' fill it with black
    DDS_Backbuffer.BltColorFill rec_pos, 0
    
    For L = 1 To Layers.MaskAnim
        For X = 0 To MAX_MAP_X
            For Y = 0 To MAX_MAP_Y
                Call BltMapTile(X, Y, L)
            Next
        Next
    Next
    
    ' blt player
    
    For L = Layers.MaskAnim To Layers.Layer_Count
        For X = 0 To MAX_MAP_X
            For Y = 0 To MAX_MAP_Y
                Call BltMapTile(X, Y, L)
            Next
        Next
    Next
    
    
    TexthDC = DDS_Backbuffer.GetDC

    ' BLT TEXT
    'Call DrawText(TexthDC, 64, 64, "LOL!", QBColor(Pink))
    Call DrawRandomText

    DDS_Backbuffer.ReleaseDC TexthDC
    
    ' Get rec
    With rec
        .Top = 32
        .Bottom = 416
        .Left = 32
        .Right = 512
    End With
    
    ' rec_pos
    With rec_pos
        .Bottom = 384
        .Right = 480
    End With
    
    ' Render
    DX7.GetWindowRect frmMain.picScreen.hwnd, rec_pos
    DDS_Primary.Blt rec_pos, DDS_Backbuffer, rec, DDBLT_WAIT


' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "Render_Graphics", Err.Description)
Err.Clear
End Sub

Public Sub DestroyDirectX()
' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    
    Set DDS_Primary = Nothing
    Set DDS_Backbuffer = Nothing
    Set DD = Nothing
    
' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "DestroyDirectX", Err.Description)
Err.Clear
End Sub

Public Sub InitDirectX()
' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:
    
    ' Clear DirectX7
    'Call DestroyDirectX
    
    Set DD = DX7.DirectDrawCreate(vbNullString)
    
    DD.SetCooperativeLevel frmMain.hwnd, DDSCL_NORMAL
    
    With DDSD_Primary
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
        .lBackBufferCount = 1 ' One Backbuffer
    End With
    Set DDS_Primary = DD.CreateSurface(DDSD_Primary)
    
    Set DD_Clip = DD.CreateClipper(0)

    DD_Clip.SetHWnd frmMain.picScreen.hwnd
    DDS_Primary.SetClipper DD_Clip
    
    Call InitSurfaces
    
' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "InitDirectX", Err.Description)
Err.Clear
End Sub

Public Sub InitSurfaces()
' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    DDSD_Temp.lFlags = DDSD_CAPS
    DDSD_Temp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    
    Set DDS_Backbuffer = Nothing
    
    With DDSD_Backbuffer
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
        .lWidth = (14 + 3) * 32
        .lHeight = (11 + 3) * 32
    End With
    Set DDS_Backbuffer = DD.CreateSurface(DDSD_Backbuffer)
    
    ' Load persistent surfaces
    
    LoadedDX7 = True

' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "InitSurfaces", Err.Description)
Err.Clear
End Sub

Public Sub BltFast(ByVal dx As Long, ByVal dy As Long, ByRef ddS As DirectDrawSurface7, srcRECT As RECT, trans As CONST_DDBLTFASTFLAGS)
' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:
    
    If Not ddS Is Nothing Then
        Call DDS_Backbuffer.BltFast(dx, dy, ddS, srcRECT, trans)
    End If
    
' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "BltFast", Err.Description)
Err.Clear
End Sub

Public Function BltToDC(ByRef Surface As DirectDrawSurface7, sRECT As DxVBLib.RECT, dRECT As DxVBLib.RECT, ByRef picBox As VB.PictureBox, Optional Clear As Boolean = True) As Boolean
' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:
    
    If Clear Then
        picBox.Cls
    End If

    Call Surface.BltToDC(picBox.hDC, sRECT, dRECT)
    picBox.Refresh
    BltToDC = True
    
' Error Handler
Exit Function
errorhandler:
Call ReportError(Err.Number, "BltToDC", Err.Description)
Err.Clear
End Function

Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal X As Long, ByVal Y As Long)
Dim TmpR As RECT
Dim TmpDDSD As DDSURFACEDESC2
Dim TmpColorKey As DDCOLORKEY

    With TmpR
        .Left = X
        .Top = Y
        .Right = X
        .Bottom = Y
    End With

    TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

    With TmpColorKey
        .low = TheSurface.GetLockedPixel(X, Y)
        .high = .low
    End With

    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey
    TheSurface.Unlock TmpR
    
End Sub

Public Sub InitDDSurf(FileName As String, ByRef SurfDesc As DDSURFACEDESC2, ByRef Surf As DirectDrawSurface7)
    
    ' Set path
    FileName = App.Path & "\graphics\" & FileName & ".bmp"

    ' Destroy surface if it exist
    If Not Surf Is Nothing Then
        Set Surf = Nothing
        Call ZeroMemory(ByVal VarPtr(SurfDesc), LenB(SurfDesc))
    End If

    ' set flags
    SurfDesc.lFlags = DDSD_CAPS
    SurfDesc.ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
    
    ' init object
    Set Surf = DD.CreateSurfaceFromFile(FileName, SurfDesc)
    
    ' Set mask
    Call SetMaskColorFromPixel(Surf, 0, 0)

End Sub

Public Sub LoadGraphics()
Dim i As Long

    For i = 1 To 255
        If FileExist(App.Path & "\graphics\tilesets\" & i & ".bmp") = True Then
            If DDS_Tileset(i) Is Nothing Then
                Call InitDDSurf("tilesets\" & i, DDSD_Tileset(i), DDS_Tileset(i))
            End If
        Else
            NumTileset = i - 1
            Exit For
        End If
    Next

End Sub

Public Sub BltMapTile(ByVal X As Long, ByVal Y As Long, ByVal Layer As Long)
Dim rec As DxVBLib.RECT
Dim mapNum As Long
Dim TilesetNum As Byte

Player(MyIndex).Map = 1
mapNum = Player(MyIndex).Map

    If Map(mapNum).Tile(X, Y).Layer(Layer).X > 0 Or Map(mapNum).Tile(X, Y).Layer(Layer).Y > 0 Then ' There has to be an image set to blt
        'set the rec
        
        With rec
            rec.Top = Map(mapNum).Tile(X, Y).Layer(Layer).Y * 32
            rec.Bottom = .Top + 32
            rec.Left = Map(mapNum).Tile(X, Y).Layer(Layer).X * 32
            rec.Right = .Left + 32
        End With
        
        TilesetNum = Map(mapNum).Tile(X, Y).Layer(Layer).tileset
        If TilesetNum = 0 Then Exit Sub
        
        If X = 1 And Y = 0 Then
            Call BltFast(X * 32, Y * 32, DDS_Tileset(TilesetNum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            Call BltFast(X * 32, Y * 32, DDS_Tileset(TilesetNum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
        
        
    End If
End Sub

