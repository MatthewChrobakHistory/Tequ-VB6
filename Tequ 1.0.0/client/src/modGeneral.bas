Attribute VB_Name = "modGeneral"
Option Explicit

' Main timer
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' Halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal wMilliseconds As Long)

' API declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

' Clearing UDT's
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' Shell Executing for opening up files
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Loading and Saving Text
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

' Master Object
Public DX7 As New DirectX7

Sub Main()
Dim Answer As Integer
' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    Call InitClient
    Call LoadOptions
    frmMenu.Visible = True
    
    If Options.InstallRuntimes = True Then
        Answer = MsgBox("Is this your first time using a Tequ program? If so, you'll need to download and register some DLL's." & vbCrLf & vbCrLf & "Don't worry! You can use the Origins Runtimes installer to install and register all the necessary files. Would you like to run the installer?", vbYesNo, "DLL Installer Prompt")
        If Answer = vbYes Then
            MsgBox "If the game doesn't work after installing the runtimes, then run the Runtimes.exe file located in the same folder as the client.", , "DLL Installer Prompt"
            Call ShellExecute(frmMain.hwnd, "runas", App.Path & "\runtimes.exe", "", App.Path, vbNormalFocus)
            Options.InstallRuntimes = False
            SaveOptions
        ElseIf Answer = vbNo Then
            Options.InstallRuntimes = False
            SaveOptions
        End If
    End If
    
' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "Main", Err.Description)
Err.Clear
End Sub

Sub InitGame(Optional ByVal Multiplayer As Boolean = False)
Dim Resolved As Boolean
Dim Tick As Long, tmr3000 As Long

    If Multiplayer = True Then
    
        If Trim$(frmMenu.txtPass.Text) <> Trim$(frmMenu.txtRetype.Text) Then
            MsgBox "Passwords don't match."
            Exit Sub
        End If
        
        ' setup the socket
        With frmMain.Socket
            .Close
            .RemoteHost = Options.IP
            .RemotePort = Options.Port
            .Connect
            SocketState = State_Connecting
        End With
        Call InitMessages
        
        frmMenu.Hide
        
        'connection failed
        tmr3000 = timeGetTime + 3000
        Do While Resolved = False
            Tick = timeGetTime
            If tmr3000 < Tick Then
                If frmMain.Socket.State <> 7 Then
                    DestroyGame
                    MsgBox "Connection failed!", vbCritical, "Failed to connect"
                    frmMenu.Show
                    Resolved = True
                End If
            End If
            If frmMain.Socket.State = 7 Then
                Resolved = True
                frmMenu.Show
            End If
            DoEvents
            Sleep 1
        Loop
        
    Else
        MyIndex = 1
        frmMain.Caption = "Tequ Client"
        Call EnterGame
    End If

End Sub

Sub InitClient()

    Call LoadFonts
    
    ' Setup the error notification system
    frmMenu.lblErrorNotification.Top = -15
    frmMenu.lblErrorNotification.Left = frmMenu.Width / 16 - frmMenu.lblErrorNotification.Width

End Sub

Sub DestroyClient()

    Call SaveOptions
    End
    
End Sub

Sub DestroyGame()
' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    If Options.OnlineMode = False Then
        Call SavePlayer(Trim$(Player(MyIndex).Name))
        Call SaveData
    End If

    With frmMenu
        .picCreate.Visible = False
        .picHomeScreen.Visible = True
    End With

    SocketState = State_Closed
    frmMain.Socket.Close
    Running = False
    Call DestroyDirectX
    frmMenu.Show
    Unload frmMain
    
' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "DestroyGame", Err.Description)
Err.Clear
End Sub

Sub GameLoop()
Dim tmr1000 As Long, tmr500 As Long, tmr250 As Long, tmr50 As Long
Dim Tick As Long
Dim LoopCount As Long
Dim Cnctd2Srvr As Boolean

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    Do While Running = True
        Tick = timeGetTime
        
        If tmr1000 < Tick Then
            ' FPS
            frmMain.Caption = "Tequ Client: " & GetWinSockState(frmMain.Socket.State) & "     CPS:" & LoopCount
            LoopCount = 0
            
            tmr1000 = Tick + 1000
        End If
        
        If tmr500 < Tick Then
        
            tmr500 = Tick + 500
        End If
        
        If tmr250 < Tick Then
        
            tmr250 = Tick + 250
        End If
        
        If tmr50 < Tick Then
        
            ' Error alert
            If ErrorAlert = True Then
                If frmMain.Visible = True Then
                    ' If it waited for 5 seconds and it's not at the bottom yet
                    If AlertMsgWait = 5000 And frmMain.lblErrorNotification.Top <> -15 Then
                        frmMain.lblErrorNotification.Top = frmMain.lblErrorNotification.Top - 1
                        ' If it reached the bottom
                        If frmMain.lblErrorNotification.Top = -15 Then
                            ' Clear the data
                            AlertMsgWait = 0
                            ErrorAlert = False
                        End If
                    End If
                    ' If it didn't reach the top yet
                    If frmMain.lblErrorNotification.Top < 15 And AlertMsgWait = 0 Then
                        frmMain.lblErrorNotification.Top = frmMain.lblErrorNotification.Top + 1
                    End If
                    ' If its at the top
                    If frmMain.lblErrorNotification.Top >= 15 Then
                        AlertMsgWait = AlertMsgWait + 50
                    End If
                Else
                    If frmMenu.Visible = True Then
                        ' If it waited for 5 seconds and it's not at the bottom yet
                        If AlertMsgWait = 5000 And frmMenu.lblErrorNotification.Top <> -15 Then
                            frmMenu.lblErrorNotification.Top = frmMenu.lblErrorNotification.Top - 1
                            ' If it reached the bottom
                            If frmMenu.lblErrorNotification.Top = -15 Then
                                ' Clear the data
                                AlertMsgWait = 0
                                ErrorAlert = False
                            End If
                        End If
                        ' If it didn't reach the top yet
                        If frmMenu.lblErrorNotification.Top < 1 And AlertMsgWait = 0 Then
                            frmMenu.lblErrorNotification.Top = frmMenu.lblErrorNotification.Top + 1
                        End If
                        ' If its at the top
                        If frmMenu.lblErrorNotification.Top >= 1 Then
                            AlertMsgWait = AlertMsgWait + 50
                        End If
                            
                    End If
                End If
            End If
            
            tmr50 = Tick + 50
        End If
        
        Call Render_Graphics
        
        ' Disconnection
        If Options.OnlineMode = True Then
            If SocketState <> State_Connected And frmMain.Socket.State = 7 Then
                SocketState = State_Connected
            End If
            If SocketState = State_Connected And frmMain.Socket.State <> 7 Then
                DestroyGame
                MsgBox "Disconnected from the server.", vbCritical, "Server Disconnection"
            End If
        End If
        
        LoopCount = LoopCount + 1
        
        DoEvents
        Sleep 1
    Loop

' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "GameLoop", Err.Description)
Err.Clear
End Sub

' Checks if a directory exists, if it doesn't, it makes it
Public Sub CheckDir(ByVal Path As String, ByVal Directory As String)

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    If LCase$(Dir$(Path & Directory, vbDirectory)) <> Directory Then Call MkDir(Path & Directory)
    
' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "CheckDir", Err.Description)
Err.Clear
End Sub

' Checks to see if a file exists
Public Function FileExist(ByVal FileName As String) As Boolean

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    If LenB(Dir$(FileName)) > 0 Then FileExist = True

' Error Handler
Exit Function
errorhandler:
Call ReportError(Err.Number, "FileExist", Err.Description)
Err.Clear
End Function

' Retrieves a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

' Error Handler
Exit Function
errorhandler:
Call ReportError(Err.Number, "GetVar", Err.Description)
Err.Clear
End Function

' Writes a string in a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    Call WritePrivateProfileString$(Header, Var, Value, File)
    
' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "PutVar", Err.Description)
Err.Clear
End Sub

Public Function GetWinSockState(ByVal State As Byte) As String

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:
    
    Select Case State
        Case sckClosed
            GetWinSockState = " Connection closed "
        Case sckOpen
            GetWinSockState = " Open "
        Case sckListening
            GetWinSockState = " Listening for incoming connections "
        Case sckConnectionPending
            GetWinSockState = " Connection pending "
        Case sckResolvingHost
            GetWinSockState = " Resolving remote host name "
        Case sckHostResolved
            GetWinSockState = " Remote host name successfully resolved "
        Case sckConnecting
            GetWinSockState = " Connecting to remote host "
        Case sckConnected
            GetWinSockState = " Connected to remote host "
        Case sckClosing
            GetWinSockState = " Connection is closing "
        Case sckError
            GetWinSockState = " Error occured "
        Case Else
            GetWinSockState = " Null "
    End Select

' Error Handler
Exit Function
errorhandler:
Call ReportError(Err.Number, "GetWinSockState", Err.Description)
Err.Clear
End Function

Public Sub ReportError(ByVal Number As Long, ByVal Source As String, ByVal Descr As String)
Dim FileName As String
Dim f As String

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    FileName = App.Path & "\errs\" & WeekdayName(Weekday(Now), True) & " " & MonthName(Month(Now), True) & " " & Day(Now) & "  -  " & Hour(Now) & ";" & Minute(Now) & ";" & Second(Now) & " - RuntimeError " & Number & " - " & Descr & ".txt"
    ' Set the filename to the variable so we can open it if we want to.
    LastErrorDir = FileName
    f = FreeFile
    Open FileName For Output As #f
    Print #f, , "A Runtime Error of type " & Number & " occured on " & WeekdayName(Weekday(Now), False) & " " & MonthName(Month(Now), False) & " " & Day(Now) & " around " & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now) & ". Apparantly, the problem is '" & Descr & "'. It happened in the sub or function called '" & Source & "'."
    Close #f
    ' Send a notification that an error occured
    Call DisplayErrorNotification

' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "ReportError", Err.Description)
Err.Clear
End Sub

Public Sub DisplayErrorNotification()

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    ' If frmMain isn't visible and frmMenu isn't, you really messed up, so just open up the file!
    If frmMain.Visible = False And frmMenu.Visible = False Then
        Call OpenLastErrorReport
        End
    End If
    
    ErrorAlert = True
    
' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "DisplayErrorNotification", Err.Description)
Err.Clear
End Sub

Public Sub OpenLastErrorReport()

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    If FileExist(LastErrorDir) = False Then Exit Sub ' Exit out in case the file doesn't exist
    Call ShellExecute(frmMenu.hwnd, "open", LastErrorDir, vbNullString, vbNullString, 1)

' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "OpenLastErrorReport", Err.Description)
Err.Clear
End Sub

Public Sub EnterGame()

    ' Setup the error notification system
    frmMain.lblErrorNotification.Top = -15
    frmMain.lblErrorNotification.Left = frmMain.Width / 16 - frmMain.lblErrorNotification.Width
    
    ' Load up DirectX7. Clear it first
    Call DestroyDirectX
    Call InitDirectX
    
    Call SetFont(GameFont, FONT_SIZE)
    
    Call LoadGraphics

    Call LoadData
    
    Running = True
    frmMenu.Visible = False
    frmMain.Visible = True
    frmMain.cmbText.ListIndex = GetFontNumber(Options.GameFont)
    Call GameLoop

End Sub
