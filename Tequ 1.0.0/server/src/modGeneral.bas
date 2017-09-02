Attribute VB_Name = "modGeneral"
Option Explicit

' Main timer
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' Halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal wMilliseconds As Long)

' API declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

' Clearing UDT's
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' Shell Executing for opening up files
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Loading and Saving Text
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Sub Main()
' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    ' Initializes the server and socket
    Call InitServer
    
    ' Loads the binary
    Call LoadData
    
    Map(1).Name = "LUCKY!"
    Map(1).MapPlayer(5) = 1
    
    ' Packets
    Call InitMessages
    
    frmServer.Show
    Call GameLoop
    
' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "Main", Err.Description)
Err.Clear
End Sub

Sub InitServer()
Dim i As Long
' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    ' Get the listening socket ready to go
    frmServer.Socket(0).Close
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = 7001
    
    For i = 1 To MAX_PLAYERS
        Load frmServer.Socket(i)
    Next
    
    frmServer.Socket(0).Listen

    frmServer.lblErrorNotification.Top = -15
    frmServer.lblErrorNotification.Left = frmServer.Width / 16 - frmServer.lblErrorNotification.Width
    Running = True
    
' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "InitServer", Err.Description)
Err.Clear
End Sub

Sub DestroyServer()
' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    Call SaveData

    Running = False
    Unload frmServer
    
' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "", Err.Description)
Err.Clear
End Sub

Sub GameLoop()
Dim tmr1000 As Long, tmr500 As Long, tmr250 As Long, tmr50 As Long
Dim Tick As Long
Dim LoopCount As Long
Dim i As Long
Dim LsIn As Long

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    Do While Running = True
        Tick = timeGetTime
        
        If tmr1000 < Tick Then
            ' FPS
            frmServer.Caption = "Server CPS: " & LoopCount
            LoopCount = 0
            
            tmr1000 = Tick + 1000
        End If
        
        If tmr500 < Tick Then
                If frmServer.lstIndex.ListIndex > 0 Then LsIn = frmServer.lstIndex.ListIndex
                frmServer.lstIndex.Clear
                For i = 1 To MAX_PLAYERS
                    frmServer.lstIndex.AddItem i & ": " & frmServer.Socket(i).RemoteHostIP & "  " & Trim$(Player(i).Name)
                Next
                frmServer.lstIndex.ListIndex = LsIn
            
            tmr500 = Tick + 500
        End If
        
        If tmr250 < Tick Then
            
            tmr250 = Tick + 250
        End If
        
        If tmr50 < Tick Then
        
            ' Error alert
            If ErrorAlert = True Then
                ' If it waited for 5 seconds and it's not at the bottom yet
                If AlertMsgWait = 5000 And frmServer.lblErrorNotification.Top <> -15 Then
                    frmServer.lblErrorNotification.Top = frmServer.lblErrorNotification.Top - 1
                    ' If it reached the bottom
                    If frmServer.lblErrorNotification.Top = -15 Then
                        ' Clear the data
                        AlertMsgWait = 0
                        ErrorAlert = False
                    End If
                End If
                ' If it didn't reach the top yet
                If frmServer.lblErrorNotification.Top < 15 And AlertMsgWait = 0 Then
                    frmServer.lblErrorNotification.Top = frmServer.lblErrorNotification.Top + 1
                End If
                ' If its at the top
                If frmServer.lblErrorNotification.Top >= 15 Then
                    AlertMsgWait = AlertMsgWait + 50
                End If
            End If
            
            tmr50 = Tick + 50
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
Exit Sub
errorhandler:
Call ReportError(Err.Number, "GetWinSockState", Err.Description)
Err.Clear
End Function

Public Sub ReportError(ByVal Number As Long, ByVal Source As String, ByVal Descr As String)
Dim FileName As String
Dim F As String

' Error Handler
If Options.Debug = True Then On Error GoTo errorhandler:

    FileName = App.Path & "\errs\" & WeekdayName(Weekday(Now), True) & " " & MonthName(Month(Now), True) & " " & Day(Now) & "  -  " & Hour(Now) & ";" & Minute(Now) & ";" & Second(Now) & " - RuntimeError " & Number & " - " & Descr & ".txt"
    ' Set the filename to the variable so we can open it if we want to.
    LastErrorDir = FileName
    F = FreeFile
    Open FileName For Output As #F
    Print #F, , "A Runtime Error of type " & Number & " occured on " & WeekdayName(Weekday(Now), False) & " " & MonthName(Month(Now), False) & " " & Day(Now) & " around " & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now) & ". Apparantly, the problem is '" & Descr & "'. It happened in the sub or function called '" & Source & "'."
    Close #F
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

    ' If frmServer isn't visible, you really messed up, so just open up the file!
    If frmServer.Visible = False And frmMenu.Visible = False Then
        Call OpenLastErrorReport
        Call DestroyGame
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
    Call ShellExecute(frmMenu.hWnd, "open", LastErrorDir, vbNullString, vbNullString, 1)

' Error Handler
Exit Sub
errorhandler:
Call ReportError(Err.Number, "OpenLastErrorReport", Err.Description)
Err.Clear
End Sub

Public Sub MakePlayer(ByVal index As Long, ByVal Name As String, ByVal Password As String)

    With Player(index)
        .Name = Name
        .Password = Password
        .Map = BOOT_MAP
        .x = BOOT_X
        .y = BOOT_Y
        .Sprite = 1
    End With
    
    Call SavePlayer(index, Name)

End Sub
