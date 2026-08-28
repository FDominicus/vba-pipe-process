VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' (c) 2013 randomnoun. All Rights Reserved. This work is licensed under a
' BSD Simplified License. (http://www.randomnoun.com/bsd-simplified.html)

' see http://www.randomnoun.com/wp/2013/10/07/its-all-made-out-of-pipes/
' for further documentation for this class


' event declarations

' is triggered when data is read from stdout
Public Event StdoutAvailable(strText As String)

' is triggered when data is read from stderr
Public Event StderrAvailable(strText As String)

' is triggered after process wait interval
Public Event Tick()


' Windows structure/constant/functions

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThStdoutRead As Long
    dwProcessId As Long
    dwThStdoutReadId As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Const WAIT_INFINITE         As Long = (-1&)
Private Const STARTF_USESHOWWINDOW  As Long = &H1
Private Const STARTF_USESTDHANDLES  As Long = &H100
Private Const SW_HIDE               As Long = &H0
Private Const SW_NORMAL             As Long = &H1
Private Const HANDLE_FLAG_INHERIT   As Long = &H1
Private Const HANDLE_FLAG_PROTECT_FROM_CLOSE As Long = &H2
Private Const DUPLICATE_SAME_ACCESS         As Long = &H2
Private Const NORMAL_PRIORITY_CLASS As Long = &H20
Private Const CREATE_UNICODE_ENVIRONMENT As Long = &H400
Private Const CREATE_NEW_PROCESS_GROUP As Long = &H200

Private Const CSIDL_WINDOWS         As Long = &H24
Private Const MAX_PATH              As Long = 260
Private Const SHGFP_TYPE_CURRENT    As Long = &H0
Private Const SHGFP_TYPE_DEFAULT    As Long = &H1

Private Const ERROR_NO_DATA         As Long = 232 ' winerror.h

Private Declare Function CreatePipe Lib "kernel32" (phStdoutReadPipe As Long, phStdoutWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Function SetHandleInformation Lib "kernel32" (ByVal hObject As Long, ByVal dwMask As Long, ByVal dwFlags As Long) As Long
Private Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long
Private Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpArguments As String, lpProcessAttributes As Any, lpThStdoutReadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long


' class variables

Dim boolSeparateStdoutStderr As Boolean
Dim boolShowWindow As Boolean
Dim strApplicationName As String
Dim strArguments As String
Dim strStdIn As String
Dim strStdOut As String
Dim strStdErr As String
Dim lngExitCode As Long
Dim lngErrNumber As Long
Dim strErrDescription As String
Dim boolIsRunning As Boolean
Dim boolHasRun As Boolean      ' i.e. whether exit codes/numbers in this class are valid
Dim boolTerminate As Boolean   ' set by event handlers to terminate process early
Dim lngStdInBufferPos

Dim lngTickWaitMillis As Long
Dim lngTimeoutMillis As Long     ' 0=no timeout
Dim lngTickCount As Long
Dim lngLastOutputTickCount As Long

Dim hProcess As Long  ' used to terminate process in class destructor

' property getters/setters

Public Property Let SeparateStdoutStderr(bSeparateStdoutStderr As Boolean)
  If boolIsRunning Then
    Err.Raise 1, "clsProcess", "Cannot set SeparateStdoutStderr property while process is running"
  End If
  boolSeparateStdoutStderr = bSeparateStdoutStderr
End Property
                       
Public Property Get SeparateStdoutStderr() As Boolean
  SeparateStdoutStderr = boolSeparateStdoutStderr
End Property

Public Property Let ShowWindow(bShowWindow As Boolean)
  If boolIsRunning Then
    Err.Raise 1, "clsProcess", "Cannot set ShowWindow property while process is running"
  End If
  boolShowWindow = bShowWindow
End Property
                       
Public Property Get ShowWindow() As Boolean
  ShowWindow = boolShowWindow
End Property

Public Property Let ApplicationName(sApplicationName As String)
  If boolIsRunning Then
    Err.Raise 1, "clsProcess", "Cannot set ApplicationName property while process is running"
  End If
  strApplicationName = sApplicationName
End Property
                       
Public Property Get ApplicationName() As String
  ApplicationName = strApplicationName
End Property

Public Property Let Arguments(sArguments As String)
  If boolIsRunning Then
    Err.Raise 1, "clsProcess", "Cannot set Arguments property while process is running"
  End If
  strArguments = sArguments
End Property
                       
Public Property Get Arguments() As String
  Arguments = strArguments
End Property

Public Property Let StdIn(sStdIn As String)
  If boolIsRunning Then
    Err.Raise 1, "clsProcess", "Cannot set StdIn property while process is running; use SendInput instead"
  End If
  strStdIn = sStdIn
End Property
                       
Public Property Get StdIn() As String
  StdIn = strStdIn
End Property

' stdout/stderr are read-only
Public Property Get StdOut() As String
  StdOut = strStdOut
End Property

Public Property Get StdErr() As String
  StdErr = strStdErr
End Property

' exit codes/numbers/description are read-only
Public Property Get ExitCode() As Long
  If boolIsRunning Then
    Err.Raise 1, "clsProcess", "Cannot get ExitCode property while process is running"
  ElseIf Not boolHasRun Then
    Err.Raise 1, "clsProcess", "Cannot get ExitCode property before process has been run"
  End If
  ExitCode = lngExitCode
End Property

Public Property Get errNumber() As Long
  If boolIsRunning Then
    Err.Raise 1, "clsProcess", "Cannot get ErrNumber property while process is running"
  ElseIf Not boolHasRun Then
    Err.Raise 1, "clsProcess", "Cannot get ErrNumber property before process has been run"
  End If
  errNumber = lngErrNumber
End Property

Public Property Get errDescription() As String
  If boolIsRunning Then
    Err.Raise 1, "clsProcess", "Cannot get ErrDescription property while process is running"
  ElseIf Not boolHasRun Then
    Err.Raise 1, "clsProcess", "Cannot get ErrDescription property before process has been run"
  End If
  errDescription = strErrDescription
End Property


Public Property Get Terminated() As Boolean
  Terminated = boolTerminate
End Property

Public Property Get TickCount() As Long
  TickCount = lngTickCount
End Property

Public Property Get LastOutputTickCount() As Long
  LastOutputTickCount = lngLastOutputTickCount
End Property

' can get/set tick wait millis or tick timeout millis
Public Property Let TickWaitMillis(lTickWaitMillis As Long)
  If boolIsRunning Then
    Err.Raise 1, "clsProcess", "Cannot set TickWaitMillis property while process is running"
  End If
  lngTickWaitMillis = lTickWaitMillis
End Property
                       
Public Property Get TickWaitMillis() As Long
  TickWaitMillis = lngTickWaitMillis
End Property

Public Property Let TimeoutMillis(lTimeoutMillis As Long)
  'If boolIsRunning Then
  '  Err.Raise 1, "clsProcess", "Cannot set TimeoutMillis property while process is running"
  'End If
  lngTimeoutMillis = lTimeoutMillis
End Property
                       
Public Property Get TimeoutMillis() As Long
  TimeoutMillis = lngTimeoutMillis
End Property


' class methods

Private Sub Class_Initialize()
  boolIsRunning = False
  boolHasRun = False
  boolTerminate = False
  lngExitCode = 0
  lngErrNumber = 0
  lngStdInBufferPos = 1 ' next char to send
  strErrDescription = ""
  lngTickWaitMillis = 100
  lngTickCount = 0
  lngLastOutputTickCount = 0
  lngTimeoutMillis = 0
End Sub

Private Sub Class_Terminate()
  ' Debug.Print "Class termination"
  If boolIsRunning And (hProcess <> 0) Then
    ' you'd better hope this isn't another process by the time this runs
    Debug.Print "Terminating still-running process"
    TerminateProcess hProcess, -1
    boolIsRunning = False
  End If
End Sub

' possibly allow access to boolIsRunning, boolHasRun variables
' in some kind of enumerated state, no doubt

' may be called within stdout/stderr event handlers to pass additional data to a process
Public Sub SendInput(strText As String)
  strStdIn = strStdIn & strText
End Sub

Public Sub Terminate()
  boolTerminate = True
End Sub

                      
' same parameters as fncRedirectProcess
' (you can use this method to assist migrating code from fncRedirectProcess to clsProcess)
Public Function ExecuteParams(sApplicationName As String, sArguments As String, _
  Optional sStdIn As String = "", Optional bShowWindow As Boolean = False, _
  Optional bSeparateStdoutStderr As Boolean = False) As RedirectProcessResult
  
  strApplicationName = sApplicationName
  strArguments = sArguments
  strStdIn = sStdIn
  boolShowWindow = bShowWindow
  boolSeparateStdoutStderr = bSeparateStdoutStderr
  Execute
End Function


Public Sub Execute()

  ' TODO: add ACLs to this
  ' TODO: add elements to process environment block

  Dim saProcess       As SECURITY_ATTRIBUTES
  Dim saThread        As SECURITY_ATTRIBUTES
  Dim saPipe          As SECURITY_ATTRIBUTES
  Dim tProcessInfo    As PROCESS_INFORMATION
  Dim tStartupInfo    As STARTUPINFO
  Dim hStdoutReadTmp As Long, hStderrReadTmp As Long, hStdinWriteTmp As Long
  Dim hStdoutRead As Long, hStderrRead As Long, hStdinRead As Long
  Dim hStdoutWrite As Long, hStderrWrite As Long, hStdinWrite As Long
  Dim lngReadBytes    As Long
  Dim abytReadBuf()   As Byte
  Dim strWriteBuf     As String
  Dim lngResult, lngResult2 As Long
  Dim strFullArguments  As String
  Dim lngSizeOf       As Long
  Dim strPath         As String
  Dim strEnv          As String
  Dim abytEnv()       As Byte
  
  If boolIsRunning Then
    Err.Raise 1, "clsProcess", "Cannot call Execute while process is running"
  End If
  
  lngExitCode = -1
  lngErrNumber = 0
  strErrDescription = ""
  lngStdInBufferPos = 1
  lngTickCount = 0
  lngLastOutputTickCount = 0
  
  ' input validation
  If (strApplicationName = "") Then subSetError 100, "Missing strApplicationName"
  If (lngErrNumber <> 0) Then Exit Sub
  
  ' assign default security descriptor associated with access token of the calling process.
  saPipe.nLength = Len(saPipe)
  saPipe.bInheritHandle = 1&
  saPipe.lpSecurityDescriptor = 0&
  
  saProcess.nLength = Len(saProcess)
  saProcess.bInheritHandle = 1&
  saProcess.lpSecurityDescriptor = 0&
  
  saThread.nLength = Len(saThread)
  saThread.bInheritHandle = 1&
  saThread.lpSecurityDescriptor = 0&

  ' create pipes
  
  ' for stdout (and possibly stderr)
  If (CreatePipe(hStdoutReadTmp, hStdoutWrite, saPipe, 0&) = 0&) Then
    subSetError 1, "CreatePipe failes on tmp stdout"
    Exit Sub
  End If
  
  ' for stderr
  If boolSeparateStdoutStderr Then
    ' separate stdout/stderr pipes
    If (CreatePipe(hStderrReadTmp, hStderrWrite, saPipe, 0&) = 0&) Then
      subSetError 2, "CreatePipe failed on tmp stderr"
      Exit Sub
    End If
  Else
    ' create a duplicate of the stdout handle here for stderr
    ' (in case child decides to close one of them)
    If (DuplicateHandle(GetCurrentProcess(), hStdoutWrite, GetCurrentProcess(), hStderrWrite, 0, True, DUPLICATE_SAME_ACCESS) = 0&) Then
      subSetError 3, "DuplicateHandle failed on stdout/stderr"
      Exit Sub
    End If
  End If
  
  ' for stdin
  If (CreatePipe(hStdinRead, hStdinWriteTmp, saPipe, 0&) = 0&) Then
    subSetError 4, "CreatePipe failed on tmp stdin"
    Exit Sub
  End If
  
  ' duplicate handles: get the "real" handles from the tmp handles, with Properties
  ' set to FALSE. this gives us closeable handles to the pipes.
  ' see http://support.microsoft.com/kb/190351
  If (DuplicateHandle(GetCurrentProcess(), hStdoutReadTmp, GetCurrentProcess(), hStdoutRead, 0, False, DUPLICATE_SAME_ACCESS) = 0&) Then
    subSetError 5, "DuplicateHandle failed on stdout"
    Exit Sub
  End If
  
  If boolSeparateStdoutStderr Then
    If (DuplicateHandle(GetCurrentProcess(), hStderrReadTmp, GetCurrentProcess(), hStderrRead, 0, False, DUPLICATE_SAME_ACCESS) = 0&) Then
      subSetError 6, "DuplicateHandle failed on stderr"
      Exit Sub
    End If
  End If
  If (DuplicateHandle(GetCurrentProcess(), hStdinWriteTmp, GetCurrentProcess(), hStdinWrite, 0, False, DUPLICATE_SAME_ACCESS) = 0&) Then
    subSetError 7, "DuplicateHandle failed on stdin"
    Exit Sub
  End If

  ' Close inheritable copies of the handles we do not want to be inherited.
  If (CloseHandle(hStdoutReadTmp) = 0) Then
    subSetError 8, "CloseHandle failed on tmp stdout"
    Exit Sub
  End If
  ' again, probably don't do this if we're using same handle for stdout/err
  If boolSeparateStdoutStderr Then
    If (CloseHandle(hStderrReadTmp) = 0) Then
      subSetError 9, "CloseHandle failed on tmp stderr"
      Exit Sub
    End If
  End If
  If (CloseHandle(hStdinWriteTmp) = 0) Then
    subSetError 10, "CloseHandle failed on tmp stdin"
    Exit Sub
  End If
  
  tStartupInfo.cb = Len(tStartupInfo)
  GetStartupInfo tStartupInfo
  tStartupInfo.cb = Len(tStartupInfo)
  tStartupInfo.hStdOutput = hStdoutWrite
  tStartupInfo.hStdError = hStderrWrite
  tStartupInfo.hStdInput = hStdinRead
  tStartupInfo.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
  tStartupInfo.wShowWindow = IIf(boolShowWindow, SW_NORMAL, SW_HIDE)
  
  ' full command sent to CreateProcess
  strFullArguments = """" & strApplicationName & """" & " " & strArguments
  
  ' define SystemRoot environment variable.
  ' if this isn't here, then TCP network applications will fail with the error
  '   The requested service provider could not be loaded or initialized.
  ' because it can't load mswsock.dll (the path to it contains "%SystemRoot" in the protocol section of the winsock registry)
  ' the other env vars that always seem to be set by CreateProcess (or possibly cmd.exe, which I used for testing) are
  '   "COMSPEC=C:\WINDOWS\system32\cmd.exe" & Chr(0) & _
  '   "PATHEXT=.COM;.EXE;.BAT;.CMD;.VBS;.JS;.WS" & Chr(0) & _
  '   "PROMPT=$P$G" & Chr(0) &
  strPath = String(MAX_PATH, 0)
  If (SHGetFolderPath(0, CSIDL_WINDOWS, 0, SHGFP_TYPE_CURRENT, strPath) <> 0) Then
    subSetError 11, "SHGetFolderPath failed"
    Exit Sub
  End If
  strPath = Left(strPath, InStr(1, strPath, Chr(0)) - 1) ' just up to null terminator
  strEnv = "SystemRoot=" & strPath & Chr(0) & Chr(0)  ' sz terminator + env block terminator
  abytEnv = StrConv(strEnv, vbFromUnicode)            ' encode as byte array
    
  ' if the executable module is a 16-bit application, lpApplicationName should be NULL
  lngResult = CreateProcess(0&, strFullArguments, saProcess, saThread, 1&, _
    NORMAL_PRIORITY_CLASS, abytEnv(0), vbNullString, tStartupInfo, tProcessInfo)
  If (lngResult = 0&) Then
    subSetError 12, "CreateProcess failed"
    Exit Sub
  End If
  
  boolIsRunning = True
  hProcess = tProcessInfo.hProcess

  ' XXX: possibly terminate process if errors occur from here on

  ' Close pipe handles (do not continue to modify the parent).
  ' We need to make sure that no handles to the write end of the output pipe are maintained
  ' in this process or else the pipe will not close when the child process exits.
  ' Probably not an issue since we don't use a blocking ReadFile later, but this is consistent
  ' with the microsoft kb article above.
  If (CloseHandle(hStdoutWrite) = 0) Then
    subSetError 13, "CloseHandle failed"
    Exit Sub
  End If
  If (CloseHandle(hStdinRead) = 0) Then
    subSetError 14, "CloseHandle failed"
    Exit Sub
  End If
  If (CloseHandle(hStderrWrite) = 0) Then
    subSetError 15, "CloseHandle failed"
    Exit Sub
  End If
  

  ' first first stdin block; TODO: may exceed buffer size ?
  If (WriteFile(hStdinWrite, ByVal strStdIn, Len(strStdIn), lngResult, ByVal 0&) = 0) Then
    subSetError 16, "WriteFile failed"
    Exit Sub
  End If
  
  ' Debug.Print "wrote " & lngResult & " bytes to stdin; buflen was " & Len(strStdIn)
  lngStdInBufferPos = lngResult + 1
  
  lngResult = WaitForSingleObject(tProcessInfo.hProcess, lngTickWaitMillis)
  Do
    DoEvents
    Select Case lngResult
      Case 258& ' 500ms timeout
        ' keep on trucking.
        lngResult = WaitForSingleObject(tProcessInfo.hProcess, lngTickWaitMillis)
      Case &H80, &HFFFFFFFF  ' abandoned / failed
        subSetError 17, "Wait abandoned/failed (" & lngResult & ")"
        Exit Sub
      Case 0
        ' wait complete
      Case Else
        subSetError 18, "WaitForSingleObject failed (" & lngResult & ")"
        lngExitCode = -5
        Exit Sub
    End Select
    
    ' pump the i/o stream pipes
    lngSizeOf = GetFileSize(hStdoutRead, 0&)
    If (lngSizeOf > 0) Then
      ReDim abytReadBuf(lngSizeOf - 1)
      If ReadFile(hStdoutRead, abytReadBuf(0), UBound(abytReadBuf) + 1, lngReadBytes, ByVal 0&) = 0 Then
        subSetError 19, "ReadFile failed"
        Exit Sub
      Else
        'Debug.Print "read-stdout: " & StrConv(abytReadBuf, vbUnicode)
        strStdOut = strStdOut & StrConv(abytReadBuf, vbUnicode)
        RaiseEvent StdoutAvailable(StrConv(abytReadBuf, vbUnicode))
        lngLastOutputTickCount = lngTickCount
      End If
    End If
    
    If boolSeparateStdoutStderr Then
      lngSizeOf = GetFileSize(hStderrRead, 0&)
      If (lngSizeOf > 0) Then
        ReDim abytReadBuf(lngSizeOf - 1)
        If ReadFile(hStderrRead, abytReadBuf(0), UBound(abytReadBuf) + 1, lngReadBytes, ByVal 0&) = 0 Then
          subSetError 20, "Read failed"
          Exit Sub
        Else
          'Debug.Print "read-stderr: " & StrConv(abytReadBuf, vbUnicode)
          strStdErr = strStdErr & StrConv(abytReadBuf, vbUnicode)
          RaiseEvent StderrAvailable(StrConv(abytReadBuf, vbUnicode))
          lngLastOutputTickCount = lngTickCount
        End If
      End If
    End If
    
    ' we send input after reading stdout/stderr since those events probably
    ' created the input in the first place
    ' stdStdin had better not change whilst this is executing...
    If Len(strStdIn) >= lngStdInBufferPos Then
      ' Debug.Print "sending '" & Mid(strStdIn, lngStdInBufferPos) & "'"
      strWriteBuf = Mid(strStdIn, lngStdInBufferPos) & Chr(0)
      If (WriteFile(hStdinWrite, ByVal strWriteBuf, Len(strWriteBuf) - 1, lngResult2, ByVal 0&) = 0) Then ' lngresult2==#bytes written
        If Err.LastDllError = ERROR_NO_DATA Then ' Pipe was closed (normal exit path).
          lngResult = 0
        Else
          subSetError 21, "WriteFile failed"
          Exit Sub
        End If
      End If
      lngStdInBufferPos = lngStdInBufferPos + lngResult2
    End If
    
    ' tick events will be at least lngTickWaitMillis apart + time taken to process streams and events above
    RaiseEvent Tick
    lngTickCount = lngTickCount + 1
    If (lngTimeoutMillis > 0) Then
      If ((lngTickCount - lngLastOutputTickCount) * lngTickWaitMillis) > lngTimeoutMillis Then
        boolTerminate = True
      End If
    End If
    If boolTerminate Then
      TerminateProcess tProcessInfo.hProcess, -1
    End If
    
  Loop Until lngResult = 0
  
  Call GetExitCodeProcess(tProcessInfo.hProcess, lngExitCode)
  lngErrNumber = 0
  strErrDescription = ""
  boolIsRunning = False: boolHasRun = True
  
  'Debug.Print "============="
  'Debug.Print "stdout: " & fncRedirectProcess.strStdOut
  'Debug.Print "stderr: " & fncRedirectProcess.strStdErr
  
  ' not too concerned about reporting error conditions from here on
  CloseHandle tProcessInfo.hThStdoutRead
  CloseHandle tProcessInfo.hProcess
  CloseHandle hStdoutRead
  CloseHandle hStderrRead
  CloseHandle hStdinWrite

End Sub


Private Sub subSetError(errNumber As Long, errDescription As String)
  lngErrNumber = errNumber
  strErrDescription = errDescription & " (" & Err.LastDllError & ")"
  boolIsRunning = False: boolHasRun = True
End Sub



