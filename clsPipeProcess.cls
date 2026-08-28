VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPipeProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' Copyright (c) 2013 randomnoun (knoxg)   - original 32-bit implementation
' Copyright (c) 2026 Friedrich Dominicus, Q-Software Solutions GmbH - 64-bit port
'
' This work is licensed under a BSD 2-Clause "Simplified" License.
' See LICENSE in the repository root.
'
' Original:  https://www.randomnoun.com/wp/2013/10/08/its-all-made-out-of-pipes/
' This fork: https://github.com/FDominicus/vba-pipe-process
'
'==============================================================================
' 64-BIT PORT
'
' All API declarations rewritten with #If VBA7 / PtrSafe; every HANDLE and
' pointer is LongPtr. The structures are what matter: STARTUPINFO grows from
' 68 to 104 bytes, and one member of the wrong width makes CreateProcess
' receive a malformed structure - Office then dies with no error message.
'
' Verified against the Microsoft documentation and the mingw-w64 headers.
'
' Fixed: the original declares lpReserved2 As Byte. It is a pointer (LPBYTE).
' On 32-bit that works by accident, because VBA pads the single byte with
' three more. On 64-bit it does not.
'
' Fixed: Dim lngResult, lngResult2 As Long declared lngResult as a Variant -
' in VBA only the last name in such a list receives the type.
'
' NOT changed: the write end of the stdin pipe is never closed. A program
' that reads stdin will therefore wait for more input after the StdIn string
' has been consumed. Send what ends it - "exit" for cmd - or set a timeout.
'
' Cosmetic, left alone to keep the diff readable: PROCESS_INFORMATION.
' hThStdoutRead is hThread, dwThStdoutReadId is dwThreadId, and
' CreateProcess takes lpThStdoutReadAttributes where the API says
' lpThreadAttributes. A global search/replace of "read" in the original
' caught the middle of "hThread". Also lpCurrentDriectory.
'
'------------------------------------------------------------------------------
' CHANGES 2026-08-28 - measured against the canonical C pattern (KB 190351
' and the implementations that copy it: git, python subprocess, wine)
'
' FIXED - the child process had exactly one environment variable.
'   The original passed a block containing only SystemRoot. A non-NULL
'   lpEnvironment REPLACES the environment; the child had no PATH, TEMP,
'   APPDATA or JAVA_HOME. gpg looks for its keyring under %APPDATA%\gnupg
'   and would have reported "no key" after a successful import; Java needs
'   %TEMP%. Now NULL, so the child inherits Access's environment. Use
'   AddEnvironment to add to it - that path builds the FULL block.
'
' FIXED - every failure path leaked handles.
'   21 x "subSetError n" followed by Exit Sub, none of them closing
'   anything. One CleanUp label now, reached by success and failure alike;
'   CloseHandleSafe zeroes each handle so nothing is closed twice.
'   A failure after CreateProcess also terminates the child instead of
'   leaving it orphaned with closed pipes.
'
' Error codes 8, 9, 10, 11, 13, 14, 15 are no longer raised. Seven API
' declarations and two locals went with them.
'
' FIXED - polling with GetFileSize, and no final read.
'   GetFileSize is not documented for anonymous pipes; PeekNamedPipe is the
'   call for this. And the loop only ever read what one tick found, so
'   whatever a tool wrote between the last tick and its exit was lost - for
'   a validator that is the verdict itself. PumpPipe now peeks while the
'   process runs and reads to the end after it stops.
'
'   Checked against the Rust standard library (library/std/src/sys):
'   Rust does not use CreatePipe for stdio at all. Anonymous pipes cannot do
'   overlapped operations, so it builds a named pipe with a unique name and
'   FILE_FLAG_OVERLAPPED and waits on both ends with WaitForMultipleObjects.
'   Cleaner, but it needs pointers that stay valid for the whole operation -
'   Rust calls that "super unsafe" in its own comment, and in VBA it is not
'   defensible. Polling is the right answer for an anonymous pipe; only the
'   call was wrong.
'
' NOT an issue here, though it looks like one - handle inheritance.
'   bInheritHandles is 1 with no handle list, so the child inherits every
'   inheritable handle this process holds. Rust does not solve that with
'   PROC_THREAD_ATTRIBUTE_HANDLE_LIST but with a global lock around the
'   whole spawn (see KB 315939): the race is two threads spawning at once,
'   where one child inherits the other's pipes. VBA is single-threaded, so
'   it cannot arise from Access.
'
' STILL OPEN:
'   - CreateProcessA is ANSI: paths outside the system code page will fail.
'==============================================================================


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
    lpSecurityDescriptor As LongPtr
    bInheritHandle As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThStdoutRead As LongPtr
    dwProcessId As Long
    dwThStdoutReadId As Long
End Type

' 64 Bit: 104 Bytes, 32 Bit: 68. Jeder Zeiger MUSS LongPtr sein, sonst
' bekommt CreateProcess eine falsch grosse Struktur und Access stirbt
' ohne Meldung.
'
' Im Original stand lpReserved2 As Byte. Auf 32 Bit ging das nur durch
' Zufall gut: 1 Byte plus 3 Fuellbytes ergaben versehentlich die richtigen
' vier. Es ist ein Zeiger (LPBYTE).
Private Type STARTUPINFO
    cb As Long
    lpReserved As LongPtr
    lpDesktop As LongPtr
    lpTitle As LongPtr
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
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
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

Private Const ERROR_NO_DATA         As Long = 232 ' winerror.h
Private Const DRAIN_BUFFER          As Long = 8192

Private Declare PtrSafe Function CreatePipe Lib "kernel32" (phStdoutReadPipe As LongPtr, phStdoutWritePipe As LongPtr, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare PtrSafe Function SetHandleInformation Lib "kernel32" (ByVal hObject As LongPtr, ByVal dwMask As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As LongPtr, ByVal hSourceHandle As LongPtr, ByVal hTargetProcessHandle As LongPtr, lpTargetHandle As LongPtr, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
Private Declare PtrSafe Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As LongPtr, ByVal lpArguments As String, lpProcessAttributes As Any, lpThStdoutReadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, lpExitCode As Long) As Long


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

Dim hProcess As LongPtr   ' used to terminate process in class destructor

' Ergaenzungen zur Umgebung des Kindprozesses. Nothing = erben, siehe
' AddEnvironment.
Dim m_Env As Object

' property getters/setters

Public Property Let SeparateStdoutStderr(bSeparateStdoutStderr As Boolean)
  If boolIsRunning Then
    Err.Raise 1, "clsPipeProcess", "Cannot set SeparateStdoutStderr property while process is running"
  End If
  boolSeparateStdoutStderr = bSeparateStdoutStderr
End Property
                       
Public Property Get SeparateStdoutStderr() As Boolean
  SeparateStdoutStderr = boolSeparateStdoutStderr
End Property

Public Property Let ShowWindow(bShowWindow As Boolean)
  If boolIsRunning Then
    Err.Raise 1, "clsPipeProcess", "Cannot set ShowWindow property while process is running"
  End If
  boolShowWindow = bShowWindow
End Property
                       
Public Property Get ShowWindow() As Boolean
  ShowWindow = boolShowWindow
End Property

Public Property Let ApplicationName(sApplicationName As String)
  If boolIsRunning Then
    Err.Raise 1, "clsPipeProcess", "Cannot set ApplicationName property while process is running"
  End If
  strApplicationName = sApplicationName
End Property
                       
Public Property Get ApplicationName() As String
  ApplicationName = strApplicationName
End Property

Public Property Let Arguments(sArguments As String)
  If boolIsRunning Then
    Err.Raise 1, "clsPipeProcess", "Cannot set Arguments property while process is running"
  End If
  strArguments = sArguments
End Property
                       
Public Property Get Arguments() As String
  Arguments = strArguments
End Property

Public Property Let StdIn(sStdIn As String)
  If boolIsRunning Then
    Err.Raise 1, "clsPipeProcess", "Cannot set StdIn property while process is running; use SendInput instead"
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
    Err.Raise 1, "clsPipeProcess", "Cannot get ExitCode property while process is running"
  ElseIf Not boolHasRun Then
    Err.Raise 1, "clsPipeProcess", "Cannot get ExitCode property before process has been run"
  End If
  ExitCode = lngExitCode
End Property

Public Property Get errNumber() As Long
  If boolIsRunning Then
    Err.Raise 1, "clsPipeProcess", "Cannot get ErrNumber property while process is running"
  ElseIf Not boolHasRun Then
    Err.Raise 1, "clsPipeProcess", "Cannot get ErrNumber property before process has been run"
  End If
  errNumber = lngErrNumber
End Property

Public Property Get errDescription() As String
  If boolIsRunning Then
    Err.Raise 1, "clsPipeProcess", "Cannot get ErrDescription property while process is running"
  ElseIf Not boolHasRun Then
    Err.Raise 1, "clsPipeProcess", "Cannot get ErrDescription property before process has been run"
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
    Err.Raise 1, "clsPipeProcess", "Cannot set TickWaitMillis property while process is running"
  End If
  lngTickWaitMillis = lTickWaitMillis
End Property
                       
Public Property Get TickWaitMillis() As Long
  TickWaitMillis = lngTickWaitMillis
End Property

Public Property Let TimeoutMillis(lTimeoutMillis As Long)
  'If boolIsRunning Then
  '  Err.Raise 1, "clsPipeProcess", "Cannot set TimeoutMillis property while process is running"
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
  Dim hStdoutReadTmp As LongPtr, hStderrReadTmp As LongPtr, hStdinWriteTmp As LongPtr
  Dim hStdoutRead As LongPtr, hStderrRead As LongPtr, hStdinRead As LongPtr
  Dim hStdoutWrite As LongPtr, hStderrWrite As LongPtr, hStdinWrite As LongPtr
  Dim strWriteBuf     As String
  Dim lngResult As Long, lngResult2 As Long
  Dim strFullArguments  As String
  Dim abytEnv()       As Byte
  
  If boolIsRunning Then
    Err.Raise 1, "clsPipeProcess", "Cannot call Execute while process is running"
  End If
  
  lngExitCode = -1
  lngErrNumber = 0
  strErrDescription = ""
  lngStdInBufferPos = 1
  lngTickCount = 0
  lngLastOutputTickCount = 0
  
  ' input validation
  If (strApplicationName = "") Then subSetError 100, "Missing strApplicationName"
  If (lngErrNumber <> 0) Then GoTo CleanUp
  
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
    GoTo CleanUp
  End If
  
  ' for stderr
  If boolSeparateStdoutStderr Then
    ' separate stdout/stderr pipes
    If (CreatePipe(hStderrReadTmp, hStderrWrite, saPipe, 0&) = 0&) Then
      subSetError 2, "CreatePipe failed on tmp stderr"
      GoTo CleanUp
    End If
  Else
    ' create a duplicate of the stdout handle here for stderr
    ' (in case child decides to close one of them)
    If (DuplicateHandle(GetCurrentProcess(), hStdoutWrite, GetCurrentProcess(), hStderrWrite, 0, True, DUPLICATE_SAME_ACCESS) = 0&) Then
      subSetError 3, "DuplicateHandle failed on stdout/stderr"
      GoTo CleanUp
    End If
  End If
  
  ' for stdin
  If (CreatePipe(hStdinRead, hStdinWriteTmp, saPipe, 0&) = 0&) Then
    subSetError 4, "CreatePipe failed on tmp stdin"
    GoTo CleanUp
  End If
  
  ' duplicate handles: get the "real" handles from the tmp handles, with Properties
  ' set to FALSE. this gives us closeable handles to the pipes.
  ' see http://support.microsoft.com/kb/190351
  If (DuplicateHandle(GetCurrentProcess(), hStdoutReadTmp, GetCurrentProcess(), hStdoutRead, 0, False, DUPLICATE_SAME_ACCESS) = 0&) Then
    subSetError 5, "DuplicateHandle failed on stdout"
    GoTo CleanUp
  End If
  
  If boolSeparateStdoutStderr Then
    If (DuplicateHandle(GetCurrentProcess(), hStderrReadTmp, GetCurrentProcess(), hStderrRead, 0, False, DUPLICATE_SAME_ACCESS) = 0&) Then
      subSetError 6, "DuplicateHandle failed on stderr"
      GoTo CleanUp
    End If
  End If
  If (DuplicateHandle(GetCurrentProcess(), hStdinWriteTmp, GetCurrentProcess(), hStdinWrite, 0, False, DUPLICATE_SAME_ACCESS) = 0&) Then
    subSetError 7, "DuplicateHandle failed on stdin"
    GoTo CleanUp
  End If

  ' Close inheritable copies of the handles we do not want to be inherited.
  CloseHandleSafe hStdoutReadTmp
  ' again, probably don't do this if we're using same handle for stdout/err
  If boolSeparateStdoutStderr Then
    CloseHandleSafe hStderrReadTmp
  End If
  CloseHandleSafe hStdinWriteTmp
  
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
  
  ' UMGEBUNG DES KINDPROZESSES
  '
  ' Das Original uebergab hier einen Block, der NUR SystemRoot enthielt -
  ' als Abhilfe dafuer, dass TCP-Anwendungen mswsock.dll nicht laden
  ' konnten. Ein lpEnvironment ungleich NULL ERSETZT die Umgebung aber
  ' vollstaendig. Der Kindprozess hatte danach kein PATH, kein TEMP, kein
  ' APPDATA, kein JAVA_HOME.
  '
  ' Fuer uns ist das nicht theoretisch: gpg sucht den Schluesselbund unter
  ' %APPDATA%\gnupg, Java legt Arbeitsdateien unter %TEMP% an. Beides waere
  ' still ausgefallen - gpg haette gemeldet, es gebe keinen Schluessel.
  '
  ' NULL bedeutet: der Kindprozess erbt die Umgebung von Access, samt
  ' SystemRoot. Das ist der Normalfall in jeder C-Fassung und braucht den
  ' Umweg des Originals nicht.
  '
  ' Nur wenn mit AddEnvironment etwas ergaenzt wurde, bauen wir einen Block -
  ' dann aber die VOLLSTAENDIGE Umgebung plus die Ergaenzungen.
  If m_Env Is Nothing Then
    lngResult = CreateProcess(0&, strFullArguments, saProcess, saThread, 1&, _
      NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, tStartupInfo, tProcessInfo)
  Else
    abytEnv = BuildEnvironmentBlock()
    lngResult = CreateProcess(0&, strFullArguments, saProcess, saThread, 1&, _
      NORMAL_PRIORITY_CLASS, abytEnv(0), vbNullString, tStartupInfo, tProcessInfo)
  End If
  If (lngResult = 0&) Then
    subSetError 12, "CreateProcess failed"
    GoTo CleanUp
  End If
  
  boolIsRunning = True
  hProcess = tProcessInfo.hProcess

  ' XXX: possibly terminate process if errors occur from here on

  ' Close pipe handles (do not continue to modify the parent).
  ' We need to make sure that no handles to the write end of the output pipe are maintained
  ' in this process or else the pipe will not close when the child process exits.
  ' Probably not an issue since we don't use a blocking ReadFile later, but this is consistent
  ' with the microsoft kb article above.
  CloseHandleSafe hStdoutWrite
  CloseHandleSafe hStdinRead
  CloseHandleSafe hStderrWrite
  

  ' first first stdin block; TODO: may exceed buffer size ?
  If (WriteFile(hStdinWrite, ByVal strStdIn, Len(strStdIn), lngResult, ByVal 0&) = 0) Then
    subSetError 16, "WriteFile failed"
    GoTo CleanUp
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
        GoTo CleanUp
      Case 0
        ' wait complete
      Case Else
        subSetError 18, "WaitForSingleObject failed (" & lngResult & ")"
        lngExitCode = -5
        GoTo CleanUp
    End Select
    
    ' pump the i/o stream pipes - non-blocking, see PumpPipe
    PumpPipe hStdoutRead, strStdOut, False, False
    If boolSeparateStdoutStderr Then
      PumpPipe hStderrRead, strStdErr, True, False
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
          GoTo CleanUp
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
  
  ' Nachlesen bis zum Ende.
  '
  ' Der Durchlauf oben liest nur, was im Augenblick des Ticks dasteht.
  ' Zwischen dem letzten Tick und dem Prozessende kann noch etwas
  ' geschrieben worden sein - bei einem Werkzeug, das erst am Schluss sein
  ' Urteil ausgibt, ist das genau die Zeile, auf die es ankommt.
  '
  ' Jetzt ist Blockieren ausgeschlossen: unsere Schreibenden sind seit
  ' CreateProcess zu, der Kindprozess ist beendet, also ist die Pipe
  ' gebrochen. ReadFile liefert entweder den Rest oder faellt mit
  ' ERROR_BROKEN_PIPE zurueck.
  PumpPipe hStdoutRead, strStdOut, False, True
  If boolSeparateStdoutStderr Then
    PumpPipe hStderrRead, strStdErr, True, True
  End If

  Call GetExitCodeProcess(tProcessInfo.hProcess, lngExitCode)
  lngErrNumber = 0
  strErrDescription = ""
  boolIsRunning = False: boolHasRun = True

CleanUp:
  ' EIN Ausgang fuer Erfolg und Fehler.
  '
  ' Vorher stand hinter jedem subSetError ein Exit Sub - 21 Stellen, keine
  ' davon schloss ein Handle. Bei einem Fehlversuch faellt das nicht auf,
  ' ueber einen Stapellauf schon.
  '
  ' Die Reihenfolge ist absichtlich: erst den Prozess, dann die Handles.

  ' Nach einem Abbruch laeuft der Kindprozess sonst verwaist weiter, mit
  ' geschlossenen Pipes. Auf dem Erfolgsweg ist lngErrNumber 0 und der
  ' Prozess ohnehin beendet.
  If lngErrNumber <> 0 And tProcessInfo.hProcess <> 0 Then
    TerminateProcess tProcessInfo.hProcess, -1
    boolIsRunning = False
  End If

  CloseHandleSafe tProcessInfo.hThStdoutRead   ' = hThread, siehe Kopf
  CloseHandleSafe tProcessInfo.hProcess
  hProcess = 0

  CloseHandleSafe hStdoutReadTmp
  CloseHandleSafe hStderrReadTmp
  CloseHandleSafe hStdinWriteTmp
  CloseHandleSafe hStdoutWrite
  CloseHandleSafe hStderrWrite
  CloseHandleSafe hStdinRead
  CloseHandleSafe hStdoutRead
  CloseHandleSafe hStderrRead
  CloseHandleSafe hStdinWrite

End Sub


' Ergaenzt eine Umgebungsvariable fuer den Kindprozess. Ohne Aufruf erbt
' er die Umgebung von Access unveraendert - das ist der Regelfall.
'
' Scripting.Dictionary statt Collection, weil eine VBA-Collection ihre
' Schluessel nicht herausgibt. Spaet gebunden, damit die Klasse ohne
' zusaetzlichen Verweis auskommt.
Public Sub AddEnvironment(sName As String, sValue As String)
  If boolIsRunning Then
    Err.Raise 1, "clsPipeProcess", "Cannot set environment while process is running"
  End If
  If m_Env Is Nothing Then
    Set m_Env = CreateObject("Scripting.Dictionary")
    m_Env.CompareMode = 1                      ' TextCompare: NAME wie name
  End If
  m_Env(sName) = sValue
End Sub


Public Sub ClearEnvironment()
  Set m_Env = Nothing
End Sub


' Vollstaendige Umgebung von Access plus die Ergaenzungen, als ANSI-Block.
'
' Environ$(n) liest die Umgebung des eigenen Prozesses und erspart das
' Hantieren mit GetEnvironmentStrings und Zeigerarithmetik. Namen, die mit
' "=" beginnen, sind die versteckten Laufwerks-Arbeitsverzeichnisse; die
' laesst Environ$ ohnehin aus.
'
' ANSI ist richtig, weil CreateProcessA benutzt wird und
' CREATE_UNICODE_ENVIRONMENT nicht gesetzt ist.
Private Function BuildEnvironmentBlock() As Byte()

  Dim dicAll  As Object
  Dim i       As Long
  Dim lngPos  As Long
  Dim sItem   As String
  Dim vKey    As Variant
  Dim sBlock  As String

  Set dicAll = CreateObject("Scripting.Dictionary")
  dicAll.CompareMode = 1

  i = 1
  Do
    sItem = Environ$(i)
    If Len(sItem) = 0 Then Exit Do
    lngPos = InStr(sItem, "=")
    If lngPos > 1 Then
      dicAll(Left$(sItem, lngPos - 1)) = Mid$(sItem, lngPos + 1)
    End If
    i = i + 1
  Loop

  ' Ergaenzungen ueberschreiben Geerbtes.
  For Each vKey In m_Env.Keys
    dicAll(CStr(vKey)) = m_Env(vKey)
  Next vKey

  For Each vKey In dicAll.Keys
    sBlock = sBlock & CStr(vKey) & "=" & dicAll(vKey) & Chr$(0)
  Next vKey
  sBlock = sBlock & Chr$(0)

  BuildEnvironmentBlock = StrConv(sBlock, vbFromUnicode)
End Function


' Holt, was in der Pipe steht, und legt es an sTarget an.
'
' bDrainToEnd = False (waehrend der Laufzeit):
'   PeekNamedPipe fragt, wieviel bereitliegt, und nur soviel wird gelesen.
'   Ohne diese Frage wuerde ReadFile blockieren, bis etwas kommt - und die
'   Schleife stuende still, solange das Werkzeug rechnet.
'
'   Die Vorfassung fragte mit GetFileSize. Das ist fuer anonyme Pipes nicht
'   dokumentiert; PeekNamedPipe ist der vorgesehene Aufruf. Der Grund, dass
'   ueberhaupt gefragt werden MUSS, ist die anonyme Pipe selbst: sie kann
'   keine ueberlappenden Operationen. Rusts Standardbibliothek weicht
'   deshalb auf eine benannte Pipe mit FILE_FLAG_OVERLAPPED aus und wartet
'   mit WaitForMultipleObjects auf beide Seiten. Das ist der sauberere Weg,
'   verlangt aber Zeiger, die ueber die ganze Dauer der Operation gueltig
'   bleiben - in VBA nicht verantwortbar. Umfragen ist fuer diese Bauform
'   also richtig, nur der Aufruf war falsch.
'
' bDrainToEnd = True (nach Prozessende):
'   Dann ist die Pipe gebrochen und ReadFile kann nicht mehr blockieren.
'   Gelesen wird, bis nichts mehr kommt - das ist der Abschluss, den jede
'   C-Fassung hat und der hier fehlte.
'
' Fehler beim Lesen brechen den Vorgang NICHT ab. Eine gebrochene Pipe ist
' der Normalfall am Ende, kein Zwischenfall.
Private Sub PumpPipe(ByVal hPipe As LongPtr, ByRef sTarget As String, _
                     ByVal bIsStderr As Boolean, ByVal bDrainToEnd As Boolean)

  Dim lngAvail    As Long
  Dim lngRead     As Long
  Dim lngLeft     As Long
  Dim abytBuf()   As Byte
  Dim sChunk      As String

  If hPipe = 0 Then Exit Sub

  ' Immer erst fragen, nie ungeprueft lesen. Der erste Entwurf las im
  ' Nachlese-Zweig direkt - das haengt, wenn ein Enkelprozess das
  ' Schreibende geerbt hat und die Pipe deshalb NICHT gebrochen ist.
  Do
    lngAvail = 0
    If PeekNamedPipe(hPipe, ByVal 0&, 0&, lngRead, lngAvail, lngLeft) = 0 Then
      ' Gebrochene Pipe. Waehrend der Laufzeit heisst das: fertig.
      ' Danach heisst es: es kann noch Gepuffertes drinstehen, das die
      ' Abfrage nicht mehr meldet. ReadFile kann hier nicht blockieren,
      ' eine gebrochene Pipe antwortet sofort - entweder mit dem Rest oder
      ' mit null Bytes. Die Schleife endet dann von selbst.
      If Not bDrainToEnd Then Exit Do
      lngAvail = DRAIN_BUFFER
    ElseIf lngAvail <= 0 Then
      Exit Do
    End If

    ReDim abytBuf(lngAvail - 1)
    lngRead = 0
    If ReadFile(hPipe, abytBuf(0), lngAvail, lngRead, ByVal 0&) = 0 Then Exit Do
    If lngRead <= 0 Then Exit Do
    If lngRead < lngAvail Then ReDim Preserve abytBuf(lngRead - 1)

    sChunk = StrConv(abytBuf, vbUnicode)
    sTarget = sTarget & sChunk
    If bIsStderr Then
      RaiseEvent StderrAvailable(sChunk)
    Else
      RaiseEvent StdoutAvailable(sChunk)
    End If
    lngLastOutputTickCount = lngTickCount
  Loop
End Sub


' Schliesst ein Handle nur, wenn es offen ist, und setzt es auf 0.
' Das Nullen ist der Punkt: ein zweites CloseHandle auf denselben Wert
' schliesst irgendwann ein fremdes Handle, das die Nummer inzwischen
' wiederbekommen hat. Deshalb ByRef.
Private Sub CloseHandleSafe(ByRef h As LongPtr)
  If h <> 0 Then
    CloseHandle h
    h = 0
  End If
End Sub


Private Sub subSetError(errNumber As Long, errDescription As String)
  lngErrNumber = errNumber
  strErrDescription = errDescription & " (" & Err.LastDllError & ")"
  boolIsRunning = False: boolHasRun = True
End Sub





