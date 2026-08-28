Attribute VB_Name = "mdlProcess"
Option Compare Database
Option Explicit

' (c) 2013 randomnoun. All Rights Reserved. This work is licensed under a
' BSD Simplified License. (http://www.randomnoun.com/bsd-simplified.html)

'' A module to allow Win32 processes to be created, executed, and have their stdin/stdout/stderr
' streams redirected.
'
' <p>This example opens a command shell (cmd), then sends commands to view environment
' variables within the shell (set), attempts to invoke an invalid program (arg), and exit
' the shell (exit). The instructions to the command shell are passed in via stdin:
'
' <pre>
' Dim rpr As RedirectProcessResult
' rpr = fncRedirectProcess "cmd", "", "set" & vbCrLf & "arg" & vbCrLf & "exit" & vbCrLf, false, False
' If rpr.lngErrNumber<>0 Or rpr.lngExitCode<>0 then
'   MsgBox "Error occurred: " & rpr.strErrDescription
' Else
'   MsgBox "Command succeeded. stdout=" & rpr.strStdOut & ", stderr=" & rpr.strStdErr
' Endif
' </pre>
'
' <p>Note that the environment block from the calling process is *not* supplied to the process, since
' this doesn't seem to actually work. The %SystemRoot% environment variable will be set.
'
' @blog http://www.randomnoun.com/wp/2013/10/07/its-all-made-out-of-pipes/
' @author knoxg
' @version $Id$

' see http://www.randomnoun.com/wp/2013/10/07/its-all-made-out-of-pipes/
' for further documentation for this class

' Uses code from
'  http://pastebin.com/CszKUpNS - redirect stdout
'  http://support.microsoft.com/kb/252652 - get folder via CSIDL value
'  http://msdn.microsoft.com/en-us/library/windows/desktop/ms682425%28v=vs.85%29.aspx - CreateProcess
'  http://support.microsoft.com/kb/173085 - reading/writing pipes
'  http://msdn.microsoft.com/en-us/library/ms682499.aspx - readin/writing pipes (2)
'  http://support.microsoft.com/kb/q129796 - determine when process is complete
'  http://support.microsoft.com/kb/190351 - duplicate handles during redirection

' result user-defined type

Public Type RedirectProcessResult
  strCommandLine As String
  strStdOut As String
  strStdErr As String
  lngExitCode As Long ' process exit code if invocation successful
  lngErrNumber As Long ' invocation error code; 0=success
  strErrDescription As String ' if lngError <> 0
End Type


' Windows structures/constants/declarations

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
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThStdoutReadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

' this works
' fncRedirectProcess "cmd", "", "echo hello" & vbcrlf & "exit"& vbcrlf
' but cvsnt doesn't.  ah. now it does.

'' So this thing kicks off a process and captures the output in some intelligible form
'
' @param strApplicationName
' @param strArguments
' @param strStdin
' @param boolShowWindow
' @param boolSeparateStdoutStderr

Public Function fncRedirectProcess(strApplicationName As String, strArguments As String, _
  Optional strStdIn As String = "", Optional boolShowWindow As Boolean = False, _
  Optional boolSeparateStdoutStderr As Boolean = False) As RedirectProcessResult

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
  Dim strFullCommandLine  As String
  Dim lngExitCode     As Long
  Dim lngSizeOf       As Long
  Dim strPath         As String
  Dim strEnv          As String
  Dim abytEnv()       As Byte

  fncRedirectProcess.lngExitCode = -1
  fncRedirectProcess.lngErrNumber = 0
  fncRedirectProcess.strErrDescription = ""
  
  ' input validation
  If (strApplicationName = "") Then subSetError fncRedirectProcess, 100, "Missing strApplicationName"
  If (fncRedirectProcess.lngErrNumber <> 0) Then Exit Function
  
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
    subSetError fncRedirectProcess, 1, "CreatePipe failes on tmp stdout"
    Exit Function
  End If
  
  ' for stderr
  If boolSeparateStdoutStderr Then
    ' separate stdout/stderr pipes
    If (CreatePipe(hStderrReadTmp, hStderrWrite, saPipe, 0&) = 0&) Then
      subSetError fncRedirectProcess, 2, "CreatePipe failed on tmp stderr"
      Exit Function
    End If
  Else
    ' create a duplicate of the stdout handle here for stderr
    ' (in case child decides to close one of them)
    If (DuplicateHandle(GetCurrentProcess(), hStdoutWrite, GetCurrentProcess(), hStderrWrite, 0, True, DUPLICATE_SAME_ACCESS) = 0&) Then
      subSetError fncRedirectProcess, 3, "DuplicateHandle failed on stdout/stderr"
      Exit Function
    End If
  End If
  
  ' for stdin
  If (CreatePipe(hStdinRead, hStdinWriteTmp, saPipe, 0&) = 0&) Then
    subSetError fncRedirectProcess, 4, "CreatePipe failed on tmp stdin"
    Exit Function
  End If
  
  ' duplicate handles: get the "real" handles from the tmp handles, with Properties
  ' set to FALSE. this gives us closeable handles to the pipes.
  ' see http://support.microsoft.com/kb/190351
  If (DuplicateHandle(GetCurrentProcess(), hStdoutReadTmp, GetCurrentProcess(), hStdoutRead, 0, False, DUPLICATE_SAME_ACCESS) = 0&) Then
    subSetError fncRedirectProcess, 5, "DuplicateHandle failed on stdout"
    Exit Function
  End If
  
  If boolSeparateStdoutStderr Then
    If (DuplicateHandle(GetCurrentProcess(), hStderrReadTmp, GetCurrentProcess(), hStderrRead, 0, False, DUPLICATE_SAME_ACCESS) = 0&) Then
      subSetError fncRedirectProcess, 6, "DuplicateHandle failed on stderr"
      Exit Function
    End If
  End If
  If (DuplicateHandle(GetCurrentProcess(), hStdinWriteTmp, GetCurrentProcess(), hStdinWrite, 0, False, DUPLICATE_SAME_ACCESS) = 0&) Then
    subSetError fncRedirectProcess, 7, "DuplicateHandle failed on stdin"
    Exit Function
  End If

  ' Close inheritable copies of the handles we do not want to be inherited.
  If (CloseHandle(hStdoutReadTmp) = 0) Then
    subSetError fncRedirectProcess, 8, "CloseHandle failed on tmp stdout"
    Exit Function
  End If
  ' again, probably don't do this if we're using same handle for stdout/err
  If boolSeparateStdoutStderr Then
    If (CloseHandle(hStderrReadTmp) = 0) Then
      subSetError fncRedirectProcess, 9, "CloseHandle failed on tmp stderr"
      Exit Function
    End If
  End If
  If (CloseHandle(hStdinWriteTmp) = 0) Then
    subSetError fncRedirectProcess, 10, "CloseHandle failed on tmp stdin"
    Exit Function
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
  strFullCommandLine = """" & strApplicationName & """" & " " & strArguments
  fncRedirectProcess.strCommandLine = strFullCommandLine
  
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
    subSetError fncRedirectProcess, 11, "SHGetFolderPath failed"
    Exit Function
  End If
  strPath = Left(strPath, InStr(1, strPath, Chr(0)) - 1) ' just up to null terminator
  strEnv = "SystemRoot=" & strPath & Chr(0) & Chr(0)  ' sz terminator + env block terminator
  abytEnv = StrConv(strEnv, vbFromUnicode)            ' encode as byte array
    
  ' if the executable module is a 16-bit application, lpApplicationName should be NULL
  lngResult = CreateProcess(0&, strFullCommandLine, saProcess, saThread, 1&, _
    NORMAL_PRIORITY_CLASS, abytEnv(0), vbNullString, tStartupInfo, tProcessInfo)
  If (lngResult = 0&) Then
    subSetError fncRedirectProcess, 12, "CreateProcess failed"
    Exit Function
  End If
  
  ' XXX: possibly terminate process if errors occur from here on

  ' Close pipe handles (do not continue to modify the parent).
  ' We need to make sure that no handles to the write end of the output pipe are maintained
  ' in this process or else the pipe will not close when the child process exits.
  ' Probably not an issue since we don't use a blocking ReadFile later, but this is consistent
  ' with the microsoft kb article above.
  If (CloseHandle(hStdoutWrite) = 0) Then
    subSetError fncRedirectProcess, 13, "CloseHandle failed"
    Exit Function
  End If
  If (CloseHandle(hStdinRead) = 0) Then
    subSetError fncRedirectProcess, 14, "CloseHandle failed"
    Exit Function
  End If
  If (CloseHandle(hStderrWrite) = 0) Then
    subSetError fncRedirectProcess, 15, "CloseHandle failed"
    Exit Function
  End If

  ' first first stdin block; TODO: may exceed buffer size ?
  If (WriteFile(hStdinWrite, ByVal strStdIn, Len(strStdIn), lngResult, ByVal 0&) = 0) Then
    subSetError fncRedirectProcess, 16, "WriteFile failed"
    Exit Function
  End If
  
  lngResult = WaitForSingleObject(tProcessInfo.hProcess, 100)  ' 100ms
  Do
    DoEvents
    Select Case lngResult
      Case 258& ' 500ms timeout
        ' keep on trucking.
        lngResult = WaitForSingleObject(tProcessInfo.hProcess, 100)
      Case &H80, &HFFFFFFFF  ' abandoned / failed
        subSetError fncRedirectProcess, 17, "Wait abandoned/failed (" & lngResult & ")"
        Exit Function
      Case 0
        ' wait complete
      Case Else
        subSetError fncRedirectProcess, 18, "WaitForSingleObject failed (" & lngResult & ")"
        lngExitCode = -5
        Exit Function
    End Select
    
    ' pump the i/o stream pipes
    lngSizeOf = GetFileSize(hStdoutRead, 0&)
    If (lngSizeOf > 0) Then
      ReDim abytReadBuf(lngSizeOf - 1)
      If ReadFile(hStdoutRead, abytReadBuf(0), UBound(abytReadBuf) + 1, lngReadBytes, ByVal 0&) = 0 Then
        subSetError fncRedirectProcess, 19, "ReadFile failed"
        Exit Function
      Else
        'Debug.Print "read-stdout: " & StrConv(abytReadBuf, vbUnicode)
        fncRedirectProcess.strStdOut = fncRedirectProcess.strStdOut & StrConv(abytReadBuf, vbUnicode)
      End If
    End If
    
    If boolSeparateStdoutStderr Then
      lngSizeOf = GetFileSize(hStderrRead, 0&)
      If (lngSizeOf > 0) Then
        ReDim abytReadBuf(lngSizeOf - 1)
        If ReadFile(hStderrRead, abytReadBuf(0), UBound(abytReadBuf) + 1, lngReadBytes, ByVal 0&) = 0 Then
          subSetError fncRedirectProcess, 20, "Read failed"
          Exit Function
        Else
          'Debug.Print "read-stderr: " & StrConv(abytReadBuf, vbUnicode)
          fncRedirectProcess.strStdErr = fncRedirectProcess.strStdErr & StrConv(abytReadBuf, vbUnicode)
        End If
      End If
    End If
    
  Loop Until lngResult = 0
  
  
  Call GetExitCodeProcess(tProcessInfo.hProcess, lngExitCode)
  fncRedirectProcess.lngExitCode = lngExitCode
  fncRedirectProcess.lngErrNumber = 0
  fncRedirectProcess.strErrDescription = ""
  
  'Debug.Print "============="
  'Debug.Print "stdout: " & fncRedirectProcess.strStdOut
  'Debug.Print "stderr: " & fncRedirectProcess.strStdErr
  
  ' not too concerned about reporting error conditions from here on
  CloseHandle tProcessInfo.hThStdoutRead
  CloseHandle tProcessInfo.hProcess
  CloseHandle hStdoutRead
  CloseHandle hStderrRead
  CloseHandle hStdinWrite
  
End Function


Private Sub subSetError(ByRef rdr As RedirectProcessResult, errNumber As Long, errDescription As String)
  rdr.lngErrNumber = errNumber
  rdr.strErrDescription = errDescription & " (" & Err.LastDllError & ")"
End Sub

