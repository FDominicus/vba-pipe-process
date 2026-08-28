VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmProcessTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Copyright (c) 2013 randomnoun (knoxg)   - original
' Copyright (c) 2026 Friedrich Dominicus, Q-Software Solutions GmbH - 64-bit port
'
' BSD 2-Clause "Simplified" License. See LICENSE in the repository root.
'
' Code behind the test form from the original dbProcess.mdb. Ported to 64-bit
' along with the class; the three declarations here belong to the animation
' control, which is decoration only.
'
' Changed from the original: the p2 test waited for an English error message
' from cmd before sending "exit" through stdin. On a non-English Windows that
' never matched, so cmd kept waiting and the window stayed open. It now checks
' whether anything arrived at all.


Dim WithEvents p As clsPipeProcess   ' dumps std/stderr to debug.print
Attribute p.VB_VarHelpID = -1
Dim WithEvents p2 As clsPipeProcess  ' as above, and adds input on various conditions
Attribute p2.VB_VarHelpID = -1
Dim WithEvents p3 As clsPipeProcess  ' as above, manually terminates after six seconds
Attribute p3.VB_VarHelpID = -1

#If VBA7 Then
Dim hAnimationControl As LongPtr    ' handle to animation control
#Else
Dim hAnimationControl As Long
#End If
Dim boolHandledPause As Boolean

' Windows functions and constants
#If VBA7 Then
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function SendMessageVal Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function SendMessageVal Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

Const ANIMATE_CLASS = "SysAnimate32" '  commctrl.h
Const WS_CHILD = &H40000000
Const WS_VISIBLE = &H10000000
Const ACS_CENTER = &H1
Const ACS_TRANSPARENT = &H2
Const ACM_OPEN = (&H400 + 100)
Const ACM_PLAY = (&H400 + 101)
Const ACM_STOP = (&H400 + 102)

Private Sub Form_Open(Cancel As Integer)
  hAnimationControl = subCreateAnimationControlFromDll("shell32.dll", 161, 0, 0, 300, 60)
End Sub


' http://support.microsoft.com/kb/189170
Private Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
  MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function

#If VBA7 Then
Private Function subCreateAnimationControlFromDll(strDll As String, lngResourceId As Long, _
  lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long) As LongPtr

  Dim hAnimationControl As LongPtr
  Dim hAviDll As LongPtr
#Else
Private Function subCreateAnimationControlFromDll(strDll As String, lngResourceId As Long, _
  lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long) As Long

  Dim hAnimationControl As Long
  Dim hAviDll As Long
#End If

  ' this link has an alternative implementation that uses the ActiveX Microsoft Animation control
  '   http://support.microsoft.com/kb/163440
  ' although that didn't work for me, so I'm doing things manually
  subCreateAnimationControlFromDll = 0
  hAviDll = LoadLibrary(strDll)
  If (hAviDll = 0) Then
    Debug.Print "Could not load library"
    Exit Function
  End If

  ' NB: this control doesn't repaint in some circumstances, but is good enough for my purposes
  hAnimationControl = CreateWindowEx(0, ANIMATE_CLASS, &H0, _
    WS_CHILD Or WS_VISIBLE Or ACS_TRANSPARENT Or ACS_CENTER, _
    lngLeft, lngTop, lngWidth, lngHeight, Me.hwnd, &H0, hAviDll, &H0)

  ' resource id is in lo word http://msdn.microsoft.com/en-us/library/windows/desktop/bb761897%28v=vs.85%29.aspx
  If (SendMessageVal(hAnimationControl, ACM_OPEN, &H0, MakeDWord(lngResourceId, 0)) = 0) Then  '161
    Debug.Print "Could not load avi"
    Exit Function
  End If
  subCreateAnimationControlFromDll = hAnimationControl
End Function

#If VBA7 Then
Private Sub subStartAnimation(hAnimationControl As LongPtr)
#Else
Private Sub subStartAnimation(hAnimationControl As Long)
#End If
  SendMessageVal hAnimationControl, ACM_PLAY, -1, MakeDWord(0, -1) ' -1=repeat indefinitely, (0,-1)=play all frames
End Sub

#If VBA7 Then
Private Sub subStopAnimation(hAnimationControl As LongPtr)
#Else
Private Sub subStopAnimation(hAnimationControl As Long)
#End If
  SendMessageVal hAnimationControl, ACM_STOP, &H0, &H0
End Sub

Private Sub cmdTest_Click()
  ' for now
  ' add an animation control
  ' via http://www.codeproject.com/Articles/3838/Loading-AVI-Resources-from-Shell32-dll-file
  ' and http://www.visualbasic.happycodings.com/API_and_Miscellaneous/code12.html
  ' and http://www.xtremevbtalk.com/showthread.php?t=217011
  
  ' Library loaded now create the animation control
  subStartAnimation hAnimationControl
  
  ' reset output text boxes and test state
  txtLog.Value = ""
  txtStdout.Value = ""
  txtStderr.Value = ""
  boolHandledPause = False
  
  ' run tests
  cmbCombineStdoutStderr.SetFocus ' can't disable a focussed control
  cmdTest.Caption = "Running"
  cmdTest.Enabled = False
  
  subTestClsProcessWithEvents
  subTestClsProcessSendInput
  subTestClsAutoTerminate
  subTestClsManualTerminate
  cmdTest.Caption = "Test"
  cmdTest.Enabled = True
  cmdTest.SetFocus
  subStopAnimation hAnimationControl
  
End Sub

' general logging
Private Sub subLog(strText As String)
  Debug.Print strText
  txtLog.Value = txtLog.Value & strText & vbCrLf
End Sub

' log to stdout textbox
Private Sub subLogStdout(strText As String)
  Debug.Print strText
  txtStdout.Value = txtStdout.Value & strText
End Sub

' log to stderr textbox
Private Sub subLogStderr(strText As String)
  Debug.Print strText
  txtStderr.Value = txtStderr.Value & strText
End Sub


' NB: This event will not fire unless clsPipeProcess.SeparateStdoutStderr is set to true below
Private Sub p_StderrAvailable(strText As String)
  subLog "p_StderrAvailable()"
  subLogStderr strText
End Sub

Private Sub p_StdoutAvailable(strText As String)
  subLog "p_StdoutAvailable()"
  subLogStdout strText
End Sub

Private Sub p_Tick()
  subLog "p_Tick: p.TickCount=" & p.TickCount & ", p.LastOutputTickCount=" & p.LastOutputTickCount
End Sub

' p2 runs 'pause' occasionally and has split stdout/stderr streams
' well, it doesn't now. but it did before.
Private Sub p2_StderrAvailable(strText As String)
  subLog "p2_StderrAvailable()"
  subLogStderr strText
  If Len(p2.StdErr) > 0 And Not boolHandledPause Then
    boolHandledPause = True
    subLog "p2_stderrAvailable: sending exit & vbCrLf"
    p2.SendInput "exit" & vbCrLf
  End If
End Sub

Private Sub p2_StdoutAvailable(strText As String)
  subLog "p2_StdoutAvailable()"
  subLogStdout strText

  ' we could check strText here but let's check the entire stdout buffer instead
  If Len(p2.StdOut) > 0 And Not boolHandledPause Then
    boolHandledPause = True
    subLog "p2_stdoutAvailable: sending exit & vbCrLf"
    p2.SendInput "exit" & vbCrLf
  End If
End Sub

Private Sub p2_Tick()
  subLog "p2_Tick()"
  If boolHandledPause Then
    ' maybe one of these will work
    ' this sometimes generates the text "More?" ? what's that all about ?
    ' p2.SendInput "exit" & vbCrLf
  End If
End Sub

' p3 is the same as p, but calls Terminate() in its _Tick() handler

' NB: This event will not fire unless clsPipeProcess.SeparateStdoutStderr is set to true below
Private Sub p3_StderrAvailable(strText As String)
  subLog "p3_StderrAvailable()"
  subLogStderr strText
End Sub

Private Sub p3_StdoutAvailable(strText As String)
  subLog "p3_StdoutAvailable()"
  subLogStdout strText
End Sub

Private Sub p3_Tick()
  subLog "p3_Tick: p3.TickCount=" & p3.TickCount & ", p3.LastOutputTickCount=" & p3.LastOutputTickCount
  If (p3.TickCount - p3.LastOutputTickCount) * p3.TickWaitMillis > 2000 Then ' 2 seconds
    subLog "p3_Tick: calling terminate"
    p3.Terminate
  End If
End Sub


Sub subTestClsProcessWithEvents()
  Set p = New clsPipeProcess
  p.ApplicationName = "cmd"
  p.StdIn = "set" & vbCrLf & "this-program-doesnt-exist" & vbCrLf & "exit" & vbCrLf
  p.SeparateStdoutStderr = Not cmbCombineStdoutStderr
  p.ShowWindow = chkShowWindow.Value

  subLog "===== subTestClsProcessWithEvents test start"
  p.Execute
  subLog "===== subTestClsProcessWithEvents execute complete"
  subLog "p.ApplicationName=" & p.ApplicationName    ' call this something else ?
  subLog "p.Arguments=" & p.Arguments  ' call this Argumentsargs ?
  subLog "p.ErrNumber=" & p.errNumber
  subLog "p.ErrDescription=" & p.errDescription
  subLog "p.ErrExitCode=" & p.ExitCode
  subLog "p.Terminated=" & p.Terminated
  subLog "p.StdOut=" & p.StdOut
  subLog "p.StdErr=" & p.StdErr
  subLog "===== subTestClsProcessWithEvents test complete"
  Set p = Nothing
End Sub

Sub subTestClsProcessSendInput()
  Set p2 = New clsPipeProcess
  p2.ApplicationName = "cmd"
  p2.StdIn = "set" & vbCrLf & "this-program-doesnt-exist" & vbCrLf  ' & "pause" & vbCrLf
  p2.SeparateStdoutStderr = Not cmbCombineStdoutStderr
  p2.ShowWindow = chkShowWindow.Value

  subLog "===== subTestClsProcessSendInput test start"
  boolHandledPause = False
  p2.Execute
  subLog "===== subTestClsProcessSendInput execute complete"
  subLog "p2.ApplicationName=" & p2.ApplicationName    ' call this something else ?
  subLog "p2.Arguments=" & p2.Arguments  ' call this Argumentsargs ?
  subLog "p2.ErrNumber=" & p2.errNumber
  subLog "p2.ErrDescription=" & p2.errDescription
  subLog "p2.ErrExitCode=" & p2.ExitCode
  subLog "p2.Terminated=" & p2.Terminated
  subLog "p2.StdOut=" & p2.StdOut
  subLog "p2.StdErr=" & p2.StdErr
  subLog "===== subTestClsProcessSendInput test complete" ' there shouldn't be anything in the debug window after this
  Set p2 = Nothing
End Sub

Sub subTestClsAutoTerminate()
  Set p = New clsPipeProcess
  p.ApplicationName = "cmd"
  p.StdIn = "set" & vbCrLf & "this-program-doesnt-exist" & vbCrLf & "pause" & vbCrLf
  p.SeparateStdoutStderr = Not cmbCombineStdoutStderr
  p.ShowWindow = chkShowWindow.Value
  p.TimeoutMillis = 500
  

  subLog "===== subTestClsAutoTerminate test start"
  boolHandledPause = False
  p.Execute
  subLog "===== subTestClsAutoTerminate execute complete"
  subLog "p.ApplicationName=" & p.ApplicationName    ' call this something else ?
  subLog "p.Arguments=" & p.Arguments  ' call this Argumentsargs ?
  subLog "p.ErrNumber=" & p.errNumber
  subLog "p.ErrDescription=" & p.errDescription
  subLog "p.ErrExitCode=" & p.ExitCode
  subLog "p.Terminated=" & p.Terminated
  subLog "p.StdOut=" & p.StdOut
  subLog "p.StdErr=" & p.StdErr
  subLog "===== subTestClsAutoTerminate test complete" ' there shouldn't be anything in the debug window after this
  Set p = Nothing
End Sub

Sub subTestClsManualTerminate()
  Set p3 = New clsPipeProcess
  p3.ApplicationName = "cmd"
  p3.StdIn = "set" & vbCrLf & "this-program-doesnt-exist" & vbCrLf & "pause" & vbCrLf
  p3.SeparateStdoutStderr = Not cmbCombineStdoutStderr
  p3.ShowWindow = chkShowWindow.Value
  p3.TimeoutMillis = 0  ' no timeout

  subLog "===== subTestClsManualTerminate test start"
  boolHandledPause = False
  p3.Execute
  subLog "===== subTestClsManualTerminate execute complete"
  subLog "p3.ApplicationName=" & p3.ApplicationName    ' call this something else ?
  subLog "p3.Arguments=" & p3.Arguments  ' call this Argumentsargs ?
  subLog "p3.ErrNumber=" & p3.errNumber
  subLog "p3.ErrDescription=" & p3.errDescription
  subLog "p3.ErrExitCode=" & p3.ExitCode
  subLog "p3.Terminated=" & p3.Terminated
  subLog "p3.StdOut=" & p3.StdOut
  subLog "p3.StdErr=" & p3.StdErr
  subLog "===== subTestClsManualTerminate test complete" ' there shouldn't be anything in the debug window after this
  Set p3 = Nothing
End Sub




