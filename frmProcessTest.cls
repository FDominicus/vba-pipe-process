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

' (c) 2013 randomnoun. All Rights Reserved. This work is licensed under a
' BSD Simplified License. (http://www.randomnoun.com/bsd-simplified.html)

'==============================================================================
' 64-BIT-FASSUNG - Umbau 28.08.2026
'
' GEAENDERT
'   clsProcess -> QSS_C_PipeProcess (3 Deklarationen, 4 New)
'   CreateWindowEx / LoadLibrary / SendMessageVal -> PtrSafe
'     hwnd, hMenu, hInstance, Rueckgabe von CreateWindowEx  -> LongPtr
'     Rueckgabe von LoadLibrary (HMODULE)                   -> LongPtr
'     SendMessage wParam/lParam (WPARAM/LPARAM) und Rueckgabe (LRESULT)
'                                                           -> LongPtr
'   hAnimationControl und die Parameter der Animations-Routinen -> LongPtr
'
' Die Animation ist reine Zierde. Wer beim Testen Ruhe haben will, setzt
' USE_ANIMATION auf False - dann laeuft kein CreateWindowEx, kein
' LoadLibrary, kein SendMessage, und drei Absturzquellen sind weg. Die
' Klasse wird davon nicht beruehrt.
'
' Die 32-Bit-Zweige sind geblieben, weil die Datei auch unter Access 2021
' laufen soll.
'
'------------------------------------------------------------------------------
' KORREKTUREN 28.08.2026 (zweiter Durchgang)
'
' 1  HAENGER in TestSendInput. Der Test schickte "set" und einen falschen
'    Befehl, aber KEIN "exit", und setzte keine Zeitgrenze. Er verliess
'    sich darauf, dass das Stdout-Ereignis feuert und "exit" nachschiebt.
'    Feuert es nicht, wartet cmd auf stdin - und das Schreibende der
'    stdin-Pipe wird nie geschlossen (siehe Klassenkopf). Access haengt
'    dann, bis man es abschiesst.
'    Jetzt: TimeoutMillis = 5000 als Netz. Der Test prueft weiterhin
'    SendInput; er kann nur nicht mehr stehenbleiben.
'
' 2  Not cmbCombineStdoutStderr bricht mit Fehler 94, wenn das
'    Steuerelement Null ist - Not Null ist Null, und Null geht nicht in
'    ein Boolean. Jetzt ueber Nz.
'    Dasselbe fuer chkShowWindow.
'
' 3  LoadLibrary ohne FreeLibrary. Jedes Oeffnen des Formulars liess einen
'    Modulverweis stehen. Folgenlos, weil shell32 ohnehin geladen ist -
'    aber es ist ein Leck, und Resource Leak ist im Fehlerprotokoll die
'    Spitzenkategorie. hAviDll steht jetzt auf Modulebene und wird in
'    Form_Close freigegeben.
'
' 4  CreateWindowEx bekam fuer lpWindowName (ByVal String) den Wert &H0.
'    VBA macht daraus die Zeichenkette "0", nicht NULL. Folgenlos, weil
'    ein Animationsfenster keinen Titel zeigt - aber vbNullString ist das,
'    was gemeint war, und kostet nichts.
'
' 5  hAnimationControl war dreifach verdeckt: einmal auf Modulebene,
'    einmal als lokale Variable in der erzeugenden Funktion, einmal als
'    Parameter der beiden Animations-Routinen. Es lief, weil die Funktion
'    ihren Wert zurueckgibt - lesbar war es nicht. Die Locals heissen
'    jetzt anders.
'
' 6  Form_Open pruefte den Rueckgabewert nicht. Bei einem Fehlschlag gingen
'    ACM_PLAY und ACM_STOP an ein Fenster mit Handle 0.
'
' 7  Dim auf Modulebene -> Private. In einem Formularmodul ist der
'    Unterschied fuer WithEvents-Variablen kosmetisch, aber Dim sagt nichts
'    ueber die Sichtbarkeit.
'
' 8  "Cls" aus den Namen der Testprozeduren - die Klasse heisst nicht mehr
'    so.
'
' NICHT geaendert, bewusst:
'   MakeDWord rechnet HiWord * &H10000 in Long. Bei HiWord <= 32767 kann
'   das nicht ueberlaufen, und -1 ergibt korrekt &HFFFF0000. Sieht nach
'   Fehler aus, ist keiner.
'==============================================================================

' Auf False setzen, wenn beim Testen nichts ausser der Klasse laufen soll.
Private Const USE_ANIMATION As Boolean = True

' Zeitgrenze fuer die Tests, die auf stdin warten. Netz, keine Vorgabe.
Private Const SAFETY_TIMEOUT_MS As Long = 5000

Private WithEvents p As clsPipeProcess   ' dumps std/stderr to debug.print
Attribute p.VB_VarHelpID = -1
Private WithEvents p2 As clsPipeProcess  ' as above, and adds input on various conditions
Attribute p2.VB_VarHelpID = -1
Private WithEvents p3 As clsPipeProcess  ' as above, manually terminates after six seconds
Attribute p3.VB_VarHelpID = -1

#If VBA7 Then
Private hAnimationControl As LongPtr    ' handle to animation control
Private hAviDll As LongPtr              ' Modul, aus dem die AVI stammt
#Else
Private hAnimationControl As Long
Private hAviDll As Long
#End If
Private boolHandledPause As Boolean

' Windows functions and constants
#If VBA7 Then
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function SendMessageVal Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function SendMessageVal Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

Private Const ANIMATE_CLASS   As String = "SysAnimate32"   ' commctrl.h
Private Const WS_CHILD        As Long = &H40000000
Private Const WS_VISIBLE      As Long = &H10000000
Private Const ACS_CENTER      As Long = &H1
Private Const ACS_TRANSPARENT As Long = &H2
Private Const ACM_OPEN        As Long = (&H400 + 100)
Private Const ACM_PLAY        As Long = (&H400 + 101)
Private Const ACM_STOP        As Long = (&H400 + 102)


Private Sub Form_Open(Cancel As Integer)
  If Not USE_ANIMATION Then Exit Sub

  hAnimationControl = CreateAnimationControlFromDll("shell32.dll", 161, 0, 0, 300, 60)
  If hAnimationControl = 0 Then
    Debug.Print "Animation nicht verfuegbar - Tests laufen trotzdem."
  End If
End Sub


' Punkt 3: was geladen und erzeugt wurde, wird auch wieder freigegeben.
Private Sub Form_Close()
  If hAnimationControl <> 0 Then
    DestroyWindow hAnimationControl
    hAnimationControl = 0
  End If
  If hAviDll <> 0 Then
    FreeLibrary hAviDll
    hAviDll = 0
  End If
End Sub


' http://support.microsoft.com/kb/189170
Private Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
  MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function

#If VBA7 Then
Private Function CreateAnimationControlFromDll(strDll As String, lngResourceId As Long, _
  lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long) As LongPtr

  Dim hCtrl As LongPtr
#Else
Private Function CreateAnimationControlFromDll(strDll As String, lngResourceId As Long, _
  lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long) As Long

  Dim hCtrl As Long
#End If

  ' this link has an alternative implementation that uses the ActiveX Microsoft Animation control
  '   http://support.microsoft.com/kb/163440
  ' although that didn't work for me, so I'm doing things manually
  CreateAnimationControlFromDll = 0

  hAviDll = LoadLibrary(strDll)
  If (hAviDll = 0) Then
    Debug.Print "Could not load library"
    Exit Function
  End If

  ' NB: this control doesn't repaint in some circumstances, but is good enough for my purposes
  ' Punkt 4: vbNullString statt &H0 - letzteres wird zur Zeichenkette "0".
  hCtrl = CreateWindowEx(0, ANIMATE_CLASS, vbNullString, _
    WS_CHILD Or WS_VISIBLE Or ACS_TRANSPARENT Or ACS_CENTER, _
    lngLeft, lngTop, lngWidth, lngHeight, Me.hwnd, &H0, hAviDll, &H0)

  If hCtrl = 0 Then
    Debug.Print "Could not create animation control"
    FreeLibrary hAviDll
    hAviDll = 0
    Exit Function
  End If

  ' resource id is in lo word http://msdn.microsoft.com/en-us/library/windows/desktop/bb761897%28v=vs.85%29.aspx
  If (SendMessageVal(hCtrl, ACM_OPEN, &H0, MakeDWord(lngResourceId, 0)) = 0) Then  '161
    Debug.Print "Could not load avi"
    DestroyWindow hCtrl
    FreeLibrary hAviDll
    hAviDll = 0
    Exit Function
  End If

  CreateAnimationControlFromDll = hCtrl
End Function

#If VBA7 Then
Private Sub StartAnimation(ByVal hCtrl As LongPtr)
#Else
Private Sub StartAnimation(ByVal hCtrl As Long)
#End If
  If hCtrl = 0 Then Exit Sub
  SendMessageVal hCtrl, ACM_PLAY, -1, MakeDWord(0, -1) ' -1=repeat indefinitely, (0,-1)=play all frames
End Sub

#If VBA7 Then
Private Sub StopAnimation(ByVal hCtrl As LongPtr)
#Else
Private Sub StopAnimation(ByVal hCtrl As Long)
#End If
  If hCtrl = 0 Then Exit Sub
  SendMessageVal hCtrl, ACM_STOP, &H0, &H0
End Sub


' --- Einstellungen aus den Steuerelementen, Null-fest (Punkt 2) ------------

Private Function OptSeparateStreams() As Boolean
  OptSeparateStreams = Not CBool(Nz(cmbCombineStdoutStderr, False))
End Function

Private Function OptShowWindow() As Boolean
  OptShowWindow = CBool(Nz(chkShowWindow, False))
End Function


Private Sub cmdTest_Click()
  ' for now
  ' add an animation control
  ' via http://www.codeproject.com/Articles/3838/Loading-AVI-Resources-from-Shell32-dll-file
  ' and http://www.visualbasic.happycodings.com/API_and_Miscellaneous/code12.html
  ' and http://www.xtremevbtalk.com/showthread.php?t=217011

  On Error GoTo HandleError

  StartAnimation hAnimationControl

  ' reset output text boxes and test state
  txtLog.Value = ""
  txtStdout.Value = ""
  txtStderr.Value = ""
  boolHandledPause = False

  ' run tests
  cmbCombineStdoutStderr.SetFocus ' can't disable a focussed control
  cmdTest.Caption = "Running"
  cmdTest.Enabled = False

  TestProcessWithEvents
  TestSendInput
  TestAutoTerminate
  TestManualTerminate

HandleExit:
  cmdTest.Caption = "Test"
  cmdTest.Enabled = True
  cmdTest.SetFocus
  StopAnimation hAnimationControl
  Exit Sub

HandleError:
  ' Ohne das bleibt der Knopf nach einem Fehler auf "Running" und gesperrt.
  subLog "FEHLER " & Err.Number & ": " & Err.Description
  Resume HandleExit
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


' NB: This event will not fire unless SeparateStdoutStderr is set to true below
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
End Sub

' p3 is the same as p, but calls Terminate() in its _Tick() handler

' NB: This event will not fire unless SeparateStdoutStderr is set to true below
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


' Gemeinsame Ausgabe nach jedem Test - vier fast gleiche Bloecke waren es
' vorher, mit je einem eigenen Tippfehler im Text.
Private Sub LogResult(ByVal sName As String, ByVal proc As clsPipeProcess)
  subLog "===== " & sName & " execute complete"
  subLog "  ApplicationName = " & proc.ApplicationName
  subLog "  Arguments       = " & proc.Arguments
  subLog "  ErrNumber       = " & proc.errNumber
  subLog "  ErrDescription  = " & proc.errDescription
  subLog "  ExitCode        = " & proc.ExitCode
  subLog "  Terminated      = " & proc.Terminated
  subLog "  StdOut          = " & proc.StdOut
  subLog "  StdErr          = " & proc.StdErr
  subLog "===== " & sName & " test complete"
End Sub


' "set" gibt die Umgebung des KINDPROZESSES aus. Genau hier sieht man, ob
' die Umgebung geerbt wird: es muessen PATH, TEMP und APPDATA drinstehen.
' Vor der Korrektur vom 28.08. stand dort genau eine Zeile - SystemRoot.
Sub TestProcessWithEvents()
  Set p = New QSS_C_PipeProcess
  p.ApplicationName = "cmd"
  p.StdIn = "set" & vbCrLf & "this-program-doesnt-exist" & vbCrLf & "exit" & vbCrLf
  p.SeparateStdoutStderr = OptSeparateStreams()
  p.ShowWindow = OptShowWindow()
  p.TimeoutMillis = SAFETY_TIMEOUT_MS

  subLog "===== TestProcessWithEvents test start"
  p.Execute
  LogResult "TestProcessWithEvents", p
  Set p = Nothing
End Sub

' Punkt 1: Dieser Test schickt bewusst KEIN "exit" - das kommt aus dem
' Ereignis. Ohne Zeitgrenze wartet cmd sonst ewig auf stdin, denn das
' Schreibende der stdin-Pipe wird nie geschlossen.
Sub TestSendInput()
  Set p2 = New QSS_C_PipeProcess
  p2.ApplicationName = "cmd"
  p2.StdIn = "set" & vbCrLf & "this-program-doesnt-exist" & vbCrLf
  p2.SeparateStdoutStderr = OptSeparateStreams()
  p2.ShowWindow = OptShowWindow()
  p2.TimeoutMillis = SAFETY_TIMEOUT_MS

  subLog "===== TestSendInput test start"
  boolHandledPause = False
  p2.Execute
  LogResult "TestSendInput", p2
  ' Terminated = True heisst hier: das Ereignis hat NICHT gefeuert und die
  ' Zeitgrenze hat eingegriffen. Der Test ist dann durchgelaufen, aber
  ' SendInput hat nicht funktioniert.
  If p2.Terminated Then subLog "  ACHTUNG: SendInput hat nicht gegriffen, Zeitgrenze hat beendet."
  Set p2 = Nothing
End Sub

Sub TestAutoTerminate()
  Set p = New QSS_C_PipeProcess
  p.ApplicationName = "cmd"
  p.StdIn = "set" & vbCrLf & "this-program-doesnt-exist" & vbCrLf & "pause" & vbCrLf
  p.SeparateStdoutStderr = OptSeparateStreams()
  p.ShowWindow = OptShowWindow()
  p.TimeoutMillis = 500

  subLog "===== TestAutoTerminate test start"
  boolHandledPause = False
  p.Execute
  LogResult "TestAutoTerminate", p
  Set p = Nothing
End Sub

' Ohne Zeitgrenze - hier beendet p3_Tick von Hand. Das ist der einzige
' Test, der absichtlich kein Netz hat; das Netz ist der Ereignisbehandler.
Sub TestManualTerminate()
  Set p3 = New QSS_C_PipeProcess
  p3.ApplicationName = "cmd"
  p3.StdIn = "set" & vbCrLf & "this-program-doesnt-exist" & vbCrLf & "pause" & vbCrLf
  p3.SeparateStdoutStderr = OptSeparateStreams()
  p3.ShowWindow = OptShowWindow()
  p3.TimeoutMillis = 0  ' no timeout

  subLog "===== TestManualTerminate test start"
  boolHandledPause = False
  p3.Execute
  LogResult "TestManualTerminate", p3
  Set p3 = Nothing
End Sub


