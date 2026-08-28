# vba-pipe-process

Run external programs from VBA with **redirected stdin, stdout and stderr** —
no console window, no temporary files, works on **32-bit and 64-bit** Office.

Uses `CreateProcess` and `CreatePipe` directly. No .NET, no COM component,
no external dependency of any kind.

```vba
Dim rpr As RedirectProcessResult
rpr = fncRedirectProcess("cmd", "/c ver", "", False, True)

Debug.Print rpr.lngExitCode      ' 0
Debug.Print rpr.strStdOut        ' Microsoft Windows [Version 10.0.26100.9168]
Debug.Print rpr.strStdErr        '
```

## Why this exists

VBA gives you `Shell`, and that is all. If you need the output of a program
you have to redirect it through `cmd` into a temporary file and read the file
afterwards. That works, but it flashes a console window at the user, it leaves
files behind, and it cannot write to a process's standard input at all.

`WScript.Shell` does not solve it either:

| | hidden window | stdout | stdin | timeout | exit code |
|---|---|---|---|---|---|
| `Shell` + redirect to file | with `vbHide` | yes | no | via API | via API |
| `WScript.Shell.Run` | yes | no | no | **no** | yes |
| `WScript.Shell.Exec` | **no** | yes | yes | no | yes |
| this | yes | yes | yes | yes | yes |

`Exec` gives you the streams but always shows a window. `Run` hides the window
but has no way to read the output and waits forever. There is no combination
of the two that does everything.

## Origin

The 32-bit implementation is the work of **knoxg / randomnoun**, published in
2013 under the BSD 2-Clause licence:

> https://www.randomnoun.com/wp/2013/10/08/its-all-made-out-of-pipes/

It is included unchanged in [`original/`](original/) so you can diff it
yourself. What this repository adds is the 64-bit port, which the blog post
predates.

## What changed for 64-bit

All Windows API declarations were rewritten with `#If VBA7` / `PtrSafe`, and
every handle and pointer became `LongPtr`. The structures matter more than the
declarations — `STARTUPINFO` grows from 68 to 104 bytes, and if a single
member has the wrong width, `CreateProcess` receives a malformed structure and
Office dies without an error message.

Checked against the Microsoft documentation and the mingw-w64 headers:

```c
typedef struct _STARTUPINFOA {
  DWORD  cb;              LPSTR  lpReserved;      LPSTR  lpDesktop;
  LPSTR  lpTitle;         DWORD  dwX;             DWORD  dwY;
  DWORD  dwXSize;         DWORD  dwYSize;         DWORD  dwXCountChars;
  DWORD  dwYCountChars;   DWORD  dwFillAttribute; DWORD  dwFlags;
  WORD   wShowWindow;     WORD   cbReserved2;     LPBYTE lpReserved2;
  HANDLE hStdInput;       HANDLE hStdOutput;      HANDLE hStdError;
} STARTUPINFOA;
```

So: three `LPSTR`, one `LPBYTE` and three `HANDLE` are `LongPtr`; the eight
`DWORD` stay `Long`; the two `WORD` stay `Integer`.

**One bug fixed.** The original declares `lpReserved2 As Byte`. It is a
pointer (`LPBYTE`). On 32-bit this happens to work by accident — VBA pads the
one byte with three, which comes out the right size. On 64-bit it does not.

`PROCESS_INFORMATION` (two `HANDLE`, two `DWORD`) and `SECURITY_ATTRIBUTES`
(`DWORD`, `LPVOID`, `BOOL`) were converted the same way.

Also fixed, unrelated to bitness: `Dim lngResult, lngResult2 As Long` declares
`lngResult` as a `Variant`. Only the last name in such a list gets the type.

## Files

| File | |
|---|---|
| `mdlPipeProcess.bas` | `fncRedirectProcess()` — one call, synchronous, returns everything |
| `clsPipeProcess.cls` | class with `StdoutAvailable`, `StderrAvailable`, `Tick` events, `SendInput`, `Terminate`, idle timeout |
| `frmProcessTest.cls` | code behind the test form from the original .mdb |
| `original/` | the 2013 files, unmodified |

Import the `.bas` and `.cls` through the VBA editor: **Alt+F11**, then
*File → Import File*. The module is enough for most uses; take the class when
you need output while the process is still running, or need to send input in
response to it.

## Two things to know

**`TimeoutMillis` is an idle timeout, not a total one.** The process is killed
when *nothing* arrives on stdout or stderr for that long. A program that works
silently for a while — a validator, a compiler — will be killed even though it
is doing its job. Set it generously, or to 0 to disable it.

**`GetFileSize` is used to poll the pipe.** Microsoft does not document this
for anonymous pipes; the supported call is `PeekNamedPipe`, which explicitly
accepts the read end of a `CreatePipe` pipe and reports the available bytes in
`lpTotalBytesAvail`. In practice `GetFileSize` works, and it is left as the
original author wrote it. If a process runs but `strStdOut` stays empty, this
is the first place to look.

## Notes on the test form

The form's `p2` test waits for a specific English error message from `cmd`
before sending `exit` through stdin:

```vba
If InStr(p2.StdErr, "is not recognized as an internal or external command") > 0
```

On a non-English Windows this never matches, `exit` is never sent, `cmd` keeps
waiting for input and the window stays open. This repository checks
`Len(p2.StdErr) > 0` instead, which works in any language.

The same class of problem is worth remembering generally: **judge a process by
its exit code, not by the text it prints.** The text is localised; the exit
code is not.

## Licence

BSD 2-Clause. See [LICENSE](LICENSE). Both copyright notices must be kept.
