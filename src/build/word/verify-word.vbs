Option Explicit
' Verify a .dotm built from source.
'
' Two classes of check, and the difference between them matters:
'
'   1. STRUCTURAL (text scan) - the strings we personalised are present, the
'      upstream ones are gone, the license gate returns True, the ribbon
'      callbacks are referenced. Fast, and it works with macros disabled.
'
'   2. RUNTIME (real invocation) - the add-in is loaded the way a user loads it
'      and its public entry points are actually CALLED. This is the check that
'      was missing: a text scan proves the source is in the file, not that the
'      VBA project compiles. A module patched with
'      "CodeModule.DeleteLines + AddFromString" passes every text scan while
'      Word hides every macro of the project, and Application.Run answers
'      -2147352573 "the macro cannot be found". Only calling it catches that.
'
' Both classes count: any FAIL here exits non-zero, so CI and build-all.ps1
' stop instead of shipping a package whose buttons do nothing.
'
' Usage: cscript //nologo verify-word.vbs <path-to.dotm>
'
' Every exit goes through Thoat(), which always unloads the add-in, closes the
' document and quits Word. A bare WScript.Quit would leave a WINWORD.EXE
' behind, and the next build would fail with "used by another process".

Dim fso, sh, path, w, d, comp, cm, txt, fails, k

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")
fails = 0

If WScript.Arguments.Count < 1 Then
  WScript.Echo "usage: cscript verify-word.vbs <path-to.dotm>"
  WScript.Quit 2
End If
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then
  WScript.Echo "LOI: khong thay " & path
  WScript.Quit 1
End If

WScript.Echo "file: " & path & " (" & fso.GetFile(path).Size & " bytes)"
WScript.Echo ""
WScript.Echo "== 1. KIEM TRA CAU TRUC (quet van ban) =="

' msoAutomationSecurityForceDisable: do not execute anything while inspecting
' the source text. The runtime phase below deliberately re-enables macros.
Set w = CreateObject("Word.Application")
w.Visible = False
w.DisplayAlerts = 0
w.AutomationSecurity = 3

On Error Resume Next
Set d = w.Documents.Open(path, False, False, False)
If Err.Number <> 0 Then
  WScript.Echo "FAIL: mo file that bai - " & Err.Number & " " & Err.Description
  Err.Clear
  Thoat 1
End If
Set comp = d.VBProject.VBComponents("Module1")
If Err.Number <> 0 Then
  WScript.Echo "FAIL: khong co Module1 - " & Err.Number
  Err.Clear
  Thoat 1
End If
On Error GoTo 0

Set cm = comp.CodeModule
WScript.Echo "Module1 lines: " & cm.CountOfLines
If cm.CountOfLines < 8000 Then
  WScript.Echo "FAIL: Module1 qua it dong (" & cm.CountOfLines & " < 8000)"
  fails = fails + 1
End If
txt = cm.Lines(1, cm.CountOfLines)

' --- personalised values present / upstream values gone ---
Check "NTUThesisAddin", txt, True
Check "giawin/ntu-thesis-addin", txt, True
Check "giatwin_h_ntu_thesis_word.dotm", txt, True
Check "CN_PhienBan = ""1.0.2""", txt, True
Check "script.google.com", txt, False
Check "chuanhoa_vanbandang.dotm", txt, False
Check "OfficePartySuite", txt, False

' --- license gate neutralised ---
If InStr(txt, "CheckLicenseActivation = True") > 0 Then
  WScript.Echo "PASS: license gate returns True"
Else
  WScript.Echo "FAIL: license gate was not neutralised"
  fails = fails + 1
End If

' --- ribbon callbacks referenced ---
If InStr(txt, "RibbonChuanHoa") > 0 Then
  WScript.Echo "PASS: ribbon callbacks referenced"
Else
  WScript.Echo "FAIL: ribbon callbacks missing"
  fails = fails + 1
End If

' Unload the inspection copy and release Word before the runtime phase: the
' add-in has to be registered fresh, not inspected while open.
On Error Resume Next
d.Close 0
Set d = Nothing
w.Quit 0
Set w = Nothing
On Error GoTo 0

WScript.Echo ""
WScript.Echo "== 2. KIEM TRA RUNTIME (nap add-in va GOI THAT) =="

Set w = CreateObject("Word.Application")
w.Visible = False
w.DisplayAlerts = 0
' msoAutomationSecurityLow = 1: macros ENABLED. Required - Application.Run
' cannot find any macro while macro execution is disabled, which is why an
' earlier probe reported "macro cannot be found" on a perfectly good build.
w.AutomationSecurity = 1

On Error Resume Next
' Guard against grading the wrong file.
'
' Word auto-loads any copy of this add-in already sitting in Word\STARTUP. That
' copy exposes the SAME public procedure names, and Application.Run resolves to
' whichever one registered first -- in practice the installed one. The checks
' would then grade the OLD build: a fresh build with a real defect would report
' the previous version's answer and go green.
'
' This script does NOT try to fix that itself. Detaching the installed copy with
' `Installed = False` writes to the user's registry and leaves their add-in
' switched off after the check -- a verification step must never change the
' machine it verifies. Instead build-all.ps1 moves same-named STARTUP copies
' aside for the duration of the check; if that did not happen, say so and fail
' rather than silently reporting someone else's numbers.
Dim i, tenCuaTa, trung
tenCuaTa = LCase(fso.GetFileName(path))
trung = 0
On Error Resume Next
For i = 1 To w.AddIns.Count
  If LCase(w.AddIns(i).Name) = tenCuaTa Then trung = trung + 1
Next
Err.Clear
On Error GoTo 0
If trung > 0 Then
  WScript.Echo "FAIL: da co " & trung & " ban cung ten dang nap trong Word."
  WScript.Echo "      Ket qua se phan anh ban da cai, khong phai tep vua build."
  WScript.Echo "      Hay chay qua src/build/build-all.ps1 (tu cach ly thu muc STARTUP)."
  fails = fails + 1
  Thoat 1
End If

On Error Resume Next
Set d = w.AddIns.Add(path, True)   ' load as a real add-in, installed
If Err.Number <> 0 Then
  WScript.Echo "FAIL: khong nap duoc add-in - " & Err.Number & " " & Err.Description
  Err.Clear
  Thoat 1
End If
If Not d.Installed Then
  WScript.Echo "FAIL: add-in khong o trang thai Installed"
  fails = fails + 1
Else
  WScript.Echo "PASS: add-in da nap (" & d.Name & ")"
End If
Err.Clear
On Error GoTo 0

' A missing macro surfaces here as -2147352573; a broken project as 500.
RunCheck "CN_PhienBan", "1.0.2"
' TT_ProfileCount comes from the upstream template, not from our personalization,
' so pinning it to a literal only breaks every time upstream legitimately adds a
' profile (it went 5 -> 6 on 2026.11.13). What we must prove is that the function
' is still reachable and returns real data: a positive integer.
RunCheckPositive "TT_ProfileCount"
RunCheck "CN_ManifestUrl", "https://github.com/giawin/ntu-thesis-addin/releases/latest/download/version.json"

WScript.Echo ""
If fails = 0 Then
  WScript.Echo "OK: tat ca kiem tra dat (cau truc + runtime)"
  Thoat 0
Else
  WScript.Echo "THAT BAI: " & fails & " kiem tra khong dat"
  Thoat 1
End If

' Unload the add-in, then close and quit, whatever path got us here.
' addIn.Delete removes it from the AddIns collection so a later run starts clean.
Sub Thoat(ma)
  On Error Resume Next
  If Not d Is Nothing Then
    d.Installed = False
    d.Delete
  End If
  If Not w Is Nothing Then w.Quit 0
  On Error GoTo 0
  WScript.Quit ma
End Sub

' Invoke a public entry point and compare what it returns.
' This is the whole point of the runtime phase: Application.Run is the same
' mechanism the ribbon callbacks and the update button use, so if Word can
' find and execute these, the shipped add-in is usable.
Sub RunCheck(macroName, expected)
  Dim got
  On Error Resume Next
  got = w.Run(macroName)
  If Err.Number <> 0 Then
    WScript.Echo "FAIL: goi " & macroName & " loi " & Err.Number & " (macro khong ton tai hoac project khong bien dich)"
    fails = fails + 1
    Err.Clear
    On Error GoTo 0
    Exit Sub
  End If
  Err.Clear
  On Error GoTo 0
  If CStr(got) = expected Then
    WScript.Echo "PASS: " & macroName & "() = [" & got & "]"
  Else
    WScript.Echo "FAIL: " & macroName & "() = [" & got & "] - mong doi [" & expected & "]"
    fails = fails + 1
  End If
End Sub

' Like RunCheck, but for values owned by the upstream template: only assert the
' macro runs and returns a positive integer, never a value we do not control.
Sub RunCheckPositive(macroName)
  Dim got, n
  On Error Resume Next
  got = w.Run(macroName)
  If Err.Number <> 0 Then
    WScript.Echo "FAIL: goi " & macroName & " loi " & Err.Number & " (macro khong ton tai hoac project khong bien dich)"
    fails = fails + 1
    Err.Clear
    On Error GoTo 0
    Exit Sub
  End If
  Err.Clear
  On Error GoTo 0
  n = -1
  On Error Resume Next
  n = CLng(got)
  Err.Clear
  On Error GoTo 0
  If n > 0 Then
    WScript.Echo "PASS: " & macroName & "() = [" & got & "] (so nguyen duong)"
  Else
    WScript.Echo "FAIL: " & macroName & "() = [" & got & "] - mong doi so nguyen duong"
    fails = fails + 1
  End If
End Sub

' Structural string check that counts failures instead of only printing them.
Sub Check(needle, hay, mustExist)
  Dim hit
  hit = (InStr(hay, needle) > 0)
  If hit = mustExist Then
    WScript.Echo "PASS: " & IIf(mustExist, "co", "khong con") & " [" & needle & "]"
  Else
    WScript.Echo "FAIL: " & IIf(mustExist, "thieu", "van con") & " [" & needle & "]"
    fails = fails + 1
  End If
End Sub

Function IIf(c, a, b)
  If c Then IIf = a Else IIf = b
End Function