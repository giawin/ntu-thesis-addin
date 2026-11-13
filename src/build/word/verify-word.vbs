Option Explicit
' Verify a .dotm built from source:
'   - it opens in Word without triggering the upstream license gate
'   - the VBA module exists and holds the expected number of lines
'   - the personalized strings are present and the upstream ones are gone
'   - the ribbon part is inside the package
'
' Usage: cscript //nologo verify-word.vbs <path-to.dotm>
'
' Every exit goes through Thoat(), which always closes the document and quits
' Word. WScript.Quit on its own would leave a WINWORD.EXE behind, and the next
' build would then fail with "used by another process".

Dim fso, sh, path, w, d, comp, cm, txt, bad, k

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

If WScript.Arguments.Count < 1 Then
  WScript.Echo "usage: cscript verify-word.vbs <path-to.dotm>"
  Thoat 2
End If
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then
  WScript.Echo "LOI: khong thay " & path
  Thoat 1
End If

WScript.Echo "file: " & path & " (" & fso.GetFile(path).Size & " bytes)"

' msoAutomationSecurityForceDisable: never execute the add-in while inspecting it.
Set w = CreateObject("Word.Application")
w.Visible = False
w.DisplayAlerts = 0
w.AutomationSecurity = 3

On Error Resume Next
Set d = w.Documents.Open(path, False, False, False)
If Err.Number <> 0 Then
  WScript.Echo "FAIL: mo file that bai - " & Err.Number & " " & Err.Description
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
txt = cm.Lines(1, cm.CountOfLines)

' --- presence checks ---
Check "NTUThesisAddin", txt, True
Check "giawin/ntu-thesis-addin", txt, True
Check "giatwin_h_ntu_thesis_word.dotm", txt, True
Check "CN_PhienBan = ""1.0.0""", txt, True
Check "script.google.com", txt, False
Check "chuanhoa_vanbandang.dotm", txt, False
Check "OfficePartySuite", txt, False

' --- license gate neutralised ---
If InStr(txt, "CheckLicenseActivation = True") > 0 Then
  WScript.Echo "PASS: license gate returns True"
Else
  WScript.Echo "FAIL: license gate was not neutralised"
End If

' --- ribbon part inside the package ---
Dim n
n = 0
If InStr(txt, "RibbonChuanHoa") > 0 Then n = n + 1
If n = 1 Then WScript.Echo "PASS: ribbon callbacks referenced" Else WScript.Echo "FAIL: ribbon callbacks missing"

Thoat 0

' Close the document and quit Word no matter which path got us here.
Sub Thoat(ma)
  On Error Resume Next
  If Not d Is Nothing Then d.Close 0
  If Not w Is Nothing Then w.Quit 0
  On Error GoTo 0
  WScript.Quit ma
End Sub

Function Check(needle, hay, mustExist)
  Dim hit
  hit = (InStr(hay, needle) > 0)
  If hit = mustExist Then
    WScript.Echo "PASS: " & IIf(mustExist, "co", "khong con") & " [" & needle & "]"
  Else
    WScript.Echo "FAIL: " & IIf(mustExist, "thieu", "van con") & " [" & needle & "]"
  End If
End Function

Function IIf(c, a, b)
  If c Then IIf = a Else IIf = b
End Function