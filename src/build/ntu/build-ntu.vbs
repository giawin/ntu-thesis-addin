Option Explicit
' Build NTUThesis.dotm - the standalone "Chuan hoa luan an NTU" add-in.
'
' Pipeline:
'   1. node vba2bas.mjs  (NTU.vba with Vietnamese -> NTU.bas pure ASCII + ChrW)
'   2. Word COM: create a blank .dotm, import NTU.bas, save with macros
'   3. pack-ntu.ps1 injects the ribbon (Word never preserves customUI itself)
'
' Usage: cscript //nologo build-ntu.vbs <repo-root>

Dim fso, sh, here, repo, srcDir, src, bas, outDir, rawTmp, raw, w, d, item, md, rc, i, lst, nodeExe

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

here = fso.GetParentFolderName(WScript.ScriptFullName)
repo = fso.GetParentFolderName(fso.GetParentFolderName(fso.GetParentFolderName(here)))
If WScript.Arguments.Count > 0 Then repo = WScript.Arguments(0)

srcDir = fso.BuildPath(repo, "src\modules\ntu-thesis")
src    = fso.BuildPath(srcDir, "NTU.vba")
outDir = fso.BuildPath(srcDir, "dist")
rawTmp = fso.BuildPath(srcDir, "dist_work")

If Not fso.FileExists(src) Then
  WScript.Echo "LOI: thieu nguon " & src
  WScript.Quit 1
End If
If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir
If Not fso.FolderExists(rawTmp) Then fso.CreateFolder rawTmp

bas = fso.BuildPath(rawTmp, "NTU.bas")
raw = fso.BuildPath(rawTmp, "NTUThesis-raw.dotm")

' ---- 1. VBA -> ASCII .bas ----
nodeExe = FindNode(fso)
If nodeExe = "" Then
  WScript.Echo "LOI: khong tim thay node.exe (can de chay vba2bas.mjs)"
  WScript.Quit 1
End If
rc = sh.Run("""" & nodeExe & """ """ & fso.BuildPath(repo, "src\build\vba2bas.mjs") & """ """ & src & """ """ & bas & """", 0, True)
If rc <> 0 Then
  WScript.Echo "LOI: vba2bas.mjs that bai (exit " & rc & ") - thuong la nhay kep chua escape trong chuoi VBA"
  WScript.Quit 1
End If
WScript.Echo "bas: " & fso.GetFile(bas).Size & " bytes"

' ---- 2. Word COM ----
On Error Resume Next
sh.RegWrite "HKCU\Software\Microsoft\Office\16.0\Word\Security\AccessVBOM", 1, "REG_DWORD"
sh.RegWrite "HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM", 1, "REG_DWORD"
On Error GoTo 0

If fso.FileExists(raw) Then fso.DeleteFile raw, True

Set w = CreateObject("Word.Application")
w.Visible = False
w.DisplayAlerts = 0
w.AutomationSecurity = 3
Set d = w.Documents.Add()

ReDim lst(-1)
i = 0
For Each item In d.VBProject.VBComponents
  ReDim Preserve lst(i)
  lst(i) = item.Name
  i = i + 1
Next
For i = 0 To UBound(lst)
  If lst(i) <> "ThisDocument" Then d.VBProject.VBComponents.Remove d.VBProject.VBComponents(lst(i))
Next

On Error Resume Next
Set md = d.VBProject.VBComponents.Import(bas)
If Err.Number <> 0 Then
  WScript.Echo "LOI: import that bai - " & Err.Number & " " & Err.Description
  Err.Clear
  d.Close 0
  w.Quit
  WScript.Quit 1
End If
On Error GoTo 0

WScript.Echo "imported: " & md.Name & " / " & md.CodeModule.CountOfLines & " dong"
If md.CodeModule.CountOfLines < 200 Then
  WScript.Echo "LOI: module qua ngan, huy"
  d.Close 0
  w.Quit
  WScript.Quit 1
End If

d.SaveAs2 raw, 15
d.Close 0
w.Quit
Set d = Nothing
Set w = Nothing

If Not fso.FileExists(raw) Then
  WScript.Echo "LOI: khong tao duoc .dotm"
  WScript.Quit 1
End If
WScript.Echo "raw: " & fso.GetFile(raw).Size & " bytes -> " & raw

Function FindNode(fso)
  Dim cand, k
  FindNode = ""
  cand = Split("C:\Program Files\nodejs\node.exe|C:\Program Files (x86)\nodejs\node.exe|" & _
               fso.BuildPath(sh.ExpandEnvironmentStrings("%LOCALAPPDATA%"), "Programs\nodejs\node.exe"), "|")
  For k = 0 To UBound(cand)
    If fso.FileExists(cand(k)) Then
      FindNode = cand(k)
      Exit Function
    End If
  Next
End Function