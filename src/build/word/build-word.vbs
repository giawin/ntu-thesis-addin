Option Explicit
' Build GiaTwinH_NTU_Thesis_Word.dotm
'
' Strategy: do NOT rebuild the .dotm from scratch (that would lose the
' word/glossary/* building blocks that hold the 50 sample-document templates
' used by the Ribbon menus). Instead open the ORIGINAL .dotm and replace only
' the Module1 source text, then SaveAs2 with wdFormatXMLTemplateMacroEnabled.
'
' Why VBScript and not PowerShell: Word COM through PowerShell hangs on
' SaveAs2. cscript/VBScript is the only reliable path on this machine.
'
' Usage: cscript //nologo build-word.vbs <repo-root>
'   repo-root defaults to three levels up from this script.

Dim fso, sh, here, repo, upstream, src, outDir, outFile, codeText, w, d, comp, cm, i

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

here = fso.GetParentFolderName(WScript.ScriptFullName)
repo = fso.GetParentFolderName(fso.GetParentFolderName(fso.GetParentFolderName(here)))
If WScript.Arguments.Count > 0 Then repo = WScript.Arguments(0)

upstream = fso.BuildPath(repo, "src\core\vendor\upstream\ChuanHoa_VanBanDang.dotm")
src      = fso.BuildPath(repo, "src\core\word\Module1.vba")
outDir   = fso.BuildPath(repo, "src\core\word\dist")
outFile  = fso.BuildPath(outDir, "GiaTwinH_NTU_Thesis_Word.dotm")

If Not fso.FileExists(upstream) Then
  WScript.Echo "LOI: thieu ban goc " & upstream
  WScript.Quit 1
End If
If Not fso.FileExists(src) Then
  WScript.Echo "LOI: thieu nguon " & src & " (agent word-core phai tao truoc)"
  WScript.Quit 1
End If
If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir

' ---- 1. VBA source must be pure ASCII (VBA stores code in ANSI codepage) ----
codeText = ReadAscii(src)
If codeText = "" Then
  WScript.Echo "LOI: nguon rong"
  WScript.Quit 1
End If
WScript.Echo "source: " & Len(codeText) & " chars, " & (UBound(Split(codeText, vbCrLf)) + 1) & " lines"

' ---- 2. Trust access to the VBA project object model ----
' Word resets this to 0 on its own; set it right before every build.
On Error Resume Next
sh.RegWrite "HKCU\Software\Microsoft\Office\16.0\Word\Security\AccessVBOM", 1, "REG_DWORD"
sh.RegWrite "HKCU\Software\Microsoft\Office\16.0\Word\Security\AccessVBOM", 1, "REG_DWORD"
On Error GoTo 0

' ---- 3. Open the original, swap the code, save as the new add-in ----
If fso.FileExists(outFile) Then fso.DeleteFile outFile, True

Set w = CreateObject("Word.Application")
w.Visible = False
w.DisplayAlerts = 0
' msoAutomationSecurityForceDisable = 3: never run the original AutoExec /
' license gate / update check while we are only editing the code.
w.AutomationSecurity = 3

Set d = w.Documents.Open(upstream, False, False, False)

On Error Resume Next
Set comp = d.VBProject.VBComponents("Module1")
If Err.Number <> 0 Then
  WScript.Echo "LOI: khong truy cap duoc VBProject (AccessVBOM bi chan?). Err=" & Err.Number & " " & Err.Description
  Err.Clear
  d.Close 0
  w.Quit
  WScript.Quit 1
End If
On Error GoTo 0

Set cm = comp.CodeModule
WScript.Echo "original Module1: " & cm.CountOfLines & " lines"

' Delete every existing line, then insert the personalized source.
If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
cm.AddFromString codeText
WScript.Echo "new Module1: " & cm.CountOfLines & " lines"

If cm.CountOfLines < 8000 Then
  WScript.Echo "LOI: so dong qua it (" & cm.CountOfLines & ") - huy"
  d.Close 0
  w.Quit
  WScript.Quit 1
End If

d.SaveAs2 outFile, 15   ' wdFormatXMLTemplateMacroEnabled
d.Close 0
w.Quit
Set d = Nothing
Set w = Nothing

If Not fso.FileExists(outFile) Then
  WScript.Echo "LOI: khong tao duoc " & outFile
  WScript.Quit 1
End If
WScript.Echo "OK -> " & outFile & " (" & fso.GetFile(outFile).Size & " bytes)"

' Read a UTF-8 / ASCII file into a string, refusing any byte > 127.
Function ReadAscii(path)
  Dim st, txt, k, b
  ReadAscii = ""
  Set st = CreateObject("ADODB.Stream")
  st.Type = 2
  st.Charset = "us-ascii"
  st.Open
  st.LoadFromFile path
  txt = st.ReadText(-1)
  st.Close
  For k = 1 To Len(txt)
    b = Asc(Mid(txt, k, 1))
    If b > 127 Then
      WScript.Echo "LOI: ky tu khong phai ASCII tai vi tri " & k & " (byte " & b & ")"
      Exit Function
    End If
  Next
  txt = Replace(txt, vbCrLf, vbLf)
  ReadAscii = Replace(txt, vbLf, vbCrLf)
End Function
