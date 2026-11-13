Option Explicit
' Build GiaTwinH_NTU_Thesis_Word.dotm
'
' Strategy: do NOT rebuild the .dotm from scratch (that would lose the
' word/glossary/* building blocks that hold the 50 sample-document templates
' used by the Ribbon menus). Instead open the ORIGINAL .dotm and replace only
' the Module1 component, then SaveAs2 with wdFormatXMLTemplateMacroEnabled.
'
' HOW the component is replaced matters, and this file used to get it wrong:
'   BROKEN: cm.DeleteLines 1, cm.CountOfLines  then  cm.AddFromString text
'           -> the project is left in a state that does not compile, Word hides
'              EVERY macro of the project, and w.Run("CN_PhienBan") answers
'              -2147352573 "the macro cannot be found" even though the text is
'              all there. A text-grep verify cannot see this.
'   CORRECT: VBComponents.Remove  then  VBComponents.Import <file>.bas
'           -> the module is registered the way the VBA IDE registers it, the
'              project compiles, macros are visible. This is the same path
'              build-ntu.vbs uses for the NTU module.
' Import requires a .bas file with "Attribute VB_Name = ..." as its first line,
' so Module1.vba is copied to a temp .bas instead of being passed directly.
'
' Why VBScript and not PowerShell: Word COM through PowerShell hangs on
' SaveAs2. cscript/VBScript is the only reliable path on this machine.
'
' Usage: cscript //nologo build-word.vbs <repo-root>
'   repo-root defaults to three levels up from this script.

Dim fso, sh, here, repo, upstream, src, outDir, outFile, basePath, w, d, p, comp, cm

Const MODULE_NAME = "Module1"
Const MIN_LINES = 8000

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
  WScript.Echo "LOI: thieu nguon " & src
  WScript.Quit 1
End If
If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir

' ---- 1. Source must be pure ASCII and must carry the VB_Name attribute ----
If Not IsPureAscii(src) Then
  WScript.Echo "LOI: nguon co byte > 127 (VBA luu code theo codepage ANSI)"
  WScript.Quit 1
End If
If Not HasVbNameAttribute(src) Then
  WScript.Echo "LOI: nguon thieu dong 'Attribute VB_Name = """ & MODULE_NAME & """' o dau file"
  WScript.Quit 1
End If
WScript.Echo "source: " & fso.GetFile(src).Size & " bytes, ASCII ok, VB_Name ok"

' ---- 2. Stage the source as a .bas (Import needs that extension) ----
basePath = fso.BuildPath(sh.ExpandEnvironmentStrings("%TEMP%"), "ntu-word-build_" & MODULE_NAME & ".bas")
fso.CopyFile src, basePath, True
WScript.Echo "staged bas: " & basePath

' ---- 3. Trust access to the VBA project object model ----
' Word resets this to 0 on its own; set it right before every build.
On Error Resume Next
sh.RegWrite "HKCU\Software\Microsoft\Office\16.0\Word\Security\AccessVBOM", 1, "REG_DWORD"
sh.RegWrite "HKCU\Software\Microsoft\Office\16.0\Word\Security\AccessVBOM", 1, "REG_DWORD"
On Error GoTo 0

' ---- 4. Open the original, swap the component, save as the new add-in ----
If fso.FileExists(outFile) Then fso.DeleteFile outFile, True

Set w = CreateObject("Word.Application")
w.Visible = False
w.DisplayAlerts = 0
' msoAutomationSecurityForceDisable = 3: never run the original AutoExec /
' license gate / update check while we are only editing the code.
w.AutomationSecurity = 3

Set d = w.Documents.Open(upstream, False, False, False)

On Error Resume Next
Set p = d.VBProject
If Err.Number <> 0 Then
  WScript.Echo "LOI: khong truy cap duoc VBProject (AccessVBOM bi chan?). Err=" & Err.Number & " " & Err.Description
  Err.Clear
  Cleanup d, w
  WScript.Quit 1
End If
Set comp = p.VBComponents(MODULE_NAME)
If Err.Number <> 0 Then
  WScript.Echo "LOI: khong thay component " & MODULE_NAME & ". Err=" & Err.Number
  Err.Clear
  Cleanup d, w
  WScript.Quit 1
End If
WScript.Echo "original " & MODULE_NAME & ": " & comp.CodeModule.CountOfLines & " lines"
On Error GoTo 0

' Remove then Import -- never DeleteLines + AddFromString (see header comment).
On Error Resume Next
p.VBComponents.Remove comp
If Err.Number <> 0 Then
  WScript.Echo "LOI: khong xoa duoc component cu. Err=" & Err.Number & " " & Err.Description
  Err.Clear
  Cleanup d, w
  WScript.Quit 1
End If
p.VBComponents.Import basePath
If Err.Number <> 0 Then
  WScript.Echo "LOI: import that bai. Err=" & Err.Number & " " & Err.Description
  Err.Clear
  Cleanup d, w
  WScript.Quit 1
End If
Set comp = p.VBComponents(MODULE_NAME)
If Err.Number <> 0 Then
  WScript.Echo "LOI: sau import khong thay " & MODULE_NAME & ". Err=" & Err.Number
  Err.Clear
  Cleanup d, w
  WScript.Quit 1
End If
Set cm = comp.CodeModule
WScript.Echo "new " & MODULE_NAME & ": " & cm.CountOfLines & " lines"
Err.Clear
On Error GoTo 0

If cm.CountOfLines < MIN_LINES Then
  WScript.Echo "LOI: so dong qua it (" & cm.CountOfLines & " < " & MIN_LINES & ") - huy"
  Cleanup d, w
  WScript.Quit 1
End If

On Error Resume Next
d.SaveAs2 outFile, 15   ' wdFormatXMLTemplateMacroEnabled
If Err.Number <> 0 Then
  WScript.Echo "LOI: SaveAs2 that bai. Err=" & Err.Number & " " & Err.Description
  Err.Clear
  Cleanup d, w
  WScript.Quit 1
End If
Err.Clear
On Error GoTo 0

Cleanup d, w

If Not fso.FileExists(outFile) Then
  WScript.Echo "LOI: khong tao duoc " & outFile
  WScript.Quit 1
End If
WScript.Echo "OK -> " & outFile & " (" & fso.GetFile(outFile).Size & " bytes)"
WScript.Quit 0

' Always close the document and quit Word, including on every failure path:
' a leaked WINWORD.EXE makes the next pack step fail with
' "... because it is being used by another process".
Sub Cleanup(doc, app)
  On Error Resume Next
  If Not doc Is Nothing Then doc.Close 0
  If Not app Is Nothing Then app.Quit 0
  Set doc = Nothing
  Set app = Nothing
  On Error GoTo 0
End Sub

' True when every byte of the file is < 128.
Function IsPureAscii(path)
  Dim st, txt, k
  IsPureAscii = False
  Set st = CreateObject("ADODB.Stream")
  st.Type = 2
  st.Charset = "us-ascii"
  st.Open
  st.LoadFromFile path
  txt = st.ReadText(-1)
  st.Close
  For k = 1 To Len(txt)
    If Asc(Mid(txt, k, 1)) > 127 Then
      WScript.Echo "  byte > 127 tai vi tri " & k
      Exit Function
    End If
  Next
  IsPureAscii = True
End Function

' True when the first line is Attribute VB_Name = "<MODULE_NAME>".
Function HasVbNameAttribute(path)
  Dim st, txt, firstLine, nl
  HasVbNameAttribute = False
  Set st = CreateObject("ADODB.Stream")
  st.Type = 2
  st.Charset = "us-ascii"
  st.Open
  st.LoadFromFile path
  txt = st.ReadText(-1)
  st.Close
  nl = InStr(txt, vbCrLf)
  If nl = 0 Then nl = InStr(txt, vbLf)
  If nl = 0 Then
    firstLine = txt
  Else
    firstLine = Left(txt, nl - 1)
  End If
  firstLine = Trim(firstLine)
  If firstLine = "Attribute VB_Name = """ & MODULE_NAME & """" Then HasVbNameAttribute = True
End Function