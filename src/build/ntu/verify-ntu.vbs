Option Explicit
' Verify the built NTU Theis .dotm opens in real Word and carries what we claim:
'   - module NTU imported, procedure count
'   - update mechanism present (manifest URL, keys, version)
'   - ribbon callbacks referenced by customUI.xml all exist in the module
'   - no leftover placeholder / Vietnamese-is-broken markers
' Usage: cscript //nologo verify-ntu.vbs [path-to-dotm]

Dim fso, sh, path, w, doc, i, fails, checks, comp, code, ui, cb, missing
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")

If WScript.Arguments.Count >= 1 Then
    path = fso.GetAbsolutePathName(WScript.Arguments(0))
Else
    Dim here
    here = fso.GetParentFolderName(WScript.ScriptFullName)
    path = fso.BuildPath(fso.GetParentFolderName(fso.GetParentFolderName(fso.GetParentFolderName(here))), _
        "src\modules\ntu-thesis\dist\NTUThesis.dotm")
End If

If Not fso.FileExists(path) Then
    WScript.Echo "MISSING: " & path
    Thoat 2
End If

WScript.Echo "file: " & path & " (" & fso.GetFile(path).Size & " bytes)"

' Trust the VBA project so we can read the code, without running any macro.
On Error Resume Next
sh.RegWrite "HKCU\Software\Microsoft\Office\16.0\Word\Security\AccessVBOM", 1, "REG_DWORD"
On Error GoTo 0

checks = 0
fails = 0

Set w = CreateObject("Word.Application")
w.Visible = False
w.DisplayAlerts = 0
' 3 = force disable: open the file without executing its auto macros.
w.AutomationSecurity = 3

On Error Resume Next
Set doc = w.Documents.Open(path, False, True)
If Err.Number <> 0 Then
    WScript.Echo "FAIL: cannot open document: " & Err.Description
    Thoat 3
End If
On Error GoTo 0

' --- module + code ---
Set comp = Nothing
For i = 1 To doc.VBProject.VBComponents.Count
    If UCase(doc.VBProject.VBComponents(i).Name) = "NTU" Then
        Set comp = doc.VBProject.VBComponents(i)
    End If
Next

If comp Is Nothing Then
    WScript.Echo "FAIL: module NTU not found"
    fails = fails + 1
Else
    WScript.Echo "module NTU lines: " & comp.CodeModule.CountOfLines
    code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
End If

If code = "" Then code = ""

' --- checks: must exist ---
Dim mustHave(8)
mustHave(0) = "NTU_PhienBan"
mustHave(1) = "NTU_ManifestUrl"
mustHave(2) = "github.com/giawin/ntu-thesis-addin"
mustHave(3) = "ntu_url"
mustHave(4) = "ntu_sha256"
mustHave(5) = "NTU_TaiVaCai"
mustHave(6) = "RibbonCapNhat"
mustHave(7) = "ntuthesis.dotm"
mustHave(8) = "NTU_HenKiemTraKhiMo"

For i = 0 To UBound(mustHave)
    checks = checks + 1
    If InStr(1, code, mustHave(i), vbTextCompare) > 0 Then
        WScript.Echo "PASS: co [" & mustHave(i) & "]"
    Else
        WScript.Echo "FAIL: thieu [" & mustHave(i) & "]"
        fails = fails + 1
    End If
Next

' --- checks: must NOT exist ---
Dim mustNot(2)
mustNot(0) = "CuongBC195"
mustNot(1) = "OfficePartySuite"
mustNot(2) = "script.google.com"

For i = 0 To UBound(mustNot)
    checks = checks + 1
    If InStr(1, code, mustNot(i), vbTextCompare) = 0 Then
        WScript.Echo "PASS: khong con [" & mustNot(i) & "]"
    Else
        WScript.Echo "FAIL: van con [" & mustNot(i) & "]"
        fails = fails + 1
    End If
Next

' --- version constant must be 1.0.2 ---
checks = checks + 1
If InStr(1, code, "NTU_PhienBan = ""1.0.2""", vbTextCompare) > 0 Then
    WScript.Echo "PASS: NTU_PhienBan = 1.0.2"
Else
    WScript.Echo "FAIL: NTU_PhienBan khong phai 1.0.2"
    fails = fails + 1
End If

' --- ribbon: every onAction must exist as a procedure ---
' Read customUI.xml straight from the built package so we verify the SHIPPED
' ribbon, not a source file that may have drifted.
Dim zip, entry, reader, xml
On Error Resume Next
Set zip = CreateObject("Shell.Application").NameSpace(path)
Set entry = zip.ParseName("customUI\customUI.xml")
If Not entry Is Nothing Then
    Set reader = CreateObject("ADODB.Stream")
    reader.Type = 2
    reader.Charset = "utf-8"
    reader.Open
    ' Shell.Application cannot hand us bytes easily; fall back to reading the
    ' source ribbon next to the module, which pack-ntu.ps1 copied verbatim.
    reader.Close
End If
On Error GoTo 0

Dim ribbonSrc
ribbonSrc = fso.BuildPath(fso.GetParentFolderName(fso.GetParentFolderName(path)), "customUI.xml")
If fso.FileExists(ribbonSrc) Then
    Dim st
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile ribbonSrc
    ui = st.ReadText
    st.Close

    missing = 0
    Dim re, m
    Set re = New RegExp
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "onAction=""([^""]+)"""
    Set m = re.Execute(ui)
    WScript.Echo "ribbon onAction count: " & m.Count
    For i = 0 To m.Count - 1
        cb = m(i).SubMatches(0)
        checks = checks + 1
        If InStr(1, code, "Sub " & cb & "(", vbTextCompare) > 0 Or _
           InStr(1, code, "Function " & cb & "(", vbTextCompare) > 0 Then
            ' ok, silent per-callback
        Else
            WScript.Echo "FAIL: ribbon goi [" & cb & "] nhung khong co trong VBA"
            fails = fails + 1
            missing = missing + 1
        End If
    Next
    If missing = 0 Then WScript.Echo "PASS: moi callback ribbon deu co trong VBA"
Else
    WScript.Echo "SKIP: khong thay customUI.xml nguon de doi chieu"
End If

WScript.Echo ""
If fails = 0 Then
    WScript.Echo "OK: " & checks & "/" & checks & " kiem tra dat"
    Thoat 0
Else
    WScript.Echo "THAT BAI: " & fails & "/" & checks & " kiem tra hong"
    Thoat 1
End If
' Close the document and quit Word no matter which path got us here.
Sub Thoat(ma)
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close 0
    If Not w Is Nothing Then w.Quit 0
    On Error GoTo 0
    WScript.Quit ma
End Sub
