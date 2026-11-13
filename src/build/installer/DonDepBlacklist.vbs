Option Explicit
' Clear Word/Excel "disabled item" blacklist entries that point at OUR add-ins.
'
' Why this is needed and why it is not paranoia:
'   When Office fails to load an add-in (a corrupt build, a crash during load,
'   a security prompt dismissed the wrong way) it does NOT just skip the file --
'   it records the path in
'     HKCU\Software\Microsoft\Office\16.0\<App>\Resiliency\DisabledItems
'   and refuses to load that path again, silently, forever. The file sits in
'   STARTUP looking perfectly installed while its tab never appears.
'
'   That is exactly what happened on this machine while v1.0.0 was broken: after
'   the build was fixed, Word STILL refused to load it until the entry was
'   deleted. Anyone who installed the broken v1.0.0 first would hit the same
'   wall after upgrading -- reinstall would look successful and change nothing.
'
' Scope: only entries whose decoded path contains our file names are removed.
' Other entries (other vendors' broken add-ins, the user's own choices) are
' left untouched.
'
' Usage: cscript //nologo DonDepBlacklist.vbs
' Exit codes: 0 = cleaned or nothing to do, 1 = could not enumerate registry.

Dim fso, sh, OUR_NAMES, vers, apps, removed, inspected
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

' Lower-case fragments that identify our own files. Keep these in step with the
' file names in CaiDat.bat.
OUR_NAMES = Array("giatwinh_ntu_thesis_word.dotm", "ntuthesis.dotm", _
                  "chuanhoa_vanbandang.dotm", "chuanhoa_excel.xlam", "tuchuanhoa.dotm")

' Office 2016/2019/2021/365 all report 16.0. Older versions use unversioned
' keys; try them too rather than guessing which Office the user has.
vers = Array("16.0", "15.0", "14.0", "")
apps = Array("Word", "Excel")

removed = 0
inspected = 0

Dim v, a, subPath, result
For Each v In vers
  For Each a In apps
    If Len(v) > 0 Then
      subPath = "Software\Microsoft\Office\" & v & "\" & a & "\Resiliency\DisabledItems"
    Else
      subPath = "Software\Microsoft\Office\" & a & "\Resiliency\DisabledItems"
    End If
    result = CleanKey(subPath)
    If result < 0 Then
      WScript.Echo "LOI: khong doc duoc HKCU\" & subPath
      WScript.Quit 1
    End If
    inspected = inspected + result
  Next
Next

If removed = 0 Then
  WScript.Echo "OK: khong co muc nao cua add-in trong danh sach chan (da kiem tra " & inspected & " muc)"
Else
  WScript.Echo "OK: da go " & removed & " muc chan add-in (da kiem tra " & inspected & " muc)"
End If
WScript.Quit 0

' Enumerate the values of one DisabledItems key and delete the ones that name
' our files. Returns the number of values inspected, or -1 if the key could not
' be opened. IMPORTANT: StdRegProv takes the hive as a separate argument, so the
' path must NOT carry the HKEY_CURRENT_USER prefix -- passing it makes EnumValues
' fail and, under On Error Resume Next, fail silently as "0 entries".
Function CleanKey(keyPath)
  Dim prov, names, types, n, i, k, raw, decoded, del, filtered
  CleanKey = 0
  On Error Resume Next
  Set prov = GetObject("winmgmts:\\.\root\default:StdRegProv")
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    CleanKey = -1
    Exit Function
  End If
  names = Empty
  types = Empty
  prov.EnumValues &H80000001, keyPath, names, types
  If Err.Number <> 0 Then
    ' Key absent on this machine: nothing to clean, not an error.
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  If Not IsArray(names) Then
    On Error GoTo 0
    Exit Function
  End If

  n = UBound(names)
  For i = 0 To n
    Err.Clear
    raw = sh.RegRead("HKEY_CURRENT_USER\" & keyPath & "\" & names(i))
    If Err.Number = 0 Then
      decoded = DecodePath(raw)
      filtered = StripZeros(raw)
      del = False
      For k = 0 To UBound(OUR_NAMES)
        If InStr(LCase(decoded), OUR_NAMES(k)) > 0 Then del = True
        If InStr(LCase(filtered), OUR_NAMES(k)) > 0 Then del = True
      Next
      If del Then
        sh.RegDelete "HKEY_CURRENT_USER\" & keyPath & "\" & names(i)
        If Err.Number = 0 Then
          removed = removed + 1
          WScript.Echo "  da go: " & Shorten(decoded)
        End If
        Err.Clear
      End If
    End If
    Err.Clear
    CleanKey = CleanKey + 1
  Next
  Err.Clear
  On Error GoTo 0
End Function

' The value is REG_BINARY. Office stores a small binary header followed by the
' path as UTF-16LE, so decoding the whole buffer as UTF-16LE and searching for
' the file name works regardless of header length.
Function DecodePath(v)
  If IsArray(v) Then
    DecodePath = Replace(DecodeUtf16(v), ChrW(0), "")
  Else
    DecodePath = CStr(v)
  End If
End Function

' Fallback for entries Office stores with its packed-unicode form: keep only the
' non-zero bytes, which leaves the readable ASCII path behind.
Function StripZeros(v)
  Dim s, i, b
  s = ""
  If IsArray(v) Then
    For i = 0 To UBound(v)
      b = v(i)
      If b < 0 Then b = b + 256
      If b <> 0 Then s = s & Chr(b)
    Next
  End If
  StripZeros = s
End Function

Function DecodeUtf16(v)
  Dim s, i, lo, hi
  s = ""
  For i = 0 To UBound(v) - 1 Step 2
    lo = v(i)
    hi = v(i + 1)
    If lo < 0 Then lo = lo + 256
    If hi < 0 Then hi = hi + 256
    s = s & ChrW(lo + hi * 256)
  Next
  DecodeUtf16 = s
End Function

' Console output only needs enough to recognise the file.
Function Shorten(s)
  Dim p
  p = InStrRev(s, "\")
  If p > 0 And Len(s) > p Then
    Shorten = Mid(s, p + 1)
  Else
    Shorten = s
  End If
End Function