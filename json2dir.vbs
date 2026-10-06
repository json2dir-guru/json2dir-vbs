' json2dir: create the directory tree a JSON document describes, in the current directory.
' Implements RFC J2D-1 (https://github.com/kitsunoff/awesome-json2dir/blob/main/spec/rfc-json2dir.md)
' in VBScript for Windows Script Host:  cscript //nologo json2dir.vbs < document.json
'
' Written for Wine's cscript (see README). Wine's WSH has no WScript.StdIn/StdErr, so the script
' reads its Unix stdin/stderr via Z:\proc\<WINE_UNIX_PID>\fd\N; Wine's VBScript has no ChrW, AscW
' or Join, so the document is handled as a "byte string": text streams in ANSI mode map every byte
' to one code page 1252 character (the launcher pins the locale), Asc() gives the byte back and
' Chr() makes it. Strings from the document stay in that form (file contents are written as is);
' only names are turned into real Unicode, by writing them as UTF-16LE to a temporary file and
' reading that back in Unicode mode.
'
' Safety: Wine's FileSystemObject cannot tell a symbolic link from what it points to, and deleting
' a link through it deletes the target. So this program never deletes, overwrites or reuses
' anything: an existing entry with a name it needs is an error.

Option Explicit

Dim fso, unixPid
Set fso = CreateObject("Scripting.FileSystemObject")
unixPid = ""
On Error Resume Next
unixPid = CreateObject("WScript.Shell").Environment("PROCESS")("WINE_UNIX_PID")
On Error GoTo 0

' ---------- standard streams ----------

Sub WriteStderr(text)
  Dim t
  On Error Resume Next
  If unixPid <> "" Then
    Set t = fso.OpenTextFile("Z:\proc\" & unixPid & "\fd\2", 2, False, 0)
    t.Write text
    t.Close
  Else
    WScript.StdErr.Write text
  End If
End Sub

Sub Fail(message)
  WriteStderr "json2dir: " & message & vbLf
  WScript.Quit 1
End Sub

Function ReadStdin()
  Dim t, s, msg
  s = ""
  On Error Resume Next
  If unixPid <> "" Then
    Set t = fso.OpenTextFile("Z:\proc\" & unixPid & "\fd\0", 1, False, 0)
    If Err.Number = 0 Then
      If Not t.AtEndOfStream Then s = t.ReadAll()
      t.Close
    End If
  Else
    If Not WScript.StdIn.AtEndOfStream Then s = WScript.StdIn.ReadAll()
  End If
  If Err.Number <> 0 Then
    msg = Err.Description & " (0x" & Hex(Err.Number) & ")"
    On Error GoTo 0
    Fail "cannot read standard input: " & msg
  End If
  ReadStdin = s
End Function

' ---------- helpers ----------

Function ByteAt(s, i) ' byte value of character i (1-based) of a byte string
  ByteAt = Asc(Mid(s, i, 1))
End Function

' Concatenates parts(0..n-1) pairwise (Wine's VBScript has no Join), O(total * log n).
Function JoinParts(parts, ByVal n)
  Dim m, i
  If n = 0 Then
    JoinParts = ""
    Exit Function
  End If
  Do While n > 1
    m = 0
    For i = 0 To n - 1 Step 2
      If i + 1 < n Then parts(m) = parts(i) & parts(i + 1) Else parts(m) = parts(i)
      m = m + 1
    Next
    n = m
  Loop
  JoinParts = parts(0)
End Function

Sub AddPart(parts, cnt, s)
  If cnt > UBound(parts) Then ReDim Preserve parts(2 * cnt + 1)
  parts(cnt) = s
  cnt = cnt + 1
End Sub

' Code point -> its UTF-8 bytes as a byte string. Surrogates come out as 3 bytes ED A0..BF xx,
' which strict UTF-8 input can never contain, so lone surrogates from \u escapes stay detectable.
Function Utf8Of(cp)
  If cp < &H80 Then
    Utf8Of = Chr(cp)
  ElseIf cp < &H800 Then
    Utf8Of = Chr(&HC0 Or (cp \ 64)) & Chr(&H80 Or (cp And 63))
  ElseIf cp < 65536 Then
    Utf8Of = Chr(&HE0 Or (cp \ 4096)) & Chr(&H80 Or ((cp \ 64) And 63)) & Chr(&H80 Or (cp And 63))
  Else
    Utf8Of = Chr(&HF0 Or (cp \ 262144)) & Chr(&H80 Or ((cp \ 4096) And 63)) & _
      Chr(&H80 Or ((cp \ 64) And 63)) & Chr(&H80 Or (cp And 63))
  End If
End Function

' ---------- UTF-8 (§3) ----------

Sub CheckUtf8(s)
  Dim i, n, b0, bk, cp, need, min, k
  n = Len(s) : i = 1
  Do While i <= n
    b0 = ByteAt(s, i)
    If b0 < &H80 Then
      i = i + 1
    Else
      If b0 >= &HC2 And b0 <= &HDF Then
        cp = b0 And &H1F : need = 1 : min = &H80
      ElseIf b0 >= &HE0 And b0 <= &HEF Then
        cp = b0 And &HF : need = 2 : min = &H800
      ElseIf b0 >= &HF0 And b0 <= &HF4 Then
        cp = b0 And &H7 : need = 3 : min = 65536
      Else
        Fail "input is not valid UTF-8"
      End If
      If i + need > n Then Fail "input is not valid UTF-8"
      For k = 1 To need
        bk = ByteAt(s, i + k)
        If (bk And &HC0) <> &H80 Then Fail "input is not valid UTF-8"
        cp = cp * 64 + (bk And &H3F)
      Next
      If cp < min Or cp > 1114111 Or (cp >= 55296 And cp <= 57343) Then Fail "input is not valid UTF-8"
      i = i + need + 1
    End If
  Loop
End Sub

' ---------- JSON (RFC 8259, strict) ----------
' Every value is a two-element array: Array("o", Dictionary name->value), Array("a", items),
' Array("s", byte string), or Array("x", "number"/"true"/"false"/"null") for values §4.5 rejects.

Dim text, pos, textLen

Sub JErr(what)
  Fail "input is not valid JSON: " & what & " at offset " & (pos - 1)
End Sub

Function Ch() ' current character, "" at the end
  If pos <= textLen Then Ch = Mid(text, pos, 1) Else Ch = ""
End Function

Sub SkipWs()
  Dim c
  Do
    c = Ch()
    If c = " " Or c = vbTab Or c = vbLf Or c = vbCr Then pos = pos + 1 Else Exit Do
  Loop
End Sub

Function IsDigit(c)
  If c = "" Then IsDigit = False Else IsDigit = (InStr("0123456789", c) > 0)
End Function

Function ParseValue()
  Dim c
  SkipWs
  c = Ch()
  If c = "{" Then
    ParseValue = ParseObject()
  ElseIf c = "[" Then
    ParseValue = ParseArray()
  ElseIf c = """" Then
    ParseValue = Array("s", ParseString())
  ElseIf c = "t" Then
    ParseValue = ParseLiteral("true")
  ElseIf c = "f" Then
    ParseValue = ParseLiteral("false")
  ElseIf c = "n" Then
    ParseValue = ParseLiteral("null")
  ElseIf c = "-" Or IsDigit(c) Then
    ParseValue = ParseNumber()
  ElseIf c = "" Then
    JErr "unexpected end"
  Else
    JErr "unexpected character"
  End If
End Function

Function ParseLiteral(word)
  If Mid(text, pos, Len(word)) <> word Then JErr "bad literal"
  pos = pos + Len(word)
  ParseLiteral = Array("x", word)
End Function

Function ParseNumber()
  If Ch() = "-" Then pos = pos + 1
  If Ch() = "0" Then
    pos = pos + 1
  ElseIf IsDigit(Ch()) Then
    Do While IsDigit(Ch()) : pos = pos + 1 : Loop
  Else
    JErr "bad number"
  End If
  If Ch() = "." Then
    pos = pos + 1
    If Not IsDigit(Ch()) Then JErr "bad number"
    Do While IsDigit(Ch()) : pos = pos + 1 : Loop
  End If
  If Ch() = "e" Or Ch() = "E" Then
    pos = pos + 1
    If Ch() = "+" Or Ch() = "-" Then pos = pos + 1
    If Not IsDigit(Ch()) Then JErr "bad number"
    Do While IsDigit(Ch()) : pos = pos + 1 : Loop
  End If
  ParseNumber = Array("x", "number")
End Function

' Reads the 4 hex digits of a \u escape at pos (just after "\u"), returns the code unit.
Function Hex4()
  Dim h, i
  h = Mid(text, pos, 4)
  If Len(h) <> 4 Then JErr "bad \u escape"
  For i = 1 To 4
    If InStr("0123456789abcdefABCDEF", Mid(h, i, 1)) = 0 Then JErr "bad \u escape"
  Next
  Hex4 = CLng("&H" & h) And 65535
  pos = pos + 4
End Function

Function ParseString()
  Dim parts(), cnt, start, code, e, u, u2
  ReDim parts(15)
  cnt = 0
  pos = pos + 1 ' opening quote
  start = pos
  Do
    If pos > textLen Then JErr "unterminated string"
    code = ByteAt(text, pos)
    If code = 34 Then
      AddPart parts, cnt, Mid(text, start, pos - start)
      pos = pos + 1
      ParseString = JoinParts(parts, cnt)
      Exit Function
    End If
    If code < 32 Then JErr "control character in string"
    If code = 92 Then
      AddPart parts, cnt, Mid(text, start, pos - start)
      e = Mid(text, pos + 1, 1)
      pos = pos + 2
      Select Case e
        Case """" : AddPart parts, cnt, """"
        Case "\" : AddPart parts, cnt, "\"
        Case "/" : AddPart parts, cnt, "/"
        Case "b" : AddPart parts, cnt, Chr(8)
        Case "f" : AddPart parts, cnt, Chr(12)
        Case "n" : AddPart parts, cnt, vbLf
        Case "r" : AddPart parts, cnt, vbCr
        Case "t" : AddPart parts, cnt, vbTab
        Case "u"
          u = Hex4()
          If u >= 55296 And u <= 56319 And Mid(text, pos, 2) = "\u" Then
            pos = pos + 2
            u2 = Hex4()
            If u2 >= 56320 And u2 <= 57343 Then
              AddPart parts, cnt, Utf8Of(65536 + (u - 55296) * 1024 + (u2 - 56320))
            Else
              AddPart parts, cnt, Utf8Of(u) & Utf8Of(u2)
            End If
          Else
            AddPart parts, cnt, Utf8Of(u)
          End If
        Case Else
          pos = pos - 1
          JErr "bad escape"
      End Select
      start = pos
    Else
      pos = pos + 1
    End If
  Loop
End Function

Function ParseArray()
  Dim items(), cnt, c
  ReDim items(3)
  cnt = 0
  pos = pos + 1
  SkipWs
  If Ch() = "]" Then
    pos = pos + 1
    ParseArray = Array("a", Array())
    Exit Function
  End If
  Do
    If cnt > UBound(items) Then ReDim Preserve items(2 * cnt + 1)
    items(cnt) = ParseValue()
    cnt = cnt + 1
    SkipWs
    c = Ch()
    If c = "]" Then
      pos = pos + 1
      ReDim Preserve items(cnt - 1)
      ParseArray = Array("a", items)
      Exit Function
    End If
    If c <> "," Then JErr "expected , or ]"
    pos = pos + 1
  Loop
End Function

Function ParseObject()
  Dim d, name, v
  Set d = CreateObject("Scripting.Dictionary") ' binary compare, keeps insertion order
  pos = pos + 1
  SkipWs
  If Ch() = "}" Then
    pos = pos + 1
    ParseObject = Array("o", d)
    Exit Function
  End If
  Do
    SkipWs
    If Ch() <> """" Then JErr "expected a member name"
    name = ParseString()
    SkipWs
    If Ch() <> ":" Then JErr "expected :"
    pos = pos + 1
    v = ParseValue()
    d(name) = v ' duplicate names: the last occurrence wins (§4.2), keeping the first position
    SkipWs
    If Ch() = "}" Then
      pos = pos + 1
      ParseObject = Array("o", d)
      Exit Function
    End If
    If Ch() <> "," Then JErr "expected , or }"
    pos = pos + 1
  Loop
End Function

Function ParseJson(s)
  text = s : pos = 1 : textLen = Len(s)
  ParseJson = ParseValue()
  SkipWs
  If pos <= textLen Then JErr "trailing data"
End Function

' ---------- validation (§4, before touching the file system) ----------

Function HasLoneSurrogate(s) ' a surrogate code point in UTF-8 form: ED A0..BF xx
  Dim i
  HasLoneSurrogate = False
  i = InStr(s, Chr(&HED))
  Do While i > 0 And i < Len(s)
    If ByteAt(s, i + 1) >= &HA0 Then
      HasLoneSurrogate = True
      Exit Function
    End If
    i = InStr(i + 1, s, Chr(&HED))
  Loop
End Function

Sub CheckString(s, where)
  If HasLoneSurrogate(s) Then Fail where & ": string contains an unpaired surrogate"
End Sub

Function IsAscii(s)
  Dim i
  IsAscii = False
  For i = 1 To Len(s)
    If ByteAt(s, i) >= &H80 Then Exit Function
  Next
  IsAscii = True
End Function

Dim wideNames ' byte-string name -> Unicode name, for the names that are not plain ASCII
Set wideNames = CreateObject("Scripting.Dictionary")

Sub Validate(v, where)
  Dim items, name, child, i, members
  Select Case v(0)
    Case "s"
      CheckString v(1), where
    Case "a"
      items = v(1)
      If UBound(items) <> 1 Then Fail where & ": an array must be [""link"", target] or [""script"", content]"
      If items(0)(0) <> "s" Or items(1)(0) <> "s" Then Fail where & ": an array must be [""link"", target] or [""script"", content]"
      ' §10: WSH can neither create symbolic links nor set execute bits, so reject, never approximate.
      If items(0)(1) = "link" Then
        Fail where & ": symbolic links cannot be created from VBScript/WSH (RFC section 10)"
      ElseIf items(0)(1) = "script" Then
        Fail where & ": executable files cannot be created from VBScript/WSH (RFC section 10)"
      Else
        Fail where & ": unknown array kind """ & items(0)(1) & """"
      End If
    Case "o"
      Set members = v(1)
      For Each name In members.Keys
        If where = "." Then child = name Else child = where & "/" & name
        CheckString name, child
        If name = "" Or name = "." Or name = ".." Or InStr(name, "/") > 0 Or InStr(name, Chr(0)) > 0 Then
          Fail child & ": invalid name """ & name & """"
        End If
        ' Win32 paths treat these as separators or wildcards; they cannot be used safely in a name.
        For i = 1 To Len(name)
          If InStr("\*?""<>|", Mid(name, i, 1)) > 0 Then Fail child & ": the name cannot be represented through Win32 file APIs"
        Next
        If Not IsAscii(name) Then wideNames(name) = ""
        Validate members(name), child
      Next
    Case Else
      Fail where & ": " & v(1) & " values are not allowed"
  End Select
End Sub

' UTF-8 byte string (valid, no surrogates) -> UTF-16LE as a byte string.
Function Utf16leOf(s)
  Dim parts(), cnt, i, n, b0, cp, need, k
  ReDim parts(Len(s))
  cnt = 0 : i = 1 : n = Len(s)
  Do While i <= n
    b0 = ByteAt(s, i)
    If b0 < &H80 Then
      cp = b0 : need = 0
    ElseIf b0 < &HE0 Then
      cp = b0 And &H1F : need = 1
    ElseIf b0 < &HF0 Then
      cp = b0 And &HF : need = 2
    Else
      cp = b0 And &H7 : need = 3
    End If
    For k = 1 To need
      cp = cp * 64 + (ByteAt(s, i + k) And &H3F)
    Next
    i = i + need + 1
    If cp >= 65536 Then
      cp = cp - 65536
      parts(cnt) = Unit16(55296 + (cp \ 1024)) & Unit16(56320 + (cp And 1023))
    Else
      parts(cnt) = Unit16(cp)
    End If
    cnt = cnt + 1
  Loop
  Utf16leOf = JoinParts(parts, cnt)
End Function

Function Unit16(u)
  Unit16 = Chr(u And 255) & Chr(u \ 256)
End Function

' Fills wideNames: all names go to a temporary file as UTF-16LE, separated by "/" (which no name
' contains), and come back through a Unicode-mode text stream as real strings.
Sub WidenNames()
  Dim keys, parts(), i, n, tmp, t, s, back, msg
  keys = wideNames.Keys
  n = UBound(keys) + 1
  If n = 0 Then Exit Sub
  ReDim parts(n - 1)
  For i = 0 To n - 1
    parts(i) = Utf16leOf(keys(i))
    If i < n - 1 Then parts(i) = parts(i) & Unit16(47)
  Next
  s = JoinParts(parts, n)
  On Error Resume Next
  tmp = fso.BuildPath(fso.GetSpecialFolder(2).Path, fso.GetTempName())
  Set t = fso.CreateTextFile(tmp, False, False)
  t.Write s
  t.Close
  Set t = fso.OpenTextFile(tmp, 1, False, -1)
  back = t.ReadAll()
  t.Close
  fso.DeleteFile tmp
  If Err.Number <> 0 Then
    msg = Err.Description & " (0x" & Hex(Err.Number) & ")"
    On Error GoTo 0
    Fail "cannot convert names to Unicode: " & msg
  End If
  On Error GoTo 0
  back = Split(back, "/")
  If UBound(back) <> n - 1 Then Fail "cannot convert names to Unicode"
  For i = 0 To n - 1
    wideNames(keys(i)) = back(i)
  Next
End Sub

Function Wide(name)
  If wideNames.Exists(name) Then Wide = wideNames(name) Else Wide = name
End Function

' Wine's file lookup is case-insensitive: names differing only in case would collide.
Sub CheckCase(v, where)
  Dim seen, name, child, folded, members
  If v(0) <> "o" Then Exit Sub
  Set seen = CreateObject("Scripting.Dictionary")
  Set members = v(1)
  For Each name In members.Keys
    If where = "." Then child = name Else child = where & "/" & name
    folded = LCase(Wide(name))
    If seen.Exists(folded) Then Fail child & ": names differing only in case are not supported here"
    seen(folded) = True
    CheckCase members(name), child
  Next
End Sub

' ---------- creating the tree (§4, §5) ----------

Function CompareBytes(a, b) ' byte strings, byte order (= code point order)
  Dim i, n, x, y
  n = Len(a)
  If Len(b) < n Then n = Len(b)
  For i = 1 To n
    x = ByteAt(a, i) : y = ByteAt(b, i)
    If x <> y Then
      If x < y Then CompareBytes = -1 Else CompareBytes = 1
      Exit Function
    End If
  Next
  CompareBytes = Len(a) - Len(b)
End Function

Function SortedNames(d)
  Dim k, n, i, j, tmp
  k = d.Keys
  n = UBound(k) + 1
  For i = 1 To n - 1 ' insertion sort
    tmp = k(i)
    j = i - 1
    Do While j >= 0
      If CompareBytes(k(j), tmp) <= 0 Then Exit Do
      k(j + 1) = k(j)
      j = j - 1
    Loop
    k(j + 1) = tmp
  Next
  SortedNames = k
End Function

Sub FsFail(where, what)
  Dim msg : msg = Err.Description & " (0x" & Hex(Err.Number) & ")"
  Err.Clear
  Fail where & ": " & what & ": " & msg
End Sub

Sub WriteFile(path, data, where)
  Dim t
  On Error Resume Next
  Set t = fso.CreateTextFile(path, False, False) ' never overwrites (also not a dangling symlink)
  If Err.Number <> 0 Then FsFail where, "cannot create file"
  If Len(data) > 0 Then t.Write data
  If Err.Number <> 0 Then FsFail where, "cannot write file"
  t.Close
  If Err.Number <> 0 Then FsFail where, "cannot write file"
End Sub

Sub MakeFolder(path, where)
  On Error Resume Next
  fso.CreateFolder path
  If Err.Number <> 0 Then FsFail where, "cannot create directory"
End Sub

Function Exists(path)
  Exists = fso.FileExists(path) Or fso.FolderExists(path)
End Function

Sub Apply(dir, rel, d)
  Dim names, name, value, path, where, i
  names = SortedNames(d)
  For i = 0 To UBound(names)
    name = names(i)
    value = d(name)
    path = dir & "\" & Wide(name)
    If rel = "" Then where = name Else where = rel & "/" & name
    ' Never delete, overwrite or reuse: through Wine an existing entry may be a symbolic link that
    ' the FileSystemObject cannot recognise, and acting on it would act on its target.
    If Exists(path) Then Fail where & ": already exists (existing entries are never replaced or reused here)"
    If value(0) = "s" Then
      WriteFile path, value(1), where
    Else
      MakeFolder path, where
      Apply path, where, value(1)
    End If
  Next
End Sub

Sub Main()
  Dim s, doc
  If WScript.Arguments.Count > 0 Then
    WriteStderr "usage: json2dir < document.json" & vbLf
    WScript.Quit 2
  End If
  s = ReadStdin()
  CheckUtf8 s
  If Len(s) >= 3 Then
    If ByteAt(s, 1) = &HEF And ByteAt(s, 2) = &HBB And ByteAt(s, 3) = &HBF Then s = Mid(s, 4) ' BOM may be ignored
  End If
  doc = ParseJson(s)
  If doc(0) <> "o" Then Fail "the root of the document must be an object"
  Validate doc, "."
  WidenNames
  CheckCase doc, "."
  ' "\\?\" turns off Win32 path rewriting, so names like "a." or "con" are used as given.
  Apply "\\?\" & fso.GetAbsolutePathName("."), "", doc(1)
  WScript.Quit 0
End Sub

' Wine's cscript reports a script error nowhere and exits with status 0, so catch it here.
On Error Resume Next
Main
Fail "internal error: " & Err.Description & " (0x" & Hex(Err.Number) & ")"
