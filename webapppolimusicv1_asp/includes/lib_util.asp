<%
' ============================================================================
'  PoliMusic — Utility helpers
'  Output encoding, asset resolution, formatting, flash messages, CSRF.
' ============================================================================

' ---- Safe text -------------------------------------------------------------
' Converts any Variant to a trimmed string.
' NOTE: some ASP collections (Session("x"), Request.QueryString("x"),
' Request.ServerVariables("x")) hand back values that IsObject() flags as
' objects, so the conversion is done through CStr / string coercion and any
' failure simply yields an empty string instead of discarding the value.
Function Txt(ByVal v)
    Dim s
    If IsNull(v) Or IsEmpty(v) Then
        Txt = ""
        Exit Function
    End If
    On Error Resume Next
    s = CStr(v)
    If Err.Number <> 0 Then
        Err.Clear
        s = v & ""
        If Err.Number <> 0 Then
            Err.Clear
            s = ""
        End If
    End If
    On Error GoTo 0
    Txt = Trim(s)
End Function

Function H(ByVal v)
    H = Server.HTMLEncode(Txt(v))
End Function

Function ToInt(ByVal v, ByVal fallback)
    Dim s
    s = Txt(v)
    If s = "" Then
        ToInt = fallback
    ElseIf IsNumeric(s) Then
        ToInt = CLng(s)
    Else
        ToInt = fallback
    End If
End Function

Function IsPost()
    IsPost = (UCase(Txt(Request.ServerVariables("REQUEST_METHOD"))) = "POST")
End Function

Function QsInt(ByVal name, ByVal fallback)
    QsInt = ToInt(Request.QueryString(name), fallback)
End Function

Function QsStr(ByVal name)
    QsStr = Txt(Request.QueryString(name))
End Function

Function FormStr(ByVal name)
    FormStr = Txt(Request.Form(name))
End Function

Function FormInt(ByVal name, ByVal fallback)
    FormInt = ToInt(Request.Form(name), fallback)
End Function

' ---- Assets ----------------------------------------------------------------
' Resolves DB paths such as "~/songFiles/x.mp3" into a browser URL that is
' relative to the application root, and returns it HTML-attribute safe.
Function AssetUrl(ByVal path, ByVal fallback)
    Dim s
    s = Replace(Txt(path), "\", "/")
    If s = "" Then s = Replace(Txt(fallback), "\", "/")
    If s = "" Then
        AssetUrl = ""
        Exit Function
    End If
    If InStr(s, "://") > 0 Then
        AssetUrl = H(s)
        Exit Function
    End If
    Do While Left(s, 1) = "/"
        s = Mid(s, 2)
    Loop
    If Left(s, 2) = "~/" Then s = Mid(s, 3)
    Do While InStr(s, "../") > 0
        s = Replace(s, "../", "")
    Loop
    AssetUrl = H(s)
End Function

' Physical check so the UI can flag records whose file is not deployed yet.
Function AssetExists(ByVal path)
    Dim rel, fso
    AssetExists = False
    rel = Replace(Txt(path), "\", "/")
    If rel = "" Then Exit Function
    If InStr(rel, "://") > 0 Then
        AssetExists = True
        Exit Function
    End If
    Do While Left(rel, 1) = "/"
        rel = Mid(rel, 2)
    Loop
    If Left(rel, 2) = "~/" Then rel = Mid(rel, 3)
    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    AssetExists = fso.FileExists(Server.MapPath(rel))
    Set fso = Nothing
    If Err.Number <> 0 Then
        Err.Clear
        AssetExists = True
    End If
    On Error GoTo 0
End Function

' ---- Formatting ------------------------------------------------------------
Function FmtDate(ByVal v)
    Dim s, d
    s = Txt(v)
    If s = "" Or Not IsDate(s) Then
        FmtDate = "&mdash;"
        Exit Function
    End If
    d = CDate(s)
    FmtDate = Right("0" & Day(d), 2) & "/" & Right("0" & Month(d), 2) & "/" & Year(d)
End Function

Function FmtDateTime(ByVal v)
    Dim s, d
    s = Txt(v)
    If s = "" Or Not IsDate(s) Then
        FmtDateTime = "&mdash;"
        Exit Function
    End If
    d = CDate(s)
    FmtDateTime = Right("0" & Day(d), 2) & "/" & Right("0" & Month(d), 2) & "/" & Year(d) & " " & _
                  Right("0" & Hour(d), 2) & ":" & Right("0" & Minute(d), 2)
End Function

Function FmtNum(ByVal v)
    Dim s
    s = Txt(v)
    If s = "" Or Not IsNumeric(s) Then
        FmtNum = "0"
    Else
        FmtNum = FormatNumber(CDbl(s), 0, -1, 0, -1)
    End If
End Function

Function Short(ByVal s, ByVal maxLen)
    Dim t
    t = Txt(s)
    If Len(t) <= maxLen Then
        Short = t
    Else
        Short = Left(t, maxLen - 1) & "..."
    End If
End Function

Function Initials(ByVal s)
    Dim parts, t, out
    t = Txt(s)
    If t = "" Then
        Initials = "?"
        Exit Function
    End If
    parts = Split(t, " ")
    out = UCase(Left(parts(0), 1))
    If UBound(parts) > 0 Then out = out & UCase(Left(parts(UBound(parts)), 1))
    Initials = out
End Function

' ---- Flash messages (survive one redirect) ---------------------------------
Sub Flash(ByVal kind, ByVal message)
    Session("pm_flash") = message
    Session("pm_flash_kind") = kind
End Sub

Sub FlashOk(ByVal message)
    Flash "ok", message
End Sub

Sub FlashErr(ByVal message)
    Flash "err", message
End Sub

Function RenderFlash()
    Dim msg, kind, icon, css
    msg = Txt(Session("pm_flash"))
    kind = Txt(Session("pm_flash_kind"))
    If msg = "" Then
        RenderFlash = ""
        Exit Function
    End If
    Session("pm_flash") = ""
    Session("pm_flash_kind") = ""
    If kind = "ok" Then
        css = "alert--ok"
        icon = "check"
    ElseIf kind = "info" Then
        css = "alert--info"
        icon = "info"
    Else
        css = "alert--err"
        icon = "alert"
    End If
    RenderFlash = "<div class=""alert " & css & """ role=""status"">" & Icon(icon) & "<div>" & H(msg) & "</div></div>"
End Function

' ---- CSRF ------------------------------------------------------------------
' Token = session id + random digits. Digits only, so it is safe inside both
' an HTML attribute and a form field, and never rendered in scientific notation.
Function Csrf()
    Dim t, i, d
    t = Txt(Session("pm_csrf"))
    If Len(t) < 16 Then
        Randomize
        t = ""
        For i = 1 To 24
            d = Int(Rnd() * 10)
            t = t & CStr(d)
        Next
        t = Txt(Session.SessionID) & t
        Session("pm_csrf") = t
    End If
    Csrf = t
End Function

Function CsrfField()
    CsrfField = "<input type=""hidden"" name=""_csrf"" value=""" & H(Csrf()) & """>"
End Function

Sub RequireCsrf()
    Dim sent
    sent = FormStr("_csrf")
    If Len(sent) = 0 Or StrComp(sent, Txt(Session("pm_csrf")), 1) <> 0 Then
        Response.Status = "403 Forbidden"
        Response.Write "<!DOCTYPE html><html lang=""en""><head><meta charset=""utf-8"">"
        Response.Write "<title>403 - PoliMusic</title></head>"
        Response.Write "<body style=""background:#0d0d0d;color:#f2f4f4;font-family:Segoe UI,sans-serif;padding:64px;line-height:1.6"">"
        Response.Write "<h1 style=""color:#c0392b;font-size:26px;margin:0 0 12px"">403 &mdash; Invalid request token</h1>"
        Response.Write "<p style=""color:#bdc3c7"">Your form session expired or the request did not come from PoliMusic.</p>"
        Response.Write "<p><a style=""color:#1abc9c"" href=""default.asp"">Return to PoliMusic</a></p></body></html>"
        Response.End
    End If
End Sub

' ---- Navigation ------------------------------------------------------------
Sub Go(ByVal url)
    On Error Resume Next
    CloseDb
    On Error GoTo 0
    Response.Redirect url
    Response.End
End Sub

Function CurrentUrl()
    Dim u, q
    u = Txt(Request.ServerVariables("URL"))
    q = Txt(Request.ServerVariables("QUERY_STRING"))
    If q <> "" Then u = u & "?" & q
    CurrentUrl = u
End Function

' ---- Form helpers ----------------------------------------------------------
' Renders a multi-line validation error string (lines separated by vbLf).
Function ErrorBox(ByVal text)
    Dim arr, i, out
    If Txt(text) = "" Then
        ErrorBox = ""
        Exit Function
    End If
    arr = Split(text, vbLf)
    out = ""
    For i = 0 To UBound(arr)
        If Trim(arr(i)) <> "" Then out = out & "<div>&bull; " & H(arr(i)) & "</div>"
    Next
    ErrorBox = "<div class=""alert alert--err"" role=""alert"">" & Icon("alert") & "<div>" & out & "</div></div>"
End Function

' ---- Validation ------------------------------------------------------------
Function IsValidUsername(ByVal s)
    Dim re
    Set re = New RegExp
    re.Pattern = "^[A-Za-z0-9._-]{3,50}$"
    IsValidUsername = re.Test(Txt(s))
End Function

Function IsValidEmail(ByVal s)
    Dim re
    Set re = New RegExp
    re.Pattern = "^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$"
    IsValidEmail = re.Test(Txt(s))
End Function

' Parses the ISO value produced by <input type="date"> without relying on the
' server locale.
Function ParseIsoDate(ByVal s, ByRef ok)
    Dim p, y, m, d, dt
    ok = False
    ParseIsoDate = Null
    s = Txt(s)
    If Len(s) <> 10 Then Exit Function
    p = Split(s, "-")
    If UBound(p) <> 2 Then Exit Function
    If Not (IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2))) Then Exit Function
    y = CInt(p(0))
    m = CInt(p(1))
    d = CInt(p(2))
    On Error Resume Next
    dt = DateSerial(y, m, d)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    If Year(dt) <> y Or Month(dt) <> m Or Day(dt) <> d Then Exit Function
    ParseIsoDate = dt
    ok = True
End Function

' Locale independent timestamp for SQL Server parameters (ISO 8601).
Function SqlDate(ByVal d)
    If IsNull(d) Or IsEmpty(d) Or Not IsDate(d) Then
        SqlDate = ""
        Exit Function
    End If
    SqlDate = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2) & "T" & _
              Right("0" & Hour(d), 2) & ":" & Right("0" & Minute(d), 2) & ":" & Right("0" & Second(d), 2)
End Function

Function IsoDate(ByVal v)
    Dim s, d
    s = Txt(v)
    If s = "" Or Not IsDate(s) Then
        IsoDate = ""
    Else
        d = CDate(s)
        IsoDate = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2)
    End If
End Function

' Inline conditional (VBScript has no native IIf).
Function IIf(ByVal condition, ByVal whenTrue, ByVal whenFalse)
    If condition Then
        IIf = whenTrue
    Else
        IIf = whenFalse
    End If
End Function

Function SelIf(ByVal value, ByVal compare)
    If CStr(value) = CStr(compare) Then
        SelIf = " selected"
    Else
        SelIf = ""
    End If
End Function

Function ChkIf(ByVal flag)
    If flag Then
        ChkIf = " checked"
    Else
        ChkIf = ""
    End If
End Function
%>
