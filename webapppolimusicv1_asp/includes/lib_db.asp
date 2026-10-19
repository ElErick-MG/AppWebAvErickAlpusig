<%
' ============================================================================
'  PoliMusic — Data access layer (ADO / SQL Server 14333)
'  Every query is parameterised: no string concatenation with user input.
' ============================================================================

' ADO constants (declared locally so we do not depend on typelib includes)
Const adCmdText       = 1
Const adOpenStatic    = 3
Const adLockReadOnly  = 1
Const adUseClient     = 3
Const adInteger       = 3
Const adDouble        = 5
Const adBoolean       = 11
Const adVarWChar      = 202
Const adExecuteNoRecords = 128

Dim gConn, gConnOpen

' ---- Connection string -----------------------------------------------------
Function BuildCs(ByVal provider)
    Dim s, p
    p = LCase(provider)
    s = "Provider=" & provider & ";" & _
        "Server=" & DB_SERVER & ";" & _
        "Database=" & DB_NAME & ";" & _
        "User Id=" & DB_USER & ";" & _
        "Password=" & DB_PASSWORD & ";"
    If p = "msoledbsql" Or p = "sqlncli11" Or p = "sqlncli" Then
        s = s & "Encrypt=" & DB_ENCRYPT & ";" & _
                "TrustServerCertificate=" & DB_TRUST_CERT & ";"
    End If
    s = s & "Connect Timeout=" & DB_TIMEOUT & ";"
    BuildCs = s
End Function

Function TryOpen(ByVal provider)
    Dim c
    Set c = Server.CreateObject("ADODB.Connection")
    On Error Resume Next
    c.Open BuildCs(provider)
    If Err.Number <> 0 Then
        Application("PM_DB_ERROR") = "[" & provider & "] " & Err.Description
        Err.Clear
        Set c = Nothing
        On Error GoTo 0
        Set TryOpen = Nothing
        Exit Function
    End If
    On Error GoTo 0
    Set TryOpen = c
End Function

' Tries the modern OLE DB driver first, then falls back to legacy providers.
Function OpenConn()
    Dim cached, provs, i, c
    cached = Txt(Application("PM_PROVIDER"))
    If cached <> "" Then
        Set c = TryOpen(cached)
        If Not c Is Nothing Then
            Set OpenConn = c
            Exit Function
        End If
        Application("PM_PROVIDER") = ""
    End If

    provs = Array("MSOLEDBSQL", "SQLNCLI11", "SQLNCLI", "SQLOLEDB")
    For i = 0 To UBound(provs)
        Set c = TryOpen(provs(i))
        If Not c Is Nothing Then
            Application("PM_PROVIDER") = provs(i)
            Set OpenConn = c
            Exit Function
        End If
    Next

    Set OpenConn = Nothing
End Function

Function Conn()
    If Not gConnOpen Then
        Set gConn = OpenConn()
        gConnOpen = True
    ElseIf gConn.State = 0 Then
        Set gConn = OpenConn()
    End If
    If gConn Is Nothing Then DbFatal
    Set Conn = gConn
End Function

Sub CloseDb()
    If gConnOpen Then
        On Error Resume Next
        gConn.Close
        Set gConn = Nothing
        Err.Clear
        On Error GoTo 0
        gConnOpen = False
    End If
End Sub

' Friendly full-page error when the database cannot be reached.
Sub DbFatal()
    Dim msg
    msg = Txt(Application("PM_DB_ERROR"))
    Response.Status = "503 Service Unavailable"
    Response.Write "<!DOCTYPE html><html lang=""en""><head><meta charset=""utf-8"">"
    Response.Write "<meta name=""viewport"" content=""width=device-width,initial-scale=1"">"
    Response.Write "<title>Database unavailable - PoliMusic</title></head>"
    Response.Write "<body style=""margin:0;background:#0d0d0d;color:#f2f4f4;font-family:Segoe UI,system-ui,sans-serif;display:grid;place-items:center;min-height:100vh;padding:32px"">"
    Response.Write "<div style=""max-width:620px;border:1px solid #262626;border-radius:18px;padding:32px;background:linear-gradient(160deg,#161616,#101010)"">"
    Response.Write "<div style=""font-size:11px;letter-spacing:.22em;text-transform:uppercase;color:#c0392b;font-weight:700"">PoliMusic &middot; connection error</div>"
    Response.Write "<h1 style=""font-size:26px;margin:10px 0 12px"">Cannot reach SQL Server</h1>"
    Response.Write "<p style=""color:#bdc3c7;line-height:1.6;margin:0 0 16px"">The application could not open a connection to <b>" & H(DB_SERVER) & "</b> / <b>" & H(DB_NAME) & "</b>.</p>"
    Response.Write "<ul style=""color:#8a9296;font-size:13px;line-height:1.8;padding-left:18px;margin:0 0 18px"">"
    Response.Write "<li>Verify the SQL Server instance is listening on port 14333.</li>"
    Response.Write "<li>Verify the login <b>usr_polimusic</b> exists and the database is online.</li>"
    Response.Write "<li>Check that the OLE DB driver (MSOLEDBSQL) is installed on the web server.</li>"
    Response.Write "</ul>"
    If msg <> "" Then
        Response.Write "<pre style=""white-space:pre-wrap;background:#0a0a0a;border:1px solid #262626;border-radius:10px;padding:12px;color:#e8837a;font-size:12px;margin:0"">" & H(msg) & "</pre>"
    End If
    Response.Write "</div></body></html>"
    Response.End
End Sub

' ---- Parameter binding -----------------------------------------------------
Function ParamType(ByVal v)
    Dim vt
    If IsNull(v) Or IsEmpty(v) Then
        ParamType = adVarWChar
        Exit Function
    End If
    vt = VarType(v)
    If vt = 2 Or vt = 3 Then
        ParamType = adInteger
    ElseIf vt = 4 Or vt = 5 Then
        ParamType = adDouble
    ElseIf vt = 11 Then
        ParamType = adBoolean
    Else
        ParamType = adVarWChar
    End If
End Function

Function BuildCmd(ByVal sql, ByVal params)
    Dim cmd, i, v, t, p
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = Conn()
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    cmd.CommandTimeout = 30

    If IsArray(params) Then
        For i = 0 To UBound(params)
            v = params(i)
            t = ParamType(v)
            If t = adVarWChar Then
                Set p = cmd.CreateParameter("@p" & i, adVarWChar, 1, 4000, v)
            Else
                Set p = cmd.CreateParameter("@p" & i, t, 1, , v)
            End If
            cmd.Parameters.Append p
        Next
    End If
    Set BuildCmd = cmd
End Function

' ---- Query helpers ---------------------------------------------------------
' Returns a client-side, read-only recordset (supports .RecordCount).
' The ActiveConnection is NOT passed to Open(): the Command already carries it,
' and ADO rejects the call when both are supplied (error 800a0bb9 / 1240656).
Function DbQuery(ByVal sql, ByVal params)
    Dim rs, cmd
    Set cmd = BuildCmd(sql, params)
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly
    rs.Open cmd
    Set DbQuery = rs
End Function

' Executes INSERT/UPDATE/DELETE and returns the affected row count (-1 on error).
Function DbExec(ByVal sql, ByVal params)
    Dim cmd, ra
    Set cmd = BuildCmd(sql, params)
    cmd.Execute ra, , adExecuteNoRecords
    If IsNull(ra) Then ra = 0
    DbExec = CLng(ra)
End Function

' Same as DbExec but traps database errors so pages can show a friendly message.
Function DbExecSafe(ByVal sql, ByVal params, ByRef errMsg)
    Dim cmd, ra
    errMsg = ""
    On Error Resume Next
    Set cmd = BuildCmd(sql, params)
    If Err.Number <> 0 Then
        errMsg = Err.Description
        Err.Clear
        On Error GoTo 0
        DbExecSafe = -1
        Exit Function
    End If
    cmd.Execute ra, , adExecuteNoRecords
    If Err.Number <> 0 Then
        errMsg = Err.Description
        Err.Clear
        On Error GoTo 0
        DbExecSafe = -1
        Exit Function
    End If
    On Error GoTo 0
    If IsNull(ra) Then ra = 0
    DbExecSafe = CLng(ra)
End Function

Function DbScalar(ByVal sql, ByVal params)
    Dim rs
    Set rs = DbQuery(sql, params)
    If rs.EOF Then
        DbScalar = Null
    Else
        DbScalar = rs.Fields(0).Value
    End If
    rs.Close
    Set rs = Nothing
End Function

' Returns a detached 2D array-like snapshot is unnecessary here: pages iterate
' recordsets directly and close them, which is enough for this application.
Function DbInt(ByVal sql, ByVal params, ByVal fallback)
    Dim v
    v = DbScalar(sql, params)
    If IsNull(v) Then
        DbInt = fallback
    Else
        DbInt = CLng(v)
    End If
End Function
%>
