<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================================================
'  diagnostico.asp — comprobación del entorno (texto plano)
'  No usa includes ni funciones del proyecto: solo ASP puro, para aislar
'  cualquier problema del objeto Request / la conexión.
' ============================================================================
Response.ContentType = "text/plain"
Response.Charset = "utf-8"
Response.Expires = 0

' Solo accesible desde el propio servidor.
Dim remoteAddr
remoteAddr = CStr(Request.ServerVariables("REMOTE_ADDR"))
If remoteAddr <> "127.0.0.1" Then
    If remoteAddr <> "::1" Then
        Response.Status = "403 Forbidden"
        Response.Write "Restringido a localhost."
        Response.End
    End If
End If

Function Safe(ByVal x)
    Dim errNum
    On Error Resume Next
    Safe = ""
    Safe = CStr(x)
    errNum = Err.Number
    If errNum <> 0 Then
        Err.Clear
        Safe = "<ERROR " & errNum & ">"
    End If
    On Error GoTo 0
End Function

Function Coerce(ByVal x)
    Dim errNum
    On Error Resume Next
    Coerce = ""
    Coerce = x & ""
    errNum = Err.Number
    If errNum <> 0 Then
        Err.Clear
        Coerce = "<ERROR " & errNum & ">"
    End If
    On Error GoTo 0
End Function

Sub Dump(ByVal label, ByVal x)
    Response.Write Left(label & Space(38), 38) & " CStr=[" & Safe(x) & "]"
    Response.Write " coerce=[" & Coerce(x) & "]"
    Response.Write " type=" & TypeName(x)
    Response.Write " vt=" & VarType(x)
    Response.Write " obj=" & IsObject(x)
    Response.Write " empty=" & IsEmpty(x)
    Response.Write " null=" & IsNull(x)
    Response.Write vbCrLf
End Sub

Response.Write "PoliMusic :: diagnostico" & vbCrLf
Response.Write String(78, "-") & vbCrLf

Response.Write vbCrLf & "[1] Metodo y URL" & vbCrLf
Dump "ServerVariables('REQUEST_METHOD')", Request.ServerVariables("REQUEST_METHOD")
Dump "ServerVariables.Item('REQUEST_METHOD')", Request.ServerVariables.Item("REQUEST_METHOD")
Dump "Request('REQUEST_METHOD')", Request("REQUEST_METHOD")
Dump "ServerVariables('URL')", Request.ServerVariables("URL")
Dump "ServerVariables('QUERY_STRING')", Request.ServerVariables("QUERY_STRING")
Dump "ServerVariables('SCRIPT_NAME')", Request.ServerVariables("SCRIPT_NAME")
Dump "ServerVariables('REMOTE_ADDR')", Request.ServerVariables("REMOTE_ADDR")
Dump "ServerVariables('SERVER_NAME')", Request.ServerVariables("SERVER_NAME")
Dump "ServerVariables.Count", Request.ServerVariables.Count

Response.Write vbCrLf & "[2] QueryString (?x=hola)" & vbCrLf
Dump "QueryString('x')", Request.QueryString("x")
Dump "QueryString.Item('x')", Request.QueryString.Item("x")
Dump "QueryString.Count", Request.QueryString.Count

Response.Write vbCrLf & "[3] Form (POST y=adios)" & vbCrLf
Dump "Form('y')", Request.Form("y")
Dump "Form.Item('y')", Request.Form.Item("y")
Dump "Form.Count", Request.Form.Count
Dump "TotalBytes", Request.TotalBytes

Response.Write vbCrLf & "[4] Sesion" & vbCrLf
Dump "Session.SessionID", Session.SessionID
Dump "Session('probe')", Session("probe")
Session("probe") = "valor-guardado"
Dump "Session('probe') tras asignar", Session("probe")

Response.Write vbCrLf & "[5] Comparaciones como las del proyecto" & vbCrLf
Response.Write "UCase(SV REQUEST_METHOD) = POST  -> " & (UCase(Safe(Request.ServerVariables("REQUEST_METHOD"))) = "POST") & vbCrLf
Response.Write "Len(SV URL) > 0                  -> " & (Len(Safe(Request.ServerVariables("URL"))) > 0) & vbCrLf
Response.Write "IsObject(SV URL)                 -> " & IsObject(Request.ServerVariables("URL")) & vbCrLf
Response.Write "TypeName(SV URL)                 -> " & TypeName(Request.ServerVariables("URL")) & vbCrLf
Response.Write "VarType(SV URL)                  -> " & VarType(Request.ServerVariables("URL")) & vbCrLf

Response.Write vbCrLf & "[6] Base de datos" & vbCrLf
Dim cn, rs, prov, cs, msg, ok
ok = False
msg = ""
For Each prov In Array("MSOLEDBSQL", "SQLNCLI11", "SQLNCLI", "SQLOLEDB")
    If Not ok Then
        cs = "Provider=" & prov & ";Server=localhost,14333;Database=BDD_PoliMusic;" & _
             "User Id=usr_polimusic;Password=Politecnica1;Connect Timeout=8;"
        If prov = "MSOLEDBSQL" Or prov = "SQLNCLI11" Or prov = "SQLNCLI" Then
            cs = cs & "Encrypt=Yes;TrustServerCertificate=Yes;"
        End If
        On Error Resume Next
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open cs
        If Err.Number = 0 Then
            ok = True
            Response.Write "Proveedor OK: " & prov & vbCrLf
            Set rs = cn.Execute("SELECT COUNT(*) AS N FROM TBL_SONG")
            Response.Write "TBL_SONG filas: " & rs("N") & vbCrLf
            rs.Close
            cn.Close
        Else
            Response.Write "Proveedor " & prov & " -> ERROR " & Err.Number & ": " & Err.Description & vbCrLf
            Err.Clear
        End If
        On Error GoTo 0
    End If
Next

Response.Write vbCrLf & "[7] Rutas" & vbCrLf
Response.Write "MapPath('/') = " & Server.MapPath("/") & vbCrLf
Response.Write "MapPath('.') = " & Server.MapPath(".") & vbCrLf

Response.Write vbCrLf & "FIN" & vbCrLf
%>
