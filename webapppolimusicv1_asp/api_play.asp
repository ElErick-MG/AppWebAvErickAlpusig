<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<%
' ============================================================================
'  api_play.asp — increments TBL_SONG.PLAYS and answers with JSON
' ============================================================================
Dim apId, apPlays

Response.Buffer = True
Response.ContentType = "application/json"
Response.Charset = "utf-8"
Response.Expires = 0
Response.CacheControl = "no-cache"

If Not IsLoggedIn() Then
    Response.Status = "401 Unauthorized"
    Response.Write "{""ok"":false,""error"":""not authenticated""}"
    Response.End
End If

If Not IsPost() Then
    Response.Status = "405 Method Not Allowed"
    Response.Write "{""ok"":false,""error"":""POST required""}"
    Response.End
End If

If StrComp(FormStr("_csrf"), Txt(Session("pm_csrf")), 1) <> 0 Then
    Response.Status = "403 Forbidden"
    Response.Write "{""ok"":false,""error"":""invalid request token""}"
    Response.End
End If

apId = FormInt("id", 0)
If apId <= 0 Then
    Response.Status = "400 Bad Request"
    Response.Write "{""ok"":false,""error"":""invalid song id""}"
    Response.End
End If

On Error Resume Next
DbExec "UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 WHERE ID_SONG = ?", Array(apId)
If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Response.Status = "500 Internal Server Error"
    Response.Write "{""ok"":false,""error"":""update failed""}"
    Response.End
End If
On Error GoTo 0

apPlays = DbInt("SELECT ISNULL(PLAYS, 0) FROM TBL_SONG WHERE ID_SONG = ?", Array(apId), 0)

Response.Write "{""ok"":true,""id"":" & apId & ",""plays"":" & apPlays & "}"
CloseDb
%>
