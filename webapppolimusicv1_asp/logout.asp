<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<%
' ============================================================================
'  logout.asp — destroys the session and returns to the login screen
' ============================================================================
Dim loName

loName = CurrentUserName()
CloseDb
SignOut
Go "login.asp"
%>
