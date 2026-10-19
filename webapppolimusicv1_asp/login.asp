<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  login.asp — authenticates against TBL_USER
'  Accepts legacy plain-text passwords and SHA1 hex hashes (HASHBYTES).
' ============================================================================
Dim lgUser, lgPass, lgErr, lgNext, lgRs

PageTitle = "Sign in"
AuthTitle = "Welcome back"
AuthSub = "Sign in to reach your songs, albums and playlists."
lgErr = ""
lgNext = ""

If IsLoggedIn() Then Go "default.asp"

If IsPost() Then
    RequireCsrf
    lgUser = FormStr("username")
    lgPass = Request.Form("password") & ""
    lgNext = FormStr("next")

    If lgUser = "" Or lgPass = "" Then
        lgErr = "Type your username and your password to continue."
    ElseIf Len(lgUser) > 50 Then
        lgErr = "That username is longer than the field allows."
    Else
        Set lgRs = FindUserByCredentials(lgUser, lgPass)
        If lgRs Is Nothing Then
            lgErr = "Invalid credentials. Check your username and password and try again."
        Else
            SignIn lgRs("ID_USER"), lgRs("USERNAME"), lgRs("USER_TYPE"), lgRs("USER_PHOTO"), lgRs("EMAIL")
            lgRs.Close
            Set lgRs = Nothing
            If Len(lgNext) > 0 And InStr(lgNext, "://") = 0 And Left(lgNext, 2) <> "//" Then
                Go lgNext
            Else
                Go "default.asp"
            End If
        End If
    End If
End If
%>
<!--#include file="includes/auth_top.asp"-->

      <% If lgErr <> "" Then %>
      <div class="alert alert--err" role="alert">
        <%= Icon("alert") %><div><%= H(lgErr) %></div>
      </div>
      <% End If %>

      <%= RenderFlash() %>

      <form class="auth__form" method="post" action="login.asp" autocomplete="on" novalidate>
        <%= CsrfField() %>
        <input type="hidden" name="next" value="<%= H(QsStr("next")) %>">

        <div class="field">
          <label for="username">Username</label>
          <input id="username" name="username" type="text" value="<%= H(lgUser) %>"
                 placeholder="e.g. Erick.Alpusig" maxlength="50" required autofocus
                 autocomplete="username" spellcheck="false">
        </div>

        <div class="field">
          <label for="password">Password</label>
          <input id="password" name="password" type="password" placeholder="Your password"
                 maxlength="50" required autocomplete="current-password">
        </div>

        <div class="auth__row">
          <span><%= Icon("lock") %> Sessions expire after 60 minutes of inactivity.</span>
        </div>

        <button class="btn btn--primary btn--block" type="submit">
          <%= Icon("arrow-right") %> Sign in
        </button>
      </form>

      <div class="auth__hint">
        <b>Demo accounts</b><br>
        Administrator &mdash; <b>userAdmin</b> / <b>12345678</b><br>
        Listener &mdash; <b>Erick.Alpusig</b> / <b>12345678</b>
      </div>

      <p class="auth__switch">
        New here? <a href="register.asp">Create a listener account</a>
      </p>

<!--#include file="includes/auth_bottom.asp"-->
