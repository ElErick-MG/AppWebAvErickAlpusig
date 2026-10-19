<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  register.asp — self enrolment for normal listeners (USER_TYPE = 0)
'  The role is hard-coded server side: nobody can self-promote to admin.
' ============================================================================
Dim rgUser, rgEmail, rgBirth, rgPass, rgPass2, rgErr, rgBirthOk, rgNewId
Dim rgBirthVal, rgNewRs

PageTitle = "Create account"
AuthTitle = "Create your account"
AuthSub = "Listener accounts can stream the catalogue and manage their own playlists."

If IsLoggedIn() Then Go "default.asp"

rgErr = ""
rgUser = ""
rgEmail = ""
rgBirth = ""
rgBirthVal = ""

If IsPost() Then
    RequireCsrf
    rgUser = FormStr("username")
    rgEmail = FormStr("email")
    rgBirth = FormStr("birthday")
    rgPass = Request.Form("password") & ""
    rgPass2 = Request.Form("password2") & ""

    If Not IsValidUsername(rgUser) Then
        rgErr = rgErr & "Username must be 3 to 50 characters: letters, numbers, dots, dashes or underscores." & vbLf
    ElseIf UserExists(rgUser) Then
        rgErr = rgErr & "That username is already taken." & vbLf
    End If

    If Not IsValidEmail(rgEmail) Then
        rgErr = rgErr & "Type a valid email address (example: name@epn.edu.ec)." & vbLf
    ElseIf Len(rgEmail) > 50 Then
        rgErr = rgErr & "The email address is longer than the 50 characters the column allows." & vbLf
    ElseIf EmailExists(rgEmail) Then
        rgErr = rgErr & "That email address is already registered." & vbLf
    End If

    If Len(rgPass) < 6 Then
        rgErr = rgErr & "The password needs at least 6 characters." & vbLf
    ElseIf Len(rgPass) > 50 Then
        rgErr = rgErr & "The password cannot be longer than 50 characters." & vbLf
    ElseIf rgPass <> rgPass2 Then
        rgErr = rgErr & "The two passwords do not match." & vbLf
    End If

    If rgBirth <> "" Then
        rgBirthVal = ParseIsoDate(rgBirth, rgBirthOk)
        If Not rgBirthOk Then
            rgErr = rgErr & "The birthday is not a valid date." & vbLf
        ElseIf rgBirthVal > Now() Then
            rgErr = rgErr & "The birthday cannot be in the future." & vbLf
        End If
    End If

    If rgErr = "" Then
        If rgBirth = "" Then
            rgBirthVal = Null
        Else
            rgBirthVal = SqlDate(rgBirthVal)
        End If
        DbExec "INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) " & _
               "VALUES (?, CONVERT(VARCHAR(40), HASHBYTES('SHA1', CAST(? AS VARCHAR(50))), 2), ?, ?, '~/userImages/default.jpg', 0)", _
               Array(rgUser, rgPass, rgEmail, rgBirthVal)

        Set rgNewRs = DbQuery("SELECT ID_USER, USERNAME, USER_TYPE, USER_PHOTO, EMAIL FROM TBL_USER WHERE USERNAME = ?", Array(rgUser))
        If Not rgNewRs.EOF Then
            SignIn rgNewRs("ID_USER"), rgNewRs("USERNAME"), rgNewRs("USER_TYPE"), rgNewRs("USER_PHOTO"), rgNewRs("EMAIL")
        End If
        rgNewRs.Close
        Set rgNewRs = Nothing
        FlashOk "Welcome to PoliMusic, " & rgUser & ". Your listener account is ready."
        Go "default.asp"
    End If
End If
%>
<!--#include file="includes/auth_top.asp"-->

      <%= ErrorBox(rgErr) %>

      <form class="auth__form" method="post" action="register.asp" novalidate>
        <%= CsrfField() %>

        <div class="field">
          <label for="username">Username</label>
          <input id="username" name="username" type="text" value="<%= H(rgUser) %>"
                 placeholder="e.g. Erick.Alpusig" maxlength="50" required autofocus
                 autocomplete="username" spellcheck="false">
          <span class="field__hint">3&ndash;50 characters. Letters, numbers, dots, dashes, underscores.</span>
        </div>

        <div class="field">
          <label for="email">Email</label>
          <input id="email" name="email" type="email" value="<%= H(rgEmail) %>"
                 placeholder="name@epn.edu.ec" maxlength="50" required autocomplete="email">
        </div>

        <div class="field">
          <label for="birthday">Birthday <span class="muted">(optional)</span></label>
          <input id="birthday" name="birthday" type="date" value="<%= H(rgBirth) %>"
                 min="1900-01-01" max="<%= IsoDate(Now()) %>">
        </div>

        <div class="form-grid">
          <div class="field">
            <label for="password">Password</label>
            <input id="password" name="password" type="password" placeholder="At least 6 characters"
                   maxlength="50" required autocomplete="new-password">
          </div>
          <div class="field">
            <label for="password2">Repeat password</label>
            <input id="password2" name="password2" type="password" placeholder="Type it again"
                   maxlength="50" required autocomplete="new-password">
          </div>
        </div>

        <div class="auth__row">
          <span><%= Icon("shield") %> New accounts are always created with the listener role.</span>
        </div>

        <button class="btn btn--primary btn--block" type="submit">
          <%= Icon("check") %> Create account
        </button>
      </form>

      <p class="auth__switch">
        Already registered? <a href="login.asp">Sign in instead</a>
      </p>

<!--#include file="includes/auth_bottom.asp"-->
