<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  user_form.asp — create / edit a user (administrators only)
'  Passwords are stored as uppercase SHA1 hex; leaving the field empty on an
'  edit keeps the current hash untouched.
' ============================================================================
Dim ufId, ufUser, ufEmail, ufBirth, ufPass, ufPass2, ufPhoto, ufType, ufErr
Dim ufIsEdit, ufRs, ufBirthOk, ufBirthVal, ufIsSelf, ufDupUser, ufDupMail

RequireAdmin

ufId = QsInt("id", 0)
ufErr = ""
ufIsEdit = (ufId > 0)
ufUser = ""
ufEmail = ""
ufBirth = ""
ufPhoto = "~/userImages/default.jpg"
ufType = ROLE_USER
ufIsSelf = (ufIsEdit And ufId = CurrentUserId())

If ufIsEdit Then
    Set ufRs = DbQuery("SELECT ID_USER, USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE FROM TBL_USER WHERE ID_USER = ?", Array(ufId))
    If ufRs.EOF Then
        ufRs.Close
        Set ufRs = Nothing
        FlashErr "That user does not exist."
        Go "users.asp"
    End If
    ufUser = Txt(ufRs("USERNAME"))
    ufEmail = Txt(ufRs("EMAIL"))
    ufBirth = IsoDate(ufRs("BIRTHDAY"))
    ufPhoto = Txt(ufRs("USER_PHOTO"))
    ufType = ToInt(ufRs("USER_TYPE"), ROLE_USER)
    ufRs.Close
    Set ufRs = Nothing
End If

If IsPost() Then
    RequireCsrf
    ufId = FormInt("id", 0)
    ufIsEdit = (ufId > 0)
    ufIsSelf = (ufIsEdit And ufId = CurrentUserId())
    ufUser = FormStr("username")
    ufEmail = FormStr("email")
    ufBirth = FormStr("birthday")
    ufPass = Request.Form("password") & ""
    ufPass2 = Request.Form("password2") & ""
    ufPhoto = FormStr("photo")
    ufType = FormInt("user_type", ROLE_USER)

    If ufPhoto = "" Then ufPhoto = "~/userImages/default.jpg"
    If ufType <> ROLE_ADMIN Then ufType = ROLE_USER
    ' An administrator cannot demote their own account (avoids lock-out).
    If ufIsSelf Then ufType = ROLE_ADMIN

    If Not IsValidUsername(ufUser) Then
        ufErr = ufErr & "Username must be 3 to 50 characters: letters, numbers, dots, dashes or underscores." & vbLf
    Else
        ufDupUser = DbInt("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = ? AND ID_USER <> ?", Array(ufUser, ufId), 0)
        If ufDupUser > 0 Then ufErr = ufErr & "Another account already uses that username." & vbLf
    End If

    If Not IsValidEmail(ufEmail) Then
        ufErr = ufErr & "Type a valid email address." & vbLf
    ElseIf Len(ufEmail) > 50 Then
        ufErr = ufErr & "The email cannot exceed 50 characters." & vbLf
    Else
        ufDupMail = DbInt("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = ? AND ID_USER <> ?", Array(ufEmail, ufId), 0)
        If ufDupMail > 0 Then ufErr = ufErr & "Another account already uses that email." & vbLf
    End If

    If Not ufIsEdit Then
        If Len(ufPass) < 6 Then ufErr = ufErr & "New users need a password of at least 6 characters." & vbLf
    End If
    If ufPass <> "" Then
        If Len(ufPass) > 50 Then ufErr = ufErr & "The password cannot exceed 50 characters." & vbLf
        If ufPass <> ufPass2 Then ufErr = ufErr & "The two passwords do not match." & vbLf
    End If

    ufBirthVal = Null
    If ufBirth <> "" Then
        ufBirthVal = ParseIsoDate(ufBirth, ufBirthOk)
        If Not ufBirthOk Then ufErr = ufErr & "The birthday is not a valid date." & vbLf
    End If

    If ufErr = "" Then
        If Not IsNull(ufBirthVal) Then ufBirthVal = SqlDate(ufBirthVal)
        If ufIsEdit Then
            DbExec "UPDATE TBL_USER SET USERNAME = ?, EMAIL = ?, BIRTHDAY = ?, USER_PHOTO = ?, USER_TYPE = ? WHERE ID_USER = ?", _
                   Array(ufUser, ufEmail, ufBirthVal, ufPhoto, ufType, ufId)
            If ufPass <> "" Then
                DbExec "UPDATE TBL_USER SET PASSWORD = CONVERT(VARCHAR(40), HASHBYTES('SHA1', CAST(? AS VARCHAR(50))), 2) WHERE ID_USER = ?", _
                       Array(ufPass, ufId)
            End If
            If ufIsSelf Then Session("username") = ufUser
            If ufIsSelf Then Session("user_photo") = ufPhoto
            FlashOk "User “" & ufUser & "” updated."
        Else
            DbExec "INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) " & _
                   "VALUES (?, CONVERT(VARCHAR(40), HASHBYTES('SHA1', CAST(? AS VARCHAR(50))), 2), ?, ?, ?, ?)", _
                   Array(ufUser, ufPass, ufEmail, ufBirthVal, ufPhoto, ufType)
            FlashOk "User “" & ufUser & "” created."
        End If
        Go "users.asp"
    End If
End If

PageTitle = IIf(ufIsEdit, "Edit user", "New user")
PageEyebrow = "Administration"
ActiveNav = "users"
%>
<!--#include file="includes/layout_top.asp"-->

      <div class="page-head">
        <div>
          <p class="page-head__eyebrow"><%= IIf(ufIsEdit, "Edit record", "Create record") %></p>
          <h1><%= IIf(ufIsEdit, "Edit user", "New user") %></h1>
          <p class="page-head__sub">Accounts live in <span class="mono">TBL_USER</span>. Passwords are hashed with SHA1 by SQL Server before they are stored.</p>
        </div>
        <div class="page-head__actions">
          <a class="btn btn--ghost" href="users.asp"><%= Icon("arrow-left") %> Back to users</a>
        </div>
      </div>

      <%= ErrorBox(ufErr) %>

      <div class="panel form-card animate-in">
        <form method="post" action="user_form.asp" novalidate>
          <%= CsrfField() %>
          <input type="hidden" name="id" value="<%= IIf(ufIsEdit, ufId, 0) %>">

          <div class="form-grid">
            <div class="field">
              <label for="username">Username</label>
              <input id="username" name="username" type="text" maxlength="50" required
                     value="<%= H(ufUser) %>" placeholder="e.g. Erick.Alpusig">
            </div>

            <div class="field">
              <label for="email">Email</label>
              <input id="email" name="email" type="email" maxlength="50" required
                     value="<%= H(ufEmail) %>" placeholder="name@epn.edu.ec">
            </div>

            <div class="field">
              <label for="birthday">Birthday</label>
              <input id="birthday" name="birthday" type="date" value="<%= H(ufBirth) %>" min="1900-01-01">
            </div>

            <div class="field">
              <label for="user_type">Role</label>
              <select id="user_type" name="user_type"<%= IIf(ufIsSelf, " disabled", "") %>>
                <option value="<%= ROLE_USER %>"<%= SelIf(ROLE_USER, ufType) %>>User — listen and manage own playlists</option>
                <option value="<%= ROLE_ADMIN %>"<%= SelIf(ROLE_ADMIN, ufType) %>>Administrator — full CRUD</option>
              </select>
              <% If ufIsSelf Then %>
              <input type="hidden" name="user_type" value="<%= ROLE_ADMIN %>">
              <span class="field__hint">You cannot change your own role.</span>
              <% End If %>
            </div>

            <div class="field field--full">
              <label for="photo">Photo path</label>
              <input id="photo" name="photo" type="text" maxlength="255"
                     value="<%= H(ufPhoto) %>" placeholder="~/userImages/default.jpg">
              <span class="field__hint">Available in the deployment: default.jpg, root.png, username1.png</span>
            </div>

            <div class="field">
              <label for="password">Password <%= IIf(ufIsEdit, "", "") %></label>
              <input id="password" name="password" type="password" maxlength="50"
                     placeholder="<%= IIf(ufIsEdit, "Leave empty to keep the current one", "At least 6 characters") %>"
                     <%= IIf(ufIsEdit, "", "required") %> autocomplete="new-password">
            </div>

            <div class="field">
              <label for="password2">Repeat password</label>
              <input id="password2" name="password2" type="password" maxlength="50"
                     placeholder="Type it again" autocomplete="new-password">
            </div>
          </div>

          <div class="form-actions">
            <a class="btn btn--ghost" href="users.asp">Cancel</a>
            <button class="btn btn--primary" type="submit"><%= Icon("check") %> Save user</button>
          </div>
        </form>
      </div>

<!--#include file="includes/layout_bottom.asp"-->
