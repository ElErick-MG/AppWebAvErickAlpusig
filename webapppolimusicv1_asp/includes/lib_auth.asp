<%
' ============================================================================
'  PoliMusic — Authentication & authorisation
'  Roles: USER_TYPE = 0 (normal user)  /  USER_TYPE = 1 (administrator)
' ============================================================================

Function IsLoggedIn()
    IsLoggedIn = (Len(Txt(Session("id_user"))) > 0)
End Function

Function IsAdmin()
    IsAdmin = (IsLoggedIn() And ToInt(Session("user_type"), ROLE_USER) = ROLE_ADMIN)
End Function

Function CurrentUserId()
    CurrentUserId = ToInt(Session("id_user"), 0)
End Function

Function CurrentUserName()
    CurrentUserName = Txt(Session("username"))
End Function

Function CurrentUserPhoto()
    CurrentUserPhoto = Txt(Session("user_photo"))
End Function

' Blocks anonymous visitors and remembers where they wanted to go.
Sub RequireLogin()
    If Not IsLoggedIn() Then
        Go "login.asp?next=" & Server.URLEncode(CurrentUrl())
    End If
End Sub

' Server-side guard: normal users can never reach administration screens,
' even by typing the URL directly.
Sub RequireAdmin()
    RequireLogin
    If Not IsAdmin() Then
        FlashErr "Access denied. That area is restricted to administrators."
        Go "default.asp"
    End If
End Sub

' Owner-or-admin guard used by playlist editing.
Function CanManagePlaylist(ByVal ownerId)
    CanManagePlaylist = (IsLoggedIn() And (IsAdmin() Or CurrentUserId() = ToInt(ownerId, -1)))
End Function

Sub SignIn(ByVal idUser, ByVal userName, ByVal userType, ByVal photo, ByVal email)
    Session("id_user") = CLng(idUser)
    Session("username") = Txt(userName)
    Session("user_type") = CLng(userType)
    Session("user_photo") = Txt(photo)
    Session("user_email") = Txt(email)
    Session("pm_csrf") = ""
    Csrf()
End Sub

Sub SignOut()
    Session.Contents.RemoveAll
    Session.Abandon
End Sub

' SHA1 comparison is delegated to SQL Server (HASHBYTES) so that both the
' legacy plain-text rows and the hashed rows of the seed script can log in.
Function FindUserByCredentials(ByVal userName, ByVal plainPassword)
    Dim rs
    Set rs = DbQuery( _
        "SELECT TOP 1 ID_USER, USERNAME, USER_TYPE, USER_PHOTO, EMAIL " & _
        "FROM TBL_USER " & _
        "WHERE USERNAME = ? AND (PASSWORD = ? OR UPPER(PASSWORD) = " & _
        "CONVERT(VARCHAR(40), HASHBYTES('SHA1', CAST(? AS VARCHAR(50))), 2))", _
        Array(userName, plainPassword, plainPassword))
    If rs.EOF Then
        Set FindUserByCredentials = Nothing
    Else
        Set FindUserByCredentials = rs
    End If
End Function

Function UserExists(ByVal userName)
    UserExists = (DbInt("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = ?", Array(userName), 0) > 0)
End Function

Function EmailExists(ByVal email)
    EmailExists = (DbInt("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = ?", Array(email), 0) > 0)
End Function

' Passwords are stored as uppercase SHA1 hex, matching the seeded format.
Function Sha1Hex(ByVal plain)
    Dim v
    v = DbScalar("SELECT CONVERT(VARCHAR(40), HASHBYTES('SHA1', CAST(? AS VARCHAR(50))), 2)", Array(plain))
    Sha1Hex = Txt(v)
End Function
%>
