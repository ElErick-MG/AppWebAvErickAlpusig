<%
' ============================================================================
'  PoliMusic — Configuration
'  Included FIRST on every page. Declares global constants and page vars.
' ============================================================================
Option Explicit

' ---- Application -----------------------------------------------------------
Const APP_NAME       = "PoliMusic"
Const APP_VERSION    = "1.0.0"
Const APP_TAGLINE    = "Escuela Politecnica Nacional"

' ---- SQL Server (port 14333) ----------------------------------------------
' Server=localhost,14333;Database=BDD_PoliMusic;User Id=usr_polimusic;
' Password=Politecnica1;Encrypt=true;TrustServerCertificate=true;
Const DB_SERVER      = "localhost,14333"
Const DB_NAME        = "BDD_PoliMusic"
Const DB_USER        = "usr_polimusic"
Const DB_PASSWORD    = "Politecnica1"
Const DB_ENCRYPT     = "Yes"
Const DB_TRUST_CERT  = "Yes"
Const DB_TIMEOUT     = 8

' ---- Roles (TBL_USER.USER_TYPE) -------------------------------------------
Const ROLE_USER      = 0
Const ROLE_ADMIN     = 1

' ---- Catalog: TBL_PLAYLIST.TYPE (TBL_CATALOG parent code 2) ----------------
Const PLAYLIST_PUBLIC  = 0
Const PLAYLIST_PRIVATE = 1

' ---- Pagination ------------------------------------------------------------
Const ROWS_PER_PAGE  = 25

' ---- Page level variables (assigned by each page before layout_top) --------
Dim PageTitle, PageEyebrow, ActiveNav, PageSearch, PageSearchPlaceholder
Dim PageAllowSearch, PageBodyClass, AuthTitle, AuthSub

PageTitle = APP_NAME
PageEyebrow = ""
ActiveNav = ""
PageSearch = ""
PageSearchPlaceholder = "Search songs, albums, artists..."
PageAllowSearch = False
PageBodyClass = ""
AuthTitle = "Sign in"
AuthSub = ""

' ---- Session lifetime ------------------------------------------------------
If Len(Session("id_user") & "") > 0 Then
    Session.Timeout = 60
End If
%>
