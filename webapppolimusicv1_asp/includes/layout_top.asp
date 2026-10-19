<%
' ============================================================================
'  PoliMusic — Layout (opening): head, sidebar, topbar, player bar
'  Requires: PageTitle, ActiveNav, PageAllowSearch, PageSearch
' ============================================================================
Dim lyNavCounts, lyPls, lyPlPhoto, lyPlTitle, lyPlId, lyNavClass, lySearchQ
Dim lySongCount, lyArtistCount, lyAlbumCount, lyMyPlCount, lyIsAdminNow

lyIsAdminNow = IsAdmin()

Set lyNavCounts = DbQuery( _
    "SELECT (SELECT COUNT(*) FROM TBL_SONG) AS SONGS, " & _
    "(SELECT COUNT(*) FROM TBL_ARTIST) AS ARTISTS, " & _
    "(SELECT COUNT(*) FROM TBL_ALBUM) AS ALBUMS, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = ?) AS MYPLAYLISTS", _
    Array(CurrentUserId()))

lySongCount = 0
lyArtistCount = 0
lyAlbumCount = 0
lyMyPlCount = 0
If Not lyNavCounts.EOF Then
    lySongCount = ToInt(lyNavCounts("SONGS"), 0)
    lyArtistCount = ToInt(lyNavCounts("ARTISTS"), 0)
    lyAlbumCount = ToInt(lyNavCounts("ALBUMS"), 0)
    lyMyPlCount = ToInt(lyNavCounts("MYPLAYLISTS"), 0)
End If
lyNavCounts.Close
Set lyNavCounts = Nothing

Set lyPls = DbQuery( _
    "SELECT TOP 7 ID_PLAYLIST, TITLE, PHOTO FROM TBL_PLAYLIST WHERE ID_USER = ? " & _
    "ORDER BY CREATION_DATE DESC, ID_PLAYLIST DESC", Array(CurrentUserId()))

lySearchQ = Txt(Request.QueryString("q"))
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#0d0d0d">
<meta name="csrf-token" content="<%= H(Csrf()) %>">
<title><%= H(PageTitle) %> &middot; PoliMusic</title>
<link rel="stylesheet" href="assets/css/polimusic.css">
<link rel="icon" type="image/svg+xml" href="assets/img/favicon.svg">
</head>
<body class="<%= H(PageBodyClass) %>">
<div class="app">

  <div class="sidebar-backdrop" data-sidebar-close></div>

  <aside class="sidebar" id="sidebar" aria-label="Main navigation">
    <a class="brand" href="default.asp">
      <span class="brand__mark">
        <svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round">
          <path d="M6 18V9.5M12 18V5.5M18 18v-6"/>
        </svg>
      </span>
      <span>
        <span class="brand__name">Poli<b>Music</b></span>
        <span class="brand__tag">EPN Streaming</span>
      </span>
    </a>

    <nav class="nav">
      <div class="nav__label">Browse</div>

      <a class="nav__item <%= IIf(ActiveNav = "home", "is-active", "") %>" href="default.asp">
        <%= Icon("home") %> Home
      </a>
      <a class="nav__item <%= IIf(ActiveNav = "songs", "is-active", "") %>" href="songs.asp">
        <%= Icon("music") %> Songs
        <span class="nav__count"><%= lySongCount %></span>
      </a>
      <a class="nav__item <%= IIf(ActiveNav = "albums", "is-active", "") %>" href="albums.asp">
        <%= Icon("disc") %> Albums
        <span class="nav__count"><%= lyAlbumCount %></span>
      </a>
      <a class="nav__item <%= IIf(ActiveNav = "artists", "is-active", "") %>" href="artists.asp">
        <%= Icon("mic") %> Artists
        <span class="nav__count"><%= lyArtistCount %></span>
      </a>
      <a class="nav__item <%= IIf(ActiveNav = "playlists", "is-active", "") %>" href="playlists.asp">
        <%= Icon("list") %> Playlists
        <span class="nav__count"><%= lyMyPlCount %></span>
      </a>

      <% If lyIsAdminNow Then %>
      <div class="nav__label">Administration</div>
      <a class="nav__item <%= IIf(ActiveNav = "users", "is-active", "") %>" href="users.asp">
        <%= Icon("users") %> Users
      </a>
      <% End If %>
    </nav>

    <div class="side-playlists">
      <div class="side-playlists__head">
        <span>Your playlists</span>
        <a href="playlist_form.asp" title="Create playlist" aria-label="Create playlist">+</a>
      </div>
      <% If lyPls.EOF Then %>
        <div class="side-empty">No playlists yet. Create one to start collecting songs.</div>
      <% Else %>
        <% Do While Not lyPls.EOF %>
          <a class="side-pl" href="playlist_detail.asp?id=<%= lyPls("ID_PLAYLIST") %>">
            <img src="<%= AssetUrl(lyPls("PHOTO"), "playlistImages/default.jpg") %>" alt="">
            <span><%= H(Short(lyPls("TITLE"), 22)) %></span>
          </a>
        <% lyPls.MoveNext
           Loop %>
      <% End If %>
    </div>

    <div class="side-user">
      <%= Avatar(CurrentUserPhoto(), CurrentUserName(), "side-user__avatar") %>
      <span class="side-user__meta">
        <span class="side-user__name" title="<%= H(CurrentUserName()) %>"><%= H(CurrentUserName()) %></span>
        <span class="side-user__role"><%= IIf(lyIsAdminNow, "Administrator", "Listener") %></span>
      </span>
      <a class="side-user__out" href="logout.asp" title="Log out" aria-label="Log out"><%= Icon("logout") %></a>
    </div>
  </aside>

  <div class="main">
    <header class="topbar">
      <button class="icon-btn topbar__menu" type="button" data-sidebar-toggle aria-label="Open navigation">
        <%= Icon("menu") %>
      </button>

      <% If PageAllowSearch Then %>
      <form class="search" method="get" action="songs.asp" role="search">
        <%= Icon("search") %>
        <input type="search" name="q" value="<%= H(lySearchQ) %>"
               placeholder="<%= H(PageSearchPlaceholder) %>" aria-label="Search songs">
      </form>
      <% Else %>
      <span class="topbar__spacer"></span>
      <% End If %>

      <span class="topbar__spacer"></span>

      <div class="topbar__actions">
        <a class="btn btn--ghost btn--sm" href="logout.asp" title="Close your session">
          <%= Icon("logout") %> Log out
        </a>
        <a class="chip-user" href="playlists.asp" title="<%= H(CurrentUserName()) %>">
          <%= Avatar(CurrentUserPhoto(), CurrentUserName(), "") %>
          <span><%= H(Short(CurrentUserName(), 18)) %></span>
        </a>
      </div>
    </header>

    <main class="content">
      <%= RenderFlash() %>
