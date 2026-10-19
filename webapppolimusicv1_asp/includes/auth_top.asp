<%
' ============================================================================
'  PoliMusic — Auth layout (opening)
'  Requires: PageTitle, AuthTitle, AuthSub
' ============================================================================
%><!DOCTYPE html>
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
<body>
<div class="auth">

  <aside class="auth__aside">
    <a class="brand" href="login.asp">
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

    <div>
      <p class="hero__eyebrow">Music, reimagined</p>
      <h2 class="auth__headline">Your catalogue.<br><em>Your playlists.</em><br>One dark room.</h2>
      <p class="auth__lead">
        PoliMusic is a compact streaming platform built on ASP Classic and SQL Server:
        every song, album and artist you own, wrapped in a charcoal and terracotta interface.
      </p>

      <div class="auth__features">
        <div class="auth__feature">
          <%= Icon("waveform") %>
          <span><b>Stream the whole library.</b> Native HTML5 playback with a persistent player bar.</span>
        </div>
        <div class="auth__feature">
          <%= Icon("list") %>
          <span><b>Curate your own playlists.</b> Public or private, always yours to edit.</span>
        </div>
        <div class="auth__feature">
          <%= Icon("shield") %>
          <span><b>Roles enforced server-side.</b> Only administrators touch the catalogue.</span>
        </div>
      </div>
    </div>

    <p class="auth__foot">PoliMusic v<%= APP_VERSION %> &middot; Escuela Polit&eacute;cnica Nacional &middot; ASP Classic + SQL Server</p>
  </aside>

  <div class="auth__panel">
    <div class="auth__card">
      <h1 class="auth__title"><%= H(AuthTitle) %></h1>
      <p class="auth__sub"><%= H(AuthSub) %></p>
