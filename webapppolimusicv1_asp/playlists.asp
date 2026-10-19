<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  playlists.asp — own playlists, public playlists and (for admins) everything
'  Listeners only ever see their own private playlists.
' ============================================================================
Dim plRs, plAdmin, plCount, plIsMine, plCanEdit, plType

RequireLogin

plAdmin = IsAdmin()

PageTitle = "Playlists"
PageEyebrow = "Collections"
ActiveNav = "playlists"
PageAllowSearch = True

Set plRs = DbQuery( _
    "SELECT p.ID_PLAYLIST, p.ID_USER, p.TITLE, p.CREATION_DATE, p.TYPE, p.PHOTO, u.USERNAME, " & _
    "CASE WHEN p.ID_USER = ? THEN 1 ELSE 0 END AS IS_MINE, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS TRACKS " & _
    "FROM TBL_PLAYLIST p INNER JOIN TBL_USER u ON u.ID_USER = p.ID_USER " & _
    "WHERE (p.ID_USER = ? OR p.TYPE = 0 OR ? = 1) " & _
    "ORDER BY IS_MINE DESC, p.CREATION_DATE DESC, p.ID_PLAYLIST DESC", _
    Array(CurrentUserId(), CurrentUserId(), IIf(plAdmin, 1, 0)))

plCount = plRs.RecordCount
%>
<!--#include file="includes/layout_top.asp"-->

      <div class="page-head">
        <div>
          <p class="page-head__eyebrow">Your collections</p>
          <h1>Playlists</h1>
          <p class="page-head__sub">
            <%= IIf(plAdmin, "As an administrator you can see and manage every playlist in the system.", "You can create, rename and delete your own playlists; public ones from other listeners are listed too.") %>
          </p>
        </div>
        <div class="page-head__actions">
          <a class="btn btn--accent" href="playlist_form.asp"><%= Icon("plus") %> New playlist</a>
        </div>
      </div>

      <% If plRs.EOF Then %>
      <div class="panel">
        <%= EmptyState("list", "No playlists yet", "Create your first playlist and start adding songs from the catalogue.") %>
        <div style="text-align:center;padding:0 0 26px">
          <a class="btn btn--accent btn--sm" href="playlist_form.asp"><%= Icon("plus") %> Create playlist</a>
        </div>
      </div>
      <% Else %>
      <div class="card-grid animate-in">
        <%
        Do While Not plRs.EOF
            plIsMine = (ToInt(plRs("IS_MINE"), 0) = 1)
            plType = ToInt(plRs("TYPE"), PLAYLIST_PUBLIC)
            plCanEdit = (plIsMine Or plAdmin)
        %>
        <div class="card">
          <a class="card__cover" href="playlist_detail.asp?id=<%= plRs("ID_PLAYLIST") %>">
            <img src="<%= AssetUrl(plRs("PHOTO"), "playlistImages/default.jpg") %>" alt="" loading="lazy">
            <span class="card__badge<%= IIf(plType = PLAYLIST_PRIVATE, " card__badge--private", "") %>">
              <%= IIf(plType = PLAYLIST_PRIVATE, "Private", "Public") %>
            </span>
          </a>
          <a class="card__title" href="playlist_detail.asp?id=<%= plRs("ID_PLAYLIST") %>"><%= H(plRs("TITLE")) %></a>
          <div class="card__meta">
            <%= plRs("TRACKS") %> song<%= IIf(ToInt(plRs("TRACKS"), 0) = 1, "", "s") %>
            &middot; <%= IIf(plIsMine, "yours", H(plRs("USERNAME"))) %>
          </div>
          <div class="row-actions" style="margin-top:10px;justify-content:flex-start">
            <a class="icon-btn" href="playlist_detail.asp?id=<%= plRs("ID_PLAYLIST") %>" title="Open playlist"><%= Icon("arrow-right") %></a>
            <% If plCanEdit Then %>
            <a class="icon-btn" href="playlist_form.asp?id=<%= plRs("ID_PLAYLIST") %>" title="Edit playlist"><%= Icon("edit") %></a>
            <form method="post" action="delete.asp" style="display:inline"
                  data-confirm="Delete playlist “<%= H(plRs("TITLE")) %>”?">
              <%= CsrfField() %>
              <input type="hidden" name="entity" value="playlist">
              <input type="hidden" name="id" value="<%= plRs("ID_PLAYLIST") %>">
              <input type="hidden" name="return" value="playlists.asp">
              <button class="icon-btn" type="submit" title="Delete playlist"><%= Icon("trash") %></button>
            </form>
            <% End If %>
          </div>
        </div>
        <%
            plRs.MoveNext
        Loop
        %>
      </div>
      <% End If %>

<%
plRs.Close
Set plRs = Nothing
%>
<!--#include file="includes/layout_bottom.asp"-->
