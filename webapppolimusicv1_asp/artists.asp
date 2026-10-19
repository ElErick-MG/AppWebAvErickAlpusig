<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  artists.asp — artist catalogue (read only for listeners, CRUD for admins)
' ============================================================================
Dim arRs, arAdmin, arFullName, arCount

RequireLogin

arAdmin = IsAdmin()

PageTitle = "Artists"
PageEyebrow = "Catalogue"
ActiveNav = "artists"
PageAllowSearch = True

Set arRs = DbQuery( _
    "SELECT a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = a.ID_ARTIST) AS SONGS, " & _
    "(SELECT COUNT(DISTINCT t.ID_ALBUM) FROM TBL_PLAYER p " & _
    " INNER JOIN TBL_TRACKLIST t ON t.ID_SONG = p.ID_SONG WHERE p.ID_ARTIST = a.ID_ARTIST) AS ALBUMS " & _
    "FROM TBL_ARTIST a ORDER BY a.ARTIST_NAME, a.ARTIST_LASTNAME", Array())

arCount = arRs.RecordCount
%>
<!--#include file="includes/layout_top.asp"-->

      <div class="page-head">
        <div>
          <p class="page-head__eyebrow">People behind the music</p>
          <h1>Artists</h1>
          <p class="page-head__sub">
            <%= arCount %> artist<%= IIf(arCount = 1, "", "s") %> registered in <span class="mono">TBL_ARTIST</span>,
            linked to songs through <span class="mono">TBL_PLAYER</span>.
          </p>
        </div>
        <div class="page-head__actions">
          <% If arAdmin Then %>
          <a class="btn btn--primary" href="artist_form.asp"><%= Icon("plus") %> New artist</a>
          <% End If %>
          <a class="btn btn--ghost" href="songs.asp"><%= Icon("music") %> Songs</a>
        </div>
      </div>

      <% If arRs.EOF Then %>
      <div class="panel">
        <%= EmptyState("mic", "No artists yet", "Artists appear here as soon as they are registered in the catalogue.") %>
      </div>
      <% Else %>
      <div class="card-grid animate-in">
        <%
        Do While Not arRs.EOF
            arFullName = Trim(arRs("ARTIST_NAME") & " " & Txt(arRs("ARTIST_LASTNAME")))
        %>
        <div class="card">
          <a class="card__cover card__cover--round" href="songs.asp?q=<%= Server.URLEncode(arFullName) %>"
             title="Songs by <%= H(arFullName) %>">
            <img src="<%= AssetUrl(arRs("ARTIST_IMAGE"), "assets/img/art-placeholder.svg") %>" alt="" loading="lazy">
          </a>
          <a class="card__title" href="songs.asp?q=<%= Server.URLEncode(arFullName) %>"><%= H(arFullName) %></a>
          <div class="card__meta">
            <%= arRs("SONGS") %> song<%= IIf(ToInt(arRs("SONGS"), 0) = 1, "", "s") %>
            &middot; <%= arRs("ALBUMS") %> album<%= IIf(ToInt(arRs("ALBUMS"), 0) = 1, "", "s") %>
          </div>
          <% If arAdmin Then %>
          <div class="row-actions" style="margin-top:10px;justify-content:flex-start">
            <a class="icon-btn" href="artist_form.asp?id=<%= arRs("ID_ARTIST") %>" title="Edit artist"><%= Icon("edit") %></a>
            <form method="post" action="delete.asp" style="display:inline"
                  data-confirm="Delete artist “<%= H(arFullName) %>” and all of their song links?">
              <%= CsrfField() %>
              <input type="hidden" name="entity" value="artist">
              <input type="hidden" name="id" value="<%= arRs("ID_ARTIST") %>">
              <input type="hidden" name="return" value="artists.asp">
              <button class="icon-btn" type="submit" title="Delete artist"><%= Icon("trash") %></button>
            </form>
          </div>
          <% End If %>
        </div>
        <%
            arRs.MoveNext
        Loop
        %>
      </div>
      <% End If %>

<%
arRs.Close
Set arRs = Nothing
%>
<!--#include file="includes/layout_bottom.asp"-->
