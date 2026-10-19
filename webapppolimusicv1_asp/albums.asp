<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  albums.asp — album catalogue
' ============================================================================
Dim alRs, alAdmin, alCount

RequireLogin

alAdmin = IsAdmin()

PageTitle = "Albums"
PageEyebrow = "Catalogue"
ActiveNav = "albums"
PageAllowSearch = True

Set alRs = DbQuery( _
    "SELECT b.ID_ALBUM, b.TITLE_ALBUM, b.RELEASE_YEAR, b.ALBUM_IMAGE_PATH, " & _
    "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = b.ID_ALBUM) AS TRACKS " & _
    "FROM TBL_ALBUM b ORDER BY b.RELEASE_YEAR DESC, b.TITLE_ALBUM", Array())

alCount = alRs.RecordCount
%>
<!--#include file="includes/layout_top.asp"-->

      <div class="page-head">
        <div>
          <p class="page-head__eyebrow">Collections</p>
          <h1>Albums</h1>
          <p class="page-head__sub">
            <%= alCount %> album<%= IIf(alCount = 1, "", "s") %> and their tracklists, resolved from
            <span class="mono">TBL_ALBUM</span> and <span class="mono">TBL_TRACKLIST</span>.
          </p>
        </div>
        <div class="page-head__actions">
          <% If alAdmin Then %>
          <a class="btn btn--primary" href="album_form.asp"><%= Icon("plus") %> New album</a>
          <% End If %>
          <a class="btn btn--ghost" href="artists.asp"><%= Icon("mic") %> Artists</a>
        </div>
      </div>

      <% If alRs.EOF Then %>
      <div class="panel">
        <%= EmptyState("disc", "No albums yet", "Albums group songs into tracklists. An administrator can create the first one.") %>
      </div>
      <% Else %>
      <div class="card-grid animate-in">
        <%
        Do While Not alRs.EOF
        %>
        <div class="card">
          <a class="card__cover" href="album_detail.asp?id=<%= alRs("ID_ALBUM") %>">
            <img src="<%= AssetUrl(alRs("ALBUM_IMAGE_PATH"), "assets/img/art-placeholder.svg") %>" alt="" loading="lazy">
            <span class="card__badge"><%= alRs("TRACKS") %> tracks</span>
          </a>
          <a class="card__title" href="album_detail.asp?id=<%= alRs("ID_ALBUM") %>"><%= H(alRs("TITLE_ALBUM")) %></a>
          <div class="card__meta"><%= IIf(Txt(alRs("RELEASE_YEAR")) = "", "Year unknown", Txt(alRs("RELEASE_YEAR"))) %></div>
          <% If alAdmin Then %>
          <div class="row-actions" style="margin-top:10px;justify-content:flex-start">
            <a class="icon-btn" href="album_form.asp?id=<%= alRs("ID_ALBUM") %>" title="Edit album"><%= Icon("edit") %></a>
            <form method="post" action="delete.asp" style="display:inline"
                  data-confirm="Delete album “<%= H(alRs("TITLE_ALBUM")) %>”? Its tracklist links will be removed.">
              <%= CsrfField() %>
              <input type="hidden" name="entity" value="album">
              <input type="hidden" name="id" value="<%= alRs("ID_ALBUM") %>">
              <input type="hidden" name="return" value="albums.asp">
              <button class="icon-btn" type="submit" title="Delete album"><%= Icon("trash") %></button>
            </form>
          </div>
          <% End If %>
        </div>
        <%
            alRs.MoveNext
        Loop
        %>
      </div>
      <% End If %>

<%
alRs.Close
Set alRs = Nothing
%>
<!--#include file="includes/layout_bottom.asp"-->
