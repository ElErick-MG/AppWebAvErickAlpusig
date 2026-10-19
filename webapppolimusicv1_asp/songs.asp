<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  songs.asp — song list + HTML5 player per row
'  CRUD column and buttons are rendered only for USER_TYPE = 1 (admin) and
'  every write operation is re-checked on the server in delete.asp / song_form.asp
' ============================================================================
Dim sgQ, sgSql, sgRs, sgAdmin, sgParams, sgTotal, sgPlays, sgHasRows
Dim sgArt, sgArtist, sgPath, sgId, sgName

RequireLogin

sgAdmin = IsAdmin()
sgQ = QsStr("q")
sgTotal = 0
sgPlays = 0
sgHasRows = False

PageTitle = "Songs list"
PageEyebrow = "Library"
ActiveNav = "songs"
PageAllowSearch = True
PageSearchPlaceholder = "Filter by song, artist or album..."

sgSql = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
        "ar.NAMES AS ARTISTS, al.IMG AS ALBUM_IMAGE, al.TITLE AS ALBUM_TITLE " & _
        "FROM TBL_SONG s " & _
        "OUTER APPLY (SELECT STRING_AGG(a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME, ''), ', ') " & _
        "             WITHIN GROUP (ORDER BY a.ARTIST_NAME) AS NAMES " & _
        "             FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST " & _
        "             WHERE p.ID_SONG = s.ID_SONG) ar " & _
        "OUTER APPLY (SELECT TOP 1 b.ALBUM_IMAGE_PATH AS IMG, b.TITLE_ALBUM AS TITLE " & _
        "             FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM b ON b.ID_ALBUM = t.ID_ALBUM " & _
        "             WHERE t.ID_SONG = s.ID_SONG ORDER BY b.ID_ALBUM) al "

If sgQ <> "" Then
    sgSql = sgSql & "WHERE (s.SONG_NAME LIKE ? OR ISNULL(ar.NAMES, '') LIKE ? OR ISNULL(al.TITLE, '') LIKE ?) "
    sgParams = Array("%" & sgQ & "%", "%" & sgQ & "%", "%" & sgQ & "%")
Else
    sgParams = Array()
End If
sgSql = sgSql & "ORDER BY s.ID_SONG"

Set sgRs = DbQuery(sgSql, sgParams)
%>
<!--#include file="includes/layout_top.asp"-->

      <div class="page-head">
        <div>
          <p class="page-head__eyebrow">Welcome <%= H(CurrentUserName()) %></p>
          <h1>Songs list</h1>
          <p class="page-head__sub">
            Every track in the PoliMusic catalogue, ready to stream. Press play and the
            player bar at the bottom follows you across the whole application.
          </p>
        </div>
        <div class="page-head__actions">
          <% If sgAdmin Then %>
          <a class="btn btn--primary" href="song_form.asp"><%= Icon("plus") %> New song</a>
          <% End If %>
          <a class="btn btn--ghost" href="albums.asp"><%= Icon("disc") %> Browse albums</a>
        </div>
      </div>

      <div class="panel animate-in">
        <div class="panel__head">
          <div class="panel__title">
            <%= Icon("music") %> Song List
            <span>&middot; <%= sgRs.RecordCount %> track<%= IIf(sgRs.RecordCount = 1, "", "s") %><%= IIf(sgQ <> "", " matching “" & H(sgQ) & "”", "") %></span>
          </div>
          <div class="page-head__actions">
            <% If sgQ <> "" Then %>
            <a class="btn btn--ghost btn--sm" href="songs.asp"><%= Icon("x") %> Clear filter</a>
            <% End If %>
            <span class="pill"><%= Icon("headphones") %> <b id="pm-total-plays">live</b> counters</span>
          </div>
        </div>

        <% If sgRs.EOF Then %>
          <div class="empty">
            <%= Icon("music") %>
            <h3>No songs found</h3>
            <p>
              <% If sgQ <> "" Then %>
              Nothing matches “<%= H(sgQ) %>”. Try another word or clear the filter.
              <% Else %>
              The catalogue is empty. <% If sgAdmin Then %>Add the first song from the “New song” button.<% End If %>
              <% End If %>
            </p>
            <% If sgAdmin And sgQ = "" Then %>
            <a class="btn btn--primary btn--sm" href="song_form.asp"><%= Icon("plus") %> Add song</a>
            <% End If %>
          </div>
        <% Else %>
        <div class="table-wrap">
          <table class="grid" id="songs-table">
            <thead>
              <tr>
                <th class="col-id">ID</th>
                <th>Song Name</th>
                <th>Play</th>
                <th class="col-num">Reproductions</th>
                <% If sgAdmin Then %><th class="col-actions">Actions</th><% End If %>
              </tr>
            </thead>
            <tbody>
            <%
            Do While Not sgRs.EOF
                sgId = sgRs("ID_SONG")
                sgName = Txt(sgRs("SONG_NAME"))
                sgPath = Txt(sgRs("SONG_PATH"))
                sgArtist = Txt(sgRs("ARTISTS"))
                sgArt = AssetUrl(sgRs("ALBUM_IMAGE"), "assets/img/art-placeholder.svg")
                sgTotal = sgTotal + 1
                sgPlays = sgPlays + ToInt(sgRs("PLAYS"), 0)
                sgHasRows = True
            %>
              <tr data-song-row="<%= sgId %>">
                <td class="col-id">#<%= Right("00" & sgId, 2) %></td>
                <td>
                  <div class="cell-song">
                    <img class="cell-song__art" src="<%= sgArt %>" alt="" loading="lazy">
                    <div class="cell-song__text">
                      <div class="cell-song__name">
                        <%= H(sgName) %>
                        <span class="eq" aria-hidden="true"><i></i><i></i><i></i><i></i></span>
                      </div>
                      <div class="cell-song__sub">
                        <%= H(IIf(sgArtist = "", "Unknown artist", sgArtist)) %>
                        <%= IIf(Txt(sgRs("ALBUM_TITLE")) = "", "", " · " & H(sgRs("ALBUM_TITLE"))) %>
                      </div>
                    </div>
                  </div>
                </td>
                <td>
                  <div class="cell-play">
                    <button class="play-btn" type="button" data-play-song="<%= sgId %>"
                            aria-label="Play <%= H(sgName) %>">
                      <%= IconC("play", "ico-play") %>
                      <%= IconC("pause", "ico-pause") %>
                    </button>
                    <div class="audio-cell">
                      <%= AudioPlayer(sgId, sgPath, sgName) %>
                      <%= AudioWarning(sgPath) %>
                    </div>
                  </div>
                </td>
                <td class="col-num">
                  <span class="plays" data-plays-for="<%= sgId %>"><%= FmtNum(sgRs("PLAYS")) %></span>
                </td>
                <% If sgAdmin Then %>
                <td class="col-actions">
                  <div class="row-actions">
                    <a class="icon-btn" href="song_form.asp?id=<%= sgId %>" title="Edit song"
                       aria-label="Edit <%= H(sgName) %>"><%= Icon("edit") %></a>
                    <form method="post" action="delete.asp" style="display:inline"
                          data-confirm="Delete “<%= H(sgName) %>” from the catalogue? This cannot be undone.">
                      <%= CsrfField() %>
                      <input type="hidden" name="entity" value="song">
                      <input type="hidden" name="id" value="<%= sgId %>">
                      <button class="icon-btn" type="submit" title="Delete song"
                              aria-label="Delete <%= H(sgName) %>"><%= Icon("trash") %></button>
                    </form>
                  </div>
                </td>
                <% End If %>
              </tr>
            <%
                sgRs.MoveNext
            Loop
            %>
              <tr id="songs-empty-filter" class="is-hidden">
                <td colspan="<%= IIf(sgAdmin, 5, 4) %>">
                  <div class="empty"><p>No song on this page matches the filter.</p></div>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
        <div class="panel__head" style="border-top:1px solid var(--line-soft);border-bottom:0">
          <span class="muted"><%= sgTotal %> song<%= IIf(sgTotal = 1, "", "s") %> &middot; <%= FmtNum(sgPlays) %> reproductions in total</span>
          <span class="badge <%= IIf(sgAdmin, "badge--admin", "badge--user") %>">
            <%= IIf(sgAdmin, "Administrator view", "Listener view") %>
          </span>
        </div>
        <% End If %>
      </div>

      <p class="muted" style="margin-top:16px;font-size:12.5px">
        <%= Icon("info") %>
        Audio streams straight from <span class="mono">SONG_PATH</span>. A “File not deployed”
        marker means the row points to a file that is not present in the
        <span class="mono">songFiles</span> folder yet.
      </p>

<%
sgRs.Close
Set sgRs = Nothing
%>
<!--#include file="includes/layout_bottom.asp"-->
