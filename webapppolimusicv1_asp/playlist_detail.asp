<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  playlist_detail.asp — tracks of one playlist + add / remove songs
'  A listener may only touch their own playlists; admins may touch any.
' ============================================================================
Dim pdId, pdRs, pdTracks, pdAvail, pdCanEdit, pdOwnerId, pdErr, pdAction
Dim pdSongId, pdPos, pdPlays, pdType, pdAvailCount

RequireLogin

pdId = QsInt("id", 0)
pdErr = ""
pdAction = FormStr("action")

Set pdRs = DbQuery( _
    "SELECT p.ID_PLAYLIST, p.ID_USER, p.TITLE, p.CREATION_DATE, p.TYPE, p.PHOTO, u.USERNAME " & _
    "FROM TBL_PLAYLIST p INNER JOIN TBL_USER u ON u.ID_USER = p.ID_USER WHERE p.ID_PLAYLIST = ?", _
    Array(pdId))
If pdRs.EOF Then
    pdRs.Close
    Set pdRs = Nothing
    FlashErr "That playlist does not exist."
    Go "playlists.asp"
End If

pdOwnerId = ToInt(pdRs("ID_USER"), -1)
pdType = ToInt(pdRs("TYPE"), PLAYLIST_PUBLIC)
pdCanEdit = CanManagePlaylist(pdOwnerId)

If pdType = PLAYLIST_PRIVATE And Not pdCanEdit Then
    pdRs.Close
    Set pdRs = Nothing
    FlashErr "That playlist is private."
    Go "playlists.asp"
End If

' ---- add a song ------------------------------------------------------------
If IsPost() And pdAction = "add" Then
    RequireCsrf
    If Not pdCanEdit Then
        FlashErr "You can only edit your own playlists."
        Go "playlist_detail.asp?id=" & pdId
    End If
    pdSongId = FormInt("song_id", 0)
    If pdSongId <= 0 Then
        pdErr = "Pick a song from the list first."
    ElseIf DbInt("SELECT COUNT(*) FROM TBL_SONG WHERE ID_SONG = ?", Array(pdSongId), 0) = 0 Then
        pdErr = "That song is not in the catalogue any more."
    ElseIf DbInt("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", Array(pdId, pdSongId), 0) > 0 Then
        pdErr = "That song is already part of this playlist."
    Else
        DbExec "INSERT INTO TBL_PLAYLIST_SONG (ID_PLAYLIST, ID_SONG) VALUES (?, ?)", Array(pdId, pdSongId)
        FlashOk "Song added to “" & Txt(pdRs("TITLE")) & "”."
        Go "playlist_detail.asp?id=" & pdId
    End If
End If

PageTitle = Txt(pdRs("TITLE"))
PageEyebrow = "Playlist"
ActiveNav = "playlists"
PageAllowSearch = True

Set pdTracks = DbQuery( _
    "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, ar.NAMES AS ARTISTS, al.IMG AS ALBUM_IMAGE " & _
    "FROM TBL_PLAYLIST_SONG ps INNER JOIN TBL_SONG s ON s.ID_SONG = ps.ID_SONG " & _
    "OUTER APPLY (SELECT STRING_AGG(a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME, ''), ', ') " & _
    "             WITHIN GROUP (ORDER BY a.ARTIST_NAME) AS NAMES FROM TBL_PLAYER p " & _
    "             INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG) ar " & _
    "OUTER APPLY (SELECT TOP 1 b.ALBUM_IMAGE_PATH AS IMG FROM TBL_TRACKLIST t " & _
    "             INNER JOIN TBL_ALBUM b ON b.ID_ALBUM = t.ID_ALBUM WHERE t.ID_SONG = s.ID_SONG " & _
    "             ORDER BY b.ID_ALBUM) al " & _
    "WHERE ps.ID_PLAYLIST = ? ORDER BY s.SONG_NAME", Array(pdId))

Set pdAvail = DbQuery( _
    "SELECT ID_SONG, SONG_NAME FROM TBL_SONG WHERE ID_SONG NOT IN " & _
    "(SELECT ID_SONG FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?) ORDER BY SONG_NAME", Array(pdId))
pdAvailCount = pdAvail.RecordCount
%>
<!--#include file="includes/layout_top.asp"-->

      <div class="cover-hero animate-in">
        <img class="cover-hero__img" src="<%= AssetUrl(pdRs("PHOTO"), "playlistImages/default.jpg") %>" alt="">
        <div class="cover-hero__body">
          <div class="cover-hero__kind">Playlist &middot; <%= IIf(pdType = PLAYLIST_PRIVATE, "private", "public") %></div>
          <h1 class="cover-hero__title"><%= H(pdRs("TITLE")) %></h1>
          <div class="cover-hero__meta">
            <span><%= H(pdRs("USERNAME")) %></span>
            <span class="dot"></span>
            <span><%= pdTracks.RecordCount %> song<%= IIf(pdTracks.RecordCount = 1, "", "s") %></span>
            <span class="dot"></span>
            <span>created <%= FmtDate(pdRs("CREATION_DATE")) %></span>
            <span class="dot"></span>
            <%= PlaylistTypeBadge(pdType) %>
          </div>
          <div class="cover-hero__actions">
            <a class="btn btn--ghost btn--sm" href="playlists.asp"><%= Icon("arrow-left") %> All playlists</a>
            <% If pdCanEdit Then %>
            <a class="btn btn--primary btn--sm" href="playlist_form.asp?id=<%= pdId %>"><%= Icon("edit") %> Edit details</a>
            <form method="post" action="delete.asp" style="display:inline"
                  data-confirm="Delete playlist “<%= H(pdRs("TITLE")) %>”?">
              <%= CsrfField() %>
              <input type="hidden" name="entity" value="playlist">
              <input type="hidden" name="id" value="<%= pdId %>">
              <input type="hidden" name="return" value="playlists.asp">
              <button class="btn btn--danger btn--sm" type="submit"><%= Icon("trash") %> Delete</button>
            </form>
            <% End If %>
          </div>
        </div>
      </div>

      <%= ErrorBox(pdErr) %>

      <% If pdCanEdit Then %>
      <div class="panel animate-in" style="margin-bottom:22px">
        <form method="post" action="playlist_detail.asp?id=<%= pdId %>" class="panel__head" style="gap:12px">
          <%= CsrfField() %>
          <input type="hidden" name="action" value="add">
          <div class="field" style="flex:1 1 320px">
            <label for="song_id">Add a song to this playlist</label>
            <select id="song_id" name="song_id"<%= IIf(pdAvailCount = 0, " disabled", "") %>>
              <% If pdAvailCount = 0 Then %>
              <option value="0">Every song is already in this playlist</option>
              <% Else %>
              <option value="0">Choose a song…</option>
              <%
              Do While Not pdAvail.EOF
              %>
              <option value="<%= pdAvail("ID_SONG") %>"><%= H(pdAvail("SONG_NAME")) %></option>
              <%
                  pdAvail.MoveNext
              Loop
              %>
              <% End If %>
            </select>
          </div>
          <button class="btn btn--accent" type="submit"<%= IIf(pdAvailCount = 0, " disabled", "") %>>
            <%= Icon("plus") %> Add
          </button>
        </form>
      </div>
      <% End If %>

      <div class="panel animate-in">
        <div class="panel__head">
          <div class="panel__title"><%= Icon("list") %> Songs <span>&middot; TBL_PLAYLIST_SONG</span></div>
          <span class="pill"><%= Icon("headphones") %> <b><%= pdTracks.RecordCount %></b> tracks</span>
        </div>

        <% If pdTracks.EOF Then %>
          <%= EmptyState("music", "This playlist is empty", "Add songs with the selector above to start building the sequence.") %>
        <% Else %>
        <div class="table-wrap">
          <table class="grid">
            <thead>
              <tr>
                <th class="col-id">#</th>
                <th>Song Name</th>
                <th>Play</th>
                <th class="col-num">Reproductions</th>
                <% If pdCanEdit Then %><th class="col-actions">Actions</th><% End If %>
              </tr>
            </thead>
            <tbody>
              <%
              pdPos = 0
              pdPlays = 0
              Do While Not pdTracks.EOF
                  pdPos = pdPos + 1
                  pdPlays = pdPlays + ToInt(pdTracks("PLAYS"), 0)
              %>
              <tr data-song-row="<%= pdTracks("ID_SONG") %>">
                <td class="col-id"><%= Right("0" & pdPos, 2) %></td>
                <td>
                  <div class="cell-song">
                    <img class="cell-song__art" src="<%= AssetUrl(pdTracks("ALBUM_IMAGE"), "assets/img/art-placeholder.svg") %>" alt="" loading="lazy">
                    <div class="cell-song__text">
                      <div class="cell-song__name">
                        <%= H(pdTracks("SONG_NAME")) %>
                        <span class="eq" aria-hidden="true"><i></i><i></i><i></i><i></i></span>
                      </div>
                      <div class="cell-song__sub"><%= H(IIf(Txt(pdTracks("ARTISTS")) = "", "Unknown artist", pdTracks("ARTISTS"))) %></div>
                    </div>
                  </div>
                </td>
                <td>
                  <div class="cell-play">
                    <button class="play-btn" type="button" data-play-song="<%= pdTracks("ID_SONG") %>"
                            aria-label="Play <%= H(pdTracks("SONG_NAME")) %>">
                      <%= IconC("play", "ico-play") %>
                      <%= IconC("pause", "ico-pause") %>
                    </button>
                    <div class="audio-cell">
                      <%= AudioPlayer(pdTracks("ID_SONG"), pdTracks("SONG_PATH"), pdTracks("SONG_NAME")) %>
                      <%= AudioWarning(pdTracks("SONG_PATH")) %>
                    </div>
                  </div>
                </td>
                <td class="col-num">
                  <span class="plays" data-plays-for="<%= pdTracks("ID_SONG") %>"><%= FmtNum(pdTracks("PLAYS")) %></span>
                </td>
                <% If pdCanEdit Then %>
                <td class="col-actions">
                  <div class="row-actions">
                    <form method="post" action="delete.asp" style="display:inline"
                          data-confirm="Remove “<%= H(pdTracks("SONG_NAME")) %>” from this playlist? The song stays in the catalogue.">
                      <%= CsrfField() %>
                      <input type="hidden" name="entity" value="playlist_song">
                      <input type="hidden" name="id" value="<%= pdId %>">
                      <input type="hidden" name="song" value="<%= pdTracks("ID_SONG") %>">
                      <input type="hidden" name="return" value="playlist_detail.asp?id=<%= pdId %>">
                      <button class="icon-btn" type="submit" title="Remove from playlist"><%= Icon("x") %></button>
                    </form>
                  </div>
                </td>
                <% End If %>
              </tr>
              <%
                  pdTracks.MoveNext
              Loop
              %>
            </tbody>
          </table>
        </div>
        <div class="panel__head" style="border-top:1px solid var(--line-soft);border-bottom:0">
          <span class="muted"><%= pdTracks.RecordCount %> track<%= IIf(pdTracks.RecordCount = 1, "", "s") %> &middot; <%= FmtNum(pdPlays) %> reproductions</span>
          <span class="badge <%= IIf(pdCanEdit, "badge--user", "badge--muted") %>"><%= IIf(pdCanEdit, "You can edit this playlist", "Read only") %></span>
        </div>
        <% End If %>
      </div>

<%
pdTracks.Close
Set pdTracks = Nothing
pdAvail.Close
Set pdAvail = Nothing
pdRs.Close
Set pdRs = Nothing
%>
<!--#include file="includes/layout_bottom.asp"-->
