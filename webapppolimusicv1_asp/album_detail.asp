<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  album_detail.asp — album header + tracklist with HTML5 players
' ============================================================================
Dim adId, adRs, adTracks, adAdmin, adPlays, adPos

RequireLogin

adId = QsInt("id", 0)
adAdmin = IsAdmin()
adPlays = 0
adPos = 0

Set adRs = DbQuery( _
    "SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", _
    Array(adId))
If adRs.EOF Then
    adRs.Close
    Set adRs = Nothing
    FlashErr "That album does not exist."
    Go "albums.asp"
End If

PageTitle = Txt(adRs("TITLE_ALBUM"))
PageEyebrow = "Album"
ActiveNav = "albums"
PageAllowSearch = True

Set adTracks = DbQuery( _
    "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, ar.NAMES AS ARTISTS " & _
    "FROM TBL_TRACKLIST t INNER JOIN TBL_SONG s ON s.ID_SONG = t.ID_SONG " & _
    "OUTER APPLY (SELECT STRING_AGG(a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME, ''), ', ') " & _
    "             WITHIN GROUP (ORDER BY a.ARTIST_NAME) AS NAMES FROM TBL_PLAYER p " & _
    "             INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG) ar " & _
    "WHERE t.ID_ALBUM = ? ORDER BY s.ID_SONG", Array(adId))
%>
<!--#include file="includes/layout_top.asp"-->

      <div class="cover-hero animate-in">
        <img class="cover-hero__img" src="<%= AssetUrl(adRs("ALBUM_IMAGE_PATH"), "assets/img/art-placeholder.svg") %>" alt="">
        <div class="cover-hero__body">
          <div class="cover-hero__kind">Album</div>
          <h1 class="cover-hero__title"><%= H(adRs("TITLE_ALBUM")) %></h1>
          <div class="cover-hero__meta">
            <span><%= IIf(Txt(adRs("RELEASE_YEAR")) = "", "Year unknown", Txt(adRs("RELEASE_YEAR"))) %></span>
            <span class="dot"></span>
            <span><%= adTracks.RecordCount %> track<%= IIf(adTracks.RecordCount = 1, "", "s") %></span>
            <span class="dot"></span>
            <span class="mono">ALBUM #<%= adId %></span>
          </div>
          <div class="cover-hero__actions">
            <a class="btn btn--ghost btn--sm" href="albums.asp"><%= Icon("arrow-left") %> All albums</a>
            <% If adAdmin Then %>
            <a class="btn btn--primary btn--sm" href="album_form.asp?id=<%= adId %>"><%= Icon("edit") %> Edit album</a>
            <% End If %>
          </div>
        </div>
      </div>

      <div class="panel animate-in">
        <div class="panel__head">
          <div class="panel__title"><%= Icon("list") %> Tracklist <span>&middot; TBL_TRACKLIST</span></div>
          <span class="pill"><%= Icon("headphones") %> <b><%= adTracks.RecordCount %></b> tracks</span>
        </div>

        <% If adTracks.EOF Then %>
          <%= EmptyState("music", "This album has no tracks", "Link songs to this album from the song editor, or pick another album.") %>
        <% Else %>
        <div class="table-wrap">
          <table class="grid">
            <thead>
              <tr>
                <th class="col-id">#</th>
                <th>Song Name</th>
                <th>Play</th>
                <th class="col-num">Reproductions</th>
              </tr>
            </thead>
            <tbody>
              <%
              Do While Not adTracks.EOF
                  adPos = adPos + 1
                  adPlays = adPlays + ToInt(adTracks("PLAYS"), 0)
              %>
              <tr data-song-row="<%= adTracks("ID_SONG") %>">
                <td class="col-id"><%= Right("0" & adPos, 2) %></td>
                <td>
                  <div class="cell-song">
                    <img class="cell-song__art" src="<%= AssetUrl(adRs("ALBUM_IMAGE_PATH"), "assets/img/art-placeholder.svg") %>" alt="" loading="lazy">
                    <div class="cell-song__text">
                      <div class="cell-song__name">
                        <%= H(adTracks("SONG_NAME")) %>
                        <span class="eq" aria-hidden="true"><i></i><i></i><i></i><i></i></span>
                      </div>
                      <div class="cell-song__sub"><%= H(IIf(Txt(adTracks("ARTISTS")) = "", "Unknown artist", adTracks("ARTISTS"))) %></div>
                    </div>
                  </div>
                </td>
                <td>
                  <div class="cell-play">
                    <button class="play-btn" type="button" data-play-song="<%= adTracks("ID_SONG") %>"
                            aria-label="Play <%= H(adTracks("SONG_NAME")) %>">
                      <%= IconC("play", "ico-play") %>
                      <%= IconC("pause", "ico-pause") %>
                    </button>
                    <div class="audio-cell">
                      <%= AudioPlayer(adTracks("ID_SONG"), adTracks("SONG_PATH"), adTracks("SONG_NAME")) %>
                      <%= AudioWarning(adTracks("SONG_PATH")) %>
                    </div>
                  </div>
                </td>
                <td class="col-num">
                  <span class="plays" data-plays-for="<%= adTracks("ID_SONG") %>"><%= FmtNum(adTracks("PLAYS")) %></span>
                </td>
              </tr>
              <%
                  adTracks.MoveNext
              Loop
              %>
            </tbody>
          </table>
        </div>
        <div class="panel__head" style="border-top:1px solid var(--line-soft);border-bottom:0">
          <span class="muted"><%= adTracks.RecordCount %> track<%= IIf(adTracks.RecordCount = 1, "", "s") %> &middot; <%= FmtNum(adPlays) %> reproductions</span>
        </div>
        <% End If %>
      </div>

<%
adTracks.Close
Set adTracks = Nothing
adRs.Close
Set adRs = Nothing
%>
<!--#include file="includes/layout_bottom.asp"-->
