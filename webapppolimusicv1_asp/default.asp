<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  default.asp — landing dashboard with the welcome banner
' ============================================================================
Dim hmCounts, hmSongs, hmAlbums, hmArtists, hmMyPl
Dim hmSongCount, hmArtistCount, hmAlbumCount, hmPlayCount, hmPlCount
Dim hmArt, hmIdx

RequireLogin

PageTitle = "Home"
PageEyebrow = "Overview"
ActiveNav = "home"
PageAllowSearch = True
PageSearchPlaceholder = "Search the catalogue..."

Set hmCounts = DbQuery( _
    "SELECT (SELECT COUNT(*) FROM TBL_SONG) AS SONGS, " & _
    "(SELECT COUNT(*) FROM TBL_ARTIST) AS ARTISTS, " & _
    "(SELECT COUNT(*) FROM TBL_ALBUM) AS ALBUMS, " & _
    "(SELECT ISNULL(SUM(PLAYS), 0) FROM TBL_SONG) AS PLAYS, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = ?) AS PLAYLISTS", _
    Array(CurrentUserId()))

hmSongCount = 0
hmArtistCount = 0
hmAlbumCount = 0
hmPlayCount = 0
hmPlCount = 0
If Not hmCounts.EOF Then
    hmSongCount = ToInt(hmCounts("SONGS"), 0)
    hmArtistCount = ToInt(hmCounts("ARTISTS"), 0)
    hmAlbumCount = ToInt(hmCounts("ALBUMS"), 0)
    hmPlayCount = ToInt(hmCounts("PLAYS"), 0)
    hmPlCount = ToInt(hmCounts("PLAYLISTS"), 0)
End If
hmCounts.Close
Set hmCounts = Nothing

Set hmSongs = DbQuery( _
    "SELECT TOP 5 s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
    "ar.NAMES AS ARTISTS, al.IMG AS ALBUM_IMAGE " & _
    "FROM TBL_SONG s " & _
    "OUTER APPLY (SELECT STRING_AGG(a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME, ''), ', ') " & _
    "             WITHIN GROUP (ORDER BY a.ARTIST_NAME) AS NAMES FROM TBL_PLAYER p " & _
    "             INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG) ar " & _
    "OUTER APPLY (SELECT TOP 1 b.ALBUM_IMAGE_PATH AS IMG FROM TBL_TRACKLIST t " & _
    "             INNER JOIN TBL_ALBUM b ON b.ID_ALBUM = t.ID_ALBUM WHERE t.ID_SONG = s.ID_SONG " & _
    "             ORDER BY b.ID_ALBUM) al " & _
    "ORDER BY s.ID_SONG DESC", Array())

Set hmAlbums = DbQuery( _
    "SELECT TOP 6 b.ID_ALBUM, b.TITLE_ALBUM, b.RELEASE_YEAR, b.ALBUM_IMAGE_PATH, " & _
    "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = b.ID_ALBUM) AS TRACKS " & _
    "FROM TBL_ALBUM b ORDER BY b.ID_ALBUM", Array())

Set hmArtists = DbQuery( _
    "SELECT TOP 6 a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = a.ID_ARTIST) AS SONGS " & _
    "FROM TBL_ARTIST a ORDER BY a.ARTIST_NAME", Array())

Set hmMyPl = DbQuery( _
    "SELECT TOP 4 p.ID_PLAYLIST, p.TITLE, p.PHOTO, p.TYPE, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS TRACKS " & _
    "FROM TBL_PLAYLIST p WHERE p.ID_USER = ? ORDER BY p.CREATION_DATE DESC, p.ID_PLAYLIST DESC", _
    Array(CurrentUserId()))
%>
<!--#include file="includes/layout_top.asp"-->

      <section class="hero animate-in">
        <div class="hero__inner">
          <div>
            <p class="hero__eyebrow">Signed in as <%= IIf(IsAdmin(), "administrator", "listener") %></p>
            <h1 class="hero__title">Welcome <span><%= H(CurrentUserName()) %></span></h1>
            <p class="hero__text">
              This is your PoliMusic room: <%= hmSongCount %> tracks, <%= hmAlbumCount %> albums and
              <%= hmArtistCount %> artists, all streamed straight from SQL Server.
            </p>
            <div class="hero__actions">
              <a class="btn btn--primary" href="songs.asp"><%= Icon("play") %> Browse songs</a>
              <a class="btn btn--ghost" href="playlists.asp"><%= Icon("list") %> My playlists</a>
              <a class="btn btn--danger" href="logout.asp"><%= Icon("logout") %> Log out</a>
            </div>
          </div>

          <div class="hero__aside">
            <div class="hero__stats">
              <div class="hero__stat">
                <b><%= hmSongCount %></b>
                <span>Tracks</span>
              </div>
              <div class="hero__stat">
                <b><%= FmtNum(hmPlayCount) %></b>
                <span>Plays</span>
              </div>
              <div class="hero__stat">
                <b><%= hmPlCount %></b>
                <span>Playlists</span>
              </div>
            </div>
            <%= Avatar(CurrentUserPhoto(), CurrentUserName(), "hero__avatar") %>
          </div>
        </div>
      </section>

      <div class="stat-grid animate-in">
        <div class="stat">
          <div class="stat__label">Catalogue</div>
          <div class="stat__value"><%= hmSongCount %></div>
          <div class="stat__hint">songs in TBL_SONG</div>
        </div>
        <div class="stat stat--jade">
          <div class="stat__label">Artists</div>
          <div class="stat__value"><%= hmArtistCount %></div>
          <div class="stat__hint">linked through TBL_PLAYER</div>
        </div>
        <div class="stat">
          <div class="stat__label">Albums</div>
          <div class="stat__value"><%= hmAlbumCount %></div>
          <div class="stat__hint">with tracklists</div>
        </div>
        <div class="stat stat--jade">
          <div class="stat__label">Your role</div>
          <div class="stat__value" style="font-size:19px;letter-spacing:0"><%= IIf(IsAdmin(), "Administrator", "Listener") %></div>
          <div class="stat__hint"><%= IIf(IsAdmin(), "full catalogue CRUD", "stream + own playlists") %></div>
        </div>
      </div>

      <section class="section animate-in">
        <div class="section__head">
          <h2>Recently added songs</h2>
          <a class="section__link" href="songs.asp">View all <%= Icon("arrow-right") %></a>
        </div>

        <div class="panel">
          <% If hmSongs.EOF Then %>
            <%= EmptyState("music", "No songs yet", "The catalogue is empty. An administrator can add the first track from the songs page.") %>
          <% Else %>
          <div class="track-list">
            <%
            hmIdx = 0
            Do While Not hmSongs.EOF
                hmIdx = hmIdx + 1
                hmArt = AssetUrl(hmSongs("ALBUM_IMAGE"), "assets/img/art-placeholder.svg")
            %>
            <div class="track-row" data-song-row="<%= hmSongs("ID_SONG") %>">
              <span class="track-row__idx"><%= Right("0" & hmIdx, 2) %></span>
              <div class="cell-song">
                <img class="cell-song__art" src="<%= hmArt %>" alt="" loading="lazy">
                <div class="cell-song__text">
                  <div class="cell-song__name">
                    <%= H(hmSongs("SONG_NAME")) %>
                    <span class="eq" aria-hidden="true"><i></i><i></i><i></i><i></i></span>
                  </div>
                  <div class="cell-song__sub"><%= H(IIf(Txt(hmSongs("ARTISTS")) = "", "Unknown artist", hmSongs("ARTISTS"))) %></div>
                </div>
              </div>
              <span class="track-row__plays"><%= Icon("headphones") %> <%= FmtNum(hmSongs("PLAYS")) %></span>
              <button class="play-btn" type="button" data-play-song="<%= hmSongs("ID_SONG") %>"
                      aria-label="Play <%= H(hmSongs("SONG_NAME")) %>">
                <%= IconC("play", "ico-play") %>
                <%= IconC("pause", "ico-pause") %>
              </button>
              <div class="is-hidden"><%= AudioPlayer(hmSongs("ID_SONG"), hmSongs("SONG_PATH"), hmSongs("SONG_NAME")) %></div>
            </div>
            <%
                hmSongs.MoveNext
            Loop
            %>
          </div>
          <% End If %>
        </div>
      </section>

      <section class="section animate-in">
        <div class="section__head">
          <h2>Albums</h2>
          <a class="section__link" href="albums.asp">All albums <%= Icon("arrow-right") %></a>
        </div>
        <div class="card-grid">
          <%
          Do While Not hmAlbums.EOF
          %>
          <a class="card" href="album_detail.asp?id=<%= hmAlbums("ID_ALBUM") %>">
            <div class="card__cover">
              <img src="<%= AssetUrl(hmAlbums("ALBUM_IMAGE_PATH"), "assets/img/art-placeholder.svg") %>" alt="" loading="lazy">
              <span class="card__badge"><%= hmAlbums("TRACKS") %> tracks</span>
            </div>
            <div class="card__title"><%= H(hmAlbums("TITLE_ALBUM")) %></div>
            <div class="card__meta"><%= IIf(Txt(hmAlbums("RELEASE_YEAR")) = "", "Year unknown", Txt(hmAlbums("RELEASE_YEAR"))) %></div>
          </a>
          <%
              hmAlbums.MoveNext
          Loop
          %>
        </div>
      </section>

      <section class="section animate-in">
        <div class="section__head">
          <h2>Artists</h2>
          <a class="section__link" href="artists.asp">All artists <%= Icon("arrow-right") %></a>
        </div>
        <div class="card-grid">
          <%
          Do While Not hmArtists.EOF
          %>
          <a class="card" href="artists.asp">
            <div class="card__cover card__cover--round">
              <img src="<%= AssetUrl(hmArtists("ARTIST_IMAGE"), "assets/img/art-placeholder.svg") %>" alt="" loading="lazy">
            </div>
            <div class="card__title" style="text-align:center">
              <%= H(hmArtists("ARTIST_NAME") & " " & Txt(hmArtists("ARTIST_LASTNAME"))) %>
            </div>
            <div class="card__meta" style="text-align:center"><%= hmArtists("SONGS") %> songs</div>
          </a>
          <%
              hmArtists.MoveNext
          Loop
          %>
        </div>
      </section>

      <section class="section animate-in">
        <div class="section__head">
          <h2>Your playlists</h2>
          <a class="section__link" href="playlists.asp">Manage <%= Icon("arrow-right") %></a>
        </div>
        <% If hmMyPl.EOF Then %>
          <div class="panel">
            <%= EmptyState("list", "No playlists yet", "Create a playlist to group the tracks you keep coming back to.") %>
            <div style="text-align:center;padding:0 0 26px">
              <a class="btn btn--accent btn--sm" href="playlist_form.asp"><%= Icon("plus") %> Create playlist</a>
            </div>
          </div>
        <% Else %>
        <div class="card-grid">
          <%
          Do While Not hmMyPl.EOF
          %>
          <a class="card" href="playlist_detail.asp?id=<%= hmMyPl("ID_PLAYLIST") %>">
            <div class="card__cover">
              <img src="<%= AssetUrl(hmMyPl("PHOTO"), "playlistImages/default.jpg") %>" alt="" loading="lazy">
              <span class="card__badge<%= IIf(ToInt(hmMyPl("TYPE"), 0) = PLAYLIST_PRIVATE, " card__badge--private", "") %>">
                <%= IIf(ToInt(hmMyPl("TYPE"), 0) = PLAYLIST_PRIVATE, "Private", "Public") %>
              </span>
            </div>
            <div class="card__title"><%= H(hmMyPl("TITLE")) %></div>
            <div class="card__meta"><%= hmMyPl("TRACKS") %> songs</div>
          </a>
          <%
              hmMyPl.MoveNext
          Loop
          %>
          <a class="card" href="playlist_form.asp" style="display:grid;place-items:center;min-height:200px;border-style:dashed">
            <div style="text-align:center;color:var(--text-mute)">
              <div style="font-size:26px;line-height:1;margin-bottom:6px;color:var(--jade)">+</div>
              New playlist
            </div>
          </a>
        </div>
        <% End If %>
      </section>

<%
hmSongs.Close
Set hmSongs = Nothing
hmAlbums.Close
Set hmAlbums = Nothing
hmArtists.Close
Set hmArtists = Nothing
hmMyPl.Close
Set hmMyPl = Nothing
%>
<!--#include file="includes/layout_bottom.asp"-->
