<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  song_form.asp — create / edit a song (administrators only)
'  Also maintains TBL_PLAYER (artists) and TBL_TRACKLIST (album) links.
' ============================================================================
Dim sfId, sfName, sfPath, sfPlays, sfAlbum, sfArtists, sfErr, sfIsEdit, sfTitle
Dim sfRs, sfArtistRs, sfAlbumRs, sfNewId, sfArtistSet, sfSelAlbum, sfHasAudio
Dim sfTmp, sfI

RequireAdmin

sfId = QsInt("id", 0)
sfErr = ""
sfIsEdit = (sfId > 0)

sfName = ""
sfPath = "~/songFiles/"
sfPlays = 0
sfAlbum = 0
sfArtists = ""
sfHasAudio = False

If sfIsEdit Then
    Set sfRs = DbQuery("SELECT ID_SONG, SONG_NAME, SONG_PATH, ISNULL(PLAYS,0) AS PLAYS FROM TBL_SONG WHERE ID_SONG = ?", Array(sfId))
    If sfRs.EOF Then
        sfRs.Close
        Set sfRs = Nothing
        FlashErr "That song no longer exists."
        Go "songs.asp"
    End If
    sfName = Txt(sfRs("SONG_NAME"))
    sfPath = Txt(sfRs("SONG_PATH"))
    sfPlays = ToInt(sfRs("PLAYS"), 0)
    sfRs.Close
    Set sfRs = Nothing

    sfArtistSet = "," & Txt(DbScalar( _
        "SELECT ISNULL(STRING_AGG(CAST(ID_ARTIST AS VARCHAR(12)), ','), '') FROM TBL_PLAYER WHERE ID_SONG = ?", _
        Array(sfId))) & ","
    sfAlbum = ToInt(DbScalar("SELECT TOP 1 ID_ALBUM FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(sfId)), 0)
End If

If IsPost() Then
    RequireCsrf
    sfId = FormInt("id", 0)
    sfIsEdit = (sfId > 0)
    sfName = FormStr("song_name")
    sfPath = FormStr("song_path")
    sfPlays = FormInt("plays", 0)
    sfAlbum = FormInt("album_id", 0)
    sfArtists = FormStr("artist_ids")
    sfArtistSet = "," & sfArtists & ","

    If sfName = "" Then sfErr = sfErr & "The song name is required." & vbLf
    If Len(sfName) > 50 Then sfErr = sfErr & "The song name cannot exceed 50 characters." & vbLf
    If sfPath = "" Then sfErr = sfErr & "The audio path is required." & vbLf
    If Len(sfPath) > 255 Then sfErr = sfErr & "The audio path cannot exceed 255 characters." & vbLf
    If sfPlays < 0 Then sfErr = sfErr & "Reproductions cannot be negative." & vbLf

    If sfErr = "" Then
        If sfIsEdit Then
            DbExec "UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ?, PLAYS = ? WHERE ID_SONG = ?", _
                   Array(sfName, sfPath, sfPlays, sfId)
        Else
            sfNewId = ToInt(DbScalar( _
                "INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) OUTPUT INSERTED.ID_SONG VALUES (?, ?, ?)", _
                Array(sfName, sfPath, sfPlays)), 0)
            sfId = sfNewId
        End If

        ' refresh artist links
        DbExec "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array(sfId)
        If sfArtists <> "" Then
            sfTmp = Split(sfArtists, ",")
            For sfI = 0 To UBound(sfTmp)
                If IsNumeric(Trim(sfTmp(sfI))) And Trim(sfTmp(sfI)) <> "" Then
                    DbExec "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)", _
                           Array(sfId, CLng(Trim(sfTmp(sfI))))
                End If
            Next
        End If

        ' refresh album link (one album per song in this UI)
        DbExec "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(sfId)
        If sfAlbum > 0 Then
            DbExec "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)", Array(sfAlbum, sfId)
        End If

        FlashOk "Song “" & sfName & "” saved."
        Go "songs.asp"
    End If
End If

Set sfArtistRs = DbQuery("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME FROM TBL_ARTIST ORDER BY ARTIST_NAME", Array())
Set sfAlbumRs = DbQuery("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR FROM TBL_ALBUM ORDER BY TITLE_ALBUM", Array())

sfHasAudio = AssetExists(sfPath)

PageTitle = IIf(sfIsEdit, "Edit song", "New song")
PageEyebrow = "Catalogue"
ActiveNav = "songs"
%>
<!--#include file="includes/layout_top.asp"-->

      <div class="page-head">
        <div>
          <p class="page-head__eyebrow"><%= IIf(sfIsEdit, "Edit record", "Create record") %></p>
          <h1><%= IIf(sfIsEdit, "Edit song", "New song") %></h1>
          <p class="page-head__sub">Songs are stored in <span class="mono">TBL_SONG</span>; artists and the album are linked through <span class="mono">TBL_PLAYER</span> and <span class="mono">TBL_TRACKLIST</span>.</p>
        </div>
        <div class="page-head__actions">
          <a class="btn btn--ghost" href="songs.asp"><%= Icon("arrow-left") %> Back to songs</a>
        </div>
      </div>

      <%= ErrorBox(sfErr) %>

      <div class="panel form-card animate-in">
        <form method="post" action="song_form.asp" novalidate>
          <%= CsrfField() %>
          <input type="hidden" name="id" value="<%= IIf(sfIsEdit, sfId, 0) %>">

          <div class="form-grid">
            <div class="field field--full">
              <label for="song_name">Song name</label>
              <input id="song_name" name="song_name" type="text" maxlength="50" required
                     value="<%= H(sfName) %>" placeholder="e.g. Adventure">
            </div>

            <div class="field field--full">
              <label for="song_path">Audio path (SONG_PATH)</label>
              <input id="song_path" name="song_path" type="text" maxlength="255" required
                     value="<%= H(sfPath) %>" placeholder="~/songFiles/bensound-adventure.mp3">
              <span class="field__hint">
                Use the <span class="mono">~/</span> prefix exactly as stored in the database.
                <% If InStr(sfPath, ".") > 0 Then %>
                  &middot; <%= IIf(sfHasAudio, "<span class=""badge badge--user"">File found</span>", "<span class=""badge badge--warn"">File not deployed</span>") %>
                <% End If %>
              </span>
            </div>

            <div class="field">
              <label for="plays">Reproductions</label>
              <input id="plays" name="plays" type="number" min="0" step="1" value="<%= sfPlays %>">
            </div>

            <div class="field">
              <label for="album_id">Album</label>
              <select id="album_id" name="album_id">
                <option value="0">— No album —</option>
                <%
                If Not sfAlbumRs.EOF Then sfAlbumRs.MoveFirst
                Do While Not sfAlbumRs.EOF
                %>
                <option value="<%= sfAlbumRs("ID_ALBUM") %>"<%= SelIf(sfAlbumRs("ID_ALBUM"), sfAlbum) %>>
                  <%= H(sfAlbumRs("TITLE_ALBUM")) %> <%= IIf(Txt(sfAlbumRs("RELEASE_YEAR")) = "", "", "(" & Txt(sfAlbumRs("RELEASE_YEAR")) & ")") %>
                </option>
                <%
                    sfAlbumRs.MoveNext
                Loop
                %>
              </select>
            </div>

            <div class="field field--full">
              <label for="artist_ids">Artists <span class="muted">(Ctrl / Cmd to select several)</span></label>
              <select id="artist_ids" name="artist_ids" multiple size="6">
                <%
                If Not sfArtistRs.EOF Then sfArtistRs.MoveFirst
                Do While Not sfArtistRs.EOF
                    sfArtistSet = Txt(sfArtistSet)
                    If sfArtistSet = "" Then sfArtistSet = ","
                %>
                <option value="<%= sfArtistRs("ID_ARTIST") %>"<%= IIf(InStr(sfArtistSet, "," & sfArtistRs("ID_ARTIST") & ",") > 0, " selected", "") %>>
                  <%= H(sfArtistRs("ARTIST_NAME") & " " & Txt(sfArtistRs("ARTIST_LASTNAME"))) %>
                </option>
                <%
                    sfArtistRs.MoveNext
                Loop
                %>
              </select>
              <span class="field__hint">Artists are shown on the songs list and on the artist page.</span>
            </div>

            <% If sfIsEdit Then %>
            <div class="field field--full">
              <label>Preview</label>
              <div class="audio-cell">
                <%= AudioPlayer(sfId, sfPath, sfName) %>
                <%= AudioWarning(sfPath) %>
              </div>
            </div>
            <% End If %>
          </div>

          <div class="form-actions">
            <% If sfIsEdit Then %>
            <button class="btn btn--danger" type="submit" form="sf-delete-form"><%= Icon("trash") %> Delete</button>
            <% End If %>
            <a class="btn btn--ghost" href="songs.asp">Cancel</a>
            <button class="btn btn--primary" type="submit"><%= Icon("check") %> Save song</button>
          </div>
        </form>

        <% If sfIsEdit Then %>
        <form id="sf-delete-form" method="post" action="delete.asp"
              data-confirm="Delete “<%= H(sfName) %>” from the catalogue?">
          <%= CsrfField() %>
          <input type="hidden" name="entity" value="song">
          <input type="hidden" name="id" value="<%= sfId %>">
          <input type="hidden" name="return" value="songs.asp">
        </form>
        <% End If %>
      </div>

<%
sfArtistRs.Close
Set sfArtistRs = Nothing
sfAlbumRs.Close
Set sfAlbumRs = Nothing
%>
<!--#include file="includes/layout_bottom.asp"-->
