<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  playlist_form.asp — create / edit a playlist
'  Listeners may only edit their own; administrators may edit any.
' ============================================================================
Dim pfId, pfTitle, pfType, pfPhoto, pfErr, pfIsEdit, pfOwner, pfRs, pfNewId

RequireLogin

pfId = QsInt("id", 0)
pfErr = ""
pfIsEdit = (pfId > 0)
pfTitle = ""
pfType = PLAYLIST_PUBLIC
pfPhoto = "~/playlistImages/default.jpg"
pfOwner = CurrentUserId()

If pfIsEdit Then
    Set pfRs = DbQuery("SELECT ID_PLAYLIST, ID_USER, TITLE, TYPE, PHOTO FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(pfId))
    If pfRs.EOF Then
        pfRs.Close
        Set pfRs = Nothing
        FlashErr "That playlist does not exist."
        Go "playlists.asp"
    End If
    pfOwner = ToInt(pfRs("ID_USER"), -1)
    If Not CanManagePlaylist(pfOwner) Then
        pfRs.Close
        Set pfRs = Nothing
        FlashErr "You can only edit your own playlists."
        Go "playlists.asp"
    End If
    pfTitle = Txt(pfRs("TITLE"))
    pfType = ToInt(pfRs("TYPE"), PLAYLIST_PUBLIC)
    pfPhoto = Txt(pfRs("PHOTO"))
    pfRs.Close
    Set pfRs = Nothing
End If

If IsPost() Then
    RequireCsrf
    pfId = FormInt("id", 0)
    pfIsEdit = (pfId > 0)
    pfTitle = FormStr("title")
    pfType = FormInt("type", PLAYLIST_PUBLIC)
    pfPhoto = FormStr("photo")

    If pfTitle = "" Then pfErr = pfErr & "The playlist needs a title." & vbLf
    If Len(pfTitle) > 50 Then pfErr = pfErr & "The title cannot exceed 50 characters." & vbLf
    If pfType <> PLAYLIST_PRIVATE Then pfType = PLAYLIST_PUBLIC
    If pfPhoto = "" Then pfPhoto = "~/playlistImages/default.jpg"
    If Len(pfPhoto) > 255 Then pfErr = pfErr & "The image path cannot exceed 255 characters." & vbLf

    If pfErr = "" Then
        If pfIsEdit Then
            If Not CanManagePlaylist(ToInt(DbScalar("SELECT ID_USER FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(pfId)), -1)) Then
                FlashErr "You can only edit your own playlists."
                Go "playlists.asp"
            End If
            DbExec "UPDATE TBL_PLAYLIST SET TITLE = ?, TYPE = ?, PHOTO = ? WHERE ID_PLAYLIST = ?", _
                   Array(pfTitle, pfType, pfPhoto, pfId)
            FlashOk "Playlist “" & pfTitle & "” updated."
            Go "playlist_detail.asp?id=" & pfId
        Else
            pfNewId = ToInt(DbScalar( _
                "INSERT INTO TBL_PLAYLIST (ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO) " & _
                "OUTPUT INSERTED.ID_PLAYLIST VALUES (?, ?, ?, ?, ?)", _
                Array(CurrentUserId(), pfTitle, SqlDate(Now()), pfType, pfPhoto)), 0)
            FlashOk "Playlist “" & pfTitle & "” created. Add songs from the playlist page."
            Go "playlist_detail.asp?id=" & pfNewId
        End If
    End If
End If

PageTitle = IIf(pfIsEdit, "Edit playlist", "New playlist")
PageEyebrow = "Playlists"
ActiveNav = "playlists"
%>
<!--#include file="includes/layout_top.asp"-->

      <div class="page-head">
        <div>
          <p class="page-head__eyebrow"><%= IIf(pfIsEdit, "Edit record", "Create record") %></p>
          <h1><%= IIf(pfIsEdit, "Edit playlist", "New playlist") %></h1>
          <p class="page-head__sub">Playlists belong to the user who created them. Public playlists are visible to everyone; private ones only to their owner (and to administrators).</p>
        </div>
        <div class="page-head__actions">
          <a class="btn btn--ghost" href="playlists.asp"><%= Icon("arrow-left") %> Back to playlists</a>
        </div>
      </div>

      <%= ErrorBox(pfErr) %>

      <div class="panel form-card animate-in">
        <form method="post" action="playlist_form.asp" novalidate>
          <%= CsrfField() %>
          <input type="hidden" name="id" value="<%= IIf(pfIsEdit, pfId, 0) %>">

          <div class="form-grid">
            <div class="field field--full">
              <label for="title">Title</label>
              <input id="title" name="title" type="text" maxlength="50" required
                     value="<%= H(pfTitle) %>" placeholder="e.g. Late night terracotta">
            </div>

            <div class="field">
              <label for="type">Visibility</label>
              <select id="type" name="type">
                <option value="<%= PLAYLIST_PUBLIC %>"<%= SelIf(PLAYLIST_PUBLIC, pfType) %>>Public — everyone can listen</option>
                <option value="<%= PLAYLIST_PRIVATE %>"<%= SelIf(PLAYLIST_PRIVATE, pfType) %>>Private — only me</option>
              </select>
            </div>

            <div class="field">
              <label for="photo">Cover image path</label>
              <input id="photo" name="photo" type="text" maxlength="255"
                     value="<%= H(pfPhoto) %>" placeholder="~/playlistImages/playlist1.png">
              <span class="field__hint">Available: default.jpg, playlist1.png … playlist5.png</span>
            </div>
          </div>

          <div class="form-actions">
            <a class="btn btn--ghost" href="playlists.asp">Cancel</a>
            <button class="btn btn--primary" type="submit"><%= Icon("check") %> Save playlist</button>
          </div>
        </form>
      </div>

<!--#include file="includes/layout_bottom.asp"-->
