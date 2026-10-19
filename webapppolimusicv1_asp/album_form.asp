<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  album_form.asp — create / edit an album (administrators only)
' ============================================================================
Dim bfId, bfTitle, bfYear, bfImage, bfErr, bfIsEdit, bfRs, bfNewId, bfYearRaw

RequireAdmin

bfId = QsInt("id", 0)
bfErr = ""
bfIsEdit = (bfId > 0)
bfTitle = ""
bfYear = ""
bfImage = "~/albumImages/album1.png"

If bfIsEdit Then
    Set bfRs = DbQuery("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(bfId))
    If bfRs.EOF Then
        bfRs.Close
        Set bfRs = Nothing
        FlashErr "That album does not exist."
        Go "albums.asp"
    End If
    bfTitle = Txt(bfRs("TITLE_ALBUM"))
    bfYear = Txt(bfRs("RELEASE_YEAR"))
    bfImage = Txt(bfRs("ALBUM_IMAGE_PATH"))
    bfRs.Close
    Set bfRs = Nothing
End If

If IsPost() Then
    RequireCsrf
    bfId = FormInt("id", 0)
    bfIsEdit = (bfId > 0)
    bfTitle = FormStr("title_album")
    bfYearRaw = FormStr("release_year")
    bfImage = FormStr("album_image")
    If bfImage = "" Then bfImage = "~/albumImages/album1.png"
    bfYear = bfYearRaw

    If bfTitle = "" Then bfErr = bfErr & "The album title is required." & vbLf
    If Len(bfTitle) > 50 Then bfErr = bfErr & "The title cannot exceed 50 characters." & vbLf
    If Len(bfImage) > 255 Then bfErr = bfErr & "The image path cannot exceed 255 characters." & vbLf
    If bfYearRaw <> "" Then
        If Not IsNumeric(bfYearRaw) Then
            bfErr = bfErr & "The release year must be a number." & vbLf
        ElseIf CLng(bfYearRaw) < 1800 Or CLng(bfYearRaw) > 2100 Then
            bfErr = bfErr & "The release year must be between 1800 and 2100." & vbLf
        End If
    End If

    If bfErr = "" Then
        If bfYearRaw = "" Then
            bfYear = Null
        Else
            bfYear = CLng(bfYearRaw)
        End If

        If bfIsEdit Then
            DbExec "UPDATE TBL_ALBUM SET TITLE_ALBUM = ?, RELEASE_YEAR = ?, ALBUM_IMAGE_PATH = ? WHERE ID_ALBUM = ?", _
                   Array(bfTitle, bfYear, bfImage, bfId)
            FlashOk "Album “" & bfTitle & "” updated."
            Go "album_detail.asp?id=" & bfId
        Else
            bfNewId = ToInt(DbScalar( _
                "INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) " & _
                "OUTPUT INSERTED.ID_ALBUM VALUES (?, ?, ?)", _
                Array(bfTitle, bfYear, bfImage)), 0)
            FlashOk "Album “" & bfTitle & "” created. Add songs from the song editor."
            Go "album_detail.asp?id=" & bfNewId
        End If
    End If
End If

PageTitle = IIf(bfIsEdit, "Edit album", "New album")
PageEyebrow = "Catalogue"
ActiveNav = "albums"
%>
<!--#include file="includes/layout_top.asp"-->

      <div class="page-head">
        <div>
          <p class="page-head__eyebrow"><%= IIf(bfIsEdit, "Edit record", "Create record") %></p>
          <h1><%= IIf(bfIsEdit, "Edit album", "New album") %></h1>
          <p class="page-head__sub">Albums live in <span class="mono">TBL_ALBUM</span>; their tracklist is built from <span class="mono">TBL_TRACKLIST</span> when songs are assigned to them.</p>
        </div>
        <div class="page-head__actions">
          <a class="btn btn--ghost" href="albums.asp"><%= Icon("arrow-left") %> Back to albums</a>
        </div>
      </div>

      <%= ErrorBox(bfErr) %>

      <div class="panel form-card animate-in">
        <form method="post" action="album_form.asp" novalidate>
          <%= CsrfField() %>
          <input type="hidden" name="id" value="<%= IIf(bfIsEdit, bfId, 0) %>">

          <div class="form-grid">
            <div class="field field--full">
              <label for="title_album">Album title</label>
              <input id="title_album" name="title_album" type="text" maxlength="50" required
                     value="<%= H(bfTitle) %>" placeholder="e.g. Rock Album">
            </div>

            <div class="field">
              <label for="release_year">Release year</label>
              <input id="release_year" name="release_year" type="number" min="1800" max="2100" step="1"
                     value="<%= H(bfYear) %>" placeholder="1976">
            </div>

            <div class="field">
              <label for="album_image">Cover path</label>
              <input id="album_image" name="album_image" type="text" maxlength="255"
                     value="<%= H(bfImage) %>" placeholder="~/albumImages/album1.png">
              <span class="field__hint">Available: album1.png … album5.png inside <span class="mono">albumImages</span>.</span>
            </div>
          </div>

          <div class="form-actions">
            <% If bfIsEdit Then %>
            <button class="btn btn--danger" type="submit" form="bf-delete-form"><%= Icon("trash") %> Delete</button>
            <% End If %>
            <a class="btn btn--ghost" href="albums.asp">Cancel</a>
            <button class="btn btn--primary" type="submit"><%= Icon("check") %> Save album</button>
          </div>
        </form>

        <% If bfIsEdit Then %>
        <form id="bf-delete-form" method="post" action="delete.asp"
              data-confirm="Delete album “<%= H(bfTitle) %>”? Its tracklist links will be removed.">
          <%= CsrfField() %>
          <input type="hidden" name="entity" value="album">
          <input type="hidden" name="id" value="<%= bfId %>">
          <input type="hidden" name="return" value="albums.asp">
        </form>
        <% End If %>
      </div>

<!--#include file="includes/layout_bottom.asp"-->
