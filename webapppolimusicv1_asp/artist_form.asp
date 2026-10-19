<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  artist_form.asp — create / edit an artist (administrators only)
' ============================================================================
Dim afId, afName, afLast, afImage, afErr, afIsEdit, afRs, afNewId

RequireAdmin

afId = QsInt("id", 0)
afErr = ""
afIsEdit = (afId > 0)
afName = ""
afLast = ""
afImage = "~/ArtistImages/artist1.jpg"

If afIsEdit Then
    Set afRs = DbQuery("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(afId))
    If afRs.EOF Then
        afRs.Close
        Set afRs = Nothing
        FlashErr "That artist does not exist."
        Go "artists.asp"
    End If
    afName = Txt(afRs("ARTIST_NAME"))
    afLast = Txt(afRs("ARTIST_LASTNAME"))
    afImage = Txt(afRs("ARTIST_IMAGE"))
    afRs.Close
    Set afRs = Nothing
End If

If IsPost() Then
    RequireCsrf
    afId = FormInt("id", 0)
    afIsEdit = (afId > 0)
    afName = FormStr("artist_name")
    afLast = FormStr("artist_lastname")
    afImage = FormStr("artist_image")
    If afImage = "" Then afImage = "~/ArtistImages/artist1.jpg"

    If afName = "" Then afErr = afErr & "The artist name is required." & vbLf
    If Len(afName) > 100 Then afErr = afErr & "The name cannot exceed 100 characters." & vbLf
    If Len(afLast) > 100 Then afErr = afErr & "The last name cannot exceed 100 characters." & vbLf
    If Len(afImage) > 255 Then afErr = afErr & "The image path cannot exceed 255 characters." & vbLf

    If afErr = "" Then
        If afIsEdit Then
            DbExec "UPDATE TBL_ARTIST SET ARTIST_NAME = ?, ARTIST_LASTNAME = ?, ARTIST_IMAGE = ? WHERE ID_ARTIST = ?", _
                   Array(afName, afLast, afImage, afId)
            FlashOk "Artist “" & afName & "” updated."
        Else
            afNewId = ToInt(DbScalar( _
                "INSERT INTO TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) " & _
                "OUTPUT INSERTED.ID_ARTIST VALUES (?, ?, ?)", _
                Array(afName, afLast, afImage)), 0)
            FlashOk "Artist “" & afName & "” created. Link songs to them from the song editor."
        End If
        Go "artists.asp"
    End If
End If

PageTitle = IIf(afIsEdit, "Edit artist", "New artist")
PageEyebrow = "Catalogue"
ActiveNav = "artists"
%>
<!--#include file="includes/layout_top.asp"-->

      <div class="page-head">
        <div>
          <p class="page-head__eyebrow"><%= IIf(afIsEdit, "Edit record", "Create record") %></p>
          <h1><%= IIf(afIsEdit, "Edit artist", "New artist") %></h1>
          <p class="page-head__sub">Artists are stored in <span class="mono">TBL_ARTIST</span> and linked to songs through <span class="mono">TBL_PLAYER</span>.</p>
        </div>
        <div class="page-head__actions">
          <a class="btn btn--ghost" href="artists.asp"><%= Icon("arrow-left") %> Back to artists</a>
        </div>
      </div>

      <%= ErrorBox(afErr) %>

      <div class="panel form-card animate-in">
        <form method="post" action="artist_form.asp" novalidate>
          <%= CsrfField() %>
          <input type="hidden" name="id" value="<%= IIf(afIsEdit, afId, 0) %>">

          <div class="form-grid">
            <div class="field">
              <label for="artist_name">First name / stage name</label>
              <input id="artist_name" name="artist_name" type="text" maxlength="100" required
                     value="<%= H(afName) %>" placeholder="e.g. Juan">
            </div>

            <div class="field">
              <label for="artist_lastname">Last name</label>
              <input id="artist_lastname" name="artist_lastname" type="text" maxlength="100"
                     value="<%= H(afLast) %>" placeholder="e.g. Gabriel">
            </div>

            <div class="field field--full">
              <label for="artist_image">Photo path</label>
              <input id="artist_image" name="artist_image" type="text" maxlength="255"
                     value="<%= H(afImage) %>" placeholder="~/ArtistImages/artist1.jpg">
              <span class="field__hint">Available: artist1.jpg … artist5.png inside <span class="mono">ArtistImages</span>.</span>
            </div>
          </div>

          <div class="form-actions">
            <% If afIsEdit Then %>
            <button class="btn btn--danger" type="submit" form="af-delete-form"><%= Icon("trash") %> Delete</button>
            <% End If %>
            <a class="btn btn--ghost" href="artists.asp">Cancel</a>
            <button class="btn btn--primary" type="submit"><%= Icon("check") %> Save artist</button>
          </div>
        </form>

        <% If afIsEdit Then %>
        <form id="af-delete-form" method="post" action="delete.asp"
              data-confirm="Delete artist “<%= H(afName) %>” and all of their song links?">
          <%= CsrfField() %>
          <input type="hidden" name="entity" value="artist">
          <input type="hidden" name="id" value="<%= afId %>">
          <input type="hidden" name="return" value="artists.asp">
        </form>
        <% End If %>
      </div>

<!--#include file="includes/layout_bottom.asp"-->
