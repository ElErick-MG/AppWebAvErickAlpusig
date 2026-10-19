<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  delete.asp — single, guarded deletion endpoint (POST only)
'  Global entities require ROLE_ADMIN; playlists require ownership or admin.
' ============================================================================
Dim dlEntity, dlId, dlSongId, dlBack, dlErr, dlRows, dlOwner, dlName

RequireLogin

If Not IsPost() Then Go "default.asp"
RequireCsrf

dlEntity = LCase(FormStr("entity"))
dlId = FormInt("id", 0)
dlSongId = FormInt("song", 0)
dlBack = FormStr("return")
dlErr = ""

' Only allow same-application relative redirects.
If dlBack = "" Or InStr(dlBack, "://") > 0 Or InStr(dlBack, "..") > 0 Then
    dlBack = "default.asp"
End If

If dlId <= 0 Then
    FlashErr "Nothing was deleted: the record identifier is missing."
    Go dlBack
End If

Select Case dlEntity

    Case "song"
        RequireAdmin
        dlName = Txt(DbScalar("SELECT SONG_NAME FROM TBL_SONG WHERE ID_SONG = ?", Array(dlId)))
        DbExec "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array(dlId)
        DbExec "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(dlId)
        DbExec "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = ?", Array(dlId)
        dlRows = DbExecSafe("DELETE FROM TBL_SONG WHERE ID_SONG = ?", Array(dlId), dlErr)
        If dlErr = "" Then
            FlashOk "Song “" & dlName & "” deleted, together with its artist, album and playlist links."
        Else
            FlashErr "The song could not be deleted: " & dlErr
        End If

    Case "artist"
        RequireAdmin
        dlName = Txt(DbScalar("SELECT ARTIST_NAME + ' ' + ISNULL(ARTIST_LASTNAME, '') FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(dlId)))
        DbExec "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = ?", Array(dlId)
        dlRows = DbExecSafe("DELETE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(dlId), dlErr)
        If dlErr = "" Then
            FlashOk "Artist “" & Trim(dlName) & "” deleted."
        Else
            FlashErr "The artist could not be deleted: " & dlErr
        End If

    Case "album"
        RequireAdmin
        dlName = Txt(DbScalar("SELECT TITLE_ALBUM FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(dlId)))
        DbExec "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array(dlId)
        dlRows = DbExecSafe("DELETE FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(dlId), dlErr)
        If dlErr = "" Then
            FlashOk "Album “" & dlName & "” deleted."
        Else
            FlashErr "The album could not be deleted: " & dlErr
        End If

    Case "user"
        RequireAdmin
        If dlId = CurrentUserId() Then
            FlashErr "You cannot delete the account you are signed in with."
            Go dlBack
        End If
        dlName = Txt(DbScalar("SELECT USERNAME FROM TBL_USER WHERE ID_USER = ?", Array(dlId)))
        DbExec "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST IN (SELECT ID_PLAYLIST FROM TBL_PLAYLIST WHERE ID_USER = ?)", Array(dlId)
        DbExec "DELETE FROM TBL_PLAYLIST WHERE ID_USER = ?", Array(dlId)
        dlRows = DbExecSafe("DELETE FROM TBL_USER WHERE ID_USER = ?", Array(dlId), dlErr)
        If dlErr = "" Then
            FlashOk "User “" & dlName & "” and their playlists were deleted."
        Else
            FlashErr "The user could not be deleted: " & dlErr
        End If

    Case "playlist"
        dlOwner = ToInt(DbScalar("SELECT ID_USER FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(dlId)), -1)
        If Not CanManagePlaylist(dlOwner) Then
            FlashErr "You can only delete your own playlists."
            Go dlBack
        End If
        dlName = Txt(DbScalar("SELECT TITLE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(dlId)))
        DbExec "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?", Array(dlId)
        dlRows = DbExecSafe("DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(dlId), dlErr)
        If dlErr = "" Then
            FlashOk "Playlist “" & dlName & "” deleted."
            If InStr(dlBack, "playlist_detail.asp") > 0 Then dlBack = "playlists.asp"
        Else
            FlashErr "The playlist could not be deleted: " & dlErr
        End If

    Case "playlist_song"
        dlOwner = ToInt(DbScalar("SELECT ID_USER FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(dlId)), -1)
        If Not CanManagePlaylist(dlOwner) Then
            FlashErr "You can only edit your own playlists."
            Go dlBack
        End If
        dlRows = DbExecSafe("DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", Array(dlId, dlSongId), dlErr)
        If dlErr = "" Then
            FlashOk "Song removed from the playlist."
        Else
            FlashErr "The song could not be removed: " & dlErr
        End If

    Case Else
        FlashErr "Unknown entity. Nothing was deleted."
End Select

Go dlBack
%>
