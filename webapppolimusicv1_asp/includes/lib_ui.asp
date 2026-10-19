<%
' ============================================================================
'  PoliMusic — UI building blocks (inline SVG icon set + small renderers)
' ============================================================================

Function IconPath(ByVal name)
    Select Case LCase(Txt(name))
        Case "home"
            IconPath = "<path d='M3 10.6 12 3.2l9 7.4'/><path d='M5.4 9.4V20.8h13.2V9.4'/>"
        Case "music"
            IconPath = "<path d='M9 17.5V5.2l10-2v12.3'/><circle cx='6.2' cy='17.6' r='2.8'/><circle cx='16.2' cy='15.6' r='2.8'/>"
        Case "disc"
            IconPath = "<circle cx='12' cy='12' r='9'/><circle cx='12' cy='12' r='2.6'/>"
        Case "mic"
            IconPath = "<rect x='9' y='2.5' width='6' height='11' rx='3'/><path d='M5.2 11.4a6.8 6.8 0 0 0 13.6 0'/><path d='M12 18.2v3.3'/>"
        Case "list"
            IconPath = "<path d='M4 6.5h16M4 12h16M4 17.5h10'/>"
        Case "users"
            IconPath = "<path d='M15.5 20.5v-1.8a3.7 3.7 0 0 0-3.7-3.7H6.2a3.7 3.7 0 0 0-3.7 3.7v1.8'/><circle cx='9' cy='7.4' r='3.6'/><path d='M21.5 20.5v-1.8a3.7 3.7 0 0 0-2.8-3.6'/><path d='M15.6 3.9a3.6 3.6 0 0 1 0 7'/>"
        Case "logout"
            IconPath = "<path d='M9.5 20.5H5.2A2.2 2.2 0 0 1 3 18.3V5.7a2.2 2.2 0 0 1 2.2-2.2h4.3'/><path d='M16 16.4l4.4-4.4L16 7.6'/><path d='M20.4 12H9.4'/>"
        Case "search"
            IconPath = "<circle cx='11' cy='11' r='7'/><path d='M20.5 20.5l-4.2-4.2'/>"
        Case "play"
            IconPath = "<path d='M8 5.2v13.6L19 12z' fill='currentColor' stroke='none'/>"
        Case "pause"
            IconPath = "<path d='M7.2 4.6h3.4v14.8H7.2zM13.4 4.6h3.4v14.8h-3.4z' fill='currentColor' stroke='none'/>"
        Case "prev"
            IconPath = "<path d='M18.5 5.4v13.2L9 12z' fill='currentColor' stroke='none'/><path d='M6 5.2v13.6' stroke-width='2.4'/>"
        Case "next"
            IconPath = "<path d='M5.5 5.4v13.2L15 12z' fill='currentColor' stroke='none'/><path d='M18 5.2v13.6' stroke-width='2.4'/>"
        Case "edit"
            IconPath = "<path d='M12.5 20.5H21'/><path d='M16.6 3.7a2.2 2.2 0 0 1 3.1 3.1L7.4 19.1l-4.1 1 1-4.1z'/>"
        Case "trash"
            IconPath = "<path d='M3.5 6.5h17'/><path d='M8.5 6.5V4.2h7v2.3'/><path d='M18.6 6.5l-.9 14H6.3l-.9-14'/><path d='M10.2 11v6M13.8 11v6'/>"
        Case "plus"
            IconPath = "<path d='M12 5v14M5 12h14'/>"
        Case "volume"
            IconPath = "<path d='M11 4.8 6.4 8.9H2.6v6.2h3.8L11 19.2z'/><path d='M15 8.6a4.8 4.8 0 0 1 0 6.8'/><path d='M17.8 5.8a8.8 8.8 0 0 1 0 12.4'/>"
        Case "headphones"
            IconPath = "<path d='M3.8 15.2v-3.1a8.2 8.2 0 0 1 16.4 0v3.1'/><rect x='1.8' y='14' width='4.2' height='7' rx='2.1'/><rect x='18' y='14' width='4.2' height='7' rx='2.1'/>"
        Case "check"
            IconPath = "<path d='M20 6.4 9.4 17 4 11.6'/>"
        Case "alert"
            IconPath = "<circle cx='12' cy='12' r='9'/><path d='M12 7.6v5.2M12 16.3v.4'/>"
        Case "info"
            IconPath = "<circle cx='12' cy='12' r='9'/><path d='M12 11v5.2M12 7.7v.4'/>"
        Case "arrow-left"
            IconPath = "<path d='M19 12H5'/><path d='M11.5 18.5 5 12l6.5-6.5'/>"
        Case "arrow-right"
            IconPath = "<path d='M5 12h14'/><path d='M12.5 5.5 19 12l-6.5 6.5'/>"
        Case "lock"
            IconPath = "<rect x='4.2' y='10.2' width='15.6' height='10.4' rx='2.4'/><path d='M8.2 10.2V7.4a3.8 3.8 0 0 1 7.6 0v2.8'/>"
        Case "mail"
            IconPath = "<rect x='2.8' y='5' width='18.4' height='14' rx='2.4'/><path d='M3.4 7.2 12 13l8.6-5.8'/>"
        Case "user"
            IconPath = "<circle cx='12' cy='8' r='4'/><path d='M4.2 20.8a7.8 7.8 0 0 1 15.6 0'/>"
        Case "calendar"
            IconPath = "<rect x='3.2' y='5.2' width='17.6' height='15.6' rx='2.4'/><path d='M8 3.2v4M16 3.2v4M3.2 10.8h17.6'/>"
        Case "image"
            IconPath = "<rect x='3.2' y='4.2' width='17.6' height='15.6' rx='2.4'/><circle cx='8.6' cy='9.6' r='1.6'/><path d='M20.8 15.2 15.6 10 5 20.6'/>"
        Case "x"
            IconPath = "<path d='M18 6 6 18M6 6l12 12'/>"
        Case "menu"
            IconPath = "<path d='M3.5 6.5h17M3.5 12h17M3.5 17.5h17'/>"
        Case "folder"
            IconPath = "<path d='M3 7.2a2 2 0 0 1 2-2h3.6l2 2.2H19a2 2 0 0 1 2 2v8.4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z'/>"
        Case "shield"
            IconPath = "<path d='M12 3 20 6v6.1c0 4.8-3.3 7.8-8 8.9-4.7-1.1-8-4.1-8-8.9V6z'/><path d='M9 12.1l2.1 2.1 4-4'/>"
        Case "globe"
            IconPath = "<circle cx='12' cy='12' r='9'/><path d='M3.2 12h17.6'/><path d='M12 3a14.5 14.5 0 0 1 0 18a14.5 14.5 0 0 1 0-18'/>"
        Case "clock"
            IconPath = "<circle cx='12' cy='12' r='9'/><path d='M12 7.2V12l3.2 2'/>"
        Case "waveform"
            IconPath = "<path d='M3 12h2.2M8 7.4v9.2M12 4.2v15.6M16 8.6v6.8M20.8 12H19'/>"
        Case "star"
            IconPath = "<path d='m12 3.6 2.6 5.4 5.9.8-4.3 4.2 1 5.9-5.2-2.8-5.2 2.8 1-5.9L3.5 9.8l5.9-.8z'/>"
        Case "external"
            IconPath = "<path d='M14 4h6v6'/><path d='M20 4 11 13'/><path d='M18 14.5V19a1.8 1.8 0 0 1-1.8 1.8H5.2A1.8 1.8 0 0 1 3.4 19V7.8A1.8 1.8 0 0 1 5.2 6h4.5'/>"
        Case Else
            IconPath = "<circle cx='12' cy='12' r='9'/>"
    End Select
End Function

' Two entry points instead of an optional parameter: VBScript does not allow
' default values, and this keeps every call unambiguous.
'   Icon("play")              -> plain icon
'   IconC("play", "ico-play") -> icon with a CSS class
Function Icon(ByVal name)
    Icon = IconHtml(name, "")
End Function

Function IconC(ByVal name, ByVal cls)
    IconC = IconHtml(name, cls)
End Function

Function IconHtml(ByVal name, ByVal cls)
    Dim extra
    extra = ""
    If Txt(cls) <> "" Then extra = " class=""" & H(cls) & """"
    IconHtml = "<svg viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""1.8"" " & _
               "stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"" focusable=""false""" & extra & ">" & _
               IconPath(name) & "</svg>"
End Function

Function RoleBadge(ByVal userType)
    If ToInt(userType, ROLE_USER) = ROLE_ADMIN Then
        RoleBadge = "<span class=""badge badge--admin"">" & Icon("shield") & "Admin</span>"
    Else
        RoleBadge = "<span class=""badge badge--user"">" & Icon("user") & "User</span>"
    End If
End Function

Function PlaylistTypeBadge(ByVal plType)
    If ToInt(plType, PLAYLIST_PUBLIC) = PLAYLIST_PRIVATE Then
        PlaylistTypeBadge = "<span class=""badge badge--muted"">" & Icon("lock") & "Private</span>"
    Else
        PlaylistTypeBadge = "<span class=""badge badge--user"">" & Icon("globe") & "Public</span>"
    End If
End Function

Function Avatar(ByVal photo, ByVal name, ByVal cls)
    Dim src
    src = AssetUrl(photo, "userImages/default.jpg")
    Avatar = "<img class=""" & H(cls) & """ src=""" & src & """ alt=""" & H(name) & _
             """ loading=""lazy"" onerror=""this.onerror=null;this.src='userImages/default.jpg';"">"
End Function

' HTML5 audio element used by every song row. data-* attributes feed the
' custom player bar; the native controls stay available as requested.
Function AudioPlayer(ByVal songId, ByVal path, ByVal songName)
    Dim url
    url = AssetUrl(path, "")
    If url = "" Then
        AudioPlayer = "<span class=""audio-missing"">" & Icon("folder") & "No audio path</span>"
        Exit Function
    End If
    AudioPlayer = "<audio controls preload=""none"" src=""" & url & _
                  """ data-song-id=""" & songId & """" & _
                  " data-title=""" & H(songName) & """></audio>"
End Function

Function AudioWarning(ByVal path)
    If AssetExists(path) Then
        AudioWarning = ""
    Else
        AudioWarning = "<span class=""audio-missing"" title=""The file referenced in TBL_SONG.SONG_PATH is not present on disk"">" & _
                       Icon("alert") & "File not deployed</span>"
    End If
End Function

Function EmptyState(ByVal iconName, ByVal title, ByVal text)
    EmptyState = "<div class=""empty"">" & Icon(iconName) & "<h3>" & H(title) & "</h3><p>" & H(text) & "</p></div>"
End Function
%>
