<%
' ============================================================================
'  PoliMusic — Layout (closing): player bar, scripts, connection cleanup
' ============================================================================
If IsObject(lyPls) Then
    If Not (lyPls Is Nothing) Then
        lyPls.Close
        Set lyPls = Nothing
    End If
End If
%>
    </main>
  </div>
</div>

<!-- ---------- Persistent player bar ---------- -->
<div class="player" id="pm-player" aria-label="Now playing">
  <div class="player__now">
    <img class="player__art" id="pm-art" src="assets/img/art-placeholder.svg" alt="">
    <div class="player__meta">
      <div class="player__title" id="pm-title">Nothing playing</div>
      <div class="player__artist" id="pm-artist">Pick a song to start</div>
    </div>
  </div>

  <div class="player__center">
    <div class="player__controls">
      <button class="player__btn" type="button" id="pm-prev" aria-label="Previous song"><%= Icon("prev") %></button>
      <button class="player__btn player__btn--main" type="button" id="pm-toggle" aria-label="Play or pause">
        <span id="pm-icon-play"><%= Icon("play") %></span>
        <span id="pm-icon-pause" class="is-hidden"><%= Icon("pause") %></span>
      </button>
      <button class="player__btn" type="button" id="pm-next" aria-label="Next song"><%= Icon("next") %></button>
    </div>
    <div class="player__timeline">
      <span class="player__time" id="pm-current">0:00</span>
      <input class="range" type="range" id="pm-seek" min="0" max="1000" value="0" step="1" aria-label="Seek">
      <span class="player__time" id="pm-duration">0:00</span>
    </div>
  </div>

  <div class="player__right">
    <span class="pill" id="pm-badge"><%= Icon("waveform") %> <b>Idle</b></span>
    <button class="player__btn" type="button" id="pm-mute" aria-label="Mute"><%= Icon("volume") %></button>
    <input class="range player__vol" type="range" id="pm-volume" min="0" max="100" value="90" aria-label="Volume">
  </div>
</div>

<div class="toasts" id="pm-toasts" aria-live="polite"></div>

<script src="assets/js/polimusic.js"></script>
<%
CloseDb
%>
</body>
</html>
