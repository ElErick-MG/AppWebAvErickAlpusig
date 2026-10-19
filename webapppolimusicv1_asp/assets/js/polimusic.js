/* ==========================================================================
   PoliMusic — client behaviour
   1. Responsive sidebar        4. Play counters
   2. Custom player bar         5. Toasts / confirmations
   3. Table filtering           6. Keyboard shortcuts
   ========================================================================== */
(function () {
  "use strict";

  var doc = document;
  var csrf = (doc.querySelector('meta[name="csrf-token"]') || {}).content || "";

  /* ---------------- helpers ---------------- */
  function $(sel, root) {
    return (root || doc).querySelector(sel);
  }

  function $$(sel, root) {
    return Array.prototype.slice.call((root || doc).querySelectorAll(sel));
  }

  function fmtTime(sec) {
    if (!isFinite(sec) || sec < 0) return "0:00";
    var m = Math.floor(sec / 60);
    var s = Math.floor(sec % 60);
    return m + ":" + (s < 10 ? "0" : "") + s;
  }

  function toast(message, kind) {
    var host = $("#pm-toasts");
    if (!host) return;
    var el = doc.createElement("div");
    el.className = "toast toast--" + (kind || "ok");
    el.innerHTML = '<span aria-hidden="true">' + (kind === "err" ? "!" : "✓") + "</span><span></span>";
    el.lastChild.textContent = message;
    host.appendChild(el);
    setTimeout(function () {
      el.classList.add("is-out");
      setTimeout(function () {
        el.remove();
      }, 260);
    }, 3200);
  }
  window.PM = window.PM || {};
  window.PM.toast = toast;

  /* ---------------- 1. sidebar ---------------- */
  var sidebar = $("#sidebar");
  var backdrop = $(".sidebar-backdrop");

  function closeSidebar() {
    if (sidebar) sidebar.classList.remove("is-open");
    if (backdrop) backdrop.classList.remove("is-open");
  }

  $$("[data-sidebar-toggle]").forEach(function (btn) {
    btn.addEventListener("click", function () {
      if (!sidebar) return;
      sidebar.classList.toggle("is-open");
      if (backdrop) backdrop.classList.toggle("is-open");
    });
  });

  if (backdrop) backdrop.addEventListener("click", closeSidebar);
  doc.addEventListener("keydown", function (e) {
    if (e.key === "Escape") closeSidebar();
  });

  /* ---------------- 2. player bar ---------------- */
  var audios = $$("audio[data-song-id]");
  var player = $("#pm-player");
  var elArt = $("#pm-art");
  var elTitle = $("#pm-title");
  var elArtist = $("#pm-artist");
  var elSeek = $("#pm-seek");
  var elCur = $("#pm-current");
  var elDur = $("#pm-duration");
  var elVol = $("#pm-volume");
  var elBadge = $("#pm-badge");
  var iconPlay = $("#pm-icon-play");
  var iconPause = $("#pm-icon-pause");
  var active = null;
  var activeIndex = -1;
  var counted = {};
  var volume = 0.9;
  var seeking = false;

  function rowOf(audio) {
    return audio.closest("tr, .track-row");
  }

  function setBadge(text) {
    if (elBadge) elBadge.innerHTML = elBadge.innerHTML.replace(/<b>.*<\/b>/, "<b>" + text + "</b>");
  }

  function setPlayingUI(isPlaying) {
    if (!iconPlay || !iconPause) return;
    iconPlay.classList.toggle("is-hidden", isPlaying);
    iconPause.classList.toggle("is-hidden", !isPlaying);
    setBadge(isPlaying ? "Playing" : active ? "Paused" : "Idle");
  }

  function markRows() {
    audios.forEach(function (a) {
      var tr = rowOf(a);
      if (tr) tr.classList.toggle("is-playing", a === active && !a.paused);
    });
  }

  function activate(audio) {
    if (active === audio) return;
    if (active && !active.paused) active.pause();
    active = audio;
    activeIndex = audios.indexOf(audio);
    active.volume = volume;

    var tr = rowOf(audio);
    var name = audio.getAttribute("data-title") || "Unknown";
    var sub = tr ? $(".cell-song__sub", tr) : null;
    var art = tr ? $(".cell-song__art", tr) : null;

    if (elTitle) elTitle.textContent = name;
    if (elArtist) elArtist.textContent = sub ? sub.textContent.trim() : "PoliMusic";
    if (elArt) elArt.src = art ? art.src : "assets/img/art-placeholder.svg";
    if (elDur) elDur.textContent = fmtTime(audio.duration);
    markRows();
  }

  function countPlay(audio) {
    var id = audio.getAttribute("data-song-id");
    if (!id || counted[id]) return;
    counted[id] = true;

    var body = "_csrf=" + encodeURIComponent(csrf) + "&id=" + encodeURIComponent(id);
    fetch("api_play.asp", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body,
      credentials: "same-origin"
    })
      .then(function (r) {
        return r.ok ? r.json() : null;
      })
      .then(function (data) {
        if (!data || !data.ok) return;
        var cell = doc.querySelector('[data-plays-for="' + id + '"]');
        if (cell) {
          cell.textContent = data.plays;
          cell.classList.add("is-live");
        }
      })
      .catch(function () {
        /* counter is non critical */
      });
  }

  audios.forEach(function (audio, index) {
    audio.volume = volume;

    audio.addEventListener("play", function () {
      activate(audio);
      setPlayingUI(true);
      markRows();
      countPlay(audio);
      var tr = rowOf(audio);
      if (tr && tr.getBoundingClientRect().top < 0) {
        tr.scrollIntoView({ behavior: "smooth", block: "center" });
      }
    });

    audio.addEventListener("pause", function () {
      if (audio === active) {
        setPlayingUI(false);
        markRows();
      }
    });

    audio.addEventListener("ended", function () {
      var next = audios[index + 1];
      if (next) {
        next.play();
      } else {
        setPlayingUI(false);
        markRows();
      }
    });

    audio.addEventListener("timeupdate", function () {
      if (audio !== active) return;
      if (elCur) elCur.textContent = fmtTime(audio.currentTime);
      if (elDur) elDur.textContent = fmtTime(audio.duration);
      if (elSeek && !seeking && audio.duration) {
        elSeek.value = String(Math.round((audio.currentTime / audio.duration) * 1000));
      }
    });

    audio.addEventListener("loadedmetadata", function () {
      if (audio === active && elDur) elDur.textContent = fmtTime(audio.duration);
    });

    audio.addEventListener("error", function () {
      if (audio === active) {
        toast("Audio file not available for “" + (audio.getAttribute("data-title") || "this song") + "”.", "err");
        setPlayingUI(false);
      }
    });
  });

  /* custom transport buttons */
  var btnToggle = $("#pm-toggle");
  if (btnToggle) {
    btnToggle.addEventListener("click", function () {
      if (!active) {
        if (audios.length) {
          audios[0].play();
        } else {
          toast("No songs available on this page.", "err");
        }
        return;
      }
      if (active.paused) active.play();
      else active.pause();
    });
  }

  function step(delta) {
    if (!audios.length) return;
    var next = activeIndex < 0 ? 0 : activeIndex + delta;
    if (next < 0) next = audios.length - 1;
    if (next >= audios.length) next = 0;
    audios[next].play();
  }

  var btnPrev = $("#pm-prev");
  var btnNext = $("#pm-next");
  if (btnPrev) btnPrev.addEventListener("click", function () { step(-1); });
  if (btnNext) btnNext.addEventListener("click", function () { step(1); });

  if (elSeek) {
    elSeek.addEventListener("input", function () {
      seeking = true;
      if (active && active.duration) {
        elCur.textContent = fmtTime((elSeek.value / 1000) * active.duration);
      }
    });
    elSeek.addEventListener("change", function () {
      if (active && active.duration) {
        active.currentTime = (elSeek.value / 1000) * active.duration;
      }
      seeking = false;
    });
  }

  if (elVol) {
    elVol.addEventListener("input", function () {
      volume = Number(elVol.value) / 100;
      if (active) active.volume = volume;
      audios.forEach(function (a) {
        if (a !== active) a.volume = volume;
      });
    });
  }

  var btnMute = $("#pm-mute");
  if (btnMute) {
    btnMute.addEventListener("click", function () {
      var target = active || audios[0];
      if (!target) return;
      target.muted = !target.muted;
      btnMute.style.color = target.muted ? "var(--tomato-bright)" : "";
      setBadge(target.muted ? "Muted" : active && !active.paused ? "Playing" : "Paused");
    });
  }

  /* clicking a row play button drives its own audio element */
  $$("[data-play-song]").forEach(function (btn) {
    btn.addEventListener("click", function () {
      var id = btn.getAttribute("data-play-song");
      var audio = doc.querySelector('audio[data-song-id="' + id + '"]');
      if (!audio) return;
      if (audio.paused) audio.play();
      else audio.pause();
    });
  });

  /* ---------------- 3. table filtering ---------------- */
  var searchInput = $('.search input[name="q"]');
  var table = $("#songs-table");
  if (searchInput && table) {
    var rows = $$("tbody tr", table);
    var emptyRow = $("#songs-empty-filter");
    searchInput.addEventListener("input", function () {
      var term = searchInput.value.trim().toLowerCase();
      var visible = 0;
      rows.forEach(function (tr) {
        var hit = !term || tr.textContent.toLowerCase().indexOf(term) !== -1;
        tr.classList.toggle("is-hidden", !hit);
        if (hit) visible++;
      });
      if (emptyRow) emptyRow.classList.toggle("is-hidden", visible > 0);
    });
  }

  /* ---------------- 4. confirmations ---------------- */
  $$("form[data-confirm]").forEach(function (form) {
    form.addEventListener("submit", function (e) {
      if (!window.confirm(form.getAttribute("data-confirm"))) {
        e.preventDefault();
      }
    });
  });

  /* ---------------- 5. keyboard ---------------- */
  doc.addEventListener("keydown", function (e) {
    var tag = (e.target.tagName || "").toLowerCase();
    if (tag === "input" || tag === "textarea" || tag === "select" || e.target.isContentEditable) return;

    if (e.code === "Space" && active) {
      e.preventDefault();
      if (active.paused) active.play();
      else active.pause();
    }
    if (e.key === "/" && searchInput) {
      e.preventDefault();
      searchInput.focus();
    }
  });

  /* ---------------- 6. entrance animation ---------------- */
  $$(".animate-in").forEach(function (el, i) {
    el.style.animationDelay = Math.min(i * 40, 260) + "ms";
  });
})();
