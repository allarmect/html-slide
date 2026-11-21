/* HTML nella slide - add-in di contenuto per PowerPoint.
   L'HTML scelto viene compresso (gzip) e salvato nelle impostazioni dell'add-in,
   che PowerPoint conserva DENTRO il file .pptx, una copia per ogni riquadro. */
(function () {
  "use strict";
  var CHUNK = 180000;          // caratteri per chiave di impostazione
  var $ = function (id) { return document.getElementById(id); };
  var settings = null;

  function mostraMsg(t) { $("msg").textContent = t || ""; }

  // ---------- compressione ----------
  function bytesToB64(bytes) {
    var s = "", step = 0x8000;
    for (var i = 0; i < bytes.length; i += step) s += String.fromCharCode.apply(null, bytes.subarray(i, i + step));
    return btoa(s);
  }
  function b64ToBytes(b64) {
    var s = atob(b64), out = new Uint8Array(s.length);
    for (var i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
    return out;
  }
  function gzip(text) {
    if (typeof CompressionStream === "undefined") return Promise.resolve(null);
    var stream = new Blob([text]).stream().pipeThrough(new CompressionStream("gzip"));
    return new Response(stream).arrayBuffer().then(function (b) { return bytesToB64(new Uint8Array(b)); });
  }
  function gunzip(b64) {
    var stream = new Blob([b64ToBytes(b64)]).stream().pipeThrough(new DecompressionStream("gzip"));
    return new Response(stream).text();
  }

  // ---------- salvataggio nel documento ----------
  function salva(rec) {
    var enc = rec.payload, n = Math.ceil(enc.length / CHUNK) || 1;
    var meta = { v: 1, tipo: rec.tipo, nome: rec.nome, gz: rec.gz, parti: n };
    var vecchie = (settings.get("meta") || {}).parti || 0;
    for (var i = 0; i < Math.max(n, vecchie); i++) {
      if (i < n) settings.set("p" + i, enc.slice(i * CHUNK, (i + 1) * CHUNK));
      else settings.remove("p" + i);
    }
    settings.set("meta", meta);
    return new Promise(function (ok, ko) {
      settings.saveAsync(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) ok();
        else ko(new Error(r.error && r.error.message || "salvataggio non riuscito"));
      });
    });
  }
  function leggi() {
    var meta = settings.get("meta");
    if (!meta) return Promise.resolve(null);
    var enc = "";
    for (var i = 0; i < meta.parti; i++) enc += settings.get("p" + i) || "";
    var p = meta.gz ? gunzip(enc) : Promise.resolve(enc);
    return p.then(function (contenuto) { return { tipo: meta.tipo, nome: meta.nome, contenuto: contenuto }; });
  }
  function cancella() {
    var meta = settings.get("meta");
    if (meta) for (var i = 0; i < meta.parti; i++) settings.remove("p" + i);
    settings.remove("meta");
    return new Promise(function (ok) { settings.saveAsync(function () { ok(); }); });
  }

  // ---------- visualizzazione ----------
  function mostra(rec) {
    var f = $("vista");
    f.removeAttribute("src"); f.removeAttribute("srcdoc");
    if (rec.tipo === "url") f.src = rec.contenuto;
    else f.srcdoc = rec.contenuto;
    $("nome").textContent = rec.nome || "";
    $("scelta").hidden = true;
    f.hidden = false;
    aggiornaBarra();
  }
  function mostraScelta() {
    $("vista").hidden = true;
    $("vista").removeAttribute("srcdoc"); $("vista").removeAttribute("src");
    $("barra").hidden = true;
    $("scelta").hidden = false;
  }

  var modoPresentazione = false;
  function aggiornaBarra() {
    $("barra").hidden = modoPresentazione || $("vista").hidden;
  }
  function osservaModalita() {
    var doc = Office.context.document;
    if (!doc.getActiveViewAsync) return;
    doc.getActiveViewAsync(function (r) { modoPresentazione = r.value === "read"; aggiornaBarra(); });
    if (doc.addHandlerAsync && Office.EventType.ActiveViewChanged) {
      doc.addHandlerAsync(Office.EventType.ActiveViewChanged, function (e) {
        modoPresentazione = e.activeView === "read"; aggiornaBarra();
      });
    }
  }

  // ---------- azioni ----------
  function caricaTesto(nome, testo, tipo) {
    mostraMsg("Salvataggio nella presentazione…");
    return gzip(testo).then(function (gz) {
      return salva({ tipo: tipo, nome: nome, gz: !!gz, payload: gz || testo });
    }).then(function () {
      mostraMsg("");
      mostra({ tipo: tipo, nome: nome, contenuto: testo });
    }).catch(function (e) {
      mostraMsg("Non sono riuscito a salvarlo nel file (" + e.message + "). Il file potrebbe essere troppo grande: oltre ~2 MB compressi PowerPoint lo rifiuta.");
    });
  }

  function init() {
    settings = Office.context.document.settings;
    osservaModalita();

    $("file").addEventListener("change", function () {
      var file = this.files && this.files[0];
      if (!file) return;
      var r = new FileReader();
      r.onload = function () { caricaTesto(file.name, String(r.result), "html"); };
      r.onerror = function () { mostraMsg("Impossibile leggere il file."); };
      r.readAsText(file, "utf-8");
      this.value = "";
    });
    $("formUrl").addEventListener("submit", function (e) {
      e.preventDefault();
      var u = $("url").value.trim();
      if (!/^https:\/\//i.test(u)) { mostraMsg("Serve un indirizzo che inizi con https://"); return; }
      caricaTesto(u, u, "url");
    });
    $("ricarica").addEventListener("click", function () {
      leggi().then(function (rec) { if (rec) mostra(rec); });
    });
    $("sostituisci").addEventListener("click", function () { $("file").click(); });
    $("rimuovi").addEventListener("click", function () { cancella().then(mostraScelta); });

    leggi().then(function (rec) { rec ? mostra(rec) : mostraScelta(); })
      .catch(function (e) { mostraScelta(); mostraMsg("Contenuto salvato non leggibile: " + e.message); });
  }

  Office.onReady(function () { init(); });
})();
