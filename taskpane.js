/*
 * Chrono LOMA - volet des tâches.
 * - Connexion Microsoft 365 (MSAL, authentification imbriquée Office)
 * - Lecture des listes (Opérations, Rédacteurs, Expéditeurs) dans le fichier Excel
 * - Attribution du numéro : ligne ajoutée dans le tableau « Chrono » + numéro en tête du mail
 */
/* global Office, msal, CHRONO_CONFIG */

(function () {
  "use strict";

  var CFG = window.CHRONO_CONFIG || {};
  var SCOPES = ["Files.ReadWrite.All", "User.Read"];
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var MSAL_URLS = [
    "https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.28.1/lib/msal-browser.min.js",
    "https://cdn.jsdelivr.net/npm/@azure/msal-browser@3/lib/msal-browser.min.js",
    "https://unpkg.com/@azure/msal-browser@3/lib/msal-browser.min.js"
  ];

  var pca = null;
  var nested = false;
  var base = null; // chemin Graph du classeur
  var item = null;

  // ---------- Utilitaires d'affichage ----------
  function $(id) { return document.getElementById(id); }
  function montrer(id, oui) { $(id).classList[oui ? "remove" : "add"]("cache"); }
  function message(texte, type) { var m = $("msg"); m.className = type || "info"; m.textContent = texte; }
  function effacerMessage() { var m = $("msg"); m.className = ""; m.textContent = ""; }
  function memo(cle, val) { try { if (val === undefined) { return localStorage.getItem(cle); } localStorage.setItem(cle, val); } catch (e) { return null; } return null; }

  function charger(url) {
    return new Promise(function (ok, ko) {
      var s = document.createElement("script");
      s.src = url; s.onload = ok; s.onerror = ko;
      document.head.appendChild(s);
    });
  }
  async function chargerMsal() {
    for (var i = 0; i < MSAL_URLS.length; i++) {
      try { await charger(MSAL_URLS[i]); if (window.msal) { return; } } catch (e) { /* suivant */ }
    }
    throw new Error("Impossible de charger la bibliothèque de connexion Microsoft (MSAL).");
  }

  function officeAsync(fn) {
    return new Promise(function (ok, ko) {
      fn(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) { ok(r.value); } else { ko(new Error(r.error ? r.error.message : "Erreur Outlook")); }
      });
    });
  }

  // ---------- Connexion ----------
  async function initAuth() {
    await chargerMsal();
    nested = Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
    var conf = { auth: { clientId: CFG.CLIENT_ID, authority: "https://login.microsoftonline.com/organizations" } };
    if (nested) {
      pca = await msal.createNestablePublicClientApplication(conf);
    } else {
      conf.auth.redirectUri = window.location.origin + window.location.pathname;
      conf.cache = { cacheLocation: "localStorage" };
      pca = await msal.createStandardPublicClientApplication(conf);
    }
  }

  async function jeton(interactif) {
    var req = { scopes: SCOPES };
    if (!nested) {
      var comptes = pca.getAllAccounts();
      if (comptes.length) { req.account = comptes[0]; }
    }
    try {
      if (!nested && !req.account) { throw new Error("pas de compte"); }
      return (await pca.acquireTokenSilent(req)).accessToken;
    } catch (e) {
      if (!interactif) { throw e; }
      return (await pca.acquireTokenPopup(req)).accessToken;
    }
  }

  async function graph(methode, chemin, corps) {
    var t = await jeton(false);
    var rep = await fetch(GRAPH + chemin, {
      method: methode,
      headers: { Authorization: "Bearer " + t, "Content-Type": "application/json" },
      body: corps ? JSON.stringify(corps) : undefined
    });
    if (!rep.ok) {
      var txt = await rep.text();
      throw new Error("Microsoft 365 a répondu " + rep.status + " : " + txt.slice(0, 300));
    }
    return rep.status === 204 ? null : rep.json();
  }

  // ---------- Classeur Excel ----------
  function idPartage(url) {
    var b64 = btoa(unescape(encodeURIComponent(url.trim())));
    return "u!" + b64.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  }
  async function ouvrirClasseur() {
    if (base) { return; }
    var di = await graph("GET", "/shares/" + idPartage(CFG.FICHIER_EXCEL) + "/driveItem?$select=id,parentReference");
    base = "/drives/" + di.parentReference.driveId + "/items/" + di.id + "/workbook";
  }
  async function colonne(table, index) {
    var r = await graph("GET", base + "/tables('" + table + "')/dataBodyRange?$select=values");
    return (r.values || []).map(function (l) { return l[index]; });
  }
  async function liste(table) {
    var v = await colonne(table, 0);
    return v.map(function (x) { return String(x == null ? "" : x).trim(); }).filter(function (x) { return x; });
  }
  // Format AA-MM-NNN, numéro remis à 001 chaque mois (ex. 26-09-001)
  function prefixeMois(d) {
    var mm = String(d.getMonth() + 1);
    if (mm.length < 2) { mm = "0" + mm; }
    return String(d.getFullYear()).slice(2) + "-" + mm + "-";
  }
  async function numeroSuivant() {
    var pre = prefixeMois(new Date());
    var nums = await colonne("Chrono", 0);
    var max = 0;
    nums.forEach(function (n) {
      var s = String(n == null ? "" : n).trim();
      if (s.indexOf(pre) === 0) {
        var k = parseInt(s.slice(pre.length), 10);
        if (!isNaN(k) && k > max) { max = k; }
      }
    });
    var suite = String(max + 1);
    while (suite.length < 3) { suite = "0" + suite; }
    return pre + suite;
  }
  function dateExcel(d) {
    return (d.getTime() - d.getTimezoneOffset() * 60000) / 86400000 + 25569;
  }

  // ---------- Formulaire ----------
  function remplir(id, valeurs, cleMemo) {
    var sel = $(id);
    sel.innerHTML = "";
    var vide = document.createElement("option");
    vide.value = ""; vide.textContent = "— choisir —";
    sel.appendChild(vide);
    valeurs.forEach(function (v) {
      var o = document.createElement("option"); o.value = v; o.textContent = v; sel.appendChild(o);
    });
    var m = cleMemo ? memo(cleMemo) : null;
    if (m && valeurs.indexOf(m) >= 0) { sel.value = m; }
  }

  async function chargerFormulaire() {
    message("Lecture du tableau du chrono…", "info");
    await ouvrirClasseur();
    var res = await Promise.all([liste("Operations"), liste("Redacteurs"), liste("Expediteurs"), numeroSuivant()]);
    remplir("selOperation", res[0]);
    remplir("selRedacteur", res[1], "chrono.redacteur");
    remplir("selExpediteur", res[2], "chrono.expediteur");
    $("prochain").textContent = res[3];
    montrer("zoneConnexion", false);
    montrer("zoneForm", true);
    effacerMessage();
  }

  async function lireEtat() {
    try {
      var tout = await officeAsync(function (cb) { item.sessionData.getAllAsync(cb); });
      return (tout && tout.chrono) || "";
    } catch (e) { return ""; }
  }
  function ecrireEtat(val) {
    return officeAsync(function (cb) { item.sessionData.setAsync("chrono", val, cb); });
  }

  function destinataires() {
    function lire(champ) {
      return officeAsync(function (cb) { item[champ].getAsync(cb); }).then(function (l) {
        return (l || []).map(function (d) { return d.displayName && d.displayName !== d.emailAddress ? d.displayName + " <" + d.emailAddress + ">" : d.emailAddress; });
      }).catch(function () { return []; });
    }
    return Promise.all([lire("to"), lire("cc")]).then(function (r) {
      var t = r[0].join("; ");
      if (r[1].length) { t += (t ? " ; Cc : " : "Cc : ") + r[1].join("; "); }
      return t;
    });
  }

  async function attribuer() {
    var op = $("selOperation").value, red = $("selRedacteur").value, exp = $("selExpediteur").value;
    if (!op || !red || !exp) { message("Choisissez l'opération, le rédacteur et l'expéditeur.", "err"); return; }
    $("btnAttribuer").disabled = true; $("btnSans").disabled = true;
    try {
      var objet = await officeAsync(function (cb) { item.subject.getAsync(cb); });
      if (!objet || !objet.trim()) {
        message("Renseignez d'abord l'objet du mail, puis cliquez à nouveau sur « Attribuer un numéro ».", "err");
        return;
      }
      message("Attribution du numéro…", "info");
      var dest = await destinataires();
      var num = await numeroSuivant();
      await graph("POST", base + "/tables('Chrono')/rows/add", {
        index: null,
        values: [[num, dateExcel(new Date()), "Mail", dest, objet, op, red, exp, ""]]
      });
      var type = await officeAsync(function (cb) { item.body.getTypeAsync(cb); });
      var texte = (CFG.PREFIXE || "N/Réf. : ") + num;
      if (type === Office.CoercionType.Html) {
        await officeAsync(function (cb) {
          item.body.prependAsync('<p style="font-family:Arial,sans-serif;font-size:10pt;color:#1F3864"><b>' + texte + "</b></p>", { coercionType: Office.CoercionType.Html }, cb);
        });
      } else {
        await officeAsync(function (cb) { item.body.prependAsync(texte + "\n\n", { coercionType: Office.CoercionType.Text }, cb); });
      }
      await ecrireEtat("fait:" + num);
      memo("chrono.redacteur", red); memo("chrono.expediteur", exp);
      afficherFait(num);
      message("Numéro " + num + " attribué. Cliquez sur Envoyer.", "ok");
    } catch (e) {
      message("Le numéro n'a pas pu être attribué.\n" + e.message, "err");
    } finally {
      $("btnAttribuer").disabled = false; $("btnSans").disabled = false;
    }
  }

  async function sansChrono() {
    try {
      await ecrireEtat("non");
      message("Ce mail partira sans numéro de chrono. Cliquez sur Envoyer.", "ok");
    } catch (e) {
      message("Erreur : " + e.message, "err");
    }
  }

  function afficherFait(num) {
    $("numFait").textContent = num;
    montrer("zoneForm", false);
    montrer("zoneConnexion", false);
    montrer("zoneFait", true);
  }

  async function demarrer() {
    item = Office.context.mailbox.item;
    $("btnAttribuer").onclick = attribuer;
    $("btnSans").onclick = sansChrono;
    $("btnConnexion").onclick = async function () {
      try { await jeton(true); await chargerFormulaire(); } catch (e) { message("Connexion impossible.\n" + e.message, "err"); }
    };

    if (!CFG.CLIENT_ID || CFG.CLIENT_ID.indexOf("COLLER") === 0 || !CFG.FICHIER_EXCEL || CFG.FICHIER_EXCEL.indexOf("COLLER") === 0) {
      message("Le fichier config.js n'est pas encore rempli (ID d'application et lien du fichier Excel).", "err");
      return;
    }
    var etat = await lireEtat();
    if (etat.indexOf("fait:") === 0) { afficherFait(etat.slice(5)); return; }

    try {
      await initAuth();
    } catch (e) { message(e.message, "err"); return; }

    try {
      await jeton(false);
      await chargerFormulaire();
    } catch (e) {
      montrer("zoneConnexion", true);
      effacerMessage();
    }
    if (etat === "non") { message("Ce mail est actuellement marqué « sans chrono ». Vous pouvez encore attribuer un numéro.", "info"); }
  }

  Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) { demarrer(); }
  });
})();
