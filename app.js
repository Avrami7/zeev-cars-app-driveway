(function(){
"use strict";

/* =====================================================================
   CONFIGURATION — A ADAPTER APRES DEPLOIEMENT DU WORKER (voir DEPLOIEMENT.md)
   ===================================================================== */
const CONFIG = {
  apiBase: "https://zeev-cars-api-driveway.4wg2rh4rhp.workers.dev",
  githubOAuthClientId: "Ov23liK2WpeDu2nKUT3v",
};

const STATUT_LABELS = { COMMANDE:"Commandé — en attente de livraison", LIVRE:"Livré", NON_LIVRE:"Non livré" };
const state = { user:null, scope:null, partners:[], vehicles:[], meta:{}, dirty:false, activeTab:null, theme:"light", deletedVehicleIds:[] };

var $ = function(s){ return document.querySelector(s); };
var $all = function(s){ return Array.prototype.slice.call(document.querySelectorAll(s)); };
function uid(p){ return p + "_" + Date.now().toString(36) + Math.random().toString(36).slice(2,7); }
function fmtMoney(n){ n = Number(n)||0; return n.toLocaleString("fr-FR",{minimumFractionDigits:2,maximumFractionDigits:2}) + " €"; }
function fmtPct(n){ if (n==null) return "—"; n = Number(n)||0; return (Math.round(n*100)/100).toString().replace(".", ",") + " %"; }
function escapeAttr(s){ return String(s==null?"":s).replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;"); }
function escapeHtml(s){ return escapeAttr(s); }
function todayIso(){ return new Date().toISOString().slice(0,10); }
function addDaysIso(iso, days){ var d = iso ? new Date(iso) : new Date(); d.setDate(d.getDate()+days); return d.toISOString().slice(0,10); }
function isoToFr(iso){
  if (!iso) return "";
  var m = String(iso).trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return iso;
  return m[3] + "/" + m[2] + "/" + m[1];
}
function frToIso(fr){
  if (!fr) return "";
  var s = String(fr).trim();
  var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (!m) return s;
  var d = m[1].padStart(2,"0"), mo = m[2].padStart(2,"0"), y = m[3];
  return y + "-" + mo + "-" + d;
}

function toast(msg, kind){
  var el = $("#toast");
  el.textContent = msg; el.className = "toast" + (kind?" "+kind:"");
  el.classList.remove("hidden");
  clearTimeout(toast._t);
  toast._t = setTimeout(function(){ el.classList.add("hidden"); }, 3400);
}
function markDirty(){
  state.dirty = true;
  state.pendingCount = (state.pendingCount||0) + 1;
  var s=$("#sync-status");
  s.textContent = "● " + state.pendingCount + " modification" + (state.pendingCount>1?"s":"") + " non enregistrée" + (state.pendingCount>1?"s":"");
  s.style.color="var(--warn)";
}
function markClean(){ state.dirty = false; state.pendingCount = 0; var s=$("#sync-status"); s.textContent="Synchronisé"; s.style.color="var(--ok)"; }

/* =====================================================================
   APPELS AU WORKER — seul point de contact reseau
   ===================================================================== */
function isApiConfigured(){ return CONFIG.apiBase && !CONFIG.apiBase.includes("VOTRE-SOUS-DOMAINE"); }
async function apiCall(path, opts){
  opts = opts || {};
  if (!isApiConfigured()) throw new Error("Configuration manquante : renseignez CONFIG.apiBase (voir DEPLOIEMENT.md).");
  var headers = { "Content-Type":"application/json" };
  var res = await fetch(CONFIG.apiBase + path, { method: opts.method||"GET", headers: headers, credentials: "include", body: opts.body?JSON.stringify(opts.body):undefined });
  var j = await res.json().catch(function(){ return {}; });
  if (!res.ok){
    var err = new Error(j.error || ("Erreur réseau (" + res.status + ")."));
    err.status = res.status;
    throw err;
  }
  return j;
}
const apiLoginGithub = function(code, redirectUri){ return apiCall("/api/login-github", { method:"POST", body:{ code:code, redirectUri:redirectUri } }); };
const apiGetData = function(){ return apiCall("/api/data"); };
const apiSaveData = function(payload){ return apiCall("/api/data", { method:"POST", body:payload }); };
const apiGetCollaborators = function(){ return apiCall("/api/collaborators"); };
const apiGetPaymentDetails = function(){ return apiCall("/api/payment-details"); };
const apiLogout = function(){ return apiCall("/api/logout", { method:"POST" }); };
async function apiExportCsv(){
  if (!isApiConfigured()) throw new Error("Configuration manquante : renseignez CONFIG.apiBase.");
  var res = await fetch(CONFIG.apiBase + "/api/export", { credentials: "include" });
  if (!res.ok){ var e = await res.json().catch(function(){return {};}); throw new Error(e.error || ("Erreur réseau (" + res.status + ")." )); }
  var blob = await res.blob();
  var disposition = res.headers.get("Content-Disposition") || "";
  var m = disposition.match(/filename="([^"]+)"/);
  var filename = m ? m[1] : ("zeevcars-driveway-export-" + todayIso() + ".csv");
  var url = URL.createObjectURL(blob);
  var a = document.createElement("a"); a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}

/* =====================================================================
   IMPORT CSV — colonnes alignees sur l'export (/api/export), pour un
   aller-retour fiable Export -> modification dans Excel -> Import.
   Colonnes reconnues (position, pas obligatoirement toutes presentes) :
   id;N°;Société de leasing;Client;Conducteur;Modèle;Immat.;VIN;
   N° contrat/commande;Statut véhicule (source);Concession;SIREN;SIRET;
   Adresse fournisseur;Statut livraison;Date livraison prévue;
   Date livraison effective;Montant HT;% Commission;Commission HT;
   Statut facturation;N° facture;Date facture;Statut paiement;
   Éligibilité facturation;Pénalité...;Remarques
   ===================================================================== */
function parseCsvLine(line){
  var out = [], cur = "", inQuotes = false;
  for (var i = 0; i < line.length; i++){
    var c = line[i];
    if (inQuotes){
      if (c === '"'){
        if (line[i+1] === '"'){ cur += '"'; i++; } else { inQuotes = false; }
      } else cur += c;
    } else {
      if (c === '"') inQuotes = true;
      else if (c === ";"){ out.push(cur); cur = ""; }
      else cur += c;
    }
  }
  out.push(cur);
  return out;
}
function parseNum(s){
  if (s == null || s === "") return null;
  var n = parseFloat(String(s).replace(/\s/g,"").replace(",", "."));
  return isNaN(n) ? null : n;
}

function importCsvFile(file){
  var reader = new FileReader();
  reader.onload = function(){
    var text = String(reader.result).replace(/^\uFEFF/, "");
    var lines = text.split(/\r?\n/).filter(function(l){ return l.trim().length; });
    if (lines.length < 2){ toast("Fichier CSV vide ou invalide.", "err"); return; }

    var cols = ["id","num","societeLeasing","client","conducteur","modele","immatriculation","vin",
      "numeroContrat","statutVehiculeSource","concession","siren","siret","adresseFournisseur",
      "statutLivraison","dateLivraisonPrevue","dateLivraisonEffective","montantHT","pctCommission",
      "commissionHT","statutFacturation","numeroFacture","dateFacture","statutPaiement",
      "eligibilite","penalite","remarques"];

    var partnersByName = {};
    state.partners.forEach(function(p){ partnersByName[p.distributeur.trim().toLowerCase()] = p; });

    var added = 0, updated = 0, skipped = 0;
    for (var i = 1; i < lines.length; i++){
      var fields = parseCsvLine(lines[i]);
      if (fields.length < 10){ skipped++; continue; }
      var row = {};
      cols.forEach(function(name, idx){ row[name] = fields[idx] != null ? fields[idx].trim() : ""; });

      var partner = partnersByName[row.concession.toLowerCase()];
      if (!partner){ skipped++; continue; }

      var existing = row.id ? state.vehicles.find(function(v){ return v.id === row.id; }) : null;
      var target = existing || { id: uid("v") };

      target.partnerId = partner.id;
      target.societeLeasing = row.societeLeasing;
      target.client = row.client;
      target.conducteur = row.conducteur;
      target.modele = row.modele;
      target.immatriculation = row.immatriculation.toUpperCase();
      target.vin = row.vin.toUpperCase();
      target.numeroContrat = row.numeroContrat;
      target.statutVehiculeSource = row.statutVehiculeSource;
      target.statutLivraison = ["COMMANDE","LIVRE","NON_LIVRE"].indexOf(row.statutLivraison) !== -1 ? row.statutLivraison : (existing ? existing.statutLivraison : "COMMANDE");
      target.dateLivraisonPrevue = frToIso(row.dateLivraisonPrevue);
      target.dateLivraisonEffective = frToIso(row.dateLivraisonEffective);
      target.montantHT = parseNum(row.montantHT);
      target.statutFacturation = ["NON_FACTURE","FACTURE"].indexOf(row.statutFacturation) !== -1 ? row.statutFacturation : (existing ? existing.statutFacturation : "NON_FACTURE");
      target.numeroFacture = row.numeroFacture;
      target.dateFacture = frToIso(row.dateFacture);
      target.statutPaiement = ["A_VERIFIER","IMPAYE","PAYE"].indexOf(row.statutPaiement) !== -1 ? row.statutPaiement : (existing ? existing.statutPaiement : "A_VERIFIER");
      target.remarques = row.remarques;

      if (existing) updated++;
      else { state.vehicles.push(target); added++; }
    }

    if (added || updated) markDirty();
    toast(added + " ajouté(s), " + updated + " mis à jour" + (skipped ? ", " + skipped + " ligne(s) ignorée(s) (concession introuvable)" : "") + ". Pensez à « Enregistrer sur GitHub ».", (added||updated) ? "ok" : "err");
    renderAll();
  };
  reader.readAsText(file, "UTF-8");
}

/* =====================================================================
   LOGIN — OAuth GitHub (Authorization Code flow)
   ===================================================================== */
const OAUTH_STATE_KEY = "zcd_oauth_state";
function oauthRedirectUri(){ return window.location.origin + window.location.pathname; }

function startGithubLogin(){
  var errEl = $("#login-error"); errEl.classList.remove("show"); errEl.textContent = "";
  if (!CONFIG.githubOAuthClientId || CONFIG.githubOAuthClientId.includes("COLLEZ_ICI")){
    errEl.textContent = "Configuration manquante : renseignez CONFIG.githubOAuthClientId (voir DEPLOIEMENT.md).";
    errEl.classList.add("show"); return;
  }
  var state_ = crypto.randomUUID();
  try { sessionStorage.setItem(OAUTH_STATE_KEY, state_); } catch(e){}
  var params = new URLSearchParams({ client_id: CONFIG.githubOAuthClientId, redirect_uri: oauthRedirectUri(), scope:"read:user", state: state_ });
  window.location.href = "https://github.com/login/oauth/authorize?" + params.toString();
}

async function handleGithubCallback(){
  var params = new URLSearchParams(window.location.search);
  var code = params.get("code"), oauthState = params.get("state");
  if (!code) return false;
  var cleanUrl = oauthRedirectUri();
  window.history.replaceState({}, document.title, cleanUrl);

  var saved = null;
  try { saved = sessionStorage.getItem(OAUTH_STATE_KEY); sessionStorage.removeItem(OAUTH_STATE_KEY); } catch(e){}
  if (saved && oauthState && saved !== oauthState){ showLoginError("Connexion refusée (anomalie de sécurité OAuth). Réessayez."); return true; }

  setLoginLoading(true);
  try {
    var r = await apiLoginGithub(code, cleanUrl);
    persistSession(r.token, r.user);
    await bootApp();
  } catch(e){
    showLoginError(e.message || "Connexion impossible.");
  } finally {
    setLoginLoading(false);
  }
  return true;
}
function persistSession(token, user){
  state.user = user; state.scope = user.scope || (user.isOwner?"FULL":null);
}
function showLoginError(msg){
  var e = $("#login-error");
  e.innerHTML = escapeHtml(msg).replace(/(github\.com\/[a-zA-Z0-9-]+)/, '<a href="https://$1" target="_blank" rel="noopener">$1</a>');
  e.classList.add("show");
}
function setLoginLoading(loading){
  var btn = $("#login-btn"), label = btn.querySelector(".btn-label"), spinner = btn.querySelector(".btn-spinner");
  btn.disabled = loading;
  if (label) label.textContent = loading ? "Connexion…" : "Se connecter avec GitHub";
  if (spinner) spinner.classList.toggle("hidden", !loading);
}
async function tryBootFromCookie(){
  try {
    var data = await apiGetData();
    state.partners = data.partners || [];
    state.vehicles = data.vehicles || [];
    state.meta = data.meta || {};
    state.deletedVehicleIds = [];
    state.scope = data.scope || null;
    if (data.user) state.user = data.user;
    $("#auth-overlay").classList.add("hidden");
    $("#app-screen").classList.remove("hidden");
    renderUserBadge();
    applyScopeVisibility();
    state.activeTab = state.partners[0] ? state.partners[0].id : "all";
    renderAll();
    markClean();
    return true;
  } catch(e){
    return false; // pas de cookie valide -> ecran de connexion normal, rien a signaler
  }
}
async function performLogout(silent){
  if (!silent && state.dirty && !confirm("Des modifications non enregistrées seront perdues. Se déconnecter quand même ?")) return;
  try { await apiLogout(); } catch(e){}
  window.location.reload();
}

/* =====================================================================
   BOOT
   ===================================================================== */
function showAccessGranted(){
  return new Promise(function(resolve){
    $all("#status-dots .status-dot").forEach(function(d){ d.classList.add("access-granted"); });
    setTimeout(resolve, 550);
  });
}
async function bootApp(){
  await showAccessGranted();
  $("#auth-overlay").classList.add("hidden");
  $("#app-screen").classList.remove("hidden");
  renderUserBadge();
  try {
    var data = await apiGetData();
    state.partners = data.partners || [];
    state.vehicles = data.vehicles || [];
    state.meta = data.meta || {};
    state.deletedVehicleIds = [];
    state.scope = data.scope || state.scope;
    if (data.user) state.user = Object.assign({}, state.user, data.user);
    renderUserBadge();
    applyScopeVisibility();
    state.activeTab = state.partners[0] ? state.partners[0].id : "all";
    renderAll();
    markClean();
  } catch(e){
    toast("Erreur de chargement : " + e.message, "err");
    if (/session/i.test(e.message)) performLogout(true);
  }
}
function renderUserBadge(){
  var u = state.user || {};
  var scopeCls = state.scope === "FULL" ? "full" : "restricted";
  var scopeLbl = state.scope === "FULL" ? "FULL" : "LIVRAISON";
  $("#current-user-badge").innerHTML =
    (u.avatarUrl ? '<img class="tb-avatar" src="'+escapeAttr(u.avatarUrl)+'" alt="">' : '') +
    '<span>' + escapeHtml(u.name||u.login||"—") + '</span>' +
    '<span class="tb-scope ' + scopeCls + '">' + scopeLbl + '</span>';
}
function applyScopeVisibility(){
  var full = state.scope === "FULL";
  $("#collab-btn").classList.toggle("hidden", !(state.user && state.user.isOwner));
  $("#export-csv-btn").classList.toggle("hidden", !full);
  $("#import-csv-btn").classList.toggle("hidden", !full);
}

/* =====================================================================
   HELPERS PARTENAIRES / VEHICULES
   ===================================================================== */
function partnerById(id){ for (var i=0;i<state.partners.length;i++) if (state.partners[i].id===id) return state.partners[i]; return null; }
function vehiclesForPartner(id){ return state.vehicles.filter(function(v){ return v.partnerId===id; }); }
function commissionFor(v){
  var p = partnerById(v.partnerId);
  if (!p || p.commissionPct == null || v.montantHT == null) return null;
  return Number(v.montantHT||0) * Number(p.commissionPct||0) / 100;
}
function shortName(p){ return p.distributeur.split(" — ")[0]; }

/* =====================================================================
   NAVIGATION
   ===================================================================== */
function renderNavTabs(){
  var nameCounts = {};
  state.partners.forEach(function(p){ var n = shortName(p); nameCounts[n] = (nameCounts[n]||0) + 1; });
  var html = state.partners.map(function(p){
    var n = shortName(p);
    var label = nameCounts[n] > 1 ? n + " (" + (p.marques||[]).join("/") + ")" : n;
    return '<button class="ntab'+(state.activeTab===p.id?" active":"")+'" data-tab="'+p.id+'" title="'+escapeAttr(p.distributeur)+' ('+escapeAttr((p.marques||[]).join("/"))+')">'+escapeHtml(label)+'</button>';
  }).join("");
  html += '<button class="ntab'+(state.activeTab==="all"?" active":"")+'" data-tab="all">Tous les véhicules commandés</button>';
  html += '<button class="ntab'+(state.activeTab==="dashboard"?" active":"")+'" data-tab="dashboard">Tableau de bord</button>';
  if (state.user && state.user.isOwner){
    html += '<button class="ntab'+(state.activeTab==="admin"?" active":"")+'" data-tab="admin">⚙ Admin</button>';
  }
  $("#nav-tabs").innerHTML = html;
  $all("#nav-tabs .ntab").forEach(function(b){ b.addEventListener("click", function(){ state.activeTab = b.dataset.tab; renderAll(); }); });
  var searchEl = $("#nav-search");
  if (searchEl && searchEl.value) searchEl.dispatchEvent(new Event("input"));
}

/* =====================================================================
   PANEL PARTENAIRE
   ===================================================================== */
function kpiCard(label, value, sub, cls, id){
  return '<div class="kpi '+(cls||"")+'"><div class="k-label">'+label+'</div><div class="k-value"'+(id?' id="'+id+'"':'')+'>'+value+'</div>'+(sub?'<div class="k-sub"'+(id?' id="'+id+'-sub"':'')+'>'+sub+'</div>':'')+'</div>';
}
function statutPill(s){
  if (s === "LIVRE") return '<span class="pill ok pill-click" data-statut="LIVRE" title="Filtrer sur ce statut">Livré</span>';
  if (s === "NON_LIVRE") return '<span class="pill err pill-click" data-statut="NON_LIVRE" title="Filtrer sur ce statut">Non livré</span>';
  return '<span class="pill warn pill-click" data-statut="COMMANDE" title="Filtrer sur ce statut">Commandé</span>';
}

function computePartnerKpis(p){
  var full = state.scope === "FULL";
  var list = vehiclesForPartner(p.id);
  var caHt = list.reduce(function(s,v){ return s + Number(v.montantHT||0); }, 0);
  var commission = full ? list.reduce(function(s,v){ var c=commissionFor(v); return s+(c||0); }, 0) : null;
  var livres = list.filter(function(v){ return v.statutLivraison === "LIVRE"; });
  var nonLivres = list.length - livres.length;
  var facturables = livres.filter(function(v){ return v.statutFacturation !== "FACTURE"; });

  // Commission facturable : uniquement les vehicules effectivement livres
  // (statut LIVRE + date de livraison effective renseignee), quel que soit
  // leur statut de paiement -- distincte du total estime (toutes lignes)
  // et de l'encaisse (deja paye).
  var commissionFacturable = 0, vehiculesLivresAvecDate = 0;
  if (full){
    livres.forEach(function(v){
      if (!v.dateLivraisonEffective) return;
      if (v.statutPaiement === "PAYE") return; // deja encaisse : ne compte pas aussi en "facturable"
      var c = commissionFor(v);
      if (c == null) return;
      commissionFacturable += c;
      vehiculesLivresAvecDate++;
    });
  }

  var commissionEncaissee = 0, commissionEnAttente = 0, retardsJours = [];
  if (full){
    list.forEach(function(v){
      var c = commissionFor(v);
      if (c == null) return;
      if (v.statutPaiement === "PAYE") commissionEncaissee += c;
      else commissionEnAttente += c;
      if (v.statutFacturation === "FACTURE" && v.dateFacture && v.statutPaiement !== "PAYE"){
        var jours = Math.floor((Date.now() - new Date(v.dateFacture+"T00:00:00").getTime()) / 86400000);
        if (jours >= 0) retardsJours.push(jours);
      }
    });
  }
  var delaiMoyen = retardsJours.length ? Math.round(retardsJours.reduce(function(a,b){return a+b;},0) / retardsJours.length) : null;

  var datesLivrees = livres.map(function(v){ return v.dateLivraisonEffective; }).filter(Boolean).sort();
  var estimationAnnuelle = null, estimationSub = "historique insuffisant pour projeter";
  if (full && datesLivrees.length && commission){
    var moisSpan = Math.max(1, Math.round((new Date(datesLivrees[datesLivrees.length-1]) - new Date(datesLivrees[0])) / (30*86400000)) + 1);
    if (moisSpan >= 2){ estimationAnnuelle = commission / moisSpan * 12; estimationSub = "projection sur " + moisSpan + " mois d'historique"; }
  }
  return { full:full, list:list, caHt:caHt, commission:commission, livres:livres, nonLivres:nonLivres, facturables:facturables,
    commissionFacturable:commissionFacturable, vehiculesLivresAvecDate:vehiculesLivresAvecDate,
    commissionEncaissee:commissionEncaissee, commissionEnAttente:commissionEnAttente, delaiMoyen:delaiMoyen,
    estimationAnnuelle:estimationAnnuelle, estimationSub:estimationSub };
}

function lockedKpiCard(label){
  return '<div class="kpi locked"><div class="k-label">'+label+'</div><div class="k-value">🔒</div><div class="k-lock-hint">Accès complet requis</div></div>';
}

function renderPartnerPanel(p, searchQuery){
  var k = computePartnerKpis(p);
  var full = k.full, list = k.list, caHt = k.caHt, commission = k.commission, livres = k.livres,
      nonLivres = k.nonLivres, facturables = k.facturables, commissionEncaissee = k.commissionEncaissee,
      commissionEnAttente = k.commissionEnAttente, delaiMoyen = k.delaiMoyen,
      estimationAnnuelle = k.estimationAnnuelle, estimationSub = k.estimationSub;

  var rateBlock = full
    ? '<span class="rate-display">' + fmtPct(p.commissionPct) + '</span>'
    : '<span class="param-text">🔒 Masqué — accès restreint (Livraison uniquement)</span>';

  var simTag = !p.signedDate ? ' <span style="color:var(--tx3);font-weight:400;">(simulation)</span>' : '';
  var commissionHighlightRow = !full
    ? '<div class="kpi-highlight-row"><div class="kpi-highlight"><div class="kh-label">Commission HT totale</div><div class="kh-value">🔒</div><div class="kh-sub">Accès complet requis</div></div>' +
      '<div class="kpi-highlight"><div class="kh-label">Commission encaissée</div><div class="kh-value">🔒</div><div class="kh-sub">Accès complet requis</div></div></div>'
    : '<div class="kpi-highlight-row">' +
        '<div class="kpi-highlight"><div class="kh-label">Commission HT totale'+simTag+'</div><div class="kh-value">'+fmtMoney(commission)+'</div>' +
          '<div class="kh-sub">Facturable (Livré — paiement à vérifier) :<br><strong style="color:var(--ok)">'+fmtMoney(k.commissionFacturable)+'</strong></div></div>' +
        '<div class="kpi-highlight ok"><div class="kh-label">Commission encaissée'+simTag+'</div><div class="kh-value">'+fmtMoney(commissionEncaissee)+'</div></div>' +
      '</div>';

  var html = '<section class="tab-panel active">' +
    '<div class="print-report-header"><h2>' + escapeHtml(p.distributeur) + '</h2><div class="print-report-date">Édité le ' + isoToFr(todayIso()) + '</div></div>' +
    (full ? '' : '<div class="bandeau bandeau-warn">🔒 Compte à accès Livraison — les montants et commissions sont masqués. Contactez l\'administrateur pour un accès complet si besoin.</div>') +
    '<div class="bandeau bandeau-ok">Entité à facturer : <strong>' + escapeHtml(p.distributeur.toUpperCase()) + '</strong>' +
      (p.siren ? ' — SIREN ' + escapeHtml(p.siren) : '') + ' — SIRET ' + escapeHtml(p.siret) + '<br>' +
      escapeHtml(p.adresse) + ', ' + escapeHtml(p.cp) + ' ' + escapeHtml(p.ville) +
      (p.telephoneStandard ? ' — Tél. ' + escapeHtml(p.telephoneStandard) : '') +
      (full && p.signedDate ? ' — Convention signée le <strong>' + p.signedDate + '</strong>' + (p.contratRef ? ' (' + escapeHtml(p.contratRef) + ')' : '') : '') + '</div>' +
    (full && p.perimetreContrat ? '<div class="bandeau bandeau-warn">' + escapeHtml(p.perimetreContrat) + '</div>' : '') +
    '<div class="params-row">' +
      '<div class="card"><label>🔒 Taux de commission contractuel</label>' + rateBlock + '</div>' +
      '<div class="card"><label>Base de calcul</label><span class="param-text">LLD/LOA : montant HT total du contrat de location signé</span></div>' +
      '<div class="card"><label>Condition de déclenchement</label><span class="param-text">' + escapeHtml((full && p.conditionDeclenchement) || "Livraison effective + PV de mise à disposition signé") + '</span></div>' +
    '</div>' +
    '<div class="actions-row">' +
      '<button class="btn primary" data-open-vehicle="'+p.id+'">Ajouter un véhicule</button>' +
      (full ? '<button class="btn" data-open-invoice="'+p.id+'">Facture consolidée</button>' : '') +
      '<button class="btn" data-open-analysis="'+p.id+'">Analyse client</button>' +
    '</div>' +
    commissionHighlightRow +
    '<div class="kpi-row">' +
      kpiCard("Livrés", livres.length, "sur " + list.length + " véhicule(s) au total", "ok", "kpi-livres") +
      kpiCard("Non livrés", nonLivres, "", nonLivres ? "warn" : "", "kpi-non-livres") +
      (full ? kpiCard("Valeur de vente totale (HT)", fmtMoney(caHt), "", "", "kpi-valeur-vente") : lockedKpiCard("Valeur de vente totale (HT)")) +
    '</div>' +
    '<div class="kpi-row">' +
      (full ? kpiCard("Délai moyen de règlement", delaiMoyen != null ? (delaiMoyen + " j") : "—", delaiMoyen != null ? "sur factures impayées" : "aucune facture en retard", "", "kpi-delai-reglement") : lockedKpiCard("Délai moyen de règlement")) +
      (full ? kpiCard("Estimation annuelle", estimationAnnuelle != null ? fmtMoney(estimationAnnuelle) : fmtMoney(0), estimationSub, "", "kpi-estimation-annuelle") : lockedKpiCard("Estimation annuelle")) +
    '</div>' +
    '<div class="veh-search-toolbar">' +
      '<input type="text" id="veh-search-'+p.id+'" placeholder="Rechercher un véhicule (client, modèle, immatriculation...)" value="'+escapeAttr(searchQuery||"")+'">' +
      '<select id="veh-filter-statut-'+p.id+'"><option value="">Tous statuts</option>' +
        Object.keys(STATUT_LABELS).map(function(k){ return '<option value="'+k+'">'+STATUT_LABELS[k]+'</option>'; }).join("") + '</select>' +
    '</div>' +
    '<div class="table-wrap"><table class="data-table" id="table-partner-'+p.id+'"></table></div>' +
  '</section>';

  $("#main-content").innerHTML = html;
  bindPanelActions();
  var searchInput = $("#veh-search-"+p.id);
  var statutSelect = $("#veh-filter-statut-"+p.id);
  function refresh(){
    var q = searchInput.value, sf = statutSelect.value;
    var l = list.filter(function(v){ return !sf || v.statutLivraison === sf; });
    renderVehicleTable("#table-partner-"+p.id, l, q, false, function(statut){ statutSelect.value = statut; refresh(); }, refresh);
  }
  searchInput.addEventListener("input", refresh);
  statutSelect.addEventListener("change", refresh);
  refresh();
}

const STATUT_PAIEMENT_LABELS = { A_VERIFIER:"À vérifier", IMPAYE:"Impayé", PAYE:"Payé" };
function eligibiliteFacturation(v){
  if (v.statutLivraison !== "LIVRE") return { label:"En attente livraison", cls:"warn" };
  if (!v.statutPaiement || v.statutPaiement === "A_VERIFIER") return { label:"Livré — paiement à vérifier", cls:"info" };
  if (v.statutPaiement !== "PAYE") return { label:"En attente paiement", cls:"warn" };
  return { label:"Éligible facturation", cls:"ok" };
}

/* Tableau simplifie : 5 colonnes essentielles + acces "Details" pour le reste.
   Recherche cote client sur client/modele/immatriculation.
   onStatusClick(statut) : appele au clic sur une pastille, pour filtrage rapide.
   refreshFn() : appele apres une action groupee pour redessiner la vue courante. */
function renderVehicleTable(sel, list, query, showConcession, onStatusClick, refreshFn){
  var full = state.scope === "FULL";
  query = (query||"").trim().toLowerCase();
  var filtered = !query ? list : list.filter(function(v){
    var p2 = partnerById(v.partnerId);
    return (v.client||"").toLowerCase().indexOf(query)!==-1 ||
           (v.conducteur||"").toLowerCase().indexOf(query)!==-1 ||
           (v.modele||"").toLowerCase().indexOf(query)!==-1 ||
           (v.immatriculation||"").toLowerCase().indexOf(query)!==-1 ||
           (showConcession && p2 && p2.distributeur.toLowerCase().indexOf(query)!==-1);
  });
  var head = "<thead><tr><th class=\"chk\"><input type=\"checkbox\" class=\"row-select-all\" title=\"Tout sélectionner\"></th><th class=\"num-col\">N°</th>" + (showConcession ? "<th>Concession</th>" : "") + "<th>Client</th><th>Conducteur</th><th>Modèle / Version</th><th>Immatriculation</th><th>Statut livraison</th>" +
    (full ? "<th>Montant HT</th>" : "<th title=\"Visible uniquement pour les comptes Facturation complète\">🔒 Montant HT</th>") +
    (full ? "<th>Commission HT</th>" : "<th title=\"Visible uniquement pour les comptes Facturation complète\">🔒 Commission HT</th>") +
    "<th></th></tr></thead>";
  var body = "";
  if (!filtered.length){
    var colspan = (showConcession ? 7 : 6) + 4;
    body = '<tr class="empty-row"><td colspan="'+colspan+'">' + (query ? "Aucun véhicule ne correspond à « " + escapeHtml(query) + " »." : 'Aucun véhicule pour cette concession — cliquez sur « Ajouter un véhicule ».') + '</td></tr>';
  } else {
    filtered.forEach(function(v, idx){
      var p = partnerById(v.partnerId);
      body += '<tr data-id="'+v.id+'">' +
        '<td class="chk"><input type="checkbox" class="row-select" data-veh-id="'+v.id+'"></td>' +
        '<td class="num-col" data-label="N°">'+(idx+1)+'</td>' +
        (showConcession ? '<td data-label="Concession">'+escapeHtml(p?p.distributeur:"—")+'</td>' : '') +
        '<td data-label="Client">'+escapeHtml(v.client)+'</td>' +
        '<td data-label="Conducteur">'+escapeHtml(v.conducteur||"—")+'</td>' +
        '<td data-label="Modèle">'+escapeHtml(v.modele)+'</td>' +
        '<td data-label="Immatriculation">'+escapeHtml(v.immatriculation)+'</td>' +
        '<td data-label="Statut">'+statutPill(v.statutLivraison)+'</td>' +
        (full ? '<td class="num" data-label="Montant HT">'+(v.montantHT!=null?fmtMoney(v.montantHT):"—")+'</td>' : '<td class="num lock-cell" data-label="Montant HT" title="Visible uniquement pour les comptes Facturation complète">🔒</td>') +
        (full ? '<td class="num" data-label="Commission HT">'+(commissionFor(v)!=null?fmtMoney(commissionFor(v)):"—")+'</td>' : '<td class="num lock-cell" data-label="Commission HT" title="Visible uniquement pour les comptes Facturation complète">🔒</td>') +
        '<td class="actions"><button class="btn small" data-details-vehicle="'+v.id+'">Détails</button> <button class="btn small danger-o" data-del-vehicle="'+v.id+'">Supprimer</button></td>' +
      '</tr>';
    });
  }
  var table = $(sel);
  table.innerHTML = head + "<tbody>" + body + "</tbody>";

  // Barre d'actions groupees : creee/recreee juste avant le tableau a chaque rendu
  var barId = sel.replace("#","") + "-bulkbar";
  var oldBar = document.getElementById(barId);
  if (oldBar) oldBar.remove();
  var wrap = table.closest(".table-wrap") || table.parentElement;
  wrap.insertAdjacentHTML("beforebegin",
    '<div class="bulk-bar hidden" id="'+barId+'">' +
      '<span class="bulk-count">0 sélectionné(s)</span>' +
      '<button class="btn small" data-bulk="LIVRE">Marquer Livré</button>' +
      '<button class="btn small" data-bulk="NON_LIVRE">Marquer Non livré</button>' +
      '<button class="btn small" data-bulk="COMMANDE">Marquer Commandé</button>' +
    '</div>'
  );
  var bar = document.getElementById(barId);

  function selectedIds(){ return $all(sel+" .row-select:checked").map(function(c){ return c.dataset.vehId; }); }
  function updateBar(){
    var ids = selectedIds();
    bar.classList.toggle("hidden", !ids.length);
    bar.querySelector(".bulk-count").textContent = ids.length + " sélectionné(s)";
  }
  $all(sel+" .row-select").forEach(function(c){ c.addEventListener("change", updateBar); });
  var selectAll = table.querySelector(".row-select-all");
  selectAll.addEventListener("change", function(){
    $all(sel+" .row-select").forEach(function(c){ c.checked = selectAll.checked; });
    updateBar();
  });
  bar.querySelectorAll("[data-bulk]").forEach(function(btn){
    btn.addEventListener("click", function(){
      var ids = selectedIds();
      if (!ids.length) return;
      var statut = btn.dataset.bulk;
      if (statut === "LIVRE" && state.scope === "FULL"){
        var missing = ids.filter(function(id){
          var v = state.vehicles.find(function(x){ return x.id === id; });
          return v && (v.montantHT == null);
        });
        if (missing.length){
          toast(missing.length + " véhicule(s) sélectionné(s) n'ont pas de montant HT — impossible de les marquer « Livré » en masse. Renseignez le montant via « Détails » d'abord.", "err");
          return;
        }
      }
      ids.forEach(function(id){
        var v = state.vehicles.find(function(x){ return x.id === id; });
        if (v) v.statutLivraison = statut;
      });
      markDirty();
      toast(ids.length + " véhicule(s) marqué(s) « " + STATUT_LABELS[statut] + " ».", "ok");
      if (refreshFn) refreshFn(); else renderVehicleTable(sel, list, query, showConcession, onStatusClick, refreshFn);
    });
  });

  $all(sel+" [data-details-vehicle]").forEach(function(btn){
    btn.addEventListener("click", function(){ openVehicleDetailsModal(btn.dataset.detailsVehicle); });
  });
  $all(sel+" [data-del-vehicle]").forEach(function(btn){
    btn.addEventListener("click", function(){ deleteVehicleWithUndo(btn.dataset.delVehicle); });
  });
  if (onStatusClick){
    $all(sel+" .pill-click").forEach(function(el){
      el.addEventListener("click", function(){ onStatusClick(el.dataset.statut); });
    });
  }
}

/* Suppression avec delai de grace : retrait immediat de l'affichage,
   mais annulable pendant 5 secondes via le toast, avant tout enregistrement. */
function deleteVehicleWithUndo(vehicleId){
  var idx = state.vehicles.findIndex(function(v){ return v.id === vehicleId; });
  if (idx === -1) return;
  var removed = state.vehicles[idx];
  var removedIndex = idx;
  state.vehicles.splice(idx, 1);
  if (state.deletedVehicleIds.indexOf(vehicleId) === -1) state.deletedVehicleIds.push(vehicleId);
  markDirty();
  renderAll();
  toastWithUndo("Véhicule supprimé.", function(){
    state.vehicles.splice(removedIndex, 0, removed);
    state.deletedVehicleIds = state.deletedVehicleIds.filter(function(id){ return id !== vehicleId; });
    markDirty();
    renderAll();
  });
}

function bindPanelActions(){
  $all("[data-open-vehicle]").forEach(function(b){ b.addEventListener("click", function(){ openVehicleModal(b.dataset.openVehicle); }); });
  $all("[data-open-invoice]").forEach(function(b){ b.addEventListener("click", function(){ openInvoiceModal(b.dataset.openInvoice); }); });
  $all("[data-open-analysis]").forEach(function(b){ b.addEventListener("click", function(){ openAnalysisModal(b.dataset.openAnalysis); }); });
}

/* Toast avec bouton d'annulation (delai de grace de 5s), utilise pour les
   suppressions afin d'eviter un confirm() natif bloquant et sans recours. */
function toastWithUndo(msg, onUndo){
  var el = $("#toast");
  clearTimeout(toast._t);
  el.innerHTML = escapeHtml(msg) + ' <button type="button" class="toast-undo-btn">Annuler</button>';
  el.className = "toast";
  el.classList.remove("hidden");
  var done = false;
  var timer = setTimeout(function(){ if (!done){ done = true; el.classList.add("hidden"); } }, 5000);
  el.querySelector(".toast-undo-btn").addEventListener("click", function(){
    if (done) return;
    done = true;
    clearTimeout(timer);
    el.classList.add("hidden");
    onUndo();
  });
}

/* =====================================================================
   VUES GLOBALES
   ===================================================================== */
function renderAllVehiclesPanel(){
  var full = state.scope === "FULL";
  $("#main-content").innerHTML =
    '<section class="tab-panel active">' +
    (full ? '' : '<div class="bandeau bandeau-warn">🔒 Compte à accès Livraison — les montants et commissions sont masqués. Contactez l\'administrateur pour un accès complet si besoin.</div>') +
    '<div class="bandeau bandeau-info">Vue consolidée de tous les véhicules suivis, tous partenaires confondus.</div>' +
      '<div class="kpi-row" id="kpi-all"></div>' +
      '<div class="toolbar"><select id="all-filter-partner"><option value="">Tous les partenaires</option>' +
        state.partners.map(function(p){ return '<option value="'+p.id+'">'+escapeHtml(p.distributeur)+'</option>'; }).join("") + '</select>' +
        '<select id="all-filter-statut"><option value="">Tous statuts</option>' +
        Object.keys(STATUT_LABELS).map(function(k){ return '<option value="'+k+'">'+STATUT_LABELS[k]+'</option>'; }).join("") + '</select></div>' +
      '<div class="veh-search-toolbar"><input type="text" id="all-veh-search" placeholder="Rechercher (client, modèle, immatriculation, concession...)"></div>' +
      '<div class="table-wrap"><table class="data-table" id="table-all-vehicles"></table></div>' +
    '</section>';
  function refresh(){
    var pf = $("#all-filter-partner").value, sf = $("#all-filter-statut").value, q = $("#all-veh-search").value;
    var list = state.vehicles.filter(function(v){ if (pf && v.partnerId!==pf) return false; if (sf && v.statutLivraison!==sf) return false; return true; });
    renderVehicleTable("#table-all-vehicles", list, q, true, function(statut){ $("#all-filter-statut").value = statut; refresh(); }, refresh);
    var caHt = list.reduce(function(s,v){ return s+Number(v.montantHT||0); }, 0);
    var commission = full ? list.reduce(function(s,v){ var c=commissionFor(v); return s+(c||0); }, 0) : null;
    var livres = list.filter(function(v){ return v.statutLivraison==="LIVRE"; });
    $("#kpi-all").innerHTML = kpiCard("Véhicules réseau", list.length, "", "accent") +
      (full ? kpiCard("CA HT total", fmtMoney(caHt), "", "") + kpiCard("Commission réseau due", fmtMoney(commission), "", "warn") : lockedKpiCard("CA HT total") + lockedKpiCard("Commission réseau due")) +
      kpiCard("Véhicules livrés", livres.length, "sur "+list.length, "ok");
  }
  $("#all-filter-partner").addEventListener("change", refresh);
  $("#all-filter-statut").addEventListener("change", refresh);
  $("#all-veh-search").addEventListener("input", refresh);
  refresh();
}

function renderDashboardPanel(){
  var full = state.scope === "FULL";
  var caHt = state.vehicles.reduce(function(s,v){ return s+Number(v.montantHT||0); }, 0);
  var commission = full ? state.vehicles.reduce(function(s,v){ var c=commissionFor(v); return s+(c||0); }, 0) : null;
  var livres = state.vehicles.filter(function(v){ return v.statutLivraison==="LIVRE"; }).length;
  var commandes = state.vehicles.filter(function(v){ return v.statutLivraison==="COMMANDE"; }).length;
  var maxCa = 0;
  var rows = state.partners.map(function(p){
    var list = vehiclesForPartner(p.id);
    var ca = list.reduce(function(s,v){ return s+Number(v.montantHT||0); }, 0);
    maxCa = Math.max(maxCa, ca);
    return { label: shortName(p), ca: ca };
  }).sort(function(a,b){ return b.ca-a.ca; });
  var bars = rows.map(function(r){
    var pct = maxCa>0 ? Math.round(r.ca/maxCa*100) : 0;
    return '<div class="bar-row"><span class="bar-label" title="'+escapeAttr(r.label)+'">'+escapeHtml(r.label)+'</span><span class="bar-track"><span class="bar-fill" style="width:'+pct+'%"></span></span><span class="bar-val">'+(full?fmtMoney(r.ca):"—")+'</span></div>';
  }).join("");
  $("#main-content").innerHTML = '<section class="tab-panel active">' +
    '<div class="bandeau bandeau-info">Tableau de bord exécutif — synthèse consolidée du réseau ZEEV CARS x DRIVEWAY.</div>' +
    '<div class="kpi-row">' + kpiCard("Concessions partenaires", state.partners.length, "", "accent") +
      kpiCard("Véhicules suivis", state.vehicles.length, commandes+" en attente de livraison", "") +
      (full ? kpiCard("CA HT apporté", fmtMoney(caHt), "", "") + kpiCard("Commission réseau due", fmtMoney(commission), livres+" livré(s)", "warn") : "") + '</div>' +
    '<div class="card"><label style="margin-bottom:14px">'+(full?"CA HT apporté par concession":"Véhicules par concession (montants masqués)")+'</label><div class="bar-chart">' + (bars||'<span class="param-text">Aucune donnée.</span>') + '</div></div>' +
  '</section>';
}

/* =====================================================================
   PANEL ADMIN — ajout / retrait de concessions partenaires (propriétaire uniquement)
   Les taux de commission restent verrouilles cote Worker (PARTNER_CONTRACTS) :
   ce panel prepare l'annuaire (cote GitHub) + genere le code a coller dans le
   Worker pour activer le taux, exactement comme le fait le premier projet.
   ===================================================================== */
function renderAdminPanel(){
  if (!state.user || !state.user.isOwner){
    $("#main-content").innerHTML = '<section class="tab-panel active"><div class="bandeau bandeau-warn">Accès réservé au compte propriétaire (permission Admin GitHub sur le dépôt de données).</div></section>';
    return;
  }
  var cards = state.partners.map(function(p){
    var list = vehiclesForPartner(p.id);
    var signe = !!p.signedDate;
    return '<div class="admin-partner-card">' +
      '<div class="admin-partner-head">' +
        '<div><span class="admin-partner-key">'+escapeHtml(p.id.toUpperCase())+'</span><h4>'+escapeHtml(p.distributeur)+'</h4></div>' +
        '<button class="btn small danger-o" data-remove-partner="'+p.id+'">🗑 Retirer</button>' +
      '</div>' +
      '<div class="param-text">' +
        (p.siren ? 'SIREN '+escapeHtml(p.siren)+' — ' : '') + 'SIRET '+escapeHtml(p.siret)+'<br>' +
        escapeHtml(p.adresse)+', '+escapeHtml(p.cp)+' '+escapeHtml(p.ville)+'<br>' +
        list.length+' véhicule(s) enregistré(s)' +
      '</div>' +
      '<div style="display:flex;align-items:center;gap:8px;margin-top:10px;">' +
        '<label style="font-size:10.5px;color:var(--tx3);text-transform:uppercase;letter-spacing:.05em;">Taux</label>' +
        '<input type="number" step="0.01" min="0" value="'+(p.commissionPct!=null?p.commissionPct:0)+'" data-rate-input="'+p.id+'" style="width:80px;padding:5px 8px;background:var(--bg2);border:1px solid var(--line);border-radius:6px;color:var(--tx1);font-size:12px;">' +
        '<span class="param-text">%</span>' +
        '<button class="btn small" data-save-rate="'+p.id+'">Enregistrer le taux</button>' +
      '</div>' +
      '<span class="pill '+(signe?"ok":"warn")+'" style="margin-top:8px;display:inline-block;">'+(signe?"✓ Contrat signé":"⏳ Simulation / non signé")+'</span>' +
    '</div>';
  }).join("");

  var perimetreAlerts = state.partners.filter(function(p){
    if (!p.commissionPct || !p.marquesAutorisees || !p.marquesAutorisees.length) return false;
    var delivered = vehiclesForPartner(p.id).filter(function(v){ return v.statutLivraison === "LIVRE"; });
    return !delivered.some(function(v){
      var marque = (v.modele||"").trim().split(/\s+/)[0].toUpperCase();
      return p.marquesAutorisees.indexOf(marque) !== -1;
    });
  }).map(function(p){
    return escapeHtml(p.distributeur) + " — commission " + fmtPct(p.commissionPct) + " active sur " + escapeHtml(p.marquesAutorisees.join("/")) + ", mais aucun véhicule livré de cette marque.";
  });

  var auditLog = ((state.meta||{}).auditLog || []).slice().reverse().slice(0, 30);
  var auditRows = auditLog.map(function(a){
    return "<tr><td>"+isoToFr(a.date.slice(0,10))+" "+a.date.slice(11,16)+"</td><td>"+escapeHtml(a.user)+"</td><td>"+escapeHtml(a.client||"—")+"</td><td>"+escapeHtml(a.field)+"</td><td class='num'>"+(a.from==null?"—":escapeHtml(String(a.from)))+"</td><td class='num'>"+(a.to==null?"—":escapeHtml(String(a.to)))+"</td></tr>";
  }).join("");

  $("#main-content").innerHTML = '<section class="tab-panel active">' +
    '<div class="bandeau bandeau-err">⚠ <strong>Zone d\'administration — accès restreint au compte propriétaire.</strong><br>Les modifications effectuées ici impactent directement l\'application pour tous les utilisateurs.</div>' +
    (perimetreAlerts.length ? '<div class="bandeau bandeau-warn">⚠ <strong>Incohérence commission / marque livrée :</strong><br>' + perimetreAlerts.join("<br>") + '</div>' : '') +
    '<div class="admin-section-title">Journal d\'audit — modifications financières (30 dernières)</div>' +
    (auditRows ? '<div class="table-wrap"><table class="data-table"><thead><tr><th>Date</th><th>Utilisateur</th><th>Client</th><th>Champ</th><th>Ancienne valeur</th><th>Nouvelle valeur</th></tr></thead><tbody>' + auditRows + '</tbody></table></div>' : '<span class="param-text">Aucune modification financière enregistrée pour le moment.</span>') +
    '<div class="admin-section-title">Concessions enregistrées</div>' +
    '<div class="admin-partner-list">' + (cards || '<span class="param-text">Aucune concession enregistrée.</span>') + '</div>' +
    '<div class="admin-section-title">Ajouter une nouvelle concession / fournisseur</div>' +
    '<div class="form-section-title">① Identité</div>' +
    '<div class="form-grid">' +
      '<div class="form-row"><label>Clé interne (sans espace, ex : RENAULT_FRANCE)</label><input type="text" id="adm-key" placeholder="EX_CONCESSION"></div>' +
      '<div class="form-row"><label>Nom affiché dans l\'app</label><input type="text" id="adm-nom" placeholder="Ex : Renault France"></div>' +
    '</div>' +
    '<div class="form-grid">' +
      '<div class="form-row"><label>Raison sociale (facturation)</label><input type="text" id="adm-raison" placeholder="Ex : RENAULT SAS"></div>' +
      '<div class="form-row"><label>Forme juridique (optionnel)</label><input type="text" id="adm-forme" placeholder="Ex : SAS au capital de 5 000 000 €"></div>' +
    '</div>' +
    '<div class="form-section-title">② Coordonnées</div>' +
    '<div class="form-grid">' +
      '<div class="form-row"><label>Adresse ligne 1</label><input type="text" id="adm-adresse1" placeholder="Ex : 13-15 Quai Le Gallo"></div>' +
      '<div class="form-row"><label>Adresse ligne 2 (CP + ville, pays)</label><input type="text" id="adm-adresse2" placeholder="Ex : 92100 Boulogne-Billancourt, FR"></div>' +
    '</div>' +
    '<div class="form-grid">' +
      '<div class="form-row"><label>SIREN</label><input type="text" id="adm-siren" placeholder="Ex : 780 129 987"></div>' +
      '<div class="form-row"><label>SIRET</label><input type="text" id="adm-siret" placeholder="Ex : 780 129 987 00012"></div>' +
    '</div>' +
    '<div class="form-section-title">③ Contrat</div>' +
    '<div class="form-row" style="max-width:220px;"><label>Taux de commission (%)</label><input type="number" step="0.01" min="0" id="adm-taux" placeholder="Ex : 1.00"></div>' +
    '<div class="modal-note">Le taux est actif immédiatement dès « Créer la concession » puis « Enregistrer sur GitHub » — aucun redéploiement du Worker n\'est nécessaire.</div>' +
    '<div class="actions-row">' +
      '<button class="btn" id="adm-preview-btn">👁 Aperçu du code (optionnel, pour archivage)</button>' +
      '<button class="btn primary" id="adm-create-btn">+ Créer la concession</button>' +
    '</div>' +
    '<div id="adm-snippet-wrap" class="hidden">' +
      '<div class="admin-section-title">Extrait de code (optionnel)</div>' +
      '<div class="modal-note">Ceci est purement informatif / pour archivage dans le dépôt de code — la concession et son taux sont déjà actifs sans cette étape.</div>' +
      '<textarea id="adm-snippet" readonly style="width:100%;min-height:140px;font-family:var(--mono);font-size:11.5px;background:var(--bg2);color:var(--tx1);border:1px solid var(--line);border-radius:8px;padding:12px;"></textarea>' +
    '</div>' +
  '</section>';

  $all("[data-remove-partner]").forEach(function(btn){
    btn.addEventListener("click", function(){
      var p = partnerById(btn.dataset.removePartner);
      if (!p) return;
      var idx = state.partners.findIndex(function(x){ return x.id === p.id; });
      if (idx === -1) return;
      var removed = state.partners[idx];
      state.partners.splice(idx, 1);
      markDirty();
      renderAll();
      var nbVehicles = vehiclesForPartner(p.id).length;
      toastWithUndo("Concession retirée" + (nbVehicles ? " (" + nbVehicles + " véhicule(s) restent rattachés dans les données)" : "") + ".", function(){
        state.partners.splice(idx, 0, removed);
        markDirty();
        renderAll();
      });
    });
  });

  $all("[data-save-rate]").forEach(function(btn){
    btn.addEventListener("click", function(){
      var p = partnerById(btn.dataset.saveRate);
      if (!p) return;
      var input = $('[data-rate-input="'+p.id+'"]');
      var val = parseFloat(input.value);
      if (isNaN(val) || val < 0){ toast("Taux invalide.", "err"); return; }
      p.commissionPct = val;
      markDirty();
      toast("Taux mis à jour pour " + p.distributeur + " — pensez à « Enregistrer sur GitHub ».", "ok");
      renderAll();
    });
  });

  function buildSnippet(){
    var key = $("#adm-key").value.trim().toUpperCase().replace(/\s+/g,"_");
    var nom = $("#adm-nom").value.trim();
    var raison = $("#adm-raison").value.trim() || nom;
    var adresse1 = $("#adm-adresse1").value.trim();
    var adresse2 = $("#adm-adresse2").value.trim();
    var siren = $("#adm-siren").value.trim();
    var siret = $("#adm-siret").value.trim();
    var taux = parseFloat($("#adm-taux").value) || 0;
    if (!key || !nom){ toast("Clé interne et nom affiché sont obligatoires.", "err"); return null; }
    var id = "p_" + key.toLowerCase();
    var cpVille = adresse2.split(",")[0] || "";
    var cp = (cpVille.match(/\d{4,5}/) || [""])[0];
    var ville = cpVille.replace(cp, "").trim();
    var partner = { id: id, distributeur: nom, adresse: adresse1, cp: cp, ville: ville, siren: siren, siret: siret, telephoneStandard: "", marques: [], contactNom: "", contactEmail: "", contactTel: "", commissionPct: taux };
    var snippet =
      "// Concession creee via le panel Admin (deja active, ceci est juste une trace) :\n" +
      "{ id:\"" + id + "\", distributeur:\"" + nom.replace(/"/g,'\\"') + "\", adresse:\"" + adresse1.replace(/"/g,'\\"') + "\", cp:\"" + cp + "\", ville:\"" + ville.replace(/"/g,'\\"') + "\", siren:\"" + siren + "\", siret:\"" + siret + "\", commissionPct:" + taux + " }\n" +
      "// Raison sociale (facturation) : " + raison;
    return { partner: partner, snippet: snippet };
  }

  $("#adm-preview-btn").addEventListener("click", function(){
    var built = buildSnippet();
    if (!built) return;
    $("#adm-snippet").value = built.snippet;
    $("#adm-snippet-wrap").classList.remove("hidden");
  });

  $("#adm-create-btn").addEventListener("click", function(){
    var built = buildSnippet();
    if (!built) return;
    if (partnerById(built.partner.id)){ toast("Cette clé interne existe déjà.", "err"); return; }
    state.partners.push(built.partner);
    markDirty();
    toast("Concession ajoutée avec son taux actif. Cliquez « Enregistrer sur GitHub » pour la conserver.", "ok");
    renderAll();
  });
}

/* =====================================================================
   MODALE — AJOUTER UN VEHICULE
   ===================================================================== */
function openVehicleModal(partnerId){
  var p = partnerById(partnerId);
  $("#vehicle-modal-partner-name").textContent = p ? p.distributeur : "—";
  ["v-leasing","v-numcontrat","v-client","v-conducteur","v-modele","v-immat","v-vin","v-date-prevue","v-date-effective","v-statut-source","v-remarques"].forEach(function(id){ $("#"+id).value=""; });
  $("#v-statut").value = "COMMANDE";
  $("#vehicle-submit-btn").dataset.partnerId = partnerId;
  openModal("modal-vehicle");
}
function submitVehicle(){
  var partnerId = $("#vehicle-submit-btn").dataset.partnerId;
  var client = $("#v-client").value.trim();
  if (!client){ toast("Le champ Client est obligatoire.", "err"); return; }
  state.vehicles.push({
    id: uid("v"), partnerId: partnerId,
    societeLeasing: $("#v-leasing").value.trim(), numeroContrat: $("#v-numcontrat").value.trim(),
    client: client, conducteur: normalizeConducteurName($("#v-conducteur").value.trim()),
    modele: $("#v-modele").value.trim(), immatriculation: $("#v-immat").value.trim().toUpperCase(), vin: $("#v-vin").value.trim().toUpperCase(),
    statutLivraison: $("#v-statut").value,
    dateLivraisonPrevue: $("#v-date-prevue").value, dateLivraisonEffective: $("#v-date-effective").value,
    statutVehiculeSource: $("#v-statut-source").value.trim(), remarques: $("#v-remarques").value.trim(),
    montantHT: null, statutFacturation: "NON_FACTURE", statutPaiement: "A_VERIFIER", numeroFacture: "", dateFacture: ""
  });
  markDirty(); closeModal("modal-vehicle"); toast("Véhicule ajouté à la liste de travail.", "ok"); renderAll();
}

/* =====================================================================
   MODALE — DETAILS / MODIFICATION D'UN VEHICULE EXISTANT
   Remplace l'edition en ligne dans le tableau : tous les champs regroupes
   par section, plus lisible qu'un tableau a 16 colonnes editables.
   ===================================================================== */
function openVehicleDetailsModal(vehicleId){
  var v = state.vehicles.find(function(x){ return x.id === vehicleId; });
  if (!v) return;
  var full = state.scope === "FULL";
  $("#vd-save-btn").dataset.vehicleId = vehicleId;
  $("#vd-delete-btn").dataset.vehicleId = vehicleId;
  $("#vd-client").value = v.client || "";
  $("#vd-conducteur").value = v.conducteur || "";
  $("#vd-modele").value = v.modele || "";
  $("#vd-immat").value = v.immatriculation || "";
  $("#vd-vin").value = v.vin || "";
  $("#vd-statut").value = v.statutLivraison || "COMMANDE";
  $("#vd-date-prevue").value = v.dateLivraisonPrevue || "";
  $("#vd-date-effective").value = v.dateLivraisonEffective || "";
  $("#vd-statut-source").value = v.statutVehiculeSource || "";
  $("#vd-leasing").value = v.societeLeasing || "";
  $("#vd-numcontrat").value = v.numeroContrat || "";
  $("#vd-remarques").value = v.remarques || "";
  $("#vd-more-details").open = !!(v.societeLeasing || v.numeroContrat || v.remarques);

  $("#vd-financier-title").classList.toggle("hidden", !full);
  $("#vd-financier-fields").classList.toggle("hidden", !full);
  if (full){
    $("#vd-montant").value = v.montantHT != null ? v.montantHT : "";
    $("#vd-statut-paiement").value = v.statutPaiement || "A_VERIFIER";
    $("#vd-numfacture").value = v.numeroFacture || "";
    $("#vd-datefacture").value = v.dateFacture || "";
    updateMontantSuggestion(v.modele, vehicleId);
    $("#vd-modele").oninput = function(){ updateMontantSuggestion($("#vd-modele").value, vehicleId); };
  }
  openModal("modal-vehicle-details");
}
/* Suggere un prix catalogue en moyennant les vehicules deja saisis pour un
   modele strictement identique (hors le vehicule courant) -- aide a la
   coherence de saisie sans jamais l'imposer. */
function updateMontantSuggestion(modele, currentId){
  var el = $("#vd-montant-suggestion");
  if (!el) return;
  var key = (modele||"").trim().toLowerCase();
  if (!key){ el.textContent = ""; return; }
  var matches = state.vehicles.filter(function(v){ return v.id !== currentId && (v.modele||"").trim().toLowerCase() === key && v.montantHT != null; });
  if (!matches.length){ el.textContent = ""; return; }
  var avg = matches.reduce(function(s,v){ return s+Number(v.montantHT); }, 0) / matches.length;
  el.textContent = "Suggestion : " + fmtMoney(avg) + " (moyenne sur " + matches.length + " véhicule(s) identique(s))";
}
function saveVehicleDetails(){
  var id = $("#vd-save-btn").dataset.vehicleId;
  var v = state.vehicles.find(function(x){ return x.id === id; });
  if (!v) return;
  var client = $("#vd-client").value.trim();
  if (!client){ toast("Le champ Client est obligatoire.", "err"); return; }
  var newStatut = $("#vd-statut").value;
  var newMontant = null;
  if (state.scope === "FULL"){
    newMontant = parseFloat($("#vd-montant").value);
    if (isNaN(newMontant)) newMontant = null;
    if (newStatut === "LIVRE" && newMontant == null){
      toast("Impossible de marquer « Livré » sans montant HT renseigné (nécessaire au calcul de la commission).", "err");
      return;
    }
  }
  v.client = client;
  v.conducteur = normalizeConducteurName($("#vd-conducteur").value.trim());
  v.modele = $("#vd-modele").value.trim();
  v.immatriculation = $("#vd-immat").value.trim().toUpperCase();
  v.vin = $("#vd-vin").value.trim().toUpperCase();
  v.statutLivraison = newStatut;
  v.dateLivraisonPrevue = $("#vd-date-prevue").value;
  v.dateLivraisonEffective = $("#vd-date-effective").value;
  v.statutVehiculeSource = $("#vd-statut-source").value.trim();
  v.societeLeasing = $("#vd-leasing").value.trim();
  v.numeroContrat = $("#vd-numcontrat").value.trim();
  v.remarques = $("#vd-remarques").value.trim();
  var montantWarning = null;
  if (state.scope === "FULL"){
    if (newMontant != null) montantWarning = checkMontantOutlier(v, newMontant);
    v.montantHT = newMontant;
    v.statutPaiement = $("#vd-statut-paiement").value;
    v.numeroFacture = $("#vd-numfacture").value.trim();
    v.dateFacture = $("#vd-datefacture").value;
  }
  markDirty();
  closeModal("modal-vehicle-details");
  if (montantWarning) toast(montantWarning, "warn");
  else toast("Véhicule mis à jour.", "ok");
  renderAll();
}

/* Normalise "M.NOM", "NOM Prenom", "prenom nom"... vers "Prenom NOM".
   Heuristique simple, non garantie a 100% mais evite la plupart des
   incoherences de saisie (M.GUERARD, HAMON Benoit, Gavel Maxime...). */
function normalizeConducteurName(raw){
  if (!raw) return raw;
  var s = raw.replace(/^(mme|monsieur|mr|m)\.?\s*/i, "").trim();
  var parts = s.split(/\s+/).filter(Boolean);
  if (parts.length < 2) return parts[0] ? (parts[0].charAt(0).toUpperCase() + parts[0].slice(1).toLowerCase()) : raw;
  function isUpper(w){ return w === w.toUpperCase() && w !== w.toLowerCase(); }
  function cap(w){ return w.split(/(\s|-)/).map(function(p){ return p.length ? p.charAt(0).toUpperCase()+p.slice(1).toLowerCase() : p; }).join(""); }
  var first = parts[0], last = parts[parts.length-1];
  var firstUp = isUpper(first), lastUp = isUpper(last);
  if (firstUp && !lastUp) return cap(last) + " " + first.toUpperCase();  // SURNAME firstname
  if (lastUp && !firstUp) return cap(first) + " " + last.toUpperCase();  // Firstname SURNAME
  if (firstUp && lastUp) return cap(last) + " " + first.toUpperCase();  // SURNAME FIRSTNAME (motif de nos donnees officielles)
  return cap(first) + " " + last.toUpperCase();  // ni l'un ni l'autre : ordre naturel suppose
}

/* Avertissement non bloquant si le montant saisi devie fortement de la
   moyenne deja enregistree pour le meme modele chez le meme partenaire. */
function checkMontantOutlier(vehicle, newMontant){
  var sameModel = state.vehicles.filter(function(v){
    return v.id !== vehicle.id && v.partnerId === vehicle.partnerId && v.modele === vehicle.modele && v.montantHT != null;
  });
  if (sameModel.length < 2) return null;
  var avg = sameModel.reduce(function(s,v){ return s + v.montantHT; }, 0) / sameModel.length;
  if (avg <= 0) return null;
  var deviation = Math.abs(newMontant - avg) / avg;
  if (deviation > 0.3){
    return "⚠ Véhicule enregistré, mais ce montant (" + fmtMoney(newMontant) + ") diffère de plus de 30 % de la moyenne pour ce modèle (" + fmtMoney(avg) + ") — à vérifier.";
  }
  return null;
}

/* =====================================================================
   MODALE — FACTURE CONSOLIDEE + IMPRESSION
   ===================================================================== */
/* Suggere le prochain numero de facture en detectant le motif prefixe+numero
   sur le plus grand numero deja utilise, tous vehicules confondus. */
function suggestNextInvoiceNumber(){
  var parsed = state.vehicles.map(function(v){ return v.numeroFacture; }).filter(Boolean).map(function(n){
    var m = String(n).match(/^(.*?)(\d+)$/);
    return m ? { prefix:m[1], num:parseInt(m[2],10), width:m[2].length } : null;
  }).filter(Boolean);
  if (!parsed.length) return "";
  parsed.sort(function(a,b){ return b.num-a.num; });
  var top = parsed[0];
  return top.prefix + String(top.num+1).padStart(top.width, "0");
}
function isDuplicateInvoiceNumber(num){
  return state.vehicles.some(function(v){ return v.numeroFacture && v.numeroFacture === num; });
}
function openInvoiceModal(partnerId){
  var p = partnerById(partnerId);
  $("#invoice-modal-partner-name").textContent = p ? p.distributeur : "—";
  $("#invoice-print-btn").dataset.partnerId = partnerId;
  $("#inv-number").value = suggestNextInvoiceNumber(); $("#inv-date").value = todayIso(); $("#inv-due").value = addDaysIso(todayIso(), 15);
  var list = vehiclesForPartner(partnerId).filter(function(v){ return v.statutLivraison === "LIVRE"; });
  var head = "<thead><tr><th class='chk'></th><th>Client</th><th>Modèle / Version</th><th>Immatriculation</th><th>Montant HT</th></tr></thead>";
  var body = list.length ? list.map(function(v){
    return "<tr><td class='chk'><input type='checkbox' class='inv-chk' data-amount='"+(v.montantHT||0)+"' data-veh-id='"+v.id+"' checked></td><td>"+escapeHtml(v.client)+"</td><td>"+escapeHtml(v.modele)+"</td><td>"+escapeHtml(v.immatriculation)+"</td><td class='num'>"+fmtMoney(v.montantHT)+"</td></tr>";
  }).join("") : "<tr class='empty-row'><td colspan='5'>Aucun véhicule livré pour cette concession pour le moment.</td></tr>";
  $("#invoice-vehicle-table").innerHTML = head + "<tbody>" + body + "</tbody>";
  $all("#invoice-vehicle-table .inv-chk").forEach(function(c){ c.addEventListener("change", updateInvoiceTotal); });
  $("#inv-number").oninput = function(){
    var warn = $("#inv-number-warn");
    if (isDuplicateInvoiceNumber($("#inv-number").value.trim())) warn.textContent = "⚠ Ce numéro est déjà utilisé sur un autre véhicule.";
    else warn.textContent = "";
  };
  updateInvoiceTotal();
  openModal("modal-invoice");
}
function updateInvoiceTotal(){
  var total = $all("#invoice-vehicle-table .inv-chk:checked").reduce(function(s,c){ return s+parseFloat(c.dataset.amount||0); }, 0);
  $("#invoice-total").textContent = fmtMoney(total);
  var partnerId = $("#invoice-print-btn").dataset.partnerId;
  var p = partnerById(partnerId);
  var line = $("#invoice-commission-line");
  if (state.scope === "FULL" && p && p.commissionPct != null){
    line.textContent = " — Commission estimée (" + fmtPct(p.commissionPct) + ") : " + fmtMoney(total * p.commissionPct / 100);
  } else {
    line.textContent = "";
  }
}
async function printInvoice(){
  var partnerId = $("#invoice-print-btn").dataset.partnerId;
  var p = partnerById(partnerId);
  var rows = $all("#invoice-vehicle-table tbody tr").filter(function(tr){ var c=tr.querySelector(".inv-chk"); return c && c.checked; });
  if (!rows.length){ toast("Sélectionnez au moins un véhicule à facturer.", "err"); return; }
  var number = $("#inv-number").value.trim() || "—", dateEm = isoToFr($("#inv-date").value) || isoToFr(todayIso()), dateEch = isoToFr($("#inv-due").value) || isoToFr(addDaysIso(frToIso(dateEm) || todayIso(), 15));
  if (number !== "—" && isDuplicateInvoiceNumber(number)){
    toast("Ce numéro de facture existe déjà sur un autre véhicule. Choisissez-en un autre.", "err");
    return;
  }
  var total = 0;
  var lines = rows.map(function(tr){
    var cells = tr.querySelectorAll("td");
    total += parseFloat(tr.querySelector(".inv-chk").dataset.amount||0);
    return "<tr><td>"+escapeHtml(cells[1].textContent)+"</td><td>"+escapeHtml(cells[2].textContent)+"</td><td>"+escapeHtml(cells[3].textContent)+"</td><td style='text-align:right'>"+escapeHtml(cells[4].textContent)+"</td></tr>";
  }).join("");
  var commission = p && p.commissionPct != null ? total * p.commissionPct / 100 : null;

  // Coordonnees bancaires : recuperees a la volee depuis le Worker (secrets
  // chiffres, jamais stockees dans le JSON GitHub ni dans le state du
  // navigateur). Reserve aux comptes FULL cote serveur -- un compte
  // LIVRAISON_ONLY n'atteint jamais cette section (bouton facture absent).
  var bankBlock = "";
  try {
    var pay = await apiGetPaymentDetails();
    bankBlock = "<p style='margin-top:14px'><strong>Coordonnées de paiement</strong><br>" +
      escapeHtml(pay.beneficiaire) + "<br>IBAN " + escapeHtml(pay.iban) + " — BIC " + escapeHtml(pay.bic) + "</p>";
  } catch(e){
    bankBlock = "<p style='margin-top:14px;color:#b00'>⚠ Coordonnées de paiement non configurées — à ajouter manuellement avant envoi.</p>";
  }

  // Tamponne le numero de facture sur chaque vehicule facture : previent la
  // reutilisation du meme numero et alimente automatiquement le suivi de
  // paiement (statutFacturation, dateFacture) deja utilise ailleurs dans l'app.
  if (number !== "—"){
    rows.forEach(function(tr){
      var vehId = tr.querySelector(".inv-chk").dataset.vehId;
      var v = state.vehicles.find(function(x){ return x.id === vehId; });
      if (v){ v.numeroFacture = number; v.dateFacture = frToIso(dateEm); v.statutFacturation = "FACTURE"; }
    });
    markDirty();
  }

  $("#print-invoice-area").innerHTML =
    "<h2>Facture consolidée n° "+escapeHtml(number)+"</h2>" +
    "<p><strong>"+escapeHtml(p.distributeur.toUpperCase())+"</strong><br>"+escapeHtml(p.adresse)+", "+escapeHtml(p.cp)+" "+escapeHtml(p.ville)+"<br>SIRET "+escapeHtml(p.siret)+"</p>" +
    "<p>Date d'émission : "+dateEm+" — Date d'échéance : "+dateEch+"</p>" +
    "<table><thead><tr><th>Client</th><th>Modèle / Version</th><th>Immatriculation</th><th>Montant HT</th></tr></thead><tbody>"+lines+"</tbody></table>" +
    "<p style='text-align:right;margin-top:14px'><strong>Total HT sélectionné : "+fmtMoney(total)+"</strong>" +
    (commission!=null ? "<br>Commission ("+fmtPct(p.commissionPct)+") : "+fmtMoney(commission) : "") + "</p>" +
    bankBlock;
  document.body.classList.add("print-invoice-mode");
  window.print();
  setTimeout(function(){ document.body.classList.remove("print-invoice-mode"); }, 500);
}

/* =====================================================================
   MODALE — ANALYSE PARC CLIENT
   ===================================================================== */
/* Categorisation heuristique du vehicule a partir de son libelle "modele".
   Priorite 1 : suffixe entre parentheses explicite, ex "(SUV)", "(Berline)".
   Priorite 2 : dictionnaire de modeles connus.
   Repli : "Véhicule particulier". Non garanti a 100%, ameliorable au cas
   par cas si besoin (les categories mal detectees restent visibles et
   modifiables en renommant le modele dans la fiche du vehicule). */
function categorizeVehicle(modeleRaw){
  var modele = (modeleRaw || "").toUpperCase();
  var m = modele.match(/\(([^)]+)\)\s*$/);
  if (m){
    var tag = m[1].trim();
    if (/UTILITAIRE|VAN|FOURGON/.test(tag)) return "Utilitaire";
    if (/SUV/.test(tag)) return "SUV";
    if (/BERLINE/.test(tag)) return "Berline";
    if (/MONOSPACE/.test(tag)) return "Monospace";
    if (/CITADINE/.test(tag)) return "Citadine";
    if (/BREAK/.test(tag)) return "Break";
    if (/COUPE|COUPÉ/.test(tag)) return "Coupé";
    return tag.charAt(0) + tag.slice(1).toLowerCase();
  }
  var dict = [
    [/SYMBIOZ|AUSTRAL|KUGA|SPORTAGE|ARONA|T-ROC|TIGUAN|X1\b|X3\b|Q3\b|3008|5008|C5 AIRCROSS|KAROQ|ATECA/, "SUV"],
    [/SCENIC|TOURAN|ZAFIRA|PICASSO|VERSO/, "Monospace"],
    [/S[ée]RIE 3|S[ée]RIE 5|CLASSE C|CLASSE E|A4\b|A6\b|MODEL 3|MODEL S|INSIGNIA|TALISMAN/, "Berline"],
    [/CLIO|208\b|POLO|C3\b|YARIS|COROLLA|IBIZA|FABIA|MAZDA2/, "Citadine"],
    [/MAXUS|UTILITAIRE|FOURGON|JUMPY|EXPERT|TRAFIC|TRANSIT|BOXER|DUCATO|VIVARO/, "Utilitaire"],
  ];
  for (var i=0;i<dict.length;i++){ if (dict[i][0].test(modele)) return dict[i][1]; }
  return "Véhicule particulier";
}

/* Camembert SVG pur (sans librairie externe) + legende, a partir d'une
   liste ordonnee de categories et de leur poids en valeur. */
var PIE_COLORS = ["#1D4ED8","#16A34A","#C98A14","#DC2626","#7C3AED","#0EA5E9","#DB2777","#059669","#EA580C","#4338CA"];
function polarToCartesian(cx, cy, r, angleDeg){
  var rad = (angleDeg - 90) * Math.PI / 180;
  return { x: cx + r * Math.cos(rad), y: cy + r * Math.sin(rad) };
}
function describeArcSlice(cx, cy, r, startAngle, endAngle){
  if (endAngle - startAngle >= 359.999){ endAngle = startAngle + 359.999; } // cercle complet (1 seule categorie)
  var start = polarToCartesian(cx, cy, r, endAngle);
  var end = polarToCartesian(cx, cy, r, startAngle);
  var largeArc = (endAngle - startAngle) <= 180 ? "0" : "1";
  return ["M", cx, cy, "L", start.x.toFixed(2), start.y.toFixed(2), "A", r, r, 0, largeArc, 0, end.x.toFixed(2), end.y.toFixed(2), "Z"].join(" ");
}
function buildPieChart(cats, byCat, valeurTotale){
  var cx = 100, cy = 100, r = 90;
  var angle = 0;
  var slices = "";
  var legend = "";
  cats.forEach(function(c, i){
    var d = byCat[c];
    var pct = valeurTotale > 0 ? (d.valeur / valeurTotale * 100) : (100 / cats.length);
    var sweep = pct / 100 * 360;
    var color = PIE_COLORS[i % PIE_COLORS.length];
    slices += '<path d="'+describeArcSlice(cx, cy, r, angle, angle + sweep)+'" fill="'+color+'" stroke="var(--panel)" stroke-width="2"></path>';
    angle += sweep;
    legend += '<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">' +
      '<span style="width:12px;height:12px;border-radius:3px;background:'+color+';flex-shrink:0;"></span>' +
      '<span style="font-size:12px;color:var(--tx1);">'+escapeHtml(c)+'</span>' +
      '<span style="font-size:11.5px;color:var(--tx3);margin-left:auto;">'+(Math.round(pct*10)/10)+' %</span>' +
    '</div>';
  });
  return '<div style="display:flex;gap:28px;align-items:center;flex-wrap:wrap;margin-bottom:18px;">' +
    '<svg viewBox="0 0 200 200" width="180" height="180" style="flex-shrink:0;">'+slices+'</svg>' +
    '<div style="flex:1;min-width:180px;">'+legend+'</div>' +
  '</div>';
}

function openAnalysisModal(partnerId){
  var p = partnerById(partnerId);
  var full = state.scope === "FULL";
  $("#analysis-modal-partner-name").textContent = (p ? p.distributeur : "—") + " — Tableau de bord exécutif";
  var list = vehiclesForPartner(partnerId);

  if (!list.length){
    $("#analysis-modal-body").innerHTML = '<p class="param-text">Aucun véhicule saisi pour cette concession pour le moment.</p>';
    openModal("modal-analysis");
    return;
  }

  var withValue = list.filter(function(v){ return v.montantHT != null; });
  var valeurTotale = withValue.reduce(function(s,v){ return s + v.montantHT; }, 0);
  var valeurMoyenne = withValue.length ? valeurTotale / withValue.length : 0;

  var byCat = {};
  list.forEach(function(v){
    var cat = categorizeVehicle(v.modele);
    if (!byCat[cat]) byCat[cat] = { count:0, valeur:0, countValeur:0 };
    byCat[cat].count++;
    if (v.montantHT != null){ byCat[cat].valeur += v.montantHT; byCat[cat].countValeur++; }
  });
  var cats = Object.keys(byCat).sort(function(a,b){ return byCat[b].valeur - byCat[a].valeur; });
  var topCat = cats[0];

  var html = '';

  html += '<div class="kpi-row">' +
    kpiCard("Véhicules au parc", list.length, "", "accent") +
    (full ? kpiCard("Valeur totale du parc HT", fmtMoney(valeurTotale), "hors taxes, hors assurances", "") : lockedKpiCard("Valeur totale du parc HT")) +
    (full ? kpiCard("Valeur moyenne / véhicule", fmtMoney(valeurMoyenne), topCat||"", "warn") : lockedKpiCard("Valeur moyenne / véhicule")) +
  '</div>';

  if (full && cats.length){
    html += '<div class="admin-section-title">Répartition par catégorie</div>';
    html += buildPieChart(cats, byCat, valeurTotale);

    html += '<div class="table-wrap"><table class="data-table"><thead><tr><th>Catégorie</th><th>Véh.</th><th>Valeur totale HT</th><th>Valeur moy./véh.</th><th>% parc</th><th>% valeur</th></tr></thead><tbody>';
    cats.forEach(function(c){
      var d = byCat[c];
      var pctParc = Math.round(d.count/list.length*1000)/10;
      var pctValeur = valeurTotale > 0 ? Math.round(d.valeur/valeurTotale*1000)/10 : 0;
      var moyCat = d.countValeur ? d.valeur/d.countValeur : 0;
      html += '<tr><td>'+escapeHtml(c)+'</td><td class="num">'+d.count+'</td><td class="num">'+fmtMoney(d.valeur)+'</td><td class="num">'+fmtMoney(moyCat)+'</td><td class="num">'+pctParc+' %</td><td class="num">'+pctValeur+' %</td></tr>';
    });
    html += '<tr style="font-weight:700;background:var(--panel2);"><td>Total</td><td class="num">'+list.length+'</td><td class="num">'+fmtMoney(valeurTotale)+'</td><td class="num">'+fmtMoney(valeurMoyenne)+'</td><td class="num">100 %</td><td class="num">100 %</td></tr>';
    html += '</tbody></table></div>';

    html += '<div class="admin-section-title">📊 Synthèse commerciale — support de négociation (direction achats / gestionnaire de flotte)</div>';
    html += '<div class="modal-note">' +
      '<strong>Nom de la concession :</strong> ' + escapeHtml(p.distributeur.toUpperCase()) + '<br>' +
      '<strong>Volume total du parc :</strong> ' + list.length + ' véhicule(s)<br>' +
      '<strong>Valeur financière totale :</strong> ' + fmtMoney(valeurTotale) + '<br>' +
      '<strong>Valeur moyenne par véhicule :</strong> ' + fmtMoney(valeurMoyenne) + '<br>' +
      '<strong>Catégories identifiées :</strong> ' + cats.length + ' famille(s) distincte(s)<br>' +
      cats.map(function(c, i){ var d = byCat[c]; return '<strong>Catégorie n°'+(i+1)+' :</strong> ' + escapeHtml(c) + ' (' + d.count + ' véh. — ' + fmtMoney(d.valeur) + ')'; }).join('<br>') + '<br>' +
      '<strong>Potentiel de renouvellement :</strong> Annuel — flotte de ' + list.length + ' véhicule(s) géré(s)' +
    '</div>';
  } else if (!full){
    html += '<div class="bandeau bandeau-warn">🔒 Répartition par catégorie et synthèse commerciale réservées aux comptes à accès complet.</div>';
  }

  $("#analysis-modal-body").innerHTML = html;
  openModal("modal-analysis");
}

/* =====================================================================
   COLLABORATEURS
   ===================================================================== */
function openCollabModal(){
  openModal("modal-collab");
  var list = $("#collab-list"); list.innerHTML = "<li>Chargement...</li>";
  apiGetCollaborators().then(function(r){
    $("#collab-manage-link").href = "https://github.com/"+r.repoOwner+"/"+r.repoName+"/settings/access";
    if (!r.collaborators.length){ list.innerHTML = "<li>Aucun collaborateur trouvé.</li>"; return; }
    list.innerHTML = r.collaborators.map(function(c){
      return '<li><img class="collab-avatar" src="'+escapeAttr(c.avatarUrl)+'" alt=""><span>'+escapeHtml(c.login)+'</span><span class="pill '+(c.scope==="FULL"?"ok":"warn")+'" style="margin-left:auto">'+(c.scope||"aucun accès")+'</span></li>';
    }).join("");
  }).catch(function(e){ list.innerHTML = "<li>"+escapeHtml(e.message)+"</li>"; });
}

/* =====================================================================
   MODALES GENERIQUES / THEME / SAUVEGARDE
   ===================================================================== */
function openModal(id){ $("#"+id).classList.remove("hidden"); }
function closeModal(id){ $("#"+id).classList.add("hidden"); }

function toggleTheme(){
  state.theme = state.theme==="dark" ? "light" : "dark";
  document.documentElement.setAttribute("data-theme", state.theme);
  $("#theme-btn").textContent = state.theme==="dark" ? "Thème clair" : "Thème sombre";
  try { localStorage.setItem("zcd_theme", state.theme); } catch(e){}
}

function showConflictBanner(){
  if (document.getElementById("conflict-banner")) return;
  var el = document.createElement("div");
  el.id = "conflict-banner";
  el.className = "bandeau bandeau-err";
  el.style.cssText = "position:fixed;top:0;left:0;right:0;z-index:9999;display:flex;justify-content:center;align-items:center;gap:14px;padding:10px 16px;border-radius:0";
  el.innerHTML = "⚠ Vos modifications ne sont plus synchronisées : quelqu'un d'autre a enregistré entre-temps. <button type='button' class='btn small' id='conflict-reload-btn'>Recharger</button>";
  document.body.prepend(el);
  document.getElementById("conflict-reload-btn").addEventListener("click", reloadAll);
}
async function saveAll(){
  $("#sync-status").textContent = "Enregistrement...";
  try {
    var res = await apiSaveData({ partners: state.partners, vehicles: state.vehicles, deletedVehicleIds: state.deletedVehicleIds });
    state.deletedVehicleIds = [];
    markClean();
    if (res && res.warnings && res.warnings.length) toast("Enregistré. ⚠ " + res.warnings.join(" — "), "warn");
    else toast("Données enregistrées.", "ok");
  } catch(e){
    if (e.status === 409){ showConflictBanner(); return; }
    toast("Échec de l'enregistrement : " + e.message, "err");
    var s = $("#sync-status"); s.textContent = "Échec de l'enregistrement"; s.style.color = "var(--err)";
  }
}
async function autoSaveIfDirty(){
  if (!state.dirty) return;
  if (document.getElementById("conflict-banner")) return; // en conflit : on arrete de retenter en boucle silencieusement
  try {
    await apiSaveData({ partners: state.partners, vehicles: state.vehicles, deletedVehicleIds: state.deletedVehicleIds });
    state.deletedVehicleIds = [];
    markClean();
    toast("Sauvegarde automatique effectuée.", "ok");
  } catch(e){
    if (e.status === 409){ showConflictBanner(); return; }
    // Autre echec (reseau...) : on retentera au prochain cycle, l'utilisateur
    // garde la main via le bouton "Enregistrer sur GitHub" en cas de souci persistant.
  }
}
setInterval(autoSaveIfDirty, 60000);
async function reloadAll(){
  if (state.dirty && !confirm("Recharger écrasera vos modifications non enregistrées. Continuer ?")) return;
  // Rafraichissement complet de la page (recupere aussi tout nouveau code deploye),
  // pas seulement les donnees -- cache-buster pour eviter une version en cache.
  var url = new URL(window.location.href);
  url.searchParams.set("_r", Date.now().toString());
  window.location.href = url.toString();
}

function renderAll(){
  renderNavTabs();
  if (state.activeTab === "all") renderAllVehiclesPanel();
  else if (state.activeTab === "dashboard") renderDashboardPanel();
  else if (state.activeTab === "admin") renderAdminPanel();
  else {
    var p = partnerById(state.activeTab) || state.partners[0];
    if (p){ state.activeTab = p.id; renderPartnerPanel(p); }
    else $("#main-content").innerHTML = '<p class="param-text">Aucun partenaire dans l\'annuaire.</p>';
  }
  $("#footer-info").textContent = "ZEEV CARS x DRIVEWAY — " + state.partners.length + " partenaires · " + state.vehicles.length + " véhicule(s) suivis";
}

/* =====================================================================
   INITIALISATION
   ===================================================================== */
document.getElementById("login-btn").addEventListener("click", startGithubLogin);
document.getElementById("logout-btn").addEventListener("click", function(){ performLogout(false); });
document.getElementById("theme-btn").addEventListener("click", toggleTheme);
document.getElementById("reload-btn").addEventListener("click", reloadAll);
document.getElementById("save-btn").addEventListener("click", saveAll);
document.getElementById("print-btn").addEventListener("click", function(){ window.print(); });
window.addEventListener("afterprint", function(){ document.body.classList.remove("print-invoice-mode"); });
document.getElementById("collab-btn").addEventListener("click", openCollabModal);
document.getElementById("export-csv-btn").addEventListener("click", function(){ apiExportCsv().catch(function(e){ toast(e.message, "err"); }); });
document.getElementById("import-csv-btn").addEventListener("click", function(){ document.getElementById("import-csv-file").click(); });
document.getElementById("import-csv-file").addEventListener("change", function(e){ if (e.target.files[0]) importCsvFile(e.target.files[0]); e.target.value = ""; });
document.getElementById("vehicle-submit-btn").addEventListener("click", submitVehicle);
document.getElementById("vd-save-btn").addEventListener("click", saveVehicleDetails);
document.getElementById("vd-delete-btn").addEventListener("click", function(){
  var id = this.dataset.vehicleId;
  closeModal("modal-vehicle-details");
  deleteVehicleWithUndo(id);
});
document.getElementById("nav-search").addEventListener("input", function(){
  var q = this.value.trim().toLowerCase();
  var count = 0;
  $all("#nav-tabs .ntab").forEach(function(tab){
    var match = !q || tab.textContent.toLowerCase().indexOf(q) !== -1 || (tab.title||"").toLowerCase().indexOf(q) !== -1;
    tab.style.display = match ? "" : "none";
    if (match) count++;
  });
  $("#nav-search-count").textContent = q ? (count + " onglet(s) trouvé(s)") : "";
});
document.getElementById("invoice-select-all").addEventListener("click", function(){ $all("#invoice-vehicle-table .inv-chk").forEach(function(c){ c.checked=true; }); updateInvoiceTotal(); });
document.getElementById("invoice-print-btn").addEventListener("click", printInvoice);
$all("[data-close]").forEach(function(b){ b.addEventListener("click", function(){ closeModal(b.dataset.close); }); });
$all(".modal-overlay").forEach(function(ov){ ov.addEventListener("click", function(e){ if (e.target===ov) ov.classList.add("hidden"); }); });
window.addEventListener("beforeunload", function(e){ if (state.dirty){ e.preventDefault(); e.returnValue = ""; } });

(async function init(){
  try { var t = localStorage.getItem("zcd_theme"); if (t){ state.theme=t; document.documentElement.setAttribute("data-theme", t); document.getElementById("theme-btn").textContent = t==="dark"?"Thème clair":"Thème sombre"; } } catch(e){}
  if (!isApiConfigured()){
    toast("⚠ CONFIG.apiBase n'est pas encore renseigné — voir DEPLOIEMENT.md.", "err");
  }
  var handled = await handleGithubCallback();
  if (!handled){
    await tryBootFromCookie();
  }
})();

})();
