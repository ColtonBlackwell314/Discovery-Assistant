/* D365 Discovery Hub — extension app logic.
   Storage: chrome.storage.local (async), scoped to this extension only.
   The ONLY external network calls anywhere in this file are the Gemini API
   calls made from the Chat page, and only when you send a chat message. */

const NAV = [
  {key:"chat", name:"Chat", icon:"bi-chat-dots-fill"},
  {key:"notes", name:"Notes", icon:"bi-journal-text"},
  {key:"stakeholders", name:"Stakeholders", icon:"bi-people-fill"},
  {key:"teamstructure", name:"Team Structure", icon:"bi-diagram-3-fill"},
  {key:"transcripts", name:"Transcripts", icon:"bi-mic-fill"},
  {key:"apikey", name:"API Key", icon:"bi-key-fill"}
];
const NOTEBOOK_COLORS = ["#0f6cbd","#1d9e75","#d85a30","#993c1d","#7f77dd","#d4537e","#639922"];
const STORAGE_KEY = "d365_discovery_sessions";
const API_KEY_STORAGE = "gemini_api_key";
const WORKING_MODEL_STORAGE = "gemini_working_model";
// Google renames/retires free-tier models periodically (gemini-2.5-flash-lite
// was cut off for new API users in mid-2026). Try candidates in order and
// remember whichever one actually works, so this list is the only thing
// that ever needs updating going forward.
const MODEL_CANDIDATES = ["gemini-3.1-flash-lite", "gemini-2.5-flash", "gemini-flash-latest"];
function endpointFor(model){ return `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`; }

let allProjects = {};
let currentId = null;
let state = null;
let activeNav = "chat";
let activeItemId = null;
let activeChatId = null;
let geminiApiKey = "";
let sidebarCollapsed = true;
let expandedTreeItems = new Set(); // item ids currently expanded in the Notes rail
let noteFlushFn = null; // flushes the currently-mounted Notes editor's unsaved text; set by renderNotes, called by renderSection before it tears the page down

/* ---------------- Split-view workspace ----------------
   A browser extension can't spawn real separate OS windows the way Windows'
   own snap layouts do — this simulates the same drag-to-edge gesture inside
   the extension's own content area instead. It's only offered when this
   page is running as a full tab (isFullTabMode, resolved in init() the same
   way the existing maximize button detects that), never in the cramped
   default side panel. Each pane is a separate <iframe> pointed at this same
   index.html with ?pane=<navKey>&embed=1, which gives every pane a fully
   independent, fully interactive app instance (its own JS globals, its own
   DOM) with zero changes to any existing render function — far simpler and
   safer than threading per-pane state through the whole single-page app.
   Panes stay in sync with each other via chrome.storage.onChanged. */
let isFullTabMode = false;
let isEmbeddedPane = false;
const WORKSPACE_SLOTS = ["tl","tr","bl","br"];
let workspaceSlots = { tl:null, tr:null, bl:null, br:null }; // navKey or null per quadrant
let dragNavKey = null;

// Re-renders the current section without losing your scroll position or
// snapping the page back to the top — used for in-place edits (save, cancel,
// remove, toggling a card) instead of a plain renderSection().
function rerenderKeepScroll(){
  const mainEl = document.querySelector(".app-main");
  const top = mainEl ? mainEl.scrollTop : 0;
  renderSection();
  const mainEl2 = document.querySelector(".app-main");
  if(mainEl2) mainEl2.scrollTop = top;
}

function formatDateLabel(ts){
  if(!ts) return "";
  return new Date(ts).toLocaleDateString(undefined, { month:"short", day:"numeric", year:"numeric" });
}

function uid(prefix){ return prefix + "_" + Date.now() + "_" + Math.random().toString(36).slice(2,7); }
function parseTags(str){ return (str || "").split(",").map(t => t.trim()).filter(Boolean); }
function escapeHtml(s){ return (s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }

function newProjectState(name){
  return { name: name || "Untitled project", items: [], transcripts: [], chats: [], stakeholders: [], stakeholdersUpdatedAt: null, teamStructure: { nodes: [], edges: [] }, teamStructureUpdatedAt: null };
}

/* ---------------- storage (chrome.storage.local, async) ---------------- */
async function loadAllProjects(){
  const data = await chrome.storage.local.get(STORAGE_KEY);
  return data[STORAGE_KEY] || {};
}
async function saveAllProjects(all){ await chrome.storage.local.set({ [STORAGE_KEY]: all }); }
async function persist(){ allProjects[currentId] = state; await saveAllProjects(allProjects); }
async function loadApiKey(){ const d = await chrome.storage.local.get(API_KEY_STORAGE); return d[API_KEY_STORAGE] || ""; }
async function saveApiKey(key){ await chrome.storage.local.set({ [API_KEY_STORAGE]: key }); }

/* ---------------- .docx transcript import (local, no library) ---------------- */
async function inflateRaw(bytes){
  const ds = new DecompressionStream("deflate-raw");
  const writer = ds.writable.getWriter();
  writer.write(bytes);
  writer.close();
  const buf = await new Response(ds.readable).arrayBuffer();
  return new Uint8Array(buf);
}
function findEntryInZip(view, name){
  const bytes = new Uint8Array(view.buffer, view.byteOffset, view.byteLength);
  let eocdPos = -1;
  for(let i = bytes.length - 22; i >= 0; i--){ if(view.getUint32(i, true) === 0x06054b50){ eocdPos = i; break; } }
  if(eocdPos === -1) throw new Error("Not a valid .docx/zip file");
  const cdOffset = view.getUint32(eocdPos + 16, true);
  const entryCount = view.getUint16(eocdPos + 10, true);
  let p = cdOffset;
  for(let i = 0; i < entryCount; i++){
    if(view.getUint32(p, true) !== 0x02014b50) break;
    const method = view.getUint16(p + 10, true);
    const compSize = view.getUint32(p + 20, true);
    const nameLen = view.getUint16(p + 28, true);
    const extraLen = view.getUint16(p + 30, true);
    const commentLen = view.getUint16(p + 32, true);
    const localHeaderOffset = view.getUint32(p + 42, true);
    const fname = new TextDecoder().decode(bytes.subarray(p + 46, p + 46 + nameLen));
    if(fname === name){
      const lp = localHeaderOffset;
      const lNameLen = view.getUint16(lp + 26, true);
      const lExtraLen = view.getUint16(lp + 28, true);
      const dataStart = lp + 30 + lNameLen + lExtraLen;
      return { method, data: bytes.subarray(dataStart, dataStart + compSize) };
    }
    p += 46 + nameLen + extraLen + commentLen;
  }
  return null;
}
function docxXmlToText(xml){
  xml = xml.replace(/<w:tab\/>/g, "\t");
  xml = xml.replace(/<\/w:p>/g, "\n");
  xml = xml.replace(/<[^>]+>/g, "");
  xml = xml.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'");
  return xml.split("\n").map(l => l.trim()).filter(Boolean).join("\n");
}
async function extractDocxText(file){
  const buf = await file.arrayBuffer();
  const view = new DataView(buf);
  const entry = findEntryInZip(view, "word/document.xml");
  if(!entry) throw new Error("Could not find document content in this .docx file");
  let raw;
  if(entry.method === 0){ raw = entry.data; }
  else if(entry.method === 8){ raw = await inflateRaw(entry.data); }
  else { throw new Error("Unsupported compression in this .docx file"); }
  return docxXmlToText(new TextDecoder("utf-8").decode(raw));
}
function wireDocxImport(inputEl, statusEl, onText){
  inputEl.addEventListener("change", async () => {
    const file = inputEl.files[0];
    if(!file) return;
    statusEl.textContent = "Reading " + file.name + "...";
    try{
      const text = await extractDocxText(file);
      await onText(text, file.name);
      statusEl.textContent = "Imported " + file.name + " (" + text.length + " characters).";
    }catch(err){
      statusEl.textContent = "Couldn't read that file: " + err.message + ". Try exporting it as a plain .docx (not .doc) and re-uploading.";
    }
    inputEl.value = "";
  });
}

/* ---------------- Gemini chat (the only networked feature) ---------------- */
// Full "Folder / Sub-folder / Item" breadcrumb for an item in the Notes tree,
// so context sent to Gemini (and anything else reading this) reflects how
// notes are actually organized, not just each item's own name in isolation.
function itemPath(itemId){
  const byId = new Map((state.items || []).map(it => [it.id, it]));
  const parts = [];
  let cur = byId.get(itemId);
  const seen = new Set();
  while(cur && !seen.has(cur.id)){
    seen.add(cur.id);
    parts.unshift(cur.name);
    cur = cur.parentId ? byId.get(cur.parentId) : null;
  }
  return parts.join(" / ");
}

function buildProjectContext(){
  const notedItems = (state.items || []).filter(it => it.desc && it.desc.trim());
  const transcripts = state.transcripts || [];
  if(notedItems.length === 0 && transcripts.length === 0) return "No discovery notes have been captured for this project yet.";
  let out = "";
  notedItems.forEach(it => {
    out += `\n## Section: ${itemPath(it.id)}\n${it.desc}\n`;
  });
  if(transcripts.length){
    out += `\n## Transcripts (all client calls, this project)\n`;
    transcripts.forEach(t => { out += `- ${t.filename}: ${t.text.slice(0,4000)}\n`; });
  }
  return out.trim() || "No discovery notes have been captured for this project yet.";
}

async function loadWorkingModel(){ const d = await chrome.storage.local.get(WORKING_MODEL_STORAGE); return d[WORKING_MODEL_STORAGE] || ""; }
async function saveWorkingModel(model){ await chrome.storage.local.set({ [WORKING_MODEL_STORAGE]: model }); }

async function callGeminiWithModel(model, systemInstruction, contents, generationConfig){
  const body = { systemInstruction, contents };
  if(generationConfig) body.generationConfig = generationConfig;
  const res = await fetch(`${endpointFor(model)}?key=${encodeURIComponent(geminiApiKey)}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  const data = await res.json().catch(() => ({}));
  if(!res.ok){
    const msg = (data && data.error && data.error.message) || `Request failed (${res.status})`;
    const err = new Error(msg);
    err.status = res.status;
    throw err;
  }
  const candidate = data.candidates && data.candidates[0];
  const text = candidate && candidate.content && candidate.content.parts && candidate.content.parts.map(p => p.text).join("");
  if(!text) throw new Error("Gemini returned an empty response.");
  return text;
}

// Same model-fallback logic as callGemini, but asks Gemini to answer as pure
// JSON (responseMimeType) and parses it, for structured-extraction features
// like the Stakeholders refresh. Throws if the model never returns valid JSON.
async function callGeminiForJSON(systemInstructionText, userText){
  if(!geminiApiKey) throw new Error("No Gemini API key configured. Add one on the Overview page.");
  const systemInstruction = { parts: [{ text: systemInstructionText }] };
  const contents = [{ role: "user", parts: [{ text: userText }] }];
  const generationConfig = { responseMimeType: "application/json" };

  const preferred = await loadWorkingModel();
  const order = preferred ? [preferred, ...MODEL_CANDIDATES.filter(m => m !== preferred)] : MODEL_CANDIDATES;

  let lastErr;
  for(const model of order){
    try{
      const text = await callGeminiWithModel(model, systemInstruction, contents, generationConfig);
      if(model !== preferred) await saveWorkingModel(model);
      try{ return JSON.parse(text); }
      catch(parseErr){ throw new Error("Gemini's response wasn't valid JSON."); }
    }catch(err){
      lastErr = err;
      const unavailable = err.status === 404 || /no longer available|not found|not supported/i.test(err.message);
      if(!unavailable) throw err;
    }
  }
  throw lastErr || new Error("No working Gemini model found.");
}

async function callGemini(history){
  if(!geminiApiKey) throw new Error("No Gemini API key configured. Add one on the Overview page.");
  const systemInstruction = {
    parts: [{ text:
      `You are an assistant helping a Microsoft Dynamics 365 CRM consultant think through discovery for a specific client project called "${state.name}". ` +
      `Answer only using the discovery notes context below plus the conversation. If the notes don't cover something, say so rather than guessing. ` +
      `Discovery notes for this project:\n${buildProjectContext()}`
    }]
  };
  const contents = history.map(m => ({ role: m.role === "assistant" ? "model" : "user", parts: [{ text: m.text }] }));

  const preferred = await loadWorkingModel();
  const order = preferred ? [preferred, ...MODEL_CANDIDATES.filter(m => m !== preferred)] : MODEL_CANDIDATES;

  let lastErr;
  for(const model of order){
    try{
      const text = await callGeminiWithModel(model, systemInstruction, contents);
      if(model !== preferred) await saveWorkingModel(model);
      return text;
    }catch(err){
      lastErr = err;
      const unavailable = err.status === 404 || /no longer available|not found|not supported/i.test(err.message);
      if(!unavailable) throw err;
    }
  }
  throw lastErr || new Error("No working Gemini model found.");
}

/* ---------------- Stakeholders (AI-extracted from all notes/transcripts) ---------------- */
// Core Gemini-refresh logic, pulled out of the page UI so it can be called
// both by each page's own "Refresh from notes" button and by the single
// top-nav refresh action. Each returns nothing on success; throws on failure
// so callers can show their own error UI.
async function refreshStakeholdersFromNotes(){
  const systemInstructionText =
    `You extract a stakeholder directory from Microsoft Dynamics 365 CRM discovery notes for a project called "${state.name}". ` +
    `Read the notes and transcripts below. Identify every distinct named person mentioned. For each person, return their name, ` +
    `their job title/role if stated, the team or department they belong to (e.g. Sales, Service, IT) if stated or clearly implied by ` +
    `which section they appear under, and a list of short bullet-point responsibilities describing what they own, what they said, or what ` +
    `they're accountable for (e.g. "Primary owner of inbound lead intake", "Assigns leads to Sales Managers vs Sales Coordinators"). ` +
    `Only include information that is actually stated in the text — leave a field empty or as an empty array if it isn't mentioned, do not invent ` +
    `or guess titles/teams/responsibilities that aren't in the source text. Do not include generic mentions of "the team" or unnamed roles, only actual named people. ` +
    `Respond with ONLY a JSON array, no prose, in this exact shape: ` +
    `[{"name":"","title":"","team":"","responsibilities":["",""]}]`;
  const userText = buildProjectContext();
  const result = await callGeminiForJSON(systemInstructionText, userText);
  const list = Array.isArray(result) ? result : (Array.isArray(result.stakeholders) ? result.stakeholders : []);
  state.stakeholders = list
    .filter(p => p && p.name && p.name.trim())
    .map(p => ({
      name: String(p.name).trim(),
      title: String(p.title||"").trim(),
      team: String(p.team||"").trim(),
      responsibilities: Array.isArray(p.responsibilities) ? p.responsibilities.map(r => String(r).trim()).filter(Boolean) : []
    }));
  state.stakeholdersUpdatedAt = Date.now();
  await persist();
}

async function refreshTeamStructureFromNotes(){
  const systemInstructionText =
    `You infer an organizational/team structure diagram from Microsoft Dynamics 365 CRM discovery notes for a project called "${state.name}". ` +
    `Read the notes and transcripts below and identify teams, sub-teams, and roles, and how they report or roll up into each other. ` +
    `Return a node/edge graph. Each node has: "id" (short unique slug, no spaces), "label" (display name, can be multi-line using \\n), ` +
    `and "type" which is one of "team" (a team or sub-team), "role" (a job role, security role, or access group), or "question" ` +
    `(an open/unresolved structural question you noticed in the notes, e.g. "Combine Sales and Service reps?" — only include this type ` +
    `if the notes actually raise an open question, do not invent one). Each edge has "from" and "to" node ids, meaning "from" is the parent ` +
    `that "to" reports into / rolls up under. IMPORTANT: for "team" node labels, reuse the exact team names already used on the stakeholder ` +
    `directory below when they refer to the same team, so stakeholder-to-team matching stays consistent between the two. Only infer structure ` +
    `that is actually stated or clearly implied in the text — do not invent teams, roles, or reporting lines that aren't supported by the notes. ` +
    `If there isn't enough information to build a structure, return empty arrays. ` +
    `Respond with ONLY JSON, no prose, in this exact shape: ` +
    `{"nodes":[{"id":"","label":"","type":"team"}],"edges":[{"from":"","to":""}]}`;
  const stakeholderTeams = Array.from(new Set((state.stakeholders || []).map(p => p.team).filter(Boolean)));
  const userText = buildProjectContext() +
    (stakeholderTeams.length ? `\n\nKnown stakeholder team names already extracted (reuse these exact names for matching team nodes):\n${stakeholderTeams.join(", ")}` : "");
  const result = await callGeminiForJSON(systemInstructionText, userText);
  const nodes = Array.isArray(result.nodes) ? result.nodes : [];
  const edges = Array.isArray(result.edges) ? result.edges : [];
  const cleanNodes = nodes
    .filter(n => n && n.id && n.label)
    .map(n => ({ id: String(n.id).trim(), label: String(n.label).trim(), type: ["team","role","question"].includes(n.type) ? n.type : "team" }));
  const validIds = new Set(cleanNodes.map(n => n.id));
  const cleanEdges = edges
    .filter(e => e && validIds.has(e.from) && validIds.has(e.to) && e.from !== e.to)
    .map(e => ({ from: String(e.from).trim(), to: String(e.to).trim() }));
  state.teamStructure = { nodes: cleanNodes, edges: cleanEdges };
  state.teamStructureUpdatedAt = Date.now();
  await persist();
}

// Runs both refreshes in the order that keeps them consistent: stakeholders
// first (so team names are known), then team structure, which reuses those
// exact team names in its prompt so Team Structure's name-matching against
// the Stakeholders page actually lines up instead of drifting apart.
async function refreshAllFromNotes(onProgress){
  if(onProgress) onProgress("Refreshing stakeholders...");
  await refreshStakeholdersFromNotes();
  if(onProgress) onProgress("Refreshing team structure...");
  await refreshTeamStructureFromNotes();
}

function renderStakeholders(body){
  const lastRefreshed = state.stakeholdersUpdatedAt ? timeAgoLabel(state.stakeholdersUpdatedAt) : null;

  body.innerHTML = `
    <div class="d-flex justify-content-between align-items-start flex-wrap gap-2 mb-2">
      <div>
        <h2 style="margin:0">Stakeholders</h2>
        <div class="small text-secondary" id="stakeholdersMeta">
          ${lastRefreshed ? `Last refreshed ${lastRefreshed}` : "Not refreshed yet"}
        </div>
      </div>
      <button class="btn btn-primary btn-sm" id="refreshStakeholdersBtn"><i class="bi bi-arrow-clockwise"></i> Refresh from notes</button>
    </div>
    <div class="ai-note" id="stakeholdersStatus" style="display:none"></div>
    <div class="stakeholder-search-wrap">
      <i class="bi bi-search"></i>
      <input type="text" id="stakeholderSearch" class="form-control form-control-sm" placeholder="Search stakeholders by name, title, team, or responsibility...">
    </div>
    <div id="stakeholdersCards" class="stakeholder-grid"></div>
    <div class="empty" id="stakeholderNoResults" style="display:none">No stakeholders match "<span id="stakeholderNoResultsQuery"></span>".</div>
    <button class="btn btn-sm btn-outline-secondary mt-2" id="addStakeholderBtn"><i class="bi bi-plus-circle-fill"></i> Add stakeholder</button>
  `;

  const cardsWrap = document.getElementById("stakeholdersCards");
  const statusEl = document.getElementById("stakeholdersStatus");
  const searchInput = document.getElementById("stakeholderSearch");
  const noResultsEl = document.getElementById("stakeholderNoResults");
  let searchQuery = "";

  function groupByTeam(list){
    // expects each item to already carry its true index into state.stakeholders as _i
    const groups = new Map();
    list.forEach((p) => {
      const key = p.team && p.team.trim() ? p.team.trim() : "Unassigned";
      if(!groups.has(key)) groups.set(key, []);
      groups.get(key).push(p);
    });
    return groups;
  }

  function matchesSearch(p, q){
    if(!q) return true;
    const haystack = [p.name, p.title, p.team, ...(p.responsibilities||[])].join(" ").toLowerCase();
    return haystack.includes(q);
  }

  function renderCards(){
    const full = state.stakeholders || [];
    if(!full.length){
      cardsWrap.innerHTML = `<div class="empty">No stakeholders yet. Click "Refresh from notes" to have AI pull names, roles, teams, and responsibilities out of everything captured so far.</div>`;
      noResultsEl.style.display = "none";
      return;
    }
    const q = searchQuery.trim().toLowerCase();
    // keep original indices (_i) so edits/removes still target the right item in state.stakeholders,
    // even while a search filter is narrowing what's shown
    const list = full.map((p, i) => ({ ...p, _i: i })).filter(p => matchesSearch(p, q));
    if(!list.length){
      cardsWrap.innerHTML = "";
      noResultsEl.style.display = "";
      document.getElementById("stakeholderNoResultsQuery").textContent = searchQuery.trim();
      return;
    }
    noResultsEl.style.display = "none";
    const groups = groupByTeam(list);
    cardsWrap.innerHTML = "";
    groups.forEach((people, teamName) => {
      const section = document.createElement("div");
      section.className = "stakeholder-team-group";
      section.innerHTML = `<h3 class="stakeholder-team-heading">${escapeHtml(teamName)}</h3>
        <div class="stakeholder-card-row">
          ${people.map(p => `
            <div class="stakeholder-card" data-i="${p._i}">
              <div class="d-flex justify-content-between align-items-start">
                <input type="text" class="sh-field sh-name" value="${escapeHtml(p.name || "")}" placeholder="Name">
                <button class="btn btn-sm btn-outline-danger remove-stakeholder" data-i="${p._i}" title="Remove"><i class="bi bi-trash-fill"></i></button>
              </div>
              <div class="stakeholder-role-line">
                <span class="text-secondary small">Role:</span>
                <input type="text" class="sh-field sh-title" value="${escapeHtml(p.title || "")}" placeholder="Title / role">
              </div>
              <div class="stakeholder-team-line">
                <span class="text-secondary small">Team:</span>
                <input type="text" class="sh-field sh-team" value="${escapeHtml(p.team || "")}" placeholder="Team">
              </div>
              <div class="text-secondary small mt-2 mb-1">Responsibilities:</div>
              <ul class="stakeholder-resp-list" data-i="${p._i}">
                ${(p.responsibilities && p.responsibilities.length ? p.responsibilities : [""]).map((r, ri) => `
                  <li><input type="text" class="sh-field sh-resp" data-ri="${ri}" value="${escapeHtml(r)}" placeholder="Responsibility"></li>
                `).join("")}
              </ul>
              <button class="btn btn-sm btn-link p-0 add-resp" data-i="${p._i}">+ add line</button>
            </div>`).join("")}
        </div>`;
      cardsWrap.appendChild(section);
    });

    cardsWrap.querySelectorAll(".stakeholder-card").forEach(card => {
      const i = Number(card.dataset.i);
      const save = async () => {
        const resp = Array.from(card.querySelectorAll(".sh-resp")).map(inp => inp.value.trim()).filter(Boolean);
        state.stakeholders[i] = {
          name: card.querySelector(".sh-name").value.trim(),
          title: card.querySelector(".sh-title").value.trim(),
          team: card.querySelector(".sh-team").value.trim(),
          responsibilities: resp
        };
        await persist();
      };
      card.querySelectorAll(".sh-field").forEach(inp => { inp.onchange = save; });
      card.querySelector(".remove-stakeholder").onclick = async () => {
        state.stakeholders.splice(i, 1);
        await persist();
        renderCards();
      };
      card.querySelector(".add-resp").onclick = async () => {
        if(!state.stakeholders[i].responsibilities) state.stakeholders[i].responsibilities = [];
        state.stakeholders[i].responsibilities.push("");
        await persist();
        renderCards();
      };
    });
  }
  renderCards();

  searchInput.oninput = () => { searchQuery = searchInput.value; renderCards(); };

  document.getElementById("addStakeholderBtn").onclick = async () => {
    state.stakeholders.push({ name:"", title:"", team:"", responsibilities:[] });
    await persist();
    renderCards();
  };

  document.getElementById("refreshStakeholdersBtn").onclick = async () => {
    const btn = document.getElementById("refreshStakeholdersBtn");
    btn.disabled = true;
    btn.innerHTML = `<span class="spinner-border spinner-border-sm"></span> Refreshing...`;
    statusEl.style.display = "";
    statusEl.textContent = "Asking Gemini to scan all notes and transcripts for this project...";
    try{
      await refreshStakeholdersFromNotes();
      statusEl.style.display = "none";
      rerenderKeepScroll();
    }catch(err){
      statusEl.textContent = "Couldn't refresh stakeholders: " + err.message;
    }finally{
      btn.disabled = false;
      btn.innerHTML = `<i class="bi bi-arrow-clockwise"></i> Refresh from notes`;
    }
  };
}

/* ---------------- Team Structure (AI-inferred org chart) ----------------
   Renders a hierarchical box-and-arrow diagram from Gemini-extracted
   nodes/edges as hand-built SVG (no charting library — CSP in the extension
   blocks remote scripts, and the shape count here is small enough that a
   simple tidy-tree layout is more reliable than pulling in a library). */
function teamStructureLayout(nodes, edges){
  const childrenOf = {};
  const hasParent = new Set();
  edges.forEach(e => {
    if(!childrenOf[e.from]) childrenOf[e.from] = [];
    childrenOf[e.from].push(e.to);
    hasParent.add(e.to);
  });
  const roots = nodes.filter(n => !hasParent.has(n.id)).map(n => n.id);
  const rootSet = roots.length ? roots : nodes.map(n => n.id); // fallback: no edges at all

  let leafCounter = 0;
  const pos = {};
  const visited = new Set();
  function assign(id, depth){
    if(visited.has(id)) return; // guard against cycles
    visited.add(id);
    const kids = (childrenOf[id] || []).filter(k => nodes.some(n => n.id === k));
    if(kids.length === 0){
      pos[id] = { x: leafCounter, depth };
      leafCounter += 1;
    } else {
      kids.forEach(k => assign(k, depth + 1));
      const xs = kids.map(k => pos[k] ? pos[k].x : 0);
      pos[id] = { x: (Math.min(...xs) + Math.max(...xs)) / 2, depth };
    }
  }
  rootSet.forEach(id => assign(id, 0));
  // any node never reached (disconnected / orphaned) gets its own column at depth 0
  nodes.forEach(n => { if(!pos[n.id]){ pos[n.id] = { x: leafCounter, depth: 0 }; leafCounter += 1; } });
  return pos;
}

// Matches Stakeholders-page people to Team Structure "team" nodes by their
// free-text `team` field, so the two pages always agree on who's where
// instead of Gemini re-guessing names independently for the diagram.
function normalizeTeamText(s){ return (s||"").toLowerCase().replace(/[^a-z0-9]+/g," ").trim(); }
function stakeholderNamesForNode(nodeLabel){
  const target = normalizeTeamText(nodeLabel);
  if(!target) return [];
  return (state.stakeholders || [])
    .filter(p => {
      const team = normalizeTeamText(p.team);
      if(!team) return false;
      return team === target || team.includes(target) || target.includes(team);
    })
    .map(p => p.name)
    .filter(Boolean);
}

function renderTeamStructure(body){
  const lastRefreshed = state.teamStructureUpdatedAt ? timeAgoLabel(state.teamStructureUpdatedAt) : null;

  body.innerHTML = `
    <div class="d-flex justify-content-between align-items-start flex-wrap gap-2 mb-2">
      <div>
        <h2 style="margin:0">Team Structure</h2>
        <div class="small text-secondary" id="teamStructureMeta">
          ${lastRefreshed ? `Last refreshed ${lastRefreshed}` : "Not refreshed yet"}
        </div>
      </div>
      <button class="btn btn-primary btn-sm" id="refreshStructureBtn"><i class="bi bi-arrow-clockwise"></i> Refresh from notes</button>
    </div>
    <div class="ai-note" id="structureStatus" style="display:none"></div>
    <div class="structure-legend">
      <span><i class="legend-swatch legend-team"></i> Team</span>
      <span><i class="legend-swatch legend-role"></i> Role / security role</span>
      <span><i class="legend-swatch legend-question"></i> Open question</span>
    </div>
    <div id="structureDiagramWrap" class="structure-diagram-wrap"></div>
  `;

  const wrap = document.getElementById("structureDiagramWrap");
  const statusEl = document.getElementById("structureStatus");

  function renderDiagram(){
    const data = state.teamStructure || { nodes: [], edges: [] };
    const nodes = data.nodes || [];
    const edges = data.edges || [];
    if(!nodes.length){
      wrap.innerHTML = `<div class="empty">No team structure yet. Click "Refresh from notes" to have AI infer how teams, sub-teams, and roles fit together based on what's been captured so far.</div>`;
      return;
    }

    const pos = teamStructureLayout(nodes, edges);
    const colWidth = 190, baseBoxH = 56, boxW = 152, marginX = 40, marginY = 30;
    const MAX_NAMES_SHOWN = 4;

    // team nodes get extra box height to fit a stakeholder-name list under the label
    const namesFor = {};
    const boxHFor = {};
    nodes.forEach(n => {
      const names = n.type === "team" ? stakeholderNamesForNode(n.label) : [];
      namesFor[n.id] = names;
      boxHFor[n.id] = names.length ? baseBoxH + 16 + Math.min(names.length, MAX_NAMES_SHOWN) * 14 : baseBoxH;
    });
    // row height must clear the tallest box at each depth so rows never overlap
    const rowHeights = {};
    nodes.forEach(n => {
      const d = pos[n.id].depth;
      rowHeights[d] = Math.max(rowHeights[d] || 0, boxHFor[n.id]);
    });
    const rowGap = 40;
    const rowOffsetY = {};
    let cursorY = marginY;
    const maxDepth = Math.max(...Object.values(pos).map(p => p.depth));
    for(let d = 0; d <= maxDepth; d++){
      rowOffsetY[d] = cursorY;
      cursorY += (rowHeights[d] || baseBoxH) + rowGap;
    }
    const maxX = Math.max(...Object.values(pos).map(p => p.x));
    const svgW = (maxX + 1) * colWidth + marginX * 2;
    const svgH = cursorY;
    // Below this width, boxes/text would shrink past a readable floor if we
    // kept scaling the whole diagram down — better to let it scroll
    // horizontally at that point instead of turning into a squint test.
    const minReadableW = Math.min(svgW, (maxX + 1) * 130 + marginX * 2);

    function cx(id){ return marginX + pos[id].x * colWidth + colWidth/2; }
    function cy(id){ return rowOffsetY[pos[id].depth]; }
    function typeClass(t){ return t === "role" ? "node-role" : t === "question" ? "node-question" : "node-team"; }

    const edgeSvg = edges.filter(e => pos[e.from] && pos[e.to]).map(e => {
      const x1 = cx(e.from), y1 = cy(e.from) + boxHFor[e.from];
      const x2 = cx(e.to), y2 = cy(e.to);
      const midY = (y1 + y2) / 2;
      return `<path d="M${x1},${y1} C${x1},${midY} ${x2},${midY} ${x2},${y2}" class="structure-edge" marker-end="url(#arrowhead)"/>`;
    }).join("");

    const nodeSvg = nodes.map(n => {
      const p = pos[n.id];
      if(!p) return "";
      const boxH = boxHFor[n.id];
      const x = marginX + p.x * colWidth + (colWidth - boxW)/2;
      const y = cy(n.id);
      const cls = typeClass(n.type);
      const label = escapeHtml(n.label || n.id);
      const names = namesFor[n.id];
      const shown = names.slice(0, MAX_NAMES_SHOWN);
      const extra = names.length - shown.length;
      const namesHtml = names.length ? `
        <div class="structure-node-names">
          ${shown.map(nm => `<div class="structure-node-name">${escapeHtml(nm)}</div>`).join("")}
          ${extra > 0 ? `<div class="structure-node-name structure-node-name-more">+${extra} more</div>` : ""}
        </div>` : "";
      return `
        <g class="structure-node ${cls}" data-id="${escapeHtml(n.id)}" transform="translate(${x},${y})">
          <rect width="${boxW}" height="${boxH}" rx="10"></rect>
          <foreignObject width="${boxW}" height="${boxH}">
            <div xmlns="http://www.w3.org/1999/xhtml" class="structure-node-box">
              <div class="structure-node-label">${label}</div>
              ${namesHtml}
            </div>
          </foreignObject>
        </g>`;
    }).join("");

    wrap.innerHTML = `
      <svg viewBox="0 0 ${svgW} ${svgH}" preserveAspectRatio="xMinYMin meet" class="structure-svg" style="max-width:${svgW}px;min-width:${minReadableW}px">
        <defs>
          <marker id="arrowhead" markerWidth="8" markerHeight="8" refX="6" refY="3" orient="auto">
            <path d="M0,0 L0,6 L7,3 z" class="structure-arrow-fill"></path>
          </marker>
        </defs>
        ${edgeSvg}
        ${nodeSvg}
      </svg>`;
  }
  renderDiagram();

  document.getElementById("refreshStructureBtn").onclick = async () => {
    const btn = document.getElementById("refreshStructureBtn");
    btn.disabled = true;
    btn.innerHTML = `<span class="spinner-border spinner-border-sm"></span> Refreshing...`;
    statusEl.style.display = "";
    statusEl.textContent = "Asking Gemini to infer team structure from all notes and transcripts...";
    try{
      await refreshTeamStructureFromNotes();
      statusEl.style.display = "none";
      rerenderKeepScroll();
    }catch(err){
      statusEl.textContent = "Couldn't refresh team structure: " + err.message;
    }finally{
      btn.disabled = false;
      btn.innerHTML = `<i class="bi bi-arrow-clockwise"></i> Refresh from notes`;
    }
  };
}

/* ---------------- init / project switching ---------------- */
async function init(){
  const params = new URLSearchParams(location.search);
  isEmbeddedPane = params.get("embed") === "1";
  if(isEmbeddedPane) document.body.classList.add("embed-pane");

  allProjects = await loadAllProjects();
  geminiApiKey = await loadApiKey();
  if(Object.keys(allProjects).length === 0){
    const id = uid("s");
    allProjects[id] = newProjectState("New discovery project");
    await saveAllProjects(allProjects);
    currentId = id;
  } else {
    currentId = Object.keys(allProjects)[0];
  }
  loadCurrentState();
  if(params.get("pane") && NAV.some(s => s.key === params.get("pane"))){
    activeNav = params.get("pane");
  }
  renderProjectSwitcher();
  renderNav();
  document.getElementById("appSidebar").classList.toggle("collapsed", sidebarCollapsed);
  renderSection();
  await setupMaximizeButton(); // also resolves isFullTabMode, used below

  if(isFullTabMode && !isEmbeddedPane) setupWorkspaceDragDrop();

  // Keep every open pane (this window plus any split-view iframes) in sync:
  // when one pane edits project data, the others pick it up without a
  // manual reload.
  if(chrome.storage && chrome.storage.onChanged){
    chrome.storage.onChanged.addListener((changes, area) => {
      if(area !== "local" || !changes[STORAGE_KEY]) return;
      allProjects = changes[STORAGE_KEY].newValue || {};
      if(!allProjects[currentId]) return;
      loadCurrentState();
      renderProjectSwitcher();
      renderSection();
    });
  }

  document.getElementById("projectSwitcherBtn").onclick = () => {
    document.getElementById("projectDropdownMenu").classList.toggle("show");
  };
  document.addEventListener("click", (e) => {
    if(!e.target.closest(".dropdown")) document.getElementById("projectDropdownMenu").classList.remove("show");
  });

  document.getElementById("newSessionBtn").onclick = async () => {
    const id = uid("s");
    allProjects[id] = newProjectState("New discovery project");
    await saveAllProjects(allProjects);
    currentId = id;
    loadCurrentState();
    renderProjectSwitcher();
    renderNav(); renderSection();
  };
  document.getElementById("projectName").oninput = async (e) => {
    state.name = e.target.value;
    await persist();
    renderProjectSwitcher();
  };

  document.getElementById("sidebarToggle").onclick = () => {
    sidebarCollapsed = !sidebarCollapsed;
    document.getElementById("appSidebar").classList.toggle("collapsed", sidebarCollapsed);
  };
}

// The toolbar icon opens this same page either as a side panel or, via the
// maximize button below, as a full tab. chrome.tabs.getCurrent() resolves to
// undefined when running inside a side panel (it isn't a tab), which is how
// we tell the two contexts apart and only show "maximize" where it's useful.
async function setupMaximizeButton(){
  try{
    const tab = await chrome.tabs.getCurrent();
    isFullTabMode = !!tab;
    const btn = document.getElementById("maximizeBtn");
    if(!tab){
      btn.style.display = "";
      btn.onclick = () => { chrome.tabs.create({ url: chrome.runtime.getURL("app/index.html") }); };
    }
  }catch(e){ /* not running inside an extension tab context; leave hidden */ }
}

/* ---------------- Split-view workspace: drag-to-snap panes ---------------- */
function paneLabel(navKey){ const s = NAV.find(n => n.key === navKey); return s ? s.name : navKey; }
function paneIcon(navKey){ const s = NAV.find(n => n.key === navKey); return s ? s.icon : "bi-file-earmark"; }
function paneIframeSrc(navKey){ return `${location.pathname}?pane=${encodeURIComponent(navKey)}&embed=1`; }

// Nav items become draggable only in full-tab, non-embedded mode, so the
// side panel and any already-split-out pane never show this affordance.
// renderNav() rebuilds #navList from scratch on every call (project switch,
// refresh click, page change), which wipes any listeners attached to the
// old nodes — so this needs to be re-run every time renderNav() runs, not
// just once at startup.
function wireNavItemDragHandles(){
  if(!isFullTabMode || isEmbeddedPane) return;
  const hint = document.getElementById("snapZoneHint");
  document.querySelectorAll(".nav-link[data-key]").forEach(link => {
    link.draggable = true;
    link.classList.add("nav-link-draggable");
    link.addEventListener("dragstart", (e) => {
      dragNavKey = link.dataset.key;
      e.dataTransfer.effectAllowed = "copy";
      e.dataTransfer.setData("text/plain", dragNavKey);
    });
    link.addEventListener("dragend", () => { dragNavKey = null; if(hint) hint.style.display = "none"; });
  });
}

function setupWorkspaceDragDrop(){
  const appMain = document.getElementById("appMain");
  const hint = document.getElementById("snapZoneHint");
  wireNavItemDragHandles();

  function quadrantFromEvent(e){
    const rect = appMain.getBoundingClientRect();
    const x = (e.clientX - rect.left) / rect.width;
    const y = (e.clientY - rect.top) / rect.height;
    return (x < 0.5 ? "l" : "r") + (y < 0.5 ? "t" : "b");
  }
  const quadrantKey = { lt:"tl", rt:"tr", lb:"bl", rb:"br" };

  appMain.addEventListener("dragover", (e) => {
    if(!dragNavKey) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "copy";
    const q = quadrantKey[quadrantFromEvent(e)];
    const rect = appMain.getBoundingClientRect();
    const halves = { tl:[0,0,50,50], tr:[50,0,50,50], bl:[0,50,50,50], br:[50,50,50,50] };
    const [left, top, w, h] = halves[q];
    hint.style.display = "block";
    hint.style.left = left + "%";
    hint.style.top = top + "%";
    hint.style.width = w + "%";
    hint.style.height = h + "%";
  });
  appMain.addEventListener("dragleave", (e) => {
    if(e.target === appMain) hint.style.display = "none";
  });
  appMain.addEventListener("drop", (e) => {
    e.preventDefault();
    hint.style.display = "none";
    if(!dragNavKey) return;
    const q = quadrantKey[quadrantFromEvent(e)];
    workspaceSlots[q] = dragNavKey;
    dragNavKey = null;
    renderWorkspace();
  });
}

// Renders up to 4 panes as a CSS grid. When only 2 adjacent slots are
// filled (both top, both bottom, both left, or both right) it collapses to
// a clean 50/50 half-screen split instead of a 2x2 grid with two dead
// quadrants — matching "snap to half the screen" for the 2-pane case, and
// only actually going full 2x2 once a 3rd or 4th (or diagonal) pane is added.
function computeWorkspaceLayout(){
  const filled = WORKSPACE_SLOTS.filter(k => workspaceSlots[k]);
  if(filled.length === 0) return null;
  const has = k => !!workspaceSlots[k];
  if(filled.length <= 2 && has("tl") && has("tr") && !has("bl") && !has("br")){
    return { columns: "1fr 1fr", rows: "1fr", areas: [["tl","tr"]] };
  }
  if(filled.length <= 2 && has("bl") && has("br") && !has("tl") && !has("tr")){
    return { columns: "1fr 1fr", rows: "1fr", areas: [["bl","br"]] };
  }
  if(filled.length <= 2 && has("tl") && has("bl") && !has("tr") && !has("br")){
    return { columns: "1fr", rows: "1fr 1fr", areas: [["tl"],["bl"]] };
  }
  if(filled.length <= 2 && has("tr") && has("br") && !has("tl") && !has("bl")){
    return { columns: "1fr", rows: "1fr 1fr", areas: [["tr"],["br"]] };
  }
  if(filled.length === 1){
    const only = filled[0];
    return { columns: "1fr", rows: "1fr", areas: [[only]] };
  }
  return { columns: "1fr 1fr", rows: "1fr 1fr", areas: [["tl","tr"],["bl","br"]] };
}

function renderWorkspace(){
  const overlay = document.getElementById("workspaceOverlay");
  const sectionBody = document.getElementById("sectionBody");
  const layout = computeWorkspaceLayout();
  if(!layout){
    overlay.style.display = "none";
    overlay.innerHTML = "";
    sectionBody.style.display = "";
    renderSection(); // refresh the single-page view — it may be stale after sitting hidden
    return;
  }
  sectionBody.style.display = "none";
  overlay.style.display = "grid";
  overlay.style.gridTemplateColumns = layout.columns;
  overlay.style.gridTemplateRows = layout.rows;
  overlay.style.gridTemplateAreas = layout.areas.map(row => `"${row.join(" ")}"`).join(" ");

  overlay.innerHTML = WORKSPACE_SLOTS.filter(k => workspaceSlots[k]).map(k => `
    <div class="workspace-pane" style="grid-area:${k}">
      <div class="workspace-pane-header">
        <i class="bi ${paneIcon(workspaceSlots[k])}"></i>
        <span>${escapeHtml(paneLabel(workspaceSlots[k]))}</span>
        <button class="workspace-pane-close" data-slot="${k}" title="Close pane"><i class="bi bi-x-lg"></i></button>
      </div>
      <iframe src="${paneIframeSrc(workspaceSlots[k])}"></iframe>
    </div>`).join("");

  overlay.querySelectorAll(".workspace-pane-close").forEach(btn => {
    btn.onclick = () => {
      workspaceSlots[btn.dataset.slot] = null;
      renderWorkspace();
    };
  });
}

// One-time migration #1: the Notes page used to be a two-tier model (folders
// that only organize, sections that only hold notes/transcripts). It became a
// single Notion-style tree where any item can both hold notes AND contain
// child items. This converts old projects the first time they load after
// that change — old folders become items with no notes of their own, old
// sections become items nested under them, nothing dropped.
function migrateToUnifiedItems(state){
  if(state.items) return; // already migrated
  const items = [];
  const oldFolders = state.folders || [];
  const oldSections = state.sections || [];
  oldFolders.forEach(f => {
    items.push({ id: f.id, name: f.name, parentId: null, meetings: [], transcripts: [], ts: f.ts || Date.now() });
  });
  oldSections.forEach(s => {
    items.push({
      id: s.id, name: s.name, parentId: s.folderId || null,
      meetings: s.meetings || [], transcripts: s.transcripts || [], ts: s.ts || Date.now()
    });
  });
  state.items = items;
  delete state.folders;
  delete state.sections;
}

// One-time migration #2: each item used to keep a *history* of separately
// titled meeting notes, plus its own transcripts. That's replaced by a
// single free-text note per item (item.desc) with no history, and a flat,
// project-level transcript log with no per-item tagging. Nothing is deleted
// silently — old meeting notes are joined into the new single note (newest
// first, since they were stored newest-first) so the text is still there,
// and old per-item transcripts move into the new flat state.transcripts.
function migrateToSingleNoteAndFlatTranscripts(state){
  if(!state.transcripts) state.transcripts = [];
  (state.items || []).forEach(it => {
    if(it.desc === undefined){
      const meetings = it.meetings || [];
      it.desc = meetings.map(m => (m.title ? `## ${m.title}\n` : "") + (m.desc || "")).join("\n\n").trim();
    }
    if(it.transcripts && it.transcripts.length){
      it.transcripts.forEach(t => state.transcripts.push(t));
    }
    delete it.meetings;
    delete it.transcripts;
  });
}

function loadCurrentState(){
  state = allProjects[currentId] || newProjectState();
  migrateToUnifiedItems(state);
  migrateToSingleNoteAndFlatTranscripts(state);
  if(!state.items) state.items = [];
  if(!state.transcripts) state.transcripts = [];
  if(!state.chats) state.chats = [];
  if(!state.stakeholders) state.stakeholders = [];
  if(!state.stakeholdersUpdatedAt) state.stakeholdersUpdatedAt = null;
  if(!state.teamStructure) state.teamStructure = { nodes: [], edges: [] };
  if(!state.teamStructureUpdatedAt) state.teamStructureUpdatedAt = null;
  state.items.forEach(it => {
    if(it.parentId === undefined) it.parentId = null;
    if(it.desc === undefined) it.desc = "";
  });
  if(!activeItemId || !state.items.find(it => it.id === activeItemId)){
    activeItemId = state.items.length ? state.items[0].id : null;
  }
  activeChatId = state.chats.length ? state.chats[0].id : null;
  document.getElementById("projectName").value = state.name || "";
}

async function switchToProject(id){
  currentId = id;
  loadCurrentState();
  renderProjectSwitcher();
  renderNav();
  renderSection();
}

function renderProjectSwitcher(){
  document.getElementById("currentProjectLabel").textContent = state.name || "Untitled project";
  const container = document.getElementById("projectListContainer");
  const ids = Object.keys(allProjects);
  if(ids.length === 0){ container.innerHTML = `<div class="small text-secondary">No projects yet.</div>`; return; }
  container.innerHTML = ids.map(id => {
    const p = allProjects[id];
    // other projects in the switcher may not be migrated yet (migration runs
    // lazily, only for the project actually being opened) — fall back to the
    // old field name so the count doesn't just show 0 until you switch to it
    const secCount = p.items ? p.items.filter(it => it.desc !== undefined ? !!(it.desc && it.desc.trim()) : ((it.meetings && it.meetings.length) || (it.transcripts && it.transcripts.length))).length : (p.sections || []).length;
    return `
      <div class="project-item ${id === currentId ? "active-project" : ""}">
        <div class="switch-project-item" data-id="${id}" style="flex:1;cursor:pointer">
          <div>${p.name || "Untitled project"}</div>
          <div class="project-meta">${secCount} note section${secCount===1?"":"s"}</div>
        </div>
        <span class="remove-project-btn" data-id="${id}"><i class="bi bi-x-circle-fill"></i></span>
      </div>`;
  }).join("");
  container.querySelectorAll(".switch-project-item").forEach(el => { el.onclick = () => switchToProject(el.dataset.id); });
  container.querySelectorAll(".remove-project-btn").forEach(btn => {
    btn.onclick = async (e) => {
      e.stopPropagation();
      if(Object.keys(allProjects).length <= 1) return;
      const wasCurrent = btn.dataset.id === currentId;
      delete allProjects[btn.dataset.id];
      await saveAllProjects(allProjects);
      if(wasCurrent) currentId = Object.keys(allProjects)[0];
      loadCurrentState();
      renderProjectSwitcher();
      renderNav(); renderSection();
    };
  });
}

function goNav(key){ activeNav = key; renderNav(); renderSection(); }
function renderNav(){
  const el = document.getElementById("navList");
  const navLinkHtml = (s) => `<div class="nav-link ${s.key===activeNav?"active":""}" data-key="${s.key}"><i class="bi ${s.icon} nav-icon"></i><span class="nav-label">${s.name}</span></div>`;
  const refreshRowHtml = `<div class="nav-link" id="refreshAllNavBtn"><i class="bi bi-arrow-clockwise nav-icon"></i><span class="nav-label">Refresh</span></div>`;
  const apiKeyIdx = NAV.findIndex(s => s.key === "apikey");
  const before = apiKeyIdx === -1 ? NAV : NAV.slice(0, apiKeyIdx);
  const after = apiKeyIdx === -1 ? [] : NAV.slice(apiKeyIdx);
  el.innerHTML = before.map(navLinkHtml).join("") + refreshRowHtml + after.map(navLinkHtml).join("");
  el.querySelectorAll(".nav-link[data-key]").forEach(n => { n.onclick = () => goNav(n.dataset.key); });

  const refreshBtn = document.getElementById("refreshAllNavBtn");
  refreshBtn.onclick = async () => {
    if(!geminiApiKey){
      alert("Add a Gemini API key on the API Key page first.");
      return;
    }
    const icon = refreshBtn.querySelector("i");
    refreshBtn.classList.add("nav-link-disabled");
    icon.className = "bi bi-arrow-repeat nav-icon spin";
    refreshBtn.title = "Refreshing stakeholders...";
    try{
      await refreshAllFromNotes((msg) => { refreshBtn.title = msg; });
      // Only Stakeholders/Team Structure have anything to redraw, but
      // re-rendering whatever's active is harmless if it's neither.
      rerenderKeepScroll();
    }catch(err){
      alert("Couldn't refresh: " + err.message);
    }finally{
      refreshBtn.classList.remove("nav-link-disabled");
      icon.className = "bi bi-arrow-clockwise nav-icon";
      refreshBtn.title = "Refresh Stakeholders and Team Structure from all notes";
    }
  };

  wireNavItemDragHandles();
}

function renderSection(){
  // Autosave: whatever's currently typed in the Notes free-text editor gets
  // flushed to storage before we tear down and rebuild the page for
  // wherever we're navigating to next — covers switching nav pages,
  // switching tree items, switching projects, all of it, from one place,
  // so there's never a "click Save or lose it" moment.
  if(noteFlushFn){ noteFlushFn(); noteFlushFn = null; }
  const body = document.getElementById("sectionBody");
  body.innerHTML = "";
  if(activeNav === "notes") renderNotes(body);
  else if(activeNav === "chat") renderChat(body);
  else if(activeNav === "stakeholders") renderStakeholders(body);
  else if(activeNav === "teamstructure") renderTeamStructure(body);
  else if(activeNav === "transcripts") renderTranscripts(body);
  else if(activeNav === "apikey") renderApiKeyPage(body);
}

function timeAgoLabel(ts){
  if(!ts) return "";
  const mins = Math.max(1, Math.round((Date.now() - ts) / 60000));
  if(mins < 60) return mins + "m ago";
  const hrs = Math.round(mins/60);
  if(hrs < 24) return hrs + "h ago";
  return Math.round(hrs/24) + "d ago";
}

/* ---------------- Overview (dashboard + Gemini key settings) ---------------- */
/* ---------------- API Key (dedicated page) ---------------- */
function renderApiKeyPage(body){
  body.innerHTML = `
    <h2>API Key</h2>
    <div class="card">
      <div class="d-flex flex-wrap gap-2 align-items-center">
        <input type="password" id="apiKeyInput" placeholder="Paste your Gemini API key" style="max-width:320px">
        <button class="btn btn-sm btn-primary" id="saveKeyBtn">Save key</button>
        ${geminiApiKey ? `<button class="btn btn-sm btn-outline-danger" id="clearKeyBtn">Remove key</button>` : ""}
        <span class="small ${geminiApiKey ? "" : "text-secondary"}" id="keyStatus">${geminiApiKey ? "✓ Key configured" : "No key configured yet"}</span>
      </div>
    </div>`;

  document.getElementById("saveKeyBtn").onclick = async () => {
    const val = document.getElementById("apiKeyInput").value.trim();
    if(!val) return;
    geminiApiKey = val;
    await saveApiKey(val);
    renderSection();
  };
  const clearBtn = document.getElementById("clearKeyBtn");
  if(clearBtn){
    clearBtn.onclick = async () => { geminiApiKey = ""; await saveApiKey(""); renderSection(); };
  }
}

/* ---------------- lightweight live-markdown editor (headings + bullets) ----------------
   Not a full markdown implementation — just enough to match how Notion/Typora
   feel: type "## " then Enter and that line becomes a heading immediately.
   The block's real markdown ("## Some title") is still what gets saved, so
   the underlying data stays plain, portable markdown text. */
function mdMakeBlock(type, text){
  const div = document.createElement("div");
  div.className = "md-block md-" + type;
  div.textContent = text || "";
  return div;
}
function mdDetectPrefix(text){
  let m = text.match(/^(#{1,3})\s+(.*)$/);
  if(m) return { type: "h" + m[1].length, clean: m[2] };
  m = text.match(/^[-*]\s+(.*)$/);
  if(m) return { type: "li", clean: m[1] };
  return null;
}
function mdInitEditor(editorEl, initialMarkdown){
  editorEl.innerHTML = "";
  const lines = (initialMarkdown || "").split("\n").filter((l,i,arr) => l.trim() || arr.length === 1);
  if(lines.length === 0){
    editorEl.appendChild(mdMakeBlock("p", ""));
  } else {
    lines.forEach(line => {
      const match = mdDetectPrefix(line);
      editorEl.appendChild(match ? mdMakeBlock(match.type, match.clean) : mdMakeBlock("p", line));
    });
  }
  editorEl.addEventListener("keydown", (e) => {
    if(e.key !== "Enter" || e.shiftKey) return;
    const sel = window.getSelection();
    if(!sel.rangeCount) return;
    let node = sel.anchorNode;
    let block = node && (node.nodeType === 1 ? node : node.parentElement);
    block = block && block.closest ? block.closest(".md-block") : null;
    if(!block || !editorEl.contains(block)) return;
    e.preventDefault();
    const text = block.textContent;
    const match = mdDetectPrefix(text);
    if(match){
      block.className = "md-block md-" + match.type;
      block.textContent = match.clean;
    }
    const next = mdMakeBlock("p", "");
    block.after(next);
    const range = document.createRange();
    range.selectNodeContents(next);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
  });
}
function mdSerializeEditor(editorEl){
  const prefixFor = { h1:"# ", h2:"## ", h3:"### ", li:"- ", p:"" };
  return Array.from(editorEl.querySelectorAll(".md-block"))
    .map(b => {
      const type = (Array.from(b.classList).find(c => c.startsWith("md-") && c !== "md-block") || "md-p").slice(3);
      return (prefixFor[type] || "") + b.textContent;
    })
    .filter(line => line.trim())
    .join("\n");
}
function mdToHtml(markdown){
  return escapeHtml(markdown || "").split("\n").map(line => {
    if(/^###\s+/.test(line)) return `<h5>${line.replace(/^###\s+/,"")}</h5>`;
    if(/^##\s+/.test(line)) return `<h4>${line.replace(/^##\s+/,"")}</h4>`;
    if(/^#\s+/.test(line)) return `<h3>${line.replace(/^#\s+/,"")}</h3>`;
    if(/^[-*]\s+/.test(line)) return `<div class="md-li-render">• ${line.replace(/^[-*]\s+/,"")}</div>`;
    if(!line.trim()) return "";
    return `<div>${line}</div>`;
  }).join("");
}

/* ---------------- Discovery notes: OneNote-style folder/file hierarchy ----------------
   One level of nesting, like OneNote's section groups: a "folder" (e.g. "Falcons")
   holds one or more "files" (the actual discovery-note sections — Sales, Service).
   Sections not assigned to a folder show at the top level, same as before. */
/* ---------------- Notes page: Notion-style item tree ----------------
   Every item can hold its own meeting notes/transcripts AND contain child
   items at the same time, same as a Notion page. There's no separate
   "folder" data type any more — an item's icon is just computed at render
   time from whether it currently has children, so dropping a file onto
   another file instantly turns that target into a "folder" visually with
   nothing to migrate. Nesting is done via HTML5 drag-and-drop. */
// Sibling order is just array order (state.items.filter preserves relative
// order) — no separate sortOrder field needed. Reordering means physically
// moving an item's entry in the array next to whichever sibling it was
// dropped beside.
function itemChildren(id){ return state.items.filter(it => it.parentId === id); }
function isDescendant(candidateId, ofId){
  // true if candidateId is ofId itself or nested anywhere under it — used to
  // stop you from dragging a folder into its own child, which would create
  // a cycle the tree renderer can't recurse out of.
  if(candidateId === ofId) return true;
  const byId = new Map(state.items.map(it => [it.id, it]));
  let cur = byId.get(candidateId);
  const seen = new Set();
  while(cur && cur.parentId && !seen.has(cur.id)){
    seen.add(cur.id);
    if(cur.parentId === ofId) return true;
    cur = byId.get(cur.parentId);
  }
  return false;
}
// Moves draggedId to directly before/after targetId, adopting targetId's
// parent — this is how dragging a file to sit between two root-level items
// both reorders it AND pulls it out of whatever folder it used to be in.
function moveItemRelativeTo(draggedId, targetId, position){
  const dragged = state.items.find(it => it.id === draggedId);
  if(!dragged) return;
  state.items = state.items.filter(it => it.id !== draggedId);
  const target = state.items.find(it => it.id === targetId);
  if(!target) return;
  dragged.parentId = target.parentId;
  const targetIdx = state.items.findIndex(it => it.id === targetId);
  state.items.splice(position === "before" ? targetIdx : targetIdx + 1, 0, dragged);
}
// Nests draggedId as the last child of newParentId (or of nothing, for
// root), used both for "drop onto a row" and the root dropzone.
function moveItemInto(draggedId, newParentId){
  const dragged = state.items.find(it => it.id === draggedId);
  if(!dragged) return;
  state.items = state.items.filter(it => it.id !== draggedId);
  dragged.parentId = newParentId;
  state.items.push(dragged);
}

function renderNotes(body){
  body.innerHTML = `
    <h2>Notes</h2>
    <div class="notes-layout">
      <div class="notebook-rail-wrap">
        <div class="notebook-rail-label">Teams</div>
        <div class="notebook-rail" id="notebookRail"></div>
        <button class="add-notebook-btn w-100 mt-2" id="addRootItemBtn"><i class="bi bi-plus-circle-fill"></i> New item</button>
      </div>
      <div class="notebook-content" id="notebookContent"></div>
    </div>`;

  const rail = document.getElementById("notebookRail");

  function treeNodeHtml(item, depth){
    const children = itemChildren(item.id);
    const hasChildren = children.length > 0;
    const isExpanded = expandedTreeItems.has(item.id);
    const isActive = item.id === activeItemId;
    const icon = hasChildren ? (isExpanded ? "bi-folder2-open" : "bi-folder-fill") : "bi-file-earmark-text-fill";
    return `
      <div class="tree-node" data-id="${item.id}">
        <div class="tree-row ${isActive ? "active" : ""}" data-id="${item.id}" draggable="true" style="padding-left:${8 + depth*16}px">
          <span class="tree-caret ${hasChildren ? "" : "tree-caret-hidden"}" data-id="${item.id}">
            <i class="bi ${isExpanded ? "bi-chevron-down" : "bi-chevron-right"}"></i>
          </span>
          <i class="bi ${icon} tree-icon"></i>
          <span class="tree-name">${escapeHtml(item.name)}</span>
          <span class="nb-remove" data-id="${item.id}" title="Remove"><i class="bi bi-x-circle-fill"></i></span>
        </div>
        <div class="tree-children" style="${hasChildren && isExpanded ? "" : "display:none"}">
          ${hasChildren ? children.map(c => treeNodeHtml(c, depth+1)).join("") : ""}
        </div>
      </div>`;
  }

  const roots = itemChildren(null);
  rail.innerHTML = `
    <div class="tree-root-dropzone" id="treeRootDrop"><i class="bi bi-arrow-bar-down"></i> Drop here to move to top level</div>
    ${roots.length ? roots.map(it => treeNodeHtml(it, 0)).join("") : `<div class="empty small" style="padding:8px">No items yet.</div>`}`;

  // expand/collapse
  rail.querySelectorAll(".tree-caret:not(.tree-caret-hidden)").forEach(caret => {
    caret.onclick = (e) => {
      e.stopPropagation();
      const id = caret.dataset.id;
      if(expandedTreeItems.has(id)) expandedTreeItems.delete(id); else expandedTreeItems.add(id);
      renderSection();
    };
  });
  // select
  rail.querySelectorAll(".tree-row").forEach(row => {
    row.onclick = (e) => {
      if(e.target.closest(".nb-remove") || e.target.closest(".tree-caret")) return;
      activeItemId = row.dataset.id;
      renderSection();
    };
  });
  // remove (children are re-parented up one level, not deleted, mirroring the
  // old "remove folder, keep files" behavior)
  rail.querySelectorAll(".tree-row .nb-remove").forEach(x => {
    x.onclick = async (e) => {
      e.stopPropagation();
      const id = x.dataset.id;
      const removed = state.items.find(it => it.id === id);
      if(!removed) return;
      const hasContent = !!(removed.desc && removed.desc.trim());
      const hasKids = itemChildren(id).length > 0;
      if(hasContent || hasKids){
        const ok = confirm(`Remove "${removed.name}"? ${hasContent ? "Its note will be deleted. " : ""}${hasKids ? "Its child items will move up a level." : ""}`);
        if(!ok) return;
      }
      state.items.forEach(it => { if(it.parentId === id) it.parentId = removed.parentId; });
      state.items = state.items.filter(it => it.id !== id);
      if(activeItemId === id) activeItemId = state.items.length ? state.items[0].id : null;
      await persist();
      renderSection();
    };
  });

  // Drag-and-drop: dropping on the top/bottom sliver of a row reorders the
  // dragged item next to it (and adopts that row's parent — this is how
  // dropping between two root-level items also pulls a file out of a
  // folder); dropping on the middle of a row nests it as that row's child.
  let dragId = null;
  const DROP_EDGE = 0.28; // top/bottom fraction of row height treated as reorder zones
  function clearDropClasses(row){ row.classList.remove("tree-drop-target","tree-drop-before","tree-drop-after"); }
  rail.querySelectorAll(".tree-row").forEach(row => {
    row.addEventListener("dragstart", (e) => {
      dragId = row.dataset.id;
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", dragId);
    });
    row.addEventListener("dragover", (e) => {
      const targetId = row.dataset.id;
      if(!dragId || dragId === targetId || isDescendant(targetId, dragId)) return;
      e.preventDefault();
      const rect = row.getBoundingClientRect();
      const frac = (e.clientY - rect.top) / rect.height;
      clearDropClasses(row);
      if(frac < DROP_EDGE) row.classList.add("tree-drop-before");
      else if(frac > 1 - DROP_EDGE) row.classList.add("tree-drop-after");
      else row.classList.add("tree-drop-target");
    });
    row.addEventListener("dragleave", () => clearDropClasses(row));
    row.addEventListener("drop", async (e) => {
      e.preventDefault();
      const targetId = row.dataset.id;
      const wasBefore = row.classList.contains("tree-drop-before");
      const wasAfter = row.classList.contains("tree-drop-after");
      clearDropClasses(row);
      if(!dragId || dragId === targetId || isDescendant(targetId, dragId)){ dragId = null; return; }
      if(wasBefore || wasAfter){
        moveItemRelativeTo(dragId, targetId, wasBefore ? "before" : "after");
      } else {
        moveItemInto(dragId, targetId);
        expandedTreeItems.add(targetId);
      }
      await persist();
      renderSection();
      dragId = null;
    });
  });
  const rootDrop = document.getElementById("treeRootDrop");
  rootDrop.addEventListener("dragover", (e) => { if(dragId) e.preventDefault(); rootDrop.classList.add("tree-drop-target"); });
  rootDrop.addEventListener("dragleave", () => rootDrop.classList.remove("tree-drop-target"));
  rootDrop.addEventListener("drop", async (e) => {
    e.preventDefault();
    rootDrop.classList.remove("tree-drop-target");
    if(!dragId) return;
    moveItemInto(dragId, null);
    await persist();
    renderSection();
    dragId = null;
  });

  document.getElementById("addRootItemBtn").onclick = async () => {
    if(noteFlushFn) noteFlushFn();
    const name = prompt("New item name (e.g. Falcons, Sales, Kickoff notes):");
    if(!name || !name.trim()) return;
    const it = { id: uid("item"), name: name.trim(), parentId: null, desc: "", ts: Date.now() };
    state.items.push(it);
    activeItemId = it.id;
    await persist();
    renderSection();
  };

  const content = document.getElementById("notebookContent");
  const section = state.items.find(it => it.id === activeItemId);
  if(!section){
    content.innerHTML = `<div class="empty">No items yet. Use "New item" on the left to create your first one — drag items onto each other to nest them, just like Notion.</div>`;
    return;
  }

  const breadcrumb = itemPath(section.id);
  content.innerHTML = `
    <div class="notebook-breadcrumb">${escapeHtml(breadcrumb)}</div>
    <div class="d-flex justify-content-between align-items-center mb-2">
      <h3 style="font-size:16px;margin:0"><i class="bi bi-file-earmark-text-fill"></i> ${escapeHtml(section.name)}</h3>
      <span class="autosave-indicator" id="autosaveIndicator"></span>
    </div>
    <div id="notesDescEditor" class="big-textarea md-editor" contenteditable="true"></div>`;

  const descEditor = document.getElementById("notesDescEditor");
  const indicator = document.getElementById("autosaveIndicator");
  mdInitEditor(descEditor, section.desc);

  let saveTimer = null;
  let lastSavedText = section.desc || "";
  function flushNotesEditor(){
    if(!descEditor || !document.body.contains(descEditor)) return;
    clearTimeout(saveTimer);
    const val = mdSerializeEditor(descEditor).trim();
    if(val === lastSavedText) return;
    lastSavedText = val;
    section.desc = val;
    persist();
    if(indicator && document.body.contains(indicator)) indicator.textContent = "Saved";
  }
  noteFlushFn = flushNotesEditor;

  descEditor.addEventListener("input", () => {
    if(indicator) indicator.textContent = "Saving...";
    clearTimeout(saveTimer);
    saveTimer = setTimeout(flushNotesEditor, 700);
  });
}

/* ---------------- Transcripts (flat, project-wide call log) ----------------
   A single, flat, chronological log of every .docx transcript imported for
   this project — deliberately not tied to a specific Notes item/team, so
   it's just "every client call we've had" in one place. */
function renderTranscripts(body){
  const transcripts = state.transcripts || [];
  body.innerHTML = `
    <h2>Transcripts</h2>
    <p class="small text-secondary">Every call transcript imported for this project, newest first.</p>
    <div class="docx-drop mb-3">
      <label class="form-label small fw-semibold mb-1">Import .docx</label>
      <input type="file" accept=".docx" id="docxInput" class="form-control form-control-sm">
      <div class="ai-note" id="docxStatus"></div>
    </div>
    <div id="transcriptList"></div>`;

  const docxInput = document.getElementById("docxInput");
  const docxStatus = document.getElementById("docxStatus");
  wireDocxImport(docxInput, docxStatus, async (text, filename) => {
    state.transcripts.unshift({ id: uid("t"), filename, text, ts: Date.now() });
    await persist();
    renderSection();
  });

  const transcriptList = document.getElementById("transcriptList");
  transcriptList.innerHTML = transcripts.length ? transcripts.map(t => `
    <div class="transcript-row">
      <span class="fname"><i class="bi bi-file-earmark-word-fill"></i> ${escapeHtml(t.filename)}</span>
      <span class="text-secondary">${escapeHtml(t.text.slice(0,140))}${t.text.length>140?"…":""}</span>
      <div class="d-flex justify-content-between align-items-center mt-1">
        <span class="text-secondary">${formatDateLabel(t.ts)} &middot; ${timeAgoLabel(t.ts)}</span>
        <a href="#" class="text-danger remove-transcript" data-id="${t.id}"><i class="bi bi-trash-fill"></i></a>
      </div>
    </div>`).join("") : `<div class="empty">No transcripts imported yet.</div>`;
  transcriptList.querySelectorAll(".remove-transcript").forEach(a => {
    a.onclick = async (e) => {
      e.preventDefault();
      state.transcripts = state.transcripts.filter(t => t.id !== a.dataset.id);
      await persist();
      renderSection();
    };
  });
}

/* ---------------- markdown rendering for Gemini chat replies ----------------
   Gemini responses come back as markdown (headings, **bold**, bullet lists).
   This renders that to HTML for display; it does not affect what gets sent
   back to Gemini (the raw text is still stored in chat history). */
function mdInline(text){
  text = escapeHtml(text);
  text = text.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
  text = text.replace(/(^|[^*])\*(?!\*)([^*]+?)\*(?!\*)/g, "$1<em>$2</em>");
  text = text.replace(/`([^`]+?)`/g, "<code>$1</code>");
  return text;
}
function renderMarkdownChat(markdown){
  const lines = (markdown || "").split("\n");
  let html = "";
  let listOpen = null; // "ul" | "ol" | null
  const closeList = () => { if(listOpen){ html += listOpen === "ul" ? "</ul>" : "</ol>"; listOpen = null; } };

  lines.forEach(rawLine => {
    const line = rawLine.trim();
    let m;
    if((m = line.match(/^(#{1,6})\s+(.*)$/))){
      closeList();
      const level = Math.min(6, m[1].length + 2); // keep headings modest inside a chat bubble
      html += `<h${level} class="chat-md-h">${mdInline(m[2])}</h${level}>`;
    } else if((m = line.match(/^[-*]\s+(.*)$/))){
      if(listOpen !== "ul"){ closeList(); html += "<ul>"; listOpen = "ul"; }
      html += `<li>${mdInline(m[1])}</li>`;
    } else if((m = line.match(/^\d+\.\s+(.*)$/))){
      if(listOpen !== "ol"){ closeList(); html += "<ol>"; listOpen = "ol"; }
      html += `<li>${mdInline(m[1])}</li>`;
    } else if(!line){
      closeList();
    } else {
      closeList();
      html += `<p>${mdInline(line)}</p>`;
    }
  });
  closeList();
  return html;
}

/* ---------------- Chat (ChatGPT-style, Gemini-backed, this project's notes only) ---------------- */
function renderChat(body){
  body.innerHTML = `
    <h2>Chat</h2>
    <div class="chat-shell">
      <div class="chat-threads">
        <button class="btn btn-sm btn-primary w-100 mb-2" id="newChatBtn"><i class="bi bi-plus-circle-fill"></i> New chat</button>
        <div id="chatThreadList"></div>
      </div>
      <div class="chat-main">
        <div class="chat-messages" id="chatMessages"></div>
        <div class="chat-input-row">
          <textarea id="chatInput" rows="1" class="form-control" placeholder="Ask something about this project's discovery notes..."></textarea>
          <button class="btn btn-primary" id="sendChatBtn"><i class="bi bi-send-fill"></i></button>
        </div>
      </div>
    </div>`;

  const threadList = document.getElementById("chatThreadList");
  threadList.innerHTML = state.chats.length ? state.chats.map(c => `
    <div class="chat-thread-item ${c.id===activeChatId?"active":""}" data-id="${c.id}">
      <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(c.title)}</span>
      <span class="del-chat" data-id="${c.id}"><i class="bi bi-x-circle-fill"></i></span>
    </div>`).join("") : `<div class="small text-secondary" style="padding:6px">No chats yet.</div>`;
  threadList.querySelectorAll(".chat-thread-item").forEach(item => {
    item.onclick = (e) => { if(e.target.classList.contains("del-chat")) return; activeChatId = item.dataset.id; renderSection(); };
  });
  threadList.querySelectorAll(".del-chat").forEach(x => {
    x.onclick = async (e) => {
      e.stopPropagation();
      state.chats = state.chats.filter(c => c.id !== x.dataset.id);
      if(activeChatId === x.dataset.id) activeChatId = state.chats.length ? state.chats[0].id : null;
      await persist();
      renderSection();
    };
  });

  document.getElementById("newChatBtn").onclick = async () => {
    const chat = { id: uid("chat"), title: "New chat", ts: Date.now(), messages: [] };
    state.chats.unshift(chat);
    activeChatId = chat.id;
    await persist();
    renderSection();
  };

  const messagesEl = document.getElementById("chatMessages");
  const chat = state.chats.find(c => c.id === activeChatId);
  if(!chat){
    messagesEl.innerHTML = `<div class="chat-empty">No chat open. Click "+ New chat" to start asking questions about ${state.name || "this project"}'s discovery notes.</div>`;
  } else if(chat.messages.length === 0){
    messagesEl.innerHTML = `<div class="chat-empty">Ask anything about ${state.name || "this project"}'s discovery notes &mdash; e.g. "what did the client say about lead routing?"</div>`;
  } else {
    messagesEl.innerHTML = chat.messages.map(m => `<div class="chat-bubble ${m.role}">${m.role === "assistant" ? renderMarkdownChat(m.text) : escapeHtml(m.text)}</div>`).join("");
    messagesEl.scrollTop = messagesEl.scrollHeight;
  }

  const inputEl = document.getElementById("chatInput");
  const sendBtn = document.getElementById("sendChatBtn");
  if(!chat) { inputEl.disabled = true; sendBtn.disabled = true; }

  const sendMessage = async () => {
    const text = inputEl.value.trim();
    if(!text || !chat) return;
    if(!geminiApiKey){
      messagesEl.innerHTML += `<div class="chat-bubble error">No Gemini API key configured yet. Go to Overview to add one.</div>`;
      return;
    }
    chat.messages.push({ role:"user", text, ts: Date.now() });
    if(chat.title === "New chat") chat.title = text.slice(0,40) + (text.length>40?"…":"");
    inputEl.value = "";
    await persist();
    renderSection();

    const thinkingEl = document.createElement("div");
    thinkingEl.className = "chat-bubble assistant";
    thinkingEl.textContent = "Thinking...";
    document.getElementById("chatMessages").appendChild(thinkingEl);
    document.getElementById("chatMessages").scrollTop = 999999;
    document.getElementById("sendChatBtn").disabled = true;

    try{
      const reply = await callGemini(chat.messages);
      chat.messages.push({ role:"assistant", text: reply, ts: Date.now() });
    }catch(err){
      chat.messages.push({ role:"error", text: "Couldn't get a response: " + err.message, ts: Date.now() });
    }
    await persist();
    renderSection();
  };

  sendBtn.onclick = sendMessage;
  inputEl.addEventListener("keydown", (e) => {
    if(e.key === "Enter" && !e.shiftKey){ e.preventDefault(); sendMessage(); }
  });
}

init();
