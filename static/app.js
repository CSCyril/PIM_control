let STATE = null;
let DETAIL = null;
let currentCode = null;
let currentTabRow = null;
let editCell = null;

const $ = (sel) => document.querySelector(sel);

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
}

function fmt(v) {
  if (v === null || v === undefined || v === "") return "";
  return String(v);
}

function setStatus(msg) {
  $("#status").textContent = msg;
}

async function api(path, opts) {
  const r = await fetch(path, opts);
  if (!r.ok) {
    const t = await r.text();
    throw new Error(t || r.statusText);
  }
  return r.json();
}

async function loadState() {
  STATE = await api("/api/state");
  buildDatalists();
  renderList();
  setStatus("Fichier : " + STATE.file + " — sauvegarde de départ : " + STATE.backup +
            (STATE.last_saved ? " — dernière sauvegarde : " + STATE.last_saved : ""));
}

function colOptions(col) {
  const m = STATE.refs && STATE.refs.col_meta && STATE.refs.col_meta[col];
  return (m && m.options) ? m.options : null;
}

function buildDatalists() {
  const dlParent = $("#dl-parent");
  dlParent.innerHTML = STATE.pm_list.map(p => `<option value="${esc(p.code)}">`).join("");
  document.querySelectorAll("datalist.dyn").forEach(d => d.remove());
  const mk = (id, entries) => {
    const dl = document.createElement("datalist");
    dl.id = id;
    dl.className = "dyn";
    dl.innerHTML = entries.map(([v, l]) => `<option value="${esc(v)}">${esc(l || "")}</option>`).join("");
    document.body.appendChild(dl);
  };
  const buildFor = (cols, distinct, prefix) => {
    for (const col of cols) {
      const opts = colOptions(col);
      const entries = [];
      if (opts) for (const [c, l] of Object.entries(opts)) entries.push([c, l]);
      for (const v of (distinct[col] || [])) if (!opts || !(v in opts)) entries.push([v, null]);
      if (col === "categories" && STATE.refs.categories)
        for (const [c, l] of Object.entries(STATE.refs.categories)) entries.push([c, l]);
      if (col === "family" && STATE.refs.families)
        for (const [c, l] of Object.entries(STATE.refs.families)) entries.push([c, l]);
      if (entries.length <= 2000) mk(prefix + col, entries);
    }
  };
  buildFor(STATE.pr_hdr, STATE.distinct_pr, "dl-pr-");
  buildFor(STATE.pm_hdr, STATE.distinct_pm, "dl-pm-");
  const ri = $("#refs-info");
  if (ri && STATE.refs.loaded) {
    const sm = STATE.refs.summary;
    ri.textContent = "📖 réf: " + sm.attributes + " attr · " + sm.options + " options · " +
      sm.categories + " cat · " + sm.families + " fam · " + sm.variants + " variantes";
    ri.title = "Référentiels chargés depuis " + STATE.refs.dir;
  } else if (ri) {
    ri.textContent = "📖 pas de référentiel (dossier refs/ vide ou absent)";
  }
}

function filteredList() {
  const q = $("#search").value.trim().toLowerCase();
  const f = $("#filter").value;
  return STATE.pm_list.filter(p => {
    if (q && !p.code.toLowerCase().includes(q)) return false;
    if (f === "issues" && p.nb_issues === 0) return false;
    if (f === "nopm" && p.has_pm) return false;
    if (f === "noproduct" && p.nb_products > 0) return false;
    return true;
  });
}

function renderList() {
  const list = filteredList();
  $("#pm-count").textContent = list.length + " PM affichés / " + STATE.pm_list.length + " au total";
  $("#pm-list").innerHTML = list.map(p => `
    <li data-code="${esc(p.code)}" class="${p.code === currentCode ? "active" : ""} ${p.has_pm ? "" : "nopm"}">
      <span class="code">${esc(p.code)}</span>
      ${p.has_pm ? "" : '<span class="tag nopm-tag" title="Pas de fiche PM">sans PM</span>'}
      ${p.nb_pm_rows > 1 ? `<span class="tag dup" title="${p.nb_pm_rows} lignes PM">×${p.nb_pm_rows}</span>` : ""}
      <span class="tag">${p.nb_products} prod.</span>
      ${p.nb_issues ? `<span class="tag issue" title="Anomalies">⚠${p.nb_issues}</span>` : ""}
    </li>`).join("");
  document.querySelectorAll("#pm-list li").forEach(li => {
    li.addEventListener("click", () => openPM(li.dataset.code));
  });
}

async function openPM(code) {
  currentCode = code;
  DETAIL = await api("/api/pm/" + encodeURIComponent(code));
  currentTabRow = DETAIL.pm_rows.length ? DETAIL.pm_rows[0].row : null;
  renderList();
  renderPM();
  $("#empty").style.display = "none";
  $("#pm-view").style.display = "block";
  window.scrollTo(0, 0);
}

function renderPM() {
  const d = DETAIL;
  const head = $("#pm-head");
  head.innerHTML = `
    <h1>${esc(d.code)} ${d.has_pm ? "" : '<span class="tag nopm-tag big">fiche PM absente</span>'}
      ${d.pm_rows.length > 1 ? `<span class="tag dup big">${d.pm_rows.length} lignes PM</span>` : ""}</h1>
    <div class="head-actions">
      ${d.has_pm ? "" : '<button id="btn-create-pm">＋ Créer la fiche PM</button>'}
    </div>`;
  const bcp = $("#btn-create-pm");
  if (bcp) bcp.addEventListener("click", async () => {
    await api("/api/pm", {method: "POST", headers: {"Content-Type": "application/json"},
      body: JSON.stringify({code: d.code})});
    await afterMutation();
  });

  const tabs = $("#pm-tabs");
  if (d.pm_rows.length > 1) {
    tabs.innerHTML = d.pm_rows.map(r =>
      `<button class="tab ${r.row === currentTabRow ? "active" : ""}" data-row="${r.row}">ligne ${r.row}</button>`).join("");
    tabs.querySelectorAll(".tab").forEach(t => t.addEventListener("click", () => {
      currentTabRow = +t.dataset.row; renderPM();
    }));
  } else tabs.innerHTML = "";

  const fields = $("#pm-fields");
  const row = d.pm_rows.find(r => r.row === currentTabRow);
  if (row) {
    fields.innerHTML = `<div class="card"><h3>Attributs du PM (ligne ${row.row})</h3>
      <div class="grid">` + STATE.pm_hdr.map(h => {
        const m = STATE.refs && STATE.refs.col_meta[h];
        return `
        <div class="field">
          <label>${m && m.label ? esc(m.label) + ' <small>' + esc(h) + '</small>' : esc(h)}</label>
          <input data-pm-row="${row.row}" data-field="${esc(h)}" list="dl-pm-${esc(h)}"
                 value="${esc(fmt(row.values[h]))}">
        </div>`; }).join("") + `</div>
      <div class="row-actions"><button id="btn-del-pm-row" class="danger">🗑 Supprimer cette ligne PM</button></div>
      </div>`;
    fields.querySelectorAll("input").forEach(inp => {
      inp.addEventListener("change", async () => {
        await api("/api/pm/row/" + inp.dataset.pmRow, {method: "PUT",
          headers: {"Content-Type": "application/json"},
          body: JSON.stringify({field: inp.dataset.field, value: inp.value})});
        await afterMutation();
      });
    });
    const del = $("#btn-del-pm-row");
    if (del) del.addEventListener("click", async () => {
      if (!confirm("Supprimer la ligne PM " + row.row + " (code " + d.code + ") ? Les produits resteront orphelins.")) return;
      await api("/api/pm/row/" + row.row, {method: "DELETE"});
      await afterMutation();
    });
  } else {
    fields.innerHTML = `<div class="card"><p>Aucune ligne dans la feuille <b>PM</b> pour ce code. Les produits rattachés sont orphelins.</p></div>`;
  }

  renderChecks();
  renderProducts();
}

function renderChecks() {
  const box = $("#pm-checks");
  const d = DETAIL;
  if (!d.checks.length) {
    box.innerHTML = `<div class="card ok">✅ Aucun problème détecté sur ce PM et ses produits.</div>`;
    return;
  }
  box.innerHTML = `<div class="card"><h3>⚠ Contrôles de cohérence (${d.checks.length})</h3>
    <ul class="checks">` + d.checks.map(c => `
      <li class="${c.severity}" data-sku="${esc(c.sku || "")}">
        <b>${esc(STATE.labels[c.type] || c.type)}</b> — ${esc(c.msg)}
      </li>`).join("") + `</ul></div>`;
  box.querySelectorAll("li[data-sku]").forEach(li => {
    if (!li.dataset.sku) return;
    li.style.cursor = "pointer";
    li.title = "Aller au produit";
    li.addEventListener("click", () => {
      const tr = document.querySelector(`tr[data-row] td[data-sku="${li.dataset.sku}"]`);
      if (tr) { tr.scrollIntoView({block: "center"}); tr.classList.add("flash"); setTimeout(() => tr.classList.remove("flash"), 1500); }
    });
  });
}

function renderProducts() {
  const d = DETAIL;
  $("#prod-count").textContent = "(" + d.products.length + ")";
  const thead = $("#products-table thead");
  thead.innerHTML = "<tr><th class='sticky'>sku</th>" +
    STATE.pr_hdr.filter(h => h !== "sku").map(h => {
      const m = STATE.refs && STATE.refs.col_meta[h];
      const label = m ? (m.is_unit ? "unité de " + m.label : m.label) : null;
      const tip = esc(h) + (m && m.type ? " · type " + m.type : "") + (m && m.group_label ? " · groupe " + m.group_label : "");
      return `<th class="${h.endsWith("-unit") ? "unit" : ""}" title="${tip}">${esc(h)}${label ? `<span class="collabel">${esc(label)}${m.type ? " · " + esc(m.type) : ""}</span>` : ""}</th>`;
    }).join("") +
    "<th>contrôles</th><th></th></tr>";
  const tbody = $("#products-table tbody");
  tbody.innerHTML = d.products.map(p => {
    const warn = p.checks.filter(c => c.severity === "warning").length;
    const info = p.checks.length - warn;
    return `<tr data-row="${p.row}" class="${p.checks.length ? "has-issues" : ""}">
      <td class="sticky" data-sku="${esc(fmt(p.values.sku))}" data-field="sku" title="cliquer pour modifier">${esc(fmt(p.values.sku)) || "—"}</td>
      ${STATE.pr_hdr.filter(h => h !== "sku").map(h => {
        const m = STATE.refs && STATE.refs.col_meta[h];
        const v = fmt(p.values[h]);
        const optLabel = m && m.options && v ? m.options[v] : null;
        return `
        <td class="${h.endsWith("-unit") ? "unit" : ""} ${p.checks.some(c => c.msg.includes(h)) ? "cell-warn" : ""}"
            data-field="${esc(h)}" title="${esc(optLabel ? optLabel + " — " : "")}cliquer pour modifier">${esc(v) || '<span class="empty-cell">—</span>'}</td>`;
      }).join("")}
      <td class="checks-cell">${p.checks.length
          ? `<details><summary class="${warn ? "issue" : "info"}">⚠${p.checks.length}</summary>
             <ul>${p.checks.map(c => `<li class="${c.severity}">${esc(c.msg)}</li>`).join("")}</ul></details>`
          : '<span class="ok-dot">✓</span>'}</td>
      <td><button class="danger mini" data-del="${p.row}" title="Supprimer le produit">🗑</button></td>
    </tr>`;
  }).join("");
  tbody.querySelectorAll("td[data-field]").forEach(td => {
    td.addEventListener("click", () => startEdit(td));
  });
  tbody.querySelectorAll("button[data-del]").forEach(b => {
    b.addEventListener("click", async () => {
      if (!confirm("Supprimer le produit ligne " + b.dataset.del + " ?")) return;
      await api("/api/produit/row/" + b.dataset.del, {method: "DELETE"});
      await afterMutation();
    });
  });
  $("#btn-add-product").onclick = async () => {
    const r = await api("/api/produit", {method: "POST", headers: {"Content-Type": "application/json"},
      body: JSON.stringify({parent: currentCode})});
    await afterMutation();
    const tr = document.querySelector(`tr[data-row="${r.row}"]`);
    if (tr) { tr.scrollIntoView({block: "center"}); const td = tr.querySelector("td[data-field='sku']"); if (td) startEdit(td); }
  };
}

function startEdit(td) {
  if (editCell) return;
  const tr = td.parentElement;
  const row = tr.dataset.row;
  const field = td.dataset.field;
  const old = td.textContent;
  editCell = td;
  const isParent = field === "parent";
  const dlId = isParent ? "dl-parent" : ("dl-pr-" + field);
  td.classList.add("editing");
  td.innerHTML = `<input list="${document.getElementById(dlId) ? dlId : ""}" value="${esc(old === "—" ? "" : old)}">`;
  const inp = td.querySelector("input");
  inp.focus(); inp.select();
  let done = false;
  const finish = async (saveIt) => {
    if (done) return; done = true;
    const val = inp.value;
    editCell = null;
    if (saveIt && val !== (old === "—" ? "" : old)) {
      try {
        await api("/api/produit/row/" + row, {method: "PUT",
          headers: {"Content-Type": "application/json"},
          body: JSON.stringify({field, value: val})});
        await afterMutation();
      } catch (e) {
        alert("Erreur d'enregistrement : " + e.message);
        renderProducts();
      }
    } else {
      renderProducts();
    }
  };
  inp.addEventListener("blur", () => finish(true));
  inp.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); finish(true); }
    if (e.key === "Escape") { e.preventDefault(); finish(false); }
  });
}

async function afterMutation() {
  await loadState();
  if (currentCode) {
    DETAIL = await api("/api/pm/" + encodeURIComponent(currentCode));
    if (!DETAIL.pm_rows.length) currentTabRow = null;
    else if (!DETAIL.pm_rows.some(r => r.row === currentTabRow)) currentTabRow = DETAIL.pm_rows[0].row;
    renderList();
    renderPM();
  }
  setStatus("💾 Enregistré à " + new Date().toLocaleTimeString());
}

function currentIndex() {
  const list = filteredList();
  return list.findIndex(p => p.code === currentCode);
}

function nav(delta) {
  const list = filteredList();
  if (!list.length) return;
  let i = currentIndex();
  i = (i + delta + list.length) % list.length;
  openPM(list[i].code);
}

function navNextIssue() {
  const list = filteredList();
  if (!list.length) return;
  const start = currentIndex();
  for (let k = 1; k <= list.length; k++) {
    const p = list[(start + k) % list.length];
    if (p.nb_issues > 0) { openPM(p.code); return; }
  }
  alert("Aucun PM avec anomalie dans la liste courante.");
}

async function showDashboard() {
  $("#dashboard").style.display = "flex";
  $("#dashboard-body").innerHTML = "Chargement…";
  const groups = await api("/api/checks");
  const total = groups.reduce((a, g) => a + g.count, 0);
  $("#dashboard-body").innerHTML = `<p class="dash-total">${total} anomalie(s) détectée(s) dans ${groups.length} catégorie(s). Cliquez sur une ligne pour ouvrir le PM concerné.</p>` +
    groups.map(g => `<details open class="dash-group">
      <summary><b>${esc(g.label)}</b> <span class="tag issue">${g.count}</span></summary>
      <ul>${g.items.map(it => `<li class="${it.severity}"><a href="#" data-pm="${esc(it.pm)}">${esc(it.pm)}</a>${it.sku ? " — " + esc(it.sku) : ""} : ${esc(it.msg)}</li>`).join("")}</ul>
    </details>`).join("");
  $("#dashboard-body").querySelectorAll("a[data-pm]").forEach(a => {
    a.addEventListener("click", (e) => {
      e.preventDefault();
      $("#dashboard").style.display = "none";
      openPM(a.dataset.pm);
    });
  });
}

$("#search").addEventListener("input", renderList);
$("#filter").addEventListener("change", renderList);
$("#btn-prev").addEventListener("click", () => nav(-1));
$("#btn-next").addEventListener("click", () => nav(1));
$("#btn-next-issue").addEventListener("click", navNextIssue);
$("#btn-dashboard").addEventListener("click", showDashboard);
$("#btn-close-dashboard").addEventListener("click", () => { $("#dashboard").style.display = "none"; });
$("#btn-save").addEventListener("click", async () => {
  const r = await api("/api/save", {method: "POST"});
  setStatus("💾 Enregistré à " + r.saved_at);
});

loadState().then(() => {
  const saved = sessionStorage.getItem("lastPM");
  if (saved && STATE.pm_list.some(p => p.code === saved)) openPM(saved);
});
setInterval(() => { if (currentCode) sessionStorage.setItem("lastPM", currentCode); }, 2000);
