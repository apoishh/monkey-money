/* Monkey Money - reads /api/data (your workbook, read-only) and draws four tabs. */
"use strict";

const $ = (s) => document.querySelector(s);
const css = (v) => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const RM = (n, dp = 0) => (n < 0 ? "-" : "") + "RM" + Math.abs(n).toLocaleString("en-MY", { minimumFractionDigits: dp, maximumFractionDigits: dp });
const pct = (n, dp = 0) => (n * 100).toFixed(dp) + "%";
const ym = (iso) => iso.slice(0, 7);
const monthName = (k) => new Date(k + "-01T00:00:00").toLocaleDateString("en-GB", { month: "short", year: "numeric" });
const dayName = (iso) => new Date(iso + "T00:00:00").toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" });
const daysAgo = (iso, today) => Math.round((new Date(today) - new Date(iso)) / 86400000);
// plain arithmetic: Date#toISOString converts to UTC and shifts local midnight back a day in Malaysia (UTC+8)
const addMonths = (k, n) => { const t = +k.slice(0, 4) * 12 + (+k.slice(5, 7) - 1) + n; return `${Math.floor(t / 12)}-${String((t % 12) + 1).padStart(2, "0")}`; };
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const pill = (kind, text) => `<span class="status ${kind}">${kind === "good" ? "✓" : kind === "warning" ? "!" : "✕"} ${esc(text)}</span>`;

let D = null;          // raw data
const charts = {};     // Chart.js instances

/* ------------------------------------------------------------------ classification */
const isEarned = (t) => t.type === "Income" && t.category !== "Investments";   // not money back from investments
const isSpend = (t) => t.type === "Expenses" && t.category !== "Investment";   // not money moved into investments
const isInvested = (t) => t.type === "Expenses" && t.category === "Investment";
// before Sep 2026 card spending was logged as one bill payment, not per purchase
const catLabel = (c) => (/maybank/i.test(c) && D.card_lump_categories.includes(c) ? "Maybank card bill (lump)"
  : /cimb/i.test(c) && D.card_lump_categories.includes(c) ? "CIMB card bill (lump)" : c);
const shortDay = (iso) => new Date(iso + "T00:00:00").toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "2-digit" });

function monthsBetween(a, b) { const out = []; for (let k = a; k <= b; k = addMonths(k, 1)) out.push(k); return out; }

function window12() {
  // the last 12 COMPLETE months you tracked (same rule as the Excel Wealth Summary)
  const end = addMonths(ym(D.last_transaction), -1);
  return monthsBetween(addMonths(end, -11), end);
}

function sumBy(list, keyFn, valFn = (t) => t.amount) {
  const m = new Map();
  for (const t of list) { const k = keyFn(t); m.set(k, (m.get(k) || 0) + valFn(t)); }
  return m;
}

/* ------------------------------------------------------------------ chart helpers */
function baseOpts(extra = {}) {
  const tick = css("--text-secondary"), grid = css("--grid");
  return Object.assign({
    responsive: true, maintainAspectRatio: false, animation: false,
    interaction: { mode: "index", intersect: false },
    plugins: {
      legend: { labels: { color: tick, boxWidth: 12, boxHeight: 12, useBorderRadius: true, borderRadius: 3 } },
      tooltip: { callbacks: { label: (c) => ` ${c.dataset.label}: ${RM(c.parsed.y ?? c.parsed.x ?? 0)}` } },
    },
    scales: {
      x: { grid: { display: false }, ticks: { color: tick }, border: { color: grid } },
      y: { grid: { color: grid }, border: { display: false }, ticks: { color: tick, callback: (v) => RM(v) } },
    },
  }, extra);
}
function draw(id, config) {
  if (typeof Chart === "undefined") return; // offline: numbers and tables still render
  // keep money axes readable: flat labels, few ticks
  const sc = config.options && config.options.scales;
  if (sc) for (const ax of Object.values(sc)) { ax.ticks = Object.assign({ maxRotation: 0, autoSkip: true }, ax.ticks); }
  if (sc && config.options.indexAxis === "y" && sc.x) sc.x.ticks.maxTicksLimit = 5;
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(document.getElementById(id), config);
}
const bar = (label, data, color, extra = {}) => Object.assign({ label, data, backgroundColor: color, borderRadius: 4, borderSkipped: false, maxBarThickness: 26, borderColor: css("--surface-1"), borderWidth: 1 }, extra);
const line = (label, data, color, extra = {}) => Object.assign({ label, data, borderColor: color, backgroundColor: color, borderWidth: 2, pointRadius: 0, pointHoverRadius: 5, tension: 0.2 }, extra);

/* ------------------------------------------------------------------ OVERVIEW */
function netWorthParts() {
  const inv = D.platforms.filter((p) => p.id !== "kwsp").reduce((s, p) => s + p.value, 0);
  const kwsp = (D.platforms.find((p) => p.id === "kwsp") || { value: 0 }).value;
  const assets = [
    ["Maybank (main account)", D.cash], ["CIMB savings", D.savings], ["Investments (not KWSP)", inv], ["KWSP / EPF", kwsp],
    ...D.loans.map((l) => [l.name + " - market value", l.market_value]),
    ["Valuables & other assets", D.possessions.reduce((s, p) => s + p.value, 0)],
  ];
  const debts = [...D.loans.map((l) => [l.name + " - loan left", l.outstanding]), ["Credit cards", D.cards]];
  const A = assets.reduce((s, a) => s + a[1], 0), L = debts.reduce((s, d) => s + d[1], 0);
  return { assets, debts, A, L, nw: A - L, inv, kwsp };
}

function renderOverview() {
  const p = netWorthParts();
  $("#nw").textContent = RM(p.nw);
  const match = D.excel_net_worth != null && Math.abs(D.excel_net_worth - p.nw) < 1;
  $("#nw-sub").innerHTML = `Own ${RM(p.A)} · owe ${RM(p.L)}` + (D.excel_net_worth != null ? ` · ${match ? "same as" : "differs from"} Excel's Wealth Summary` : "");
  const liquid = D.cash + D.savings + p.inv;
  $("#tiles").innerHTML = [
    ["Maybank (main)", RM(D.cash), "tracker balance"],
    ["CIMB savings", RM(D.savings), "tracker savings"],
    ["Investments", RM(p.inv), "excl. KWSP"],
    ["KWSP / EPF", RM(p.kwsp), "retirement"],
    ["Can access", RM(liquid), "cash + savings + investments"],
    ["Debts", RM(p.L), "loans and cards"],
  ].map(([k, v, s]) => `<div class="tile"><div class="label">${k}</div><div class="v">${v}</div><div class="s">${s}</div></div>`).join("");

  $("#bank").innerHTML = D.bank_check.length ? D.bank_check.map((b) => {
    const ok = b.status === "Matches";
    return `<div class="row"><div class="k"><b>${esc(b.account)}</b><div class="s">${b.asof ? "statement " + dayName(b.asof) : "no statement yet"}</div></div>
      <div class="n small">bank ${b.bank != null ? RM(b.bank, 2) : "-"}<br>tracker ${b.tracker != null ? RM(b.tracker, 2) : "-"}</div>
      ${b.asof ? pill(ok ? "good" : "critical", ok ? "Matches" : `${b.status} by ${RM(Math.abs(b.diff), 2)}`) : ""}</div>`;
  }).join("") : `<p class="muted">Run "1 - Import statements" once to start the bank check.</p>`;

  const rows = [["Budget Tracking", D.last_transaction, 14, "Import new statements"]]
    .concat(D.platforms.filter((x) => !x.closed && x.updated).map((x) => [x.name, x.updated, 45, "Add a current-value row on the Investment sheet"]));
  $("#fresh").innerHTML = rows.map(([name, iso, lim, todo]) => {
    const d = daysAgo(iso, D.today);
    return `<div class="row"><div class="k"><b>${esc(name)}</b><div class="s">${dayName(iso)}${d > lim ? " · " + todo : ""}</div></div>
      ${pill(d > lim ? (d > lim * 3 ? "critical" : "warning") : "good", d === 0 ? "today" : d + " days old")}</div>`;
  }).join("");

  const items = [...p.assets.map((a) => [a[0], a[1]]), ...p.debts.map((d) => [d[0], -d[1]])].filter((x) => Math.abs(x[1]) > 0.5);
  draw("c-balance", {
    type: "bar",
    data: { labels: items.map((x) => x[0]), datasets: [
      bar("Own", items.map((x) => (x[1] > 0 ? x[1] : null)), css("--series-1")),
      bar("Owe", items.map((x) => (x[1] < 0 ? x[1] : null)), css("--series-2")),
    ] },
    options: baseOpts({ indexAxis: "y", interaction: { mode: "nearest", intersect: true },
      plugins: { legend: { labels: { color: css("--text-secondary"), boxWidth: 12 } }, tooltip: { callbacks: { label: (c) => ` ${RM(Math.abs(c.parsed.x))}` } } },
      scales: { x: { grid: { color: css("--grid") }, ticks: { color: css("--text-secondary"), callback: (v) => RM(v) } },
                y: { stacked: true, grid: { display: false }, ticks: { color: css("--text-secondary") } } } }),
  });
  $("#t-balance").innerHTML = `<table><tr><th>Item</th><th class="n">RM</th></tr>${items.map((x) => `<tr><td>${esc(x[0])}</td><td class="n ${x[1] < 0 ? "neg" : ""}">${RM(x[1], 2)}</td></tr>`).join("")}
    <tr><th>Net worth</th><th class="n">${RM(p.nw, 2)}</th></tr></table>`;
}

/* ------------------------------------------------------------------ SPENDING */
function periodMonths(v) {
  if (v === "12m") return window12();
  if (/^\d{4}$/.test(v)) return [...new Set(D.transactions.filter((t) => t.date.startsWith(v)).map((t) => ym(t.date)))].sort();
  return [v];
}

function setupSpending() {
  const months = [...new Set(D.transactions.map((t) => ym(t.date)))].sort().reverse();
  const years = [...new Set(months.map((m) => m.slice(0, 4)))];
  const w = window12();
  $("#sp-period").innerHTML = `<option value="12m">Last 12 full months (${monthName(w[0])} – ${monthName(w[11])})</option>` +
    years.map((y) => `<option value="${y}">${y}</option>`).join("") +
    `<optgroup label="Single month">${months.map((m) => `<option value="${m}">${monthName(m)}</option>`).join("")}</optgroup>`;
  const cats = [...new Set(D.transactions.filter(isSpend).map((t) => t.category))].sort();
  $("#sp-cat").innerHTML = `<option value="">All categories</option>` + cats.map((c) => `<option value="${esc(c)}">${esc(catLabel(c))}</option>`).join("");
  $("#sp-period").onchange = renderSpending;
  $("#sp-cat").onchange = renderSpending;
  $("#tx-search").oninput = renderTx;
}

function renderSpending() {
  const per = $("#sp-period").value, cat = $("#sp-cat").value;
  const months = periodMonths(per), set = new Set(months), n = months.length || 1;
  const inP = D.transactions.filter((t) => set.has(ym(t.date)));
  const spend = inP.filter(isSpend).filter((t) => !cat || t.category === cat);
  const inc = inP.filter(isEarned).reduce((s, t) => s + t.amount, 0);
  const out = spend.reduce((s, t) => s + t.amount, 0);
  const invd = inP.filter(isInvested).reduce((s, t) => s + t.amount, 0);
  const allSpend = inP.filter(isSpend).reduce((s, t) => s + t.amount, 0);
  const each = n > 1 ? " / month" : "";
  $("#sp-note").textContent = n > 1 ? `Figures are monthly averages over ${n} months.` : "";
  $("#sp-tiles").innerHTML = (cat ? [[cat, RM(out / n), "spent" + each]] : [
    ["Money in", RM(inc / n), "income" + each],
    ["Spent", RM(out / n), "excl. investing" + each],
    ["Invested", RM(invd / n), each ? "per month" : "this month"],
    ["Left over", RM((inc - allSpend) / n), inc ? `${pct((inc - allSpend) / inc)} of income kept` : ""],
  ]).map(([k, v, s]) => `<div class="tile"><div class="label">${esc(k)}</div><div class="v ${v.startsWith("-") ? "neg" : ""}">${v}</div><div class="s">${s}</div></div>`).join("");

  // monthly bars: 24 months up to the end of the period
  const endM = months[months.length - 1] || ym(D.last_transaction);
  const ms = monthsBetween(addMonths(endM, -23), endM);
  const byM = (f) => { const m = sumBy(D.transactions.filter(f), (t) => ym(t.date)); return ms.map((k) => +(m.get(k) || 0).toFixed(2)); };
  const ds = cat ? [bar(cat, byM((t) => isSpend(t) && t.category === cat), css("--series-2"))]
    : [bar("Money in", byM(isEarned), css("--series-1")), bar("Spent", byM(isSpend), css("--series-2"))];
  draw("c-monthly", { type: "bar", data: { labels: ms.map(monthName), datasets: ds }, options: baseOpts() });

  // categories
  $("#cat-title").textContent = n > 1 ? "Where it went (per month)" : "Where it went";
  const byCat = [...sumBy(inP.filter(isSpend), (t) => t.category)].map(([k, v]) => [k, v / n]).sort((a, b) => b[1] - a[1]);
  draw("c-cats", { type: "bar",
    data: { labels: byCat.map((c) => catLabel(c[0])), datasets: [bar("Spent", byCat.map((c) => +c[1].toFixed(2)), byCat.map((c) => (cat && c[0] !== cat ? css("--neutral") : css("--series-1"))))] },
    options: baseOpts({ indexAxis: "y", plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => ` ${RM(c.parsed.x)}` } } },
      scales: { x: { grid: { color: css("--grid") }, ticks: { color: css("--text-secondary"), callback: (v) => RM(v) } }, y: { grid: { display: false }, ticks: { color: css("--text-secondary") } } } }) });

  // budget vs actual
  const bud = {};
  for (const m of months) for (const [c, v] of Object.entries((D.budget.Expenses || {})[m] || {})) bud[c] = (bud[c] || 0) + v / n;
  const actual = Object.fromEntries(byCat);
  const bc = [...new Set([...Object.keys(bud), ...Object.keys(actual)])].filter((c) => c !== "Investment")
    .sort((a, b) => Math.max(bud[b] || 0, actual[b] || 0) - Math.max(bud[a] || 0, actual[a] || 0));
  const tb = Object.values(bud).reduce((s, v) => s + v, 0), ta = bc.reduce((s, c) => s + (actual[c] || 0), 0);
  $("#bud-title").textContent = n > 1 ? "Budget vs actual (per month)" : "Budget vs actual";
  $("#bud-note").textContent = tb ? `Planned ${RM(tb)} · spent ${RM(ta)} · ${ta > tb ? "over by " + RM(ta - tb) : "under by " + RM(tb - ta)}. Budget from the Budget Planning sheet.` : "No budget set for this period in Budget Planning.";
  draw("c-budget", { type: "bar",
    data: { labels: bc.map(catLabel), datasets: [bar("Budget", bc.map((c) => +(bud[c] || 0).toFixed(2)), css("--neutral")), bar("Actual", bc.map((c) => +(actual[c] || 0).toFixed(2)), css("--series-1"))] },
    options: baseOpts({ indexAxis: "y",
      plugins: { legend: { labels: { color: css("--text-secondary"), boxWidth: 12 } }, tooltip: { callbacks: { label: (c) => ` ${c.dataset.label}: ${RM(c.parsed.x)}` } } },
      scales: { x: { grid: { color: css("--grid") }, ticks: { color: css("--text-secondary"), callback: (v) => RM(v) } }, y: { grid: { display: false }, ticks: { color: css("--text-secondary") } } } }) });

  // top places
  const named = spend.filter((t) => t.details && !/^(transfer from a\/c|qr payment)/i.test(t.details));
  const byPlace = new Map();
  for (const t of named) {
    const k = t.details.replace(/\s+(SDN|BHD|AP)\b.*$/i, "").replace(/[-\s]+$/, "").toUpperCase().slice(0, 28);
    const e = byPlace.get(k) || { total: 0, count: 0, cat: t.category }; e.total += t.amount; e.count++; byPlace.set(k, e);
  }
  const top = [...byPlace].sort((a, b) => b[1].total - a[1].total).slice(0, 12);
  $("#merchants").innerHTML = top.length ? `<table><tr><th>Place</th><th>Category</th><th class="n">Times</th><th class="n">Total</th></tr>` +
    top.map(([k, e]) => `<tr><td>${esc(k)}</td><td>${esc(e.cat)}</td><td class="n">${e.count}</td><td class="n">${RM(e.total, 2)}</td></tr>`).join("") + `</table>`
    : `<p class="muted small">No merchant names in this period. They come from imported statements (Sep 2026 onwards).</p>`;
  renderTx();
}

function renderTx() {
  const per = $("#sp-period").value, cat = $("#sp-cat").value, q = $("#tx-search").value.trim().toLowerCase();
  const set = new Set(periodMonths(per));
  let list = D.transactions.filter((t) => set.has(ym(t.date)) && (!cat || t.category === cat));
  if (q) list = list.filter((t) => `${t.details} ${t.category} ${t.type} ${t.amount}`.toLowerCase().includes(q));
  list = list.slice().reverse();
  $("#tx").innerHTML = `<table><tr><th>Date</th><th>Details</th><th>Category</th><th class="n">RM</th></tr>` +
    list.slice(0, 400).map((t) => `<tr><td class="nw">${shortDay(t.date)}</td><td class="ell" title="${esc(t.details)}">${esc(t.details || "-")}</td><td class="ell">${esc(catLabel(t.category))}</td><td class="n">${t.type === "Income" ? "+" : t.type === "Expenses" ? "−" : "↔ "}${Math.abs(t.amount).toFixed(2)}</td></tr>`).join("") +
    `</table>` + (list.length > 400 ? `<p class="muted small">Showing 400 of ${list.length}. Search to narrow down.</p>` : "");
}

/* ------------------------------------------------------------------ INVESTMENTS */
const platColor = (i) => css(`--series-${(i % 8) + 1}`);

function setupInvestments() {
  $("#inv-pick").innerHTML = `<option value="">All platforms</option>` + D.platforms.map((p) => `<option value="${p.id}">${esc(p.name)}${p.closed ? " (closed)" : ""}</option>`).join("");
  $("#inv-pick").onchange = renderInvestments;
}

function renderInvestments() {
  const open = D.platforms.filter((p) => !p.closed);
  const val = open.reduce((s, p) => s + p.value, 0), put = open.reduce((s, p) => s + p.invested, 0);
  const nk = open.filter((p) => p.id !== "kwsp"), nkv = nk.reduce((s, p) => s + p.value, 0), nki = nk.reduce((s, p) => s + p.invested, 0);
  $("#inv-tiles").innerHTML = [
    ["Total value", RM(val), "incl. KWSP"], ["Money put in", RM(put), "incl. KWSP"],
    ["Gain", RM(val - put), put ? pct((val - put) / put, 1) + " overall" : ""],
    ["Excl. KWSP", RM(nkv), `gain ${RM(nkv - nki)} (${nki ? pct((nkv - nki) / nki, 1) : "-"})`],
  ].map(([k, v, s]) => `<div class="tile"><div class="label">${k}</div><div class="v ${v.startsWith("-") ? "neg" : ""}">${v}</div><div class="s">${s}</div></div>`).join("");

  const pick = $("#inv-pick").value;
  const idx = (p) => D.platforms.indexOf(p);
  let datasets;
  if (!pick) {
    datasets = D.platforms.filter((p) => !p.closed).map((p) => line(p.name, p.series.filter((s) => s.value != null).map((s) => ({ x: s.date, y: s.value })), platColor(idx(p)), { stepped: false, pointRadius: 2 }));
  } else {
    const p = D.platforms.find((x) => x.id === pick);
    datasets = [
      line("Current value", p.series.filter((s) => s.value != null).map((s) => ({ x: s.date, y: s.value })), platColor(idx(p)), { pointRadius: 3 }),
      line("Money put in", p.series.filter((s) => s.invested != null).map((s) => ({ x: s.date, y: s.invested })), css("--text-secondary"), { borderDash: [5, 4], stepped: true, tension: 0 }),
    ];
  }
  const opts = baseOpts({ interaction: { mode: "nearest", intersect: false } });
  opts.scales.x = { type: "category", grid: { display: false }, ticks: { color: css("--text-secondary"), maxTicksLimit: 10 } };
  // category axis over all dates so lines share one x-axis
  const allDates = [...new Set(datasets.flatMap((d) => d.data.map((p) => p.x)))].sort();
  datasets.forEach((d) => { const m = new Map(d.data.map((p) => [p.x, p.y])); d.data = allDates.map((x) => m.get(x) ?? null); d.spanGaps = true; });
  draw("c-inv", { type: "line", data: { labels: allDates.map(dayName), datasets }, options: opts });

  const alloc = open.filter((p) => p.value > 0).sort((a, b) => b.value - a.value);
  draw("c-alloc", { type: "bar",
    data: { labels: alloc.map((p) => p.name), datasets: [bar("Value", alloc.map((p) => p.value), alloc.map((p) => platColor(idx(p))))] },
    options: baseOpts({ indexAxis: "y", plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => ` ${RM(c.parsed.x)} · ${pct(c.parsed.x / val, 1)}` } } },
      scales: { x: { grid: { color: css("--grid") }, ticks: { color: css("--text-secondary"), callback: (v) => RM(v) } }, y: { grid: { display: false }, ticks: { color: css("--text-secondary") } } } }) });

  $("#inv-table").innerHTML = `<div class="scroll"><table><tr><th>Platform</th><th class="n">Put in</th><th class="n">Value</th><th class="n">Gain</th><th>Updated</th></tr>` +
    D.platforms.map((p) => {
      const g = p.value - p.invested, d = p.updated ? daysAgo(p.updated, D.today) : null;
      return `<tr><td><b>${esc(p.name)}</b><div class="small muted">${esc(p.kind)}</div></td><td class="n">${RM(p.invested)}</td><td class="n">${p.closed ? "closed" : RM(p.value)}</td>
        <td class="n ${g < 0 ? "neg" : ""}">${p.closed ? "-" : RM(g) + `<div class="small muted">${p.invested > 0 ? pct(g / p.invested, 1) : ""}</div>`}</td>
        <td>${p.closed ? "" : d == null ? "-" : pill(d > 45 ? (d > 135 ? "critical" : "warning") : "good", d + "d ago")}</td></tr>`;
    }).join("") + `</table></div>`;
}

/* ------------------------------------------------------------------ FUTURE */
let base = null;

function setupFuture() {
  const w = window12(), set = new Set(w);
  const inW = D.transactions.filter((t) => set.has(ym(t.date)));
  const income = inW.filter(isEarned).reduce((s, t) => s + t.amount, 0) / 12;
  const cats = [...sumBy(inW.filter(isSpend), (t) => t.category)].map(([k, v]) => [k, v / 12]).filter((c) => c[1] > 5).sort((a, b) => b[1] - a[1]);
  const S = D.settings;
  const dob = new Date(S.dob + "T00:00:00"), now = new Date(D.today + "T00:00:00");
  const age = (now - dob) / (365.25 * 86400000);
  base = { income: Math.round(income * 100) / 100, cats, age, dob, salaryGross: S.monthly_salary || 0,
    growth: (S.lt_growth || 0.04) * 100, epf: (S.epf_dividend || 0.055) * 100 };
  const sl = [
    ["income", "Take-home income / month", 0, Math.max(15000, Math.ceil(base.income * 2)), "any", base.income, (v) => RM(v), "your last 12 months' average"],
    ["extra", "Extra income / month", 0, 3000, 50, 0, (v) => RM(v), "side income, raise not yet in the numbers…"],
    ["raise", "Salary rise / year", 0, 10, 0.5, 5, (v) => v + "%", "applies to income and KWSP"],
    ["inflation", "Spending rises / year", 0, 8, 0.5, 3, (v) => v + "%", "prices going up"],
    ["growth", "Investment growth / year", 0, 12, 0.5, base.growth, (v) => v + "%", "from Setting - an assumption"],
    ["epf", "KWSP dividend / year", 0, 8, 0.25, base.epf, (v) => v + "%", "from Setting - an assumption"],
  ];
  $("#sliders").innerHTML = sl.map(([id, lab, mn, mx, st, val, f, hint]) =>
    `<div class="slider"><label for="s-${id}">${lab} <b id="v-${id}"></b></label><input type="range" id="s-${id}" min="${mn}" max="${mx}" step="${st}" value="${val}" data-default="${val}"><div class="s">${hint}</div></div>`).join("");
  sl.forEach(([id, , , , , , f]) => { const el = $("#s-" + id); const upd = () => { $("#v-" + id).textContent = f(+el.value); }; el.oninput = () => { upd(); renderFuture(); }; upd(); });
  $("#cuts").innerHTML = base.cats.slice(0, 10).map(([c, v], i) =>
    `<div class="slider"><label for="cut-${i}">${esc(catLabel(c))} <b id="cv-${i}"></b></label><input type="range" id="cut-${i}" min="0" max="60" step="5" value="0" data-default="0"><div class="s">now ${RM(v)} / month</div></div>`).join("");
  base.cats.slice(0, 10).forEach((_, i) => { const el = $("#cut-" + i); const upd = () => { $("#cv-" + i).textContent = el.value === "0" ? "no cut" : "-" + el.value + "%"; }; el.oninput = () => { upd(); renderFuture(); }; upd(); });
  $("#reset").onclick = () => { document.querySelectorAll("#tab-future input[type=range]").forEach((el) => { el.value = el.dataset.default; el.dispatchEvent(new Event("input")); }); };
}

function simulate() {
  const v = (id) => +$("#s-" + id).value;
  const cuts = base.cats.map(([c, amt], i) => amt * (1 - (i < 10 ? +$("#cut-" + i).value : 0) / 100));
  const spend0 = cuts.reduce((s, x) => s + x, 0);
  const p = netWorthParts();
  let pot = D.cash + D.savings + p.inv, kwsp = p.kwsp;
  const otherNet = p.nw - pot - kwsp; // property, vehicles, valuables minus loans and cards - held at today's value
  const months = Math.round((Math.max(60, base.age + 25) - base.age) * 12); // to age 60, at least 25 years
  const out = [];
  for (let m = 1; m <= months; m++) {
    const yrs = m / 12;
    const income = (v("income")) * Math.pow(1 + v("raise") / 100, Math.floor(yrs)) + v("extra");
    const spend = spend0 * Math.pow(1 + v("inflation") / 100, Math.floor(yrs));
    pot = pot * (1 + v("growth") / 100 / 12) + (income - spend);
    kwsp = kwsp * (1 + v("epf") / 100 / 12) + 0.24 * base.salaryGross * Math.pow(1 + v("raise") / 100, Math.floor(yrs));
    const d = new Date(D.today + "T00:00:00"); d.setMonth(d.getMonth() + m);
    out.push({ date: d, age: base.age + yrs, pot, kwsp, nw: pot + kwsp + otherNet, surplus: income - spend, spend });
  }
  return { out, spend0, surplus0: v("income") + v("extra") - spend0, income0: v("income") + v("extra") };
}

function renderFuture() {
  const sim = simulate(), S = D.settings, p = netWorthParts();
  const potNow = D.cash + D.savings + p.inv;
  const ageDate = (a) => { const d = new Date(base.dob); d.setFullYear(d.getFullYear() + a); return d; };
  const efTarget = (S.ef_target_months || 6) * sim.spend0;
  // Goals come from the Wealth Summary sheet (rows 48-52). The emergency fund and the goals measured on the
  // Saving sheet share one savings pot, so they are stacked in that order.
  let stack = 0;
  const goals = [];
  const add = (name, target, by, measure, stacked, nowValue) => { if (stacked) stack += target; goals.push({ name, target, by, measure, need: stacked ? stack : target, nowValue }); };
  // emergency money today = CIMB savings + Touch 'n Go (same as Excel); stocks etc. are not emergency cash
  const tng = (D.platforms.find((x) => x.id === "tng") || { value: 0 }).value;
  const list = (D.goals || []).slice().sort((a, b) => ({ emergency: 0, stacked: 1, plain: 2 }[a.kind] - { emergency: 0, stacked: 1, plain: 2 }[b.kind]));
  for (const g of list) {
    const by = new Date(g.date + "T00:00:00");
    if (g.kind === "emergency") add(`Emergency fund (${S.ef_target_months || 6} months of spending)`, efTarget, by, "pot", true, D.savings + tng);
    else add(g.name, g.target, by, g.measure, g.kind === "stacked");
  }

  const fmtD = (d) => d.toLocaleDateString("en-GB", { month: "short", year: "numeric" });
  $("#goal-note").textContent = "Goals come from the Wealth Summary sheet. The emergency fund and the Saving-sheet goals share one savings pot, so they are counted on top of each other. 'Saved' = cash + savings + investments (not KWSP).";
  $("#goals").innerHTML = goals.map((g) => {
    const cur = g.measure === "nw" ? p.nw : potNow;
    const have = g.nowValue != null ? Math.max(0, Math.min(1, g.nowValue / g.target))
      : Math.max(0, Math.min(1, (cur - (g.need - g.target)) / g.target));
    const hit = sim.out.find((r) => (g.measure === "nw" ? r.nw : r.pot) >= g.need);
    let st;
    if (have >= 1 && cur >= g.need) st = pill("good", "Done");
    else if (!hit) st = pill("critical", `Not reached by ${Math.floor(sim.out[sim.out.length - 1].age)}`);
    else if (hit.date <= g.by) st = pill("good", "On track");
    else st = pill("warning", "Late");
    return `<div class="goal"><div class="head"><b>${esc(g.name)}</b>${st}</div>
      <div class="bar"><span style="width:${(have * 100).toFixed(1)}%"></span></div>
      <div class="meta">${RM(g.target)} by ${fmtD(g.by)} · now ${pct(have)} · ${hit ? "reached " + fmtD(hit.date) + " (age " + Math.floor(hit.age) + ")" : "not reached on this plan"}</div></div>`;
  }).join("");

  const at = (a) => sim.out.find((r) => r.age >= a) || sim.out[sim.out.length - 1];
  const in10 = at(base.age + 10), last = sim.out[sim.out.length - 1];
  $("#fu-tiles").innerHTML = [
    ["Left over / month", RM(sim.surplus0), sim.income0 ? pct(sim.surplus0 / sim.income0) + " of income" : ""],
    ["Spending / month", RM(sim.spend0), "after your cuts"],
    ["Saved in 10 years", RM(in10.pot), `at age ${Math.floor(in10.age)} - cash + savings + investments`],
    [`Net worth at ${Math.floor(last.age)}`, RM(last.nw), "incl. KWSP " + RM(last.kwsp)],
  ].map(([k, v, s]) => `<div class="tile"><div class="label">${k}</div><div class="v ${v.startsWith("-") ? "neg" : ""}">${v}</div><div class="s">${s}</div></div>`).join("");

  const yearly = sim.out.filter((r, i) => i % 12 === 11);
  draw("c-proj", { type: "line",
    data: { labels: yearly.map((r) => "age " + Math.floor(r.age)), datasets: [
      line("Net worth", yearly.map((r) => Math.round(r.nw)), css("--series-1")),
      line("Savings & investments", yearly.map((r) => Math.round(r.pot)), css("--series-2")),
    ] },
    options: baseOpts() });
}

/* ------------------------------------------------------------------ boot */
function renderAll() { renderOverview(); renderSpending(); renderInvestments(); renderFuture(); }

function showTab(name) {
  document.querySelectorAll(".tabs button").forEach((b) => b.setAttribute("aria-selected", b.dataset.tab === name));
  document.querySelectorAll(".tab").forEach((s) => (s.hidden = s.id !== "tab-" + name));
  Object.values(charts).forEach((c) => c.resize());
  try { localStorage.setItem("tab", name); } catch (e) { /* private mode */ }
  if (location.hash !== "#" + name) history.replaceState(null, "", "#" + name);
}

async function boot() {
  document.querySelectorAll(".tabs button").forEach((b) => (b.onclick = () => showTab(b.dataset.tab)));
  if (typeof Chart === "undefined") {
    $("#error").hidden = false;
    $("#error").textContent = "Charts need an internet connection the first time (Chart.js loads from cdnjs). Numbers below still work.";
  } else {
    Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
  }
  // Local server -> live workbook. Anywhere else (the GitHub Pages demo) -> the exported sample data.json
  const live = ["localhost", "127.0.0.1"].includes(location.hostname) && location.protocol === "http:"
    && !new URLSearchParams(location.search).has("demo");  // ?demo = preview the static demo locally
  try {
    const res = await fetch(live ? "/api/data" : "data.json", { cache: "no-store" });
    D = await res.json();
    if (D.error) throw new Error(D.error);
  } catch (e) {
    $("#error").hidden = false;
    $("#error").textContent = "Could not read the workbook: " + e.message;
    return;
  }
  $("#source").textContent = live ? `${D.source.file} · saved ${D.source.saved} · read-only`
    : `Demo · fictional sample data · as of ${dayName(D.today)}`;
  // on GitHub Pages (USER.github.io/REPO/) point the demo banner at the repository
  const gh = location.hostname.match(/^([^.]+)\.github\.io$/), repo = location.pathname.split("/").filter(Boolean)[0];
  if (gh && repo && $("#repo-link")) $("#repo-link").href = `https://github.com/${gh[1]}/${repo}`;
  setupSpending(); setupInvestments(); setupFuture();
  renderAll();
  let start = location.hash.slice(1);
  if (!start) { try { start = localStorage.getItem("tab") || "overview"; } catch (e) { start = "overview"; } }
  showTab(["overview", "spending", "investments", "future"].includes(start) ? start : "overview");
  matchMedia("(prefers-color-scheme: dark)").addEventListener("change", renderAll);
  addEventListener("hashchange", () => { const t = location.hash.slice(1); if (["overview", "spending", "investments", "future"].includes(t)) showTab(t); });
}
boot();
