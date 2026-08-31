/**
 * parseMatches — runs INSIDE the browser page (page.evaluate) in both
 * Puppeteer and Playwright. Must not reference any outer-scope variables.
 *
 * Parses the vlr.gg /matches and /matches/results pages:
 *   - date headers use `.wf-label`
 *   - match cards are anchors `a.wf-module-item.match-item`
 *   - teams: `.match-item-vs-team` > `.match-item-vs-team-name`
 *   - scores: `.match-item-vs-team-score`
 *   - status (LIVE/Completed/Upcoming): `.match-item-status`
 *   - time/relative eta: `.match-item-eta`
 *   - event + stage: `.match-item-event` / `.match-item-event-series`
 */
export function parseMatches() {
  const clean = (s) => (s || "").replace(/\s+/g, " ").trim();

  const nodes = document.querySelectorAll("a.wf-module-item.match-item, .wf-label");
  const matches = [];
  let currentDate = null;

  const STATUS_RE = /(Completed|Upcoming|LIVE|Postponed|Showmatch)/;

  for (const node of nodes) {
    const cls = node.className || "";
    if (cls.includes("wf-label")) {
      // Date header like "Mon, August 31, 2026 Today", "Yesterday", ...
      const label = clean(node.textContent);
      if (label) {
        const noteMatch = label.match(/\b(Today|Yesterday)\b$/);
        currentDate = noteMatch ? clean(label.slice(0, noteMatch.index)) : label;
      }
      continue;
    }

    const href = node.getAttribute("href") || "";
    const idMatch = href.match(/^\/(\d+)\//);
    const url = href ? "https://www.vlr.gg" + href : null;

    // --- teams & scores ---
    const teams = [];
    node.querySelectorAll(".match-item-vs-team").forEach((teamEl) => {
      const nameEl = teamEl.querySelector(".match-item-vs-team-name");
      const scoreEl = teamEl.querySelector(".match-item-vs-team-score");
      const name = nameEl ? clean(nameEl.textContent) : null;
      const rawScore = scoreEl ? clean(scoreEl.textContent) : "";
      const score = /^\d+$/.test(rawScore) ? parseInt(rawScore, 10) : null;
      teams.push({
        name,
        score,
        winner: cls.includes("mod-winner") || teamEl.className.includes("mod-winner"),
      });
    });

    // --- status & time ---
    // On /matches, status lives in `.match-item-status` (e.g. "LIVE").
    // On /matches/results it's inline in `.match-item-eta` ("Completed 6h 55m"),
    // and the actual start time ("8:30 PM") is also inside eta.
    const statusEl = node.querySelector(".match-item-status");
    let status = statusEl ? clean(statusEl.textContent) : null;
    let time = null;       // absolute time, e.g. "8:30 PM"
    let timeUntil = null;  // relative, e.g. "6h 55m" / "2h 11m"
    const etaEl = node.querySelector(".match-item-eta");
    if (etaEl) {
      const etaText = etaEl.innerText
        ? etaEl.innerText.split("\n").map(clean).filter(Boolean).join(" ")
        : clean(etaEl.textContent);
      if (!status) {
        const s = etaText.match(STATUS_RE);
        status = s ? s[1] : null;
      }
      const abs = etaText.match(/\d{1,2}:\d{2}\s*[APap]\.?[Mm]\.?/);
      if (abs) time = clean(abs[0]);
      const rel = etaText.match(/\d+\s*[dhm]\w*(\s+\d+\s*[hm]\w*)?/);
      if (rel) timeUntil = clean(rel[0]);
    }

    // --- event & stage ---
    const eventEl = node.querySelector(".match-item-event");
    let event = null;
    let stage = null;
    if (eventEl) {
      const seriesEl = eventEl.querySelector(".match-item-event-series");
      stage = seriesEl ? clean(seriesEl.textContent) : null;
      // Remove the series text from the full event text
      const full = clean(eventEl.textContent);
      event = stage && full.endsWith(stage) ? clean(full.slice(0, full.length - stage.length)) : full;
    }

    matches.push({
      id: idMatch ? parseInt(idMatch[1], 10) : null,
      url,
      date: currentDate,
      time,
      time_until: timeUntil,
      status,
      team1: teams[0] || null,
      team2: teams[1] || null,
      event,
      stage,
    });
  }

  return matches;
}

/** Flatten parsed matches to CSV rows (shared by both scripts). */
export function toCsvRows(matches) {
  const header = [
    "id",
    "date",
    "time",
    "time_until",
    "status",
    "team1",
    "team1_score",
    "team2",
    "team2_score",
    "event",
    "stage",
    "url",
  ];
  const esc = (v) => {
    if (v === null || v === undefined) return "";
    const s = String(v);
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const rows = [header.join(",")];
  for (const m of matches) {
    rows.push(
      [
        m.id,
        esc(m.date),
        esc(m.time),
        esc(m.time_until),
        esc(m.status),
        esc(m.team1 && m.team1.name),
        m.team1 ? m.team1.score : "",
        esc(m.team2 && m.team2.name),
        m.team2 ? m.team2.score : "",
        esc(m.event),
        esc(m.stage),
        esc(m.url),
      ].join(",")
    );
  }
  return rows.join("\n");
}

/** Small CLI arg parser shared by both scripts. */
export function parseArgs(argv) {
  const args = {
    mode: "results", // "results" | "matches"
    pages: 1,
    out: null,
    csv: null,
    delay: 1500,
    headed: false,
  };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--mode") args.mode = argv[++i];
    else if (a === "--pages") args.pages = Math.max(1, parseInt(argv[++i], 10) || 1);
    else if (a === "--out") args.out = argv[++i];
    else if (a === "--csv") args.csv = argv[++i];
    else if (a === "--delay") args.delay = parseInt(argv[++i], 10) || 1500;
    else if (a === "--headed") args.headed = true;
  }
  if (!args.out) args.out = `vlr_${args.mode}.json`;
  if (!args.csv) args.csv = `vlr_${args.mode}.csv`;
  return args;
}
