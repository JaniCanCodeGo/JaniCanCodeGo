/* Grant Writing Studio frontend.
   Accessibility notes: status changes are announced via the aria-live
   region; focus moves to the progress section on submit and to results on
   completion; all dynamic content uses semantic elements. */

(function () {
  "use strict";

  const form = document.getElementById("grant-form");
  const formError = document.getElementById("form-error");
  const progressSection = document.getElementById("progress-section");
  const resultsSection = document.getElementById("results-section");
  const stepsList = document.getElementById("steps");
  const liveStatus = document.getElementById("live-status");
  const resultsDiv = document.getElementById("results");
  const runBtn = document.getElementById("run-btn");

  const STEP_LABELS = {
    website_review: "Website review & 508 spot-check",
    ein: "EIN lookup / application prep",
    business_plan: "Research, forecast, business plan & prospectus",
    trademark: "Trademark search & TEAS packet",
    grant: "Grant narrative",
  };

  const FORECAST_FIELDS = [
    "grant_target", "base_revenue", "base_expenses", "revenue_growth_pct",
    "participants", "merch_units", "merch_price", "merch_unit_cost",
  ];

  let pollTimer = null;

  form.addEventListener("submit", async function (ev) {
    ev.preventDefault();
    formError.textContent = "";

    const orgName = document.getElementById("org_name").value.trim();
    if (!orgName) {
      formError.textContent = "Organization name is required.";
      document.getElementById("org_name").focus();
      return;
    }

    const forecast = {};
    for (const id of FORECAST_FIELDS) {
      const v = document.getElementById(id).value;
      if (v !== "") forecast[id] = Number(v);
    }

    const payload = {
      org_name: orgName,
      website_url: val("website_url"),
      ein: val("ein"),
      state: val("state"),
      entity_type: val("entity_type"),
      contact_email: val("contact_email"),
      mission: val("mission"),
      programs: val("programs"),
      merchandise: val("merchandise"),
      focus_areas: val("focus_areas"),
      trademark_name: val("trademark_name"),
      mark_in_use: document.getElementById("mark_in_use").checked,
      forecast: forecast,
    };

    runBtn.disabled = true;
    stepsList.innerHTML = "";
    resultsDiv.innerHTML = "";
    resultsSection.hidden = true;
    progressSection.hidden = false;
    announce("Agent run started.");
    progressSection.querySelector("h2").setAttribute("tabindex", "-1");
    progressSection.querySelector("h2").focus();

    try {
      const resp = await fetch("/api/runs", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (!resp.ok) throw new Error("Server error " + resp.status);
      const data = await resp.json();
      poll(data.run_id);
    } catch (err) {
      formError.textContent = "Could not start the run: " + err.message;
      runBtn.disabled = false;
    }
  });

  function val(id) { return document.getElementById(id).value.trim(); }

  function announce(msg) { liveStatus.textContent = msg; }

  function poll(runId) {
    pollTimer = setInterval(async function () {
      let state;
      try {
        const resp = await fetch("/api/runs/" + runId);
        state = await resp.json();
      } catch (err) {
        return; // transient; keep polling
      }
      renderSteps(state.steps || []);
      if (state.status !== "running") {
        clearInterval(pollTimer);
        runBtn.disabled = false;
        if (state.status === "done") {
          announce("Agent run complete. Results are available below.");
          renderResults(runId, state.result);
        } else {
          announce("Agent run failed.");
          formError.textContent =
            "The run failed. Details: " + (state.error || "unknown error");
        }
      }
    }, 1500);
  }

  function renderSteps(steps) {
    // Show latest status per step, in order of first appearance.
    const latest = new Map();
    for (const s of steps) latest.set(s.step, s);
    stepsList.innerHTML = "";
    for (const [step, s] of latest) {
      const li = document.createElement("li");
      li.className = s.status;
      const label = STEP_LABELS[step] || step;
      const stateWord = s.status === "done" ? "Complete"
        : s.status === "running" ? "In progress" : s.status;
      li.textContent = label + " — " + stateWord +
        (s.detail ? ": " + s.detail : "");
      stepsList.appendChild(li);
    }
  }

  function renderResults(runId, result) {
    resultsSection.hidden = false;
    resultsDiv.innerHTML = "";

    // Downloads
    const dl = card("Generated documents");
    const list = document.createElement("ul");
    for (const f of result.files || []) {
      const li = document.createElement("li");
      const a = document.createElement("a");
      a.href = "/api/runs/" + runId + "/files/" + encodeURIComponent(f);
      a.textContent = f;
      li.appendChild(a);
      list.appendChild(li);
    }
    dl.appendChild(list);

    // EIN
    const einCard = card("EIN");
    const ein = result.ein || {};
    const p = document.createElement("p");
    if (ein.found) {
      p.textContent = "EIN found: " + ein.ein + " (source: " + ein.source + ")";
    } else {
      p.textContent = "No EIN found. An IRS Form SS-4 worksheet was " +
        "prepared (see the Agent Reports download). Apply free at the IRS " +
        "online EIN Assistant — the responsible party must submit it.";
    }
    einCard.appendChild(p);
    if ((ein.matches || []).length) {
      einCard.appendChild(table(
        "Public-record matches (ProPublica Nonprofit Explorer)",
        ["EIN", "Name", "City", "State"],
        ein.matches.map(m => [m.ein, m.name, m.city, m.state])));
    }

    // Trademark
    let tmCard = null;
    if (result.trademark) {
      const tm = result.trademark;
      tmCard = card("Trademark: “" + tm.query + "”");
      if ((tm.api_hits || []).length) {
        tmCard.appendChild(table(
          "Potentially conflicting live marks",
          ["Mark", "Serial", "Status", "Owner"],
          tm.api_hits.map(h => [h.keyword, h.serial, h.status, h.owner])));
      } else {
        addPara(tmCard, tm.api_error
          ? "Automated search note: " + tm.api_error
          : "No conflicting live marks surfaced in the automated search.");
      }
      const links = document.createElement("ul");
      for (const l of tm.manual_search_links || []) {
        const li = document.createElement("li");
        const a = document.createElement("a");
        a.href = l.url; a.textContent = l.label;
        a.rel = "noopener"; a.target = "_blank";
        const sr = document.createElement("span");
        sr.className = "visually-hidden";
        sr.textContent = " (opens in a new tab)";
        a.appendChild(sr);
        li.appendChild(a);
        links.appendChild(li);
      }
      addPara(tmCard, "Confirm manually:");
      tmCard.appendChild(links);
      addPara(tmCard, "TEAS Plus packet prepared — estimated filing fee $" +
        tm.teas_application.fees.total_usd + " (" +
        tm.teas_application.fees.classes + " class(es)). File and sign at " +
        "teas.uspto.gov. " + tm.disclaimer);
    }

    // Website accessibility snapshot
    let siteCard = null;
    const site = result.website_review || {};
    if (site.ok && site.accessibility) {
      siteCard = card("Your website — Section 508 spot-check");
      addPara(siteCard, site.accessibility.passed + " of " +
        site.accessibility.total + " heuristic checks passed. " +
        site.accessibility.note);
      siteCard.appendChild(table("Checks",
        ["Check", "WCAG", "Status", "Detail"],
        site.accessibility.checks.map(c =>
          [c.id, c.wcag, c.status, c.detail])));
    }

    for (const el of [dl, einCard, tmCard, siteCard]) {
      if (el) resultsDiv.appendChild(el);
    }
    const h = resultsSection.querySelector("h2");
    h.setAttribute("tabindex", "-1");
    h.focus();
  }

  function card(title) {
    const div = document.createElement("div");
    div.className = "card";
    const h = document.createElement("h3");
    h.textContent = title;
    div.appendChild(h);
    return div;
  }

  function addPara(parent, text) {
    const p = document.createElement("p");
    p.textContent = text;
    parent.appendChild(p);
  }

  function table(captionText, headers, rows) {
    const t = document.createElement("table");
    const cap = document.createElement("caption");
    cap.textContent = captionText;
    t.appendChild(cap);
    const thead = document.createElement("thead");
    const tr = document.createElement("tr");
    for (const htext of headers) {
      const th = document.createElement("th");
      th.scope = "col";
      th.textContent = htext;
      tr.appendChild(th);
    }
    thead.appendChild(tr);
    t.appendChild(thead);
    const tbody = document.createElement("tbody");
    for (const row of rows) {
      const trb = document.createElement("tr");
      for (const cell of row) {
        const td = document.createElement("td");
        td.textContent = cell == null ? "" : String(cell);
        trb.appendChild(td);
      }
      tbody.appendChild(trb);
    }
    t.appendChild(tbody);
    return t;
  }
})();
