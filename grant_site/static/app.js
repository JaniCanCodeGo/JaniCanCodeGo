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
    budget: "Line-item budget & justification",
    readiness: "Registrations & attachments readiness",
    trademark: "Trademark search & TEAS packet",
    grant: "Grant narrative",
    rfp: "RFP question-by-question response",
    tailor: "AI prose tailoring",
  };

  const BUDGET_ITEM_IDS = ["b_personnel", "b_travel", "b_equipment",
    "b_supplies", "b_contractual", "b_other"];

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
      mailing_address: val("mailing_address"),
      responsible_party: val("responsible_party"),
      start_date: val("start_date"),
      mission: val("mission"),
      programs: val("programs"),
      merchandise: val("merchandise"),
      focus_areas: val("focus_areas"),
      trademark_name: val("trademark_name"),
      mark_in_use: document.getElementById("mark_in_use").checked,
      forecast: forecast,
      county: val("county"),
      local_need: val("local_need"),
      funder_name: val("funder_name"),
      rfp_text: val("rfp_text"),
      budget_items: (function () {
        const items = {};
        for (const id of BUDGET_ITEM_IDS) {
          const v = document.getElementById(id).value;
          if (v !== "") items[id.slice(2)] = Number(v);
        }
        return items;
      })(),
      fringe_pct: numOrNull("fringe_pct"),
      indirect_pct: numOrNull("indirect_pct"),
      match_amount: numOrNull("match_amount"),
      have_items: Array.from(
        document.querySelectorAll("#have_items input:checked"))
        .map(cb => cb.value),
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

  function numOrNull(id) {
    const v = document.getElementById(id).value;
    return v === "" ? null : Number(v);
  }

  // Grant opportunity search widget
  const oppBtn = document.getElementById("opp_btn");
  oppBtn.addEventListener("click", async function () {
    const q = val("opp_query");
    const box = document.getElementById("opp_results");
    if (!q) { box.textContent = "Enter a search term first."; return; }
    oppBtn.disabled = true;
    box.textContent = "Searching Grants.gov and the California Grants Portal…";
    try {
      const resp = await fetch("/api/opportunities/search?q=" +
        encodeURIComponent(q));
      const data = await resp.json();
      box.innerHTML = "";
      renderOppList(box, "Federal (Grants.gov)", data.federal);
      renderOppList(box, "California Grants Portal", data.california);
      const links = document.createElement("p");
      links.append("Search manually: ");
      for (const l of data.manual_links || []) {
        const a = document.createElement("a");
        a.href = l.url; a.textContent = l.label;
        a.rel = "noopener"; a.target = "_blank";
        links.append(a, "  ");
      }
      box.appendChild(links);
    } catch (err) {
      box.textContent = "Search failed: " + err.message;
    } finally {
      oppBtn.disabled = false;
    }
  });

  function renderOppList(box, title, result) {
    const h = document.createElement("h4");
    h.textContent = title;
    box.appendChild(h);
    if (!result || !result.ok) {
      addPara(box, "Unavailable right now" +
        (result && result.error ? " (" + result.error + ")" : "") +
        " — use the manual links below.");
      return;
    }
    if (!result.results.length) { addPara(box, "No open matches found."); return; }
    const ul = document.createElement("ul");
    for (const r of result.results) {
      const li = document.createElement("li");
      const a = document.createElement("a");
      a.href = r.url; a.textContent = r.title || r.number || "Untitled";
      a.rel = "noopener"; a.target = "_blank";
      li.appendChild(a);
      const bits = [r.agency, r.close_date ? "closes " + r.close_date : ""]
        .filter(Boolean).join(" — ");
      if (bits) li.append(" (" + bits + ")");
      ul.appendChild(li);
    }
    box.appendChild(ul);
  }

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
      p.textContent = "No EIN found. A completed IRS Form SS-4 application " +
        "document with a step-by-step filing guide is in your downloads " +
        "above. Filing at the IRS online EIN Assistant is free and takes " +
        "about 10 minutes.";
    }
    einCard.appendChild(p);
    if ((ein.matches || []).length) {
      einCard.appendChild(table(
        "Public-record matches (ProPublica Nonprofit Explorer)",
        ["EIN", "Name", "City", "State"],
        ein.matches.map(m => [m.ein, m.name, m.city, m.state])));
    }

    // Readiness
    let readyCard = null;
    if (result.readiness) {
      readyCard = card("Application readiness");
      addPara(readyCard, result.readiness.ready + " of " +
        result.readiness.total + " registrations and attachments in place. " +
        "The full checklist with links is in your downloads.");
      const gaps = result.readiness.items
        .filter(i => i.status === "action_needed").map(i => i.label);
      if (gaps.length) {
        addPara(readyCard, "Blocking items: " + gaps.join("; "));
      }
    }

    // Budget
    let budgetCard = null;
    if (result.budget) {
      budgetCard = card("Project budget");
      addPara(budgetCard, "Line items total $" +
        result.budget.total.toLocaleString() +
        (result.budget.allocated_by_default_pcts
          ? " (auto-allocated from your request using standard percentages — replace with actual figures before submission)."
          : " (from your figures)."));
    }

    // RFP
    let rfpCard = null;
    if (result.rfp) {
      rfpCard = card("RFP response");
      addPara(rfpCard, result.rfp.questions.length +
        " funder question(s) extracted and answered." +
        (result.rfp.deadlines.length
          ? " Deadline(s) found: " + result.rfp.deadlines.join("; ") + "."
          : "") +
        (result.rfp.limits.length
          ? " Limits: " + result.rfp.limits.map(l => l.limit + " " + l.unit).join(", ") + "."
          : ""));
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

    for (const el of [dl, readyCard, budgetCard, rfpCard, einCard, tmCard,
                      siteCard]) {
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
