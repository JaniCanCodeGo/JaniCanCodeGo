/* Grant Writing Studio (web edition) — UI wiring.
   Runs the whole agent in the browser and offers every document as a
   download (.doc opens in Microsoft Word; .md is plain text). */

(function () {
  "use strict";

  const $ = id => document.getElementById(id);
  const form = $("grant-form"), formError = $("form-error");
  const progressSection = $("progress-section"), resultsSection = $("results-section");
  const stepsList = $("steps"), liveStatus = $("live-status"), resultsDiv = $("results");
  const runBtn = $("run-btn");

  const val = id => ($(id) ? $(id).value.trim() : "");
  const num = id => { const v = val(id); return v === "" ? null : Number(v); };

  function announce(msg) { liveStatus.textContent = msg; }
  function step(label, detail) {
    const li = document.createElement("li");
    li.className = "done";
    li.textContent = label + " — Complete" + (detail ? ": " + detail : "");
    stepsList.appendChild(li);
  }

  /* ---- document download helpers ---- */
  function docToMarkdown(doc) {
    return "# " + doc.title + "\n\n" +
      doc.sections.map(([h, b]) => "## " + h + "\n\n" + b).join("\n\n") + "\n";
  }
  function escapeHtml(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
  function docToWordHtml(doc) {
    const body = doc.sections.map(([h, b]) =>
      "<h2>" + escapeHtml(h) + "</h2>" +
      b.split("\n\n").map(p =>
        /^[A-Za-z].*  / .test(p) || p.includes("  $") // crude table detection
          ? "<pre style='font-family:Consolas,monospace;font-size:10pt'>" + escapeHtml(p) + "</pre>"
          : "<p>" + escapeHtml(p).replace(/\n/g, "<br>") + "</p>").join("")
    ).join("");
    return "<html xmlns:w='urn:schemas-microsoft-com:office:word'><head><meta charset='utf-8'>" +
      "<title>" + escapeHtml(doc.title) + "</title></head><body style=\"font-family:Calibri,Arial,sans-serif;\">" +
      "<h1>" + escapeHtml(doc.title) + "</h1>" + body + "</body></html>";
  }
  function safeName(t) { return t.replace(/[^\w\s-]/g, "").trim().replace(/\s+/g, "_").slice(0, 80); }
  function download(filename, content, mime) {
    const blob = new Blob([content], { type: mime });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
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
  function docCard(doc) {
    const c = card(doc.title);
    const row = document.createElement("p");
    const word = document.createElement("button");
    word.type = "button"; word.className = "secondary";
    word.textContent = "Download for Word (.doc)";
    word.addEventListener("click", () =>
      download(safeName(doc.title) + ".doc", docToWordHtml(doc), "application/msword"));
    const md = document.createElement("button");
    md.type = "button"; md.className = "secondary"; md.style.marginLeft = "0.5rem";
    md.textContent = "Download text (.md)";
    md.addEventListener("click", () =>
      download(safeName(doc.title) + ".md", docToMarkdown(doc), "text/markdown"));
    row.append(word, md);
    c.appendChild(row);
    const details = document.createElement("details");
    const sum = document.createElement("summary");
    sum.textContent = "Preview this document";
    details.appendChild(sum);
    for (const [h, b] of doc.sections) {
      const h4 = document.createElement("h4"); h4.textContent = h;
      const pre = document.createElement("pre");
      pre.style.whiteSpace = "pre-wrap"; pre.style.fontFamily = "inherit";
      pre.textContent = b;
      details.append(h4, pre);
    }
    c.appendChild(details);
    return c;
  }

  /* ---- opportunity search widget ---- */
  const oppBtn = $("opp_btn");
  if (oppBtn) oppBtn.addEventListener("click", async function () {
    const q = val("opp_query"), box = $("opp_results");
    if (!q) { box.textContent = "Enter a search term first."; return; }
    oppBtn.disabled = true;
    box.textContent = "Searching Grants.gov…";
    const data = await Engine.searchOpportunities(q);
    box.innerHTML = "";
    if (data.federal.ok && data.federal.results.length) {
      const ul = document.createElement("ul");
      for (const r of data.federal.results) {
        const li = document.createElement("li");
        const a = document.createElement("a");
        a.href = r.url; a.textContent = r.title; a.rel = "noopener"; a.target = "_blank";
        li.appendChild(a);
        const bits = [r.agency, r.close_date ? "closes " + r.close_date : ""].filter(Boolean).join(" — ");
        if (bits) li.append(" (" + bits + ")");
        ul.appendChild(li);
      }
      box.appendChild(ul);
    } else {
      addPara(box, "The automated search could not run from the browser — use these links:");
    }
    const links = document.createElement("p");
    links.append("Search directly: ");
    for (const l of data.manualLinks) {
      const a = document.createElement("a");
      a.href = l.url; a.textContent = l.label; a.rel = "noopener"; a.target = "_blank";
      links.append(a, "  ");
    }
    box.appendChild(links);
    oppBtn.disabled = false;
  });

  /* ---- main run ---- */
  form.addEventListener("submit", async function (ev) {
    ev.preventDefault();
    formError.textContent = "";
    const orgName = val("org_name");
    if (!orgName) {
      formError.textContent = "Organization name is required.";
      $("org_name").focus();
      return;
    }
    runBtn.disabled = true;
    stepsList.innerHTML = "";
    resultsDiv.innerHTML = "";
    resultsSection.hidden = true;
    progressSection.hidden = false;
    announce("Agent run started.");

    const org = {
      name: orgName, mission: val("mission"), programs: val("programs"),
      merchandise: val("merchandise"), contact_email: val("contact_email"),
    };
    const focus = val("focus_areas").split(",").map(s => s.trim()).filter(Boolean);

    // 1. EIN
    const einResult = await Engine.lookupEin(orgName, val("state"), val("ein"));
    step("EIN lookup / application prep", einResult.found
      ? "EIN " + einResult.ein : "No EIN confirmed — completed SS-4 document prepared");

    // 2. Research + forecast
    const pains = Engine.researchPainpoints(focus, val("local_need"), val("county"));
    const fc = Engine.buildForecast({
      grant_target: num("grant_target") ?? 25000,
      base_revenue: num("base_revenue") ?? undefined,
      base_expenses: num("base_expenses") ?? undefined,
      revenue_growth_pct: num("revenue_growth_pct") ?? undefined,
      participants: num("participants") ?? undefined,
      merch_units: num("merch_units") ?? undefined,
      merch_price: num("merch_price") ?? undefined,
      merch_unit_cost: num("merch_unit_cost") ?? undefined,
    });
    step("Research & 3-year forecast",
      pains.painpoints.length + " documented pain points; forecast built from printed assumptions");

    // 3. Budget
    const items = {};
    for (const k of ["personnel", "travel", "equipment", "supplies", "contractual", "other"]) {
      const v = num("b_" + k); if (v) items[k] = v;
    }
    const budget = Engine.buildBudget(fc.assumptions.grant_target, items,
      num("fringe_pct"), num("indirect_pct"), num("match_amount"));
    step("Line-item budget & justification", "totals " + "$" + Math.round(budget.total).toLocaleString());

    // 4. Readiness
    const have = Array.from(document.querySelectorAll("#have_items input:checked")).map(cb => cb.value);
    const readiness = Engine.assessReadiness(einResult, have);
    step("Registrations & attachments readiness", readiness.ready + "/" + readiness.total + " in place");

    // 5. Trademark
    const mark = val("trademark_name") || orgName;
    const tm = Engine.searchTrademark(mark, val("merchandise") + " " + val("programs"));
    step("Trademark knockout & TEAS packet", tm.variants.length + " variants; classes " + tm.classes.join(", "));

    // 6. Documents
    const plan = Engine.buildBusinessPlan(org, pains, fc, tm);
    const grant = Engine.buildGrantNarrative(org, pains, fc, einResult, budget);
    const budgetDoc = Engine.buildBudgetDocument(orgName, budget, fc);
    const readyDoc = Engine.buildReadinessDocument(orgName, readiness);
    const teasDoc = Engine.buildTeasDocument(tm,
      { name: orgName, email: val("contact_email"), entity_type: val("entity_type") },
      val("merchandise") || val("programs"), $("mark_in_use").checked);
    const docs = [plan, grant, budgetDoc, readyDoc, teasDoc];
    if (!einResult.found) {
      docs.push(Engine.buildSs4Document({
        legal_name: orgName, state: val("state"), county: val("county"),
        mailing_address: val("mailing_address"),
        responsible_party: val("responsible_party"),
        entity_type: val("entity_type"), start_date: val("start_date"),
        activity: val("programs"),
      }));
    }
    const rfpText = val("rfp_text");
    if (rfpText) {
      const parsed = Engine.extractRfp(rfpText);
      const grantSec = Object.fromEntries(grant.sections);
      docs.push(Engine.buildRfpResponse(parsed, {
        orgName,
        need: grantSec["Statement of Need"],
        project: grantSec["Project Description and Goals"],
        budget: Engine.requestJustification(budget, fc) + "\n\nSee the attached line-item Project Budget & Justification.",
        evaluation: grantSec["Evaluation Plan"],
        capacity: grantSec["Organizational Capacity"],
        sustainability: grantSec["Sustainability"],
      }, val("funder_name")));
      step("RFP question-by-question response", parsed.questions.length + " funder questions answered");
    }
    step("Documents generated", docs.length + " documents ready to download");

    // Render results
    resultsSection.hidden = false;
    if (einResult.error) {
      const c = card("EIN search note");
      addPara(c, einResult.error);
      resultsDiv.appendChild(c);
    }
    if (einResult.matches && einResult.matches.length > 1 && !einResult.found) {
      const c = card("Possible EIN matches found — pick yours");
      addPara(c, einResult.matches.map(m => `${m.ein} — ${m.name} (${m.city}, ${m.state})`).join("  |  "));
      addPara(c, "If one of these is your organization, enter that EIN in the form and run again.");
      resultsDiv.appendChild(c);
    }
    for (const d of docs) resultsDiv.appendChild(docCard(d));
    const tmCard = card("Confirm the trademark manually");
    const ul = document.createElement("ul");
    for (const l of tm.manualLinks) {
      const li = document.createElement("li");
      const a = document.createElement("a");
      a.href = l.url; a.textContent = l.label; a.rel = "noopener"; a.target = "_blank";
      li.appendChild(a); ul.appendChild(li);
    }
    tmCard.appendChild(ul);
    resultsDiv.appendChild(tmCard);

    announce("Agent run complete. " + docs.length + " documents are ready below.");
    const h = resultsSection.querySelector("h2");
    h.setAttribute("tabindex", "-1");
    h.focus();
    runBtn.disabled = false;
  });
})();
