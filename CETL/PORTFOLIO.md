# Clear Enough To Lead (CETL) — Founder & Project Portfolio

**Founder:** Jani
**Date:** August 2026
**Companion to:** the CETL Business Plan (v2.0)

> CETL's core teaching promise is *"you will build real things."* This portfolio exists
> to prove the founder already does — every item below is working software or a
> production-ready asset built by the founder, and together they demonstrate the exact
> path the curriculum teaches: web pages → local projects → interactive tools → working
> AI agents.

---

## 1. Working software built by the founder

### CETL Content Agent + MCP Server
A prompt-engineering agent that generates all CETL content — lesson plans, exercises,
troubleshooting FAQs, podcast outlines and show notes, member messaging, and marketing
copy — in a consistent, editorially controlled voice.

- **Stack:** Python CLI + Model Context Protocol (MCP) server integrating directly with
  Claude Desktop
- **Design choice worth noticing:** requires no API key — the agent builds structured,
  persona-governed prompts and lets Claude generate natively, which keeps operating
  costs near zero
- **What it demonstrates:** the founder ships the same category of tool the flagship
  module ("Building Mini Claude Agents") teaches

### CETL Business Plan Agent
An agent that *gathers the business's own data* (curriculum, podcast roadmap, project
manifest, verified research citations) and produces the complete business plan — then
converts it into a formatted Word document.

- **Stack:** Python; markdown-to-docx pipeline (headings, tables, footnotes,
  superscript citations); MCP tools for Claude Desktop
- **What it demonstrates:** end-to-end automation of a real business deliverable — the
  business plan accompanying this portfolio was produced by this tool
- **Discipline worth noticing:** the generation prompt forbids invented statistics —
  every market claim must cite one of 24 independently fact-checked sources

### Production document automation (professional practice)
In her professional role, the founder built and maintains document-processing automation
for a state education agency's civil-rights compliance program: parsing completed review
documents, performing live directory lookups, and generating compliance plans and formal
correspondence as finished Word documents.

- **Stack:** Python, python-docx, web scraping with graceful degradation, MCP server
  deployment
- **What it demonstrates:** the founder's tooling runs in a real institutional setting
  with real stakes — this is professional-grade capability, not hobby code

---

## 2. Product assets ready for launch

### Build With Claude — four-module curriculum
| # | Module | The artifact each student ships |
|---|---|---|
| 1 | HTML Basics | A live personal web page, built from scratch |
| 2 | Localhost Projects | A project running on her own machine — file vs. server understood |
| 3 | Before Apps | A small interactive project, end to end |
| 4 | Building Mini Claude Agents | A working personal agent solving a real problem on repeat |

Every module is scoped with explicit outcomes and is designed ADHD-first: short
sections, one concept at a time, a visible win per session, and honest treatment of the
places beginners actually get stuck.

### Clear Enough To Lead — podcast launch slate
| Episode | Theme |
|---|---|
| *The Bar Is Clear Enough* | What good-enough leadership looks like while managing a team, a brain, and a body at once |
| *Nobody Told Me This Would Happen At Work* | Leading through perimenopause symptoms that don't check your calendar |
| *The Burnout Math* | Why 40% today beats 120% today and nothing for three days |

### Verified research base
A 24-source, independently fact-checked evidence file grounding the business plan's
market analysis and product design — including the deliberate *exclusion* of widely
repeated figures that failed verification. The research discipline is itself a
differentiator: CETL's claims survive due diligence.

---

## 3. What this portfolio demonstrates to a funder

1. **Execution, pre-funding.** Everything above was built with zero outside capital, on
   ~10 hours a week, by one person — the same constraints the business plan budgets for.
2. **The product is proven on its founder.** CETL's curriculum is the documented path
   the founder herself took from zero coding background to shipped agents. The first
   case study already exists.
3. **Costs stay low because the founder automates.** The agents above replace the
   contractor spend (copywriting, document production, ops) that typically burns early
   funding.
4. **Authenticity moat.** The founder is the customer: a manager navigating ADHD and
   perimenopause while leading teams. The audience can tell — and it cannot be copied
   by a competitor hiring a curriculum writer.

---

*Live demonstrations of any item above are available on request. The business plan
(v2.0) contains the full market research, financial projections, and the $25,000
use-of-funds this portfolio supports.*
