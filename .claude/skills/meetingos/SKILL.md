---
name: meetingos
description: Turn a Granola meeting into an interactive dashboard. Pulls a meeting transcript from Granola, detects the meeting type (brainstorm, planning, sales, standup, 1:1, interview, status/review, decision), extracts decisions, action items, open questions, and key quotes, then renders a self-contained interactive HTML dashboard tailored to that meeting type. Use when the user asks to summarize, recap, "build a dashboard for", "make interactive", or "do the work from" a Granola meeting or their meeting notes.
---

# Meeting OS

Turn a recorded Granola meeting into an interactive, type-aware dashboard.

## Prerequisites

This skill relies on the **Granola MCP connector** being connected to the
session. The Granola tools are named `mcp__Granola__*`. If they are not
available, tell the user to connect Granola, then stop.

If a Granola tool's schema is not loaded, fetch it first with
`ToolSearch` (e.g. `select:mcp__Granola__list_meetings`).

## Workflow

### 1. Pick the meeting

- If the user named a specific meeting (by title, attendee, or date), find it
  with `mcp__Granola__query_granola_meetings` (best for recent / short-term
  questions) or `mcp__Granola__list_meetings` (best for a date range).
- If the user did not specify one, call `mcp__Granola__list_meetings` and show
  the 5 most recent meetings with their titles and dates, then ask which one —
  unless the request clearly means "the latest meeting," in which case take the
  most recent and say which one you picked.
- Once identified, pull the full content with
  `mcp__Granola__get_meeting_transcript` (and `mcp__Granola__get_meetings` for
  the structured notes/summary if available). Always work from the real
  transcript text — never invent content that isn't in the meeting.

### 2. Detect the meeting type

Read the transcript and classify it into ONE primary type. Use the signals
below; when ambiguous, pick the closest and note the runner-up in the output.

| Type | Signals | Dashboard emphasis |
|------|---------|--------------------|
| `brainstorm` | many ideas, "what if", divergent, few decisions | **Idea board** — ideas ranked, a "winning idea" callout |
| `planning` | dates, milestones, owners, sequencing | **Timeline** — phases/milestones with owners + dates |
| `sales` | prospect, pricing, objections, next steps | **One-pager** — needs, objections→responses, deal next steps |
| `standup` | "yesterday/today/blockers", round-robin | **Per-person board** — done / doing / blockers |
| `one_on_one` | 2 people, feedback, growth, personal | **1:1 recap** — wins, feedback, growth items, follow-ups |
| `interview` | candidate, questions, evaluation | **Scorecard** — strengths, concerns, signal per competency |
| `status_review` | progress, metrics, risks, retro | **Status board** — on-track / at-risk / blocked + metrics |
| `decision` | options weighed, a choice made | **Decision record** — options, criteria, chosen path, why |

### 3. Extract the core elements (every type)

From the transcript, pull:

- **Decisions** — what was decided, and by/with whom if stated.
- **Action items** — task, owner, and due date if mentioned. Mark owner as
  "Unassigned" rather than guessing.
- **Open questions** — anything left unresolved or explicitly parked.
- **Key quotes** — 2–5 short, verbatim, high-signal lines (attribute the
  speaker if the transcript identifies them).

Plus the **type-specific** content from the table above.

Do not fabricate. If a category is empty, render it as "None captured" — that's
a valid and useful result.

### 4. Render the dashboard

Generate a **single self-contained HTML file** (all CSS inline in a `<style>`
block, all interactivity in vanilla `<script>` — no external CDNs, fonts, or
network calls, so it opens offline by double-click).

Write it to `meetingos_output/` in the current working directory, named
`{slugified-meeting-title}_{YYYY-MM-DD}.html`. Create the folder if needed.

Build the HTML by following `reference/dashboard_template.html` — it shows the
required layout, styling, and the interactivity (tabbed sections, an
action-item checklist with localStorage persistence, copy-to-clipboard for the
action list). Adapt the highlighted "type-specific" panel to match the detected
meeting type (idea board, timeline, one-pager, etc.).

Requirements for the output:
- Header: meeting title, date, detected type (with a confidence note), attendees.
- A summary line (1–2 sentences) at the top.
- Cards/sections for Decisions, Action Items (interactive checklist), Open
  Questions, Key Quotes, and the type-specific panel.
- Responsive and readable; works as a plain file:// page.

### 5. Hand off

After writing the file:
- Tell the user the path and a one-paragraph recap (type + headline outcome).
- Surface the file with `SendUserFile` so they can open it directly.
- Offer obvious follow-ups based on type (e.g. "draft the follow-up email",
  "create calendar holds for the action items", "turn the winning idea into a
  spec"). Many of these map to other connected tools (Gmail, Google Calendar).

## Notes

- Keep everything grounded in the actual transcript. Accuracy beats polish.
- If the transcript is very long, summarize faithfully but preserve exact
  wording for decisions, action items, and quotes.
- This skill reads meeting data and writes a local HTML file only. It never
  sends anything externally unless the user explicitly asks for a follow-up.
