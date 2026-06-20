# 🐾 START HERE — Your Obsidian Second Brain

This vault is the complete setup from the "30 Obsidian workflow plugins and
setups" slides, already built for you. Folders, templates, a `CLAUDE.md`,
and starter notes are all in place. This guide gets you from zero to a
working AI-powered second brain in about 30 minutes.

> **Note:** Software can't click buttons inside the Obsidian app for you, so
> the plugin installs below are quick manual steps. Everything else (folder
> structure, templates, daily-notes config, CLAUDE.md) is already done.

---

## ✅ What's already set up for you

- **PARA folder structure** — `00 - Inbox` through `07 - Attachments`
- **`CLAUDE.md`** — tells Claude your vault structure every session
- **6 templates** — Daily Note, Project, Meeting, Person, Weekly Review
- **Core plugins pre-configured** — Daily Notes points at `04 - Notes/Daily
  Notes` and uses the Daily Note Template; new notes land in `00 - Inbox`;
  attachments go to `07 - Attachments`
- **Starter notes** — 3 example projects, an example daily note, a permanent note

---

## 1️⃣ Open this folder as a vault (2 min)

1. Download / copy this `obsidian-vault` folder somewhere permanent
   (e.g., your Desktop, or a synced folder like iCloud / Dropbox / Google Drive).
2. Open the **Obsidian** app → **Open folder as vault** → pick this folder.
3. If Obsidian asks, click **Trust author and enable plugins**.

You'll see the numbered folders and the example notes immediately.

---

## 2️⃣ Install the 5 starter plugins (10 min)

These are the "Starter Stack" from Slide 14. In Obsidian:

**Settings (⚙️) → Community plugins → Turn on community plugins → Browse**,
then search each one, click **Install**, then **Enable**:

| # | Plugin | What it does |
|---|--------|--------------|
| 1 | **Smart Connections** | AI chat with ALL your notes; finds hidden links |
| 2 | **Templater** | Auto-fills new notes with structure (supercharges templates) |
| 3 | **Dataview** | Turns your vault into a searchable database |
| 4 | **Periodic Notes** | Daily / weekly / monthly notes |
| 5 | **Tasks** | Track `- [ ]` to-dos across the whole vault |

> The templates in `06 - Templates` already use the `{{date}}` / `{{title}}`
> syntax, so they work with the built-in Templates core plugin even before
> you install Templater.

### Nice-to-have extras (from Slides 4–6)
Calendar, Kanban, Excalidraw, QuickAdd, Natural Language Dates, Advanced
Tables, Style Settings, Obsidian Git. Install any that appeal to you — see
`06 - Templates/PLUGIN-GUIDE.md` for the full list and what each one does.

---

## 3️⃣ Connect Claude to your vault (5 min)

This is the part you asked about — "I have a connector but don't know how to
set it up." There are two common ways; pick **one**:

### Option A — Claude's Obsidian connector (easiest)
1. In **Claude** (desktop app or claude.ai), open **Settings → Connectors**.
2. Find the **Obsidian** connector and click **Connect / Configure**.
3. When it asks for your **vault location / folder**, point it at the folder
   you opened in step 1 (the one containing this README and `CLAUDE.md`).
4. Approve the permission prompt. Claude can now read and write notes here.

### Option B — MCPVault (the "mcpvault" from Slide 11)
1. Install the MCP server it references (search "obsidian mcp" / "mcpvault").
2. Add it to your Claude MCP config, pointing its `vault path` at this folder.
3. Restart Claude. You'll get ~14 vault commands (create note, search, link…).

**Either way, the key step is the same: point the connector at THIS folder.**
Because `CLAUDE.md` lives in the root, Claude will automatically learn your
folder structure, naming conventions, and active projects.

> Not sure which connector you already have? Tell me its name and I'll give
> you the exact click-path.

---

## 4️⃣ Try it — your first workflows

Once connected, ask Claude things like:

- *"Read my CLAUDE.md, then create a project note for ⟨X⟩ in 01 - Projects."*
- *"Turn the notes in 00 - Inbox/Quick Capture into atomic notes and file them."*
- **Meeting Processor (Slide 7):** paste messy meeting notes →
  *"Turn this into a meeting note using my Meeting Note Template and pull out action items."*
- **Weekly Review (Slide 7):** *"Scan my daily notes from this week and draft
  a Weekly Review."*
- **Idea Cross-Pollinator (Slide 8):** *"Find unexpected connections between
  my notes."*

---

## 5️⃣ The 30-Minute timeline (Slide 15)

| When | Do this |
|------|---------|
| Hour 1 | Open vault, install the 5 starter plugins (steps 1–2 above) |
| Hour 2 | Install Smart Connections, chat with your vault |
| Hour 3 | Connect Claude (step 3), run your first query |
| Day 2 | Write 10 notes about your current projects |
| Day 3 | Set up your Weekly Review routine |
| Week 2 | Full AI second brain running 🎉 |

---

## 📁 Folder cheat-sheet (PARA)

- **00 - Inbox** — capture first, sort later
- **01 - Projects** — has a goal + an end date
- **02 - Areas** — ongoing, no end date
- **03 - Resources** — topics & references
- **04 - Notes** — Daily / People / Permanent / Reference
- **05 - Archives** — done or inactive
- **06 - Templates** — your note templates
- **07 - Attachments** — images, PDFs

Questions? Just ask me to walk you through any step.
