# Cheat-Sheet: Using Your Installed Skills

No commands, no terminal. Just talk to Claude in plain English.

## karpathy-llm-wiki — your second brain

| You want to… | Just say… |
|---|---|
| **Add a source** (PDF, article, directive, notes) | *"Add this to my wiki"* — then paste it or point to the file |
| **Ask what you know** | *"What do I know about [topic]?"* |
| **Save an answer permanently** | *"Archive that to my wiki"* |
| **Clean up / fix broken links** | *"Lint my wiki"* |

The wiki grows every time you drop something in. You never organize it by hand —
Claude maintains it for you.

- `raw/` holds the original sources (never edited).
- `wiki/` holds the clean, compiled articles, plus an `index.md` table of contents
  and a `log.md` history.

A working example lives in `demo/`, seeded with the CRR 2 / APN guidance.

## stop-slop — make writing sound human

| You want to… | Just say… |
|---|---|
| **Remove the AI voice from a draft** | *"Run stop-slop on this"* or *"make this sound human"* |

Use it on grant narratives, business posts, podcast scripts, ADHD teaching
material, and CRR/LOF letters — anywhere the writing needs to sound like you.

## Where the skills live

- **In this repo:** already installed under `.claude/skills/`. They work whenever
  you work with Claude here. Nothing to set up.
- **On your own computer (optional, for use everywhere):** see
  `.claude/skills/README.md`, or just ask Claude to walk you through it.
