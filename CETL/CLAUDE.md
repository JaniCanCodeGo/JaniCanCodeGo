# CLAUDE.md

## What this project does

Clear Enough To Lead (CETL) is a public-facing workgroup + podcast project, separate from Jani's private KungFuJay chief-of-staff project.

- **Teaching track ("Build With Claude"):** a workgroup teaching perimenopausal women with ADHD to build web projects (HTML → localhost projects → small interactive projects → mini Claude agents), with zero prior coding background assumed
- **Podcast track:** "Clear Enough To Lead" — leadership, life, and stress management, hosted by Jani

This is intentionally distinct from KungFuJay (Jani's private personal/work chief-of-staff agent). KungFuJay content and tools should never surface here, and vice versa.

## Architecture

`cetl_content_agent.py` holds the persona and all prompt templates, organized into five sections:

| Section | File location | Purpose |
|---|---|---|
| `curriculum` | `CURRICULUM_PROMPTS` | Teaching content (lesson-plan, exercise, troubleshooting-faq) |
| `podcast` | `PODCAST_PROMPTS` | Episode content (episode-outline, show-notes, interview-questions) |
| `audience` | `AUDIENCE_PROMPTS` | Messaging to workgroup members (welcome-message, encouragement-note, assignment-recap) |
| `host` | `HOST_PROMPTS` | Jani's own prep material (teaching-script, podcast-intro-script) |
| `promo` | `PROMO_PROMPTS` | Marketing/outreach (social-post, newsletter) |

`cetl_mcp_server.py` exposes one MCP tool per section, each calling the matching `build_*_prompt()` function. No API key is used anywhere — tools return a persona + request prompt; the calling Claude session generates the content natively.

## Data files (structured, for the website build)

- `data/curriculum.json` — modules with id, order, slug, summary, outcomes, status
- `data/podcast_episodes.json` — episodes with id, order, slug, topic, status, summary
- `data/site_manifest.json` — entry point describing the whole repo's data shape; other sessions/tools should read this first

Long-form written content lives in `content/` as markdown files named after each item's slug (e.g. `content/curriculum/html-basics.md` matches the `html-basics` module in `curriculum.json`).

## Content boundaries

- `content/for_audience/` — public-facing-adjacent (sent to enrolled members), not for the marketing site
- `content/for_host/` — private, Jani's own notes, never public
- `content/promo/` and `content/curriculum/` and `content/podcast/` — safe for public site use per `site_manifest.json` guidance

## Output files

CLI-generated prompts save to `outputs/{section}/` (gitignored), named by content type and timestamp.

## Voice

The full persona lives in `PERSONA` at the top of `cetl_content_agent.py`. It's public-facing: warm, direct, funny, never condescending to beginners, never toxic-positivity, professional enough for an audience she doesn't know personally. If the tone needs adjustment, that's the place to change it — do not copy KungFuJay's more casual private-voice persona here.
