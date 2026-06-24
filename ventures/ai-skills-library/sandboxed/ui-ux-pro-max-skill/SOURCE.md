# ui-ux-pro-max-skill — source & safe use

- **Upstream:** https://github.com/nextlevelbuilder/ui-ux-pro-max-skill
- **Retrieved:** 2026-06-23 (default branch `main`)
- **License:** MIT
- **What it does:** Supplies AI coding assistants with curated UI/UX design knowledge
  (color palettes, font pairings, UI styles, component specs) and ships `uipro-cli` to
  install those skill files into a project. Optional Python scripts generate
  logos/icons/mockups via Google Gemini if you provide a key.

## Good for (my ventures)
The day you build a website or landing page — for the podcast, a business idea, or
the ADHD + AI offering.

## How to run it SAFELY
Run inside a disposable sandbox when generating designs.

```bash
# inside a sandbox only
npx uipro-cli init     # copies design skill files into a project
# Gemini image-gen is optional and needs your own GEMINI_API_KEY in .env
```

## Caveats (from the audit)
- The CLI runs shell commands (`unzip`/`cp`/`Expand-Archive`) on local paths it
  generates — low risk, but it's real execution.
- `npx uipro-cli` / `npx shadcn` pull live packages from npm at run time (normal
  supply-chain consideration).
- Optional Gemini generation sends your prompts/images to Google if you opt in.
