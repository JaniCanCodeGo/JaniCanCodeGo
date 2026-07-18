# Forgotten Children ’94 — Brand Guide

**How to use this file:** paste it (or point Claude at it) at the start of any design
session — website pages, social graphics, podcast art, slide decks, Canva briefs.
It contains everything needed to reproduce the brand: tokens, type, voice, layout
patterns, and hard rules. The reference implementation is `website/index.html`.

---

## 1. Brand essence

- **Name:** The Forgotten Children of the 1994 Crime Bill (short: *Forgotten Children ’94*)
- **Tagline:** *They sentenced our parents. Nobody counted us.*
- **Mission in three verbs:** Count us. Hear us. Heal us.
- **Personality:** a protest poster with a bibliography. Dignified, direct, furious
  about the right things, never pitying, always sourced.
- **Visual reference:** golden-ratio poster style — a huge gold circle, blue duotone
  imagery, one giant condensed headline, small emotional micro-copy columns.

## 2. Voice — "Impact language"

Short declaratives. First person plural for the community ("we," "us"), first person
singular for the founder's story. Verbs over adjectives. No euphemisms, no pity, no
"at-risk youth" framing — we are survivors of policy, not statistics with legs.

**Sounds like:**
- "They sentenced our parents. Nobody counted us. So we're counting ourselves."
- "The prediction was wrong."
- "One sourced number, one lived story, every episode."

**Never sounds like:**
- "Sadly, these vulnerable children often fall through the cracks…" (pity)
- "Studies suggest a possible correlation…" (hedging — either it's sourced and we
  say it plainly, or we don't say it)
- Any statistic without a named source. Where a popular number is a myth, we say
  "that's a myth" out loud. Honesty *is* the brand.

**Structural signature:** every claim pattern is *big statement → precise number →
named source*. Headlines are single words or short phrases, often split with a
color accent on the final syllables (FORGOT/TEN, COUNT/ED, UN/BOUND, HE/ARD).

## 3. Color

Primary colors, poster-weight. Gold is the hero; use it in exactly one bold place
per composition (the circle, the active tab, the accent bar). Red is scarce:
accent syllables, warnings, the "honesty" callouts. Blue is the workhorse.

### Core palette
| Token | Hex | Role |
|---|---|---|
| Gold | `#F2A900` | Hero accent: circle motif, buttons, active states, highlights |
| Blue | `#1E5AA8` | Links, primary data, workhorse accent |
| Red | `#C8283C` | Headline accent syllable, warnings, sparing emphasis |
| Ink | `#17242E` | Text, nav/footer bars (constant in both themes) |
| Deep duotone | `#2E4756` | Imagery duotone, photographic treatments |
| Paper (light) | `#DCE4E9` | Page ground — cool blue-grey, never pure white |

### Theme tokens (CSS custom properties)
```css
:root{ /* light */
  --paper:#DCE4E9; --panel:#EAF0F3; --tile:#F3F7F9; --line:#B9C6CF;
  --ink:#17242E; --ink2:#3E5260; --muted:#5F7383;
  --gold:#F2A900; --blue:#1E5AA8; --red:#C8283C; --deep:#2E4756;
}
/* dark — redefine tokens only; components never restyle inside the media query */
@media (prefers-color-scheme: dark){ :root{
  --paper:#0F1920; --panel:#182631; --tile:#141F28; --line:#2C3D49;
  --ink:#E6EDF2; --ink2:#B9C7D1; --muted:#8499A8;
  --gold:#F2A900; --blue:#6FA8FF; --red:#E8556A; --deep:#9FB4C0;
}}
```
Nav bar, footer, and "cognition" blocks stay **constant** `#17242E` with gold
`#F2A900` in both themes (they are the ink bars; they do not flip).

### Chart palettes (colorblind-validated — do not substitute)
- **Light surfaces:** `#1E5AA8` (blue) · `#B87A00` (amber) · `#C8283C` (red) · `#0E9384` (teal)
- **Dark surfaces:** `#4C8DF0` · `#BA8400` · `#E8556A` · `#17A398`

Chart rules: single-series charts use one hue (no rainbow bars); **never encode
race as different hues** — put groups on the axis, one color on the marks; text and
labels always wear ink/muted tokens, never the data color; every chart ships a
"view data as table" fallback and a named source line.

## 4. Typography

| Role | Face | Treatment |
|---|---|---|
| Display | **Anton** (embed as data-URI woff2; fallbacks: Impact, Haettenschweiler, 'Franklin Gothic Bold', 'Arial Black') | ALL CAPS, line-height 0.92, huge — `clamp(3.6rem, 16vw, 11rem)` for heroes. Split-color final syllables in red. |
| Body | **Georgia** (serif) | 1.0625rem / 1.65 line-height; italic for kickers and quotes |
| Utility | **Helvetica Neue / Arial** | Labels & sources: 0.7–0.72rem, UPPERCASE, letter-spacing 0.13–0.14em, muted color |

Numbers in stat tiles: Helvetica semibold, 2rem, proportional figures.
Tabular figures (`font-variant-numeric: tabular-nums`) only in table columns.

## 5. Signature motifs & components

1. **The gold circle** — one large `#F2A900` circle bleeding off the top of every
   page header, behind the giant headline. It is the sun/spotlight/zero — the brand
   mark in its simplest form.
2. **Split-syllable headline** — one word, ink + red split (FORGOT<span red>TEN</span>).
3. **Micro-copy columns** — two small flanking text columns, poem-fragment tone,
   left-aligned / right-aligned, max ~11.5rem wide.
4. **Stat tile** — big number, one-line meaning, tiny source line. Grid of 3–4.
5. **Callouts** — left-bar panels: **red bar** = warnings & "what the data does NOT
   show" honesty blocks; **gold bar** = notes, boundaries, context.
6. **Timeline** — 4px ink spine, gold dots (red dots for federal events),
   Anton year labels.
7. **Provision cards** — panel cards with a 5px top bar rotating gold/blue/red.
8. **Buttons** — rectangular, no radius: gold (primary), blue (secondary),
   2px ink outline (ghost). UPPERCASE, letterspaced.

## 6. Layout

- Poster header first (circle + headline + italic kicker), content column after.
- Reading column 46rem; data-heavy pages 60rem.
- Space with flex/grid `gap`, not stacked margins. Wide tables scroll in their own
  container — the page never scrolls sideways.
- Sticky ink nav with gold active tab; footer = ink bar with three columns
  (identity / honesty policy / disclaimers + sister-project link).

## 7. Imagery

Blue duotone (`#2E4756` over paper) photographic/architectural imagery in the
poster spirit. **Never:** handcuffs, prison-bar clichés, children's faces, mugshots,
cartoon illustration. Institutional architecture, empty hallways, hands,
documents/case-file textures are the lane.

## 8. Accessibility (non-negotiable)

- Both themes always designed, token-level, with `data-theme` overrides beating the
  media query in both directions.
- Chart palettes only from §3 (they pass CVD-separation and contrast validation).
- Visible focus states (3px gold outline). `prefers-reduced-motion` respected —
  motion (like the BLS dot) is user-initiated only, never autoplaying.
- Every chart: table fallback + source line. Legends for 2+ series; none for one.

## 9. Hard content rules (brand = credibility)

1. No statistic without a named source in `research/SOURCES.md`.
2. Known myths are named as myths (the "70% of prisoners' kids" claim, etc.).
3. Three-strikes history told honestly: state laws first; the bill as accelerant.
4. Healing content always carries the we-are-not-therapists disclaimer and a
   licensed-care referral before any technique is described.
5. Denominators labeled; measurement frames never mixed in one chart.
6. Contributors own their stories: consent, anonymity on request, right to withdraw.

## 10. Quick-start snippet for any new page

```css
@font-face{font-family:"Anton";src:url(data:font/woff2;base64,…) format("woff2")} /* copy from website/index.html */
.disp{font-family:'Anton',Impact,'Arial Black',sans-serif;text-transform:uppercase;line-height:.92}
.label{font-family:'Helvetica Neue',Arial,sans-serif;font-size:.72rem;text-transform:uppercase;letter-spacing:.14em}
body{background:var(--paper);color:var(--ink);font-family:Georgia,serif;line-height:1.65}
```

**One-line brief for Claude or any designer:** *"Protest poster with a bibliography:
pale blue-grey paper, one huge gold circle, giant Anton all-caps headline with a red
syllable, Georgia body, sourced stat tiles, gold/blue/red only, both themes,
no pity, no unsourced numbers."*
