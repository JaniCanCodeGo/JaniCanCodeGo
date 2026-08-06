# CETL App

The public website/app for Clear Enough To Lead: podcast pages, the Concentration focus game, legal pages, and light/dark themes. Pure static HTML/CSS/JS — no build step, no dependencies, no framework — which matches the Build With Claude curriculum (this is exactly the kind of site the workgroup learns to make).

## Run it

Open `index.html` directly in a browser, or serve it locally (the "Localhost Projects" way):

```bash
cd CETL/app
python3 -m http.server 8000
# then visit http://localhost:8000
```

## Pages

| Page | Path |
|---|---|
| Home | `index.html` |
| Podcast episode list | `podcast/index.html` |
| Episode 1–3 detail pages | `podcast/<slug>.html` (slugs match `../data/podcast_episodes.json`) |
| Concentration game | `game/index.html` |
| Terms of Service | `legal/terms.html` |
| Privacy Policy | `legal/privacy.html` |
| Design mockups (both themes) | `design/mockups.html` |

Episode pages are hand-built from `../data/podcast_episodes.json` (the source of truth per `site_manifest.json`). When an episode's `status` changes to `published`, replace its "Audio coming soon" placeholder with a real `<audio>` element and add new episode pages following the existing ones.

## Themes

`assets/theme.js` runs synchronously in `<head>`: it applies the visitor's saved choice from `localStorage`, falling back to the OS `prefers-color-scheme`. Every page has a toggle button in the nav. All colors live as CSS custom properties at the top of `assets/styles.css` (`:root` = light, `:root[data-theme="dark"]` = dark).

## Security posture

- **No third-party code**: no CDNs, no analytics, no trackers, no external fonts. Every request is same-origin.
- **CSP on every page**: `default-src 'self'` with `object-src 'none'`, `base-uri 'none'`, `form-action 'none'`. No inline scripts or styles anywhere, so the CSP needs no `unsafe-inline`.
- **No data collection**: no forms, no cookies. The only client-side state is the theme choice and game best-score in `localStorage`, which never leave the device.
- **No dynamic HTML from data**: game cards are built with `createElement`/`textContent` (never `innerHTML`), so there is no XSS sink even if the symbol set changes.
- When deploying, also send the CSP as an HTTP header and add `frame-ancestors 'none'` to it (that directive only works as a header, not in `<meta>`), plus `X-Content-Type-Options: nosniff` and `Referrer-Policy: no-referrer`. The `<meta>` tags are a baseline; headers are stronger.

## Legal

`legal/terms.html` and `legal/privacy.html` are solid CCPA-aware starting points but are templates — have an attorney review them before launch (each page says so at the bottom).
