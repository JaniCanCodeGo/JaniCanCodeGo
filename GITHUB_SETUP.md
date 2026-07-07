# Connecting GitHub locally (Windows / PowerShell)

This guide gets GitHub working on your **local Windows machine** so you can
`clone`, `pull`, and `push` this repo (`JaniCanCodeGo/JaniCanCodeGo`) from
PowerShell. The heavy lifting is done by the included **`Setup-GitHub.ps1`**
script.

## Quick start

1. Open **PowerShell** (regular window — no admin needed).
2. `cd` to wherever you keep this repo, or to any folder if you want the
   script to clone it for you.
3. Allow the script to run for this session, then run it:

   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   .\Setup-GitHub.ps1
   ```

4. A browser window opens — **sign in to GitHub and approve access**. That's
   the "grant permissions" step. When it closes, you're connected.

Optionally pass your commit identity up front so it doesn't prompt:

```powershell
.\Setup-GitHub.ps1 -GitName "Your Name" -GitEmail "you@example.com"
```

## What the script does

| Step | Action |
|------|--------|
| 1 | Installs **Git** if missing (via `winget`) |
| 2 | Installs the **GitHub CLI** (`gh`) if missing |
| 3 | Sets your `git` name/email for commits |
| 4 | Runs `gh auth login` — opens the browser to **grant permissions** |
| 5 | Runs `gh auth setup-git` so `git` reuses your login (no tokens to manage) |
| 6 | Clones the repo, or verifies access if you're already inside it |

After it finishes, confirm it works:

```powershell
git fetch
git status
```

## Why the browser step?

Modern GitHub auth uses OAuth through the browser instead of a password.
Approving in the browser is what grants your machine permission to talk to
GitHub. The token is stored securely by the Windows credential manager, so
you only do this once.

## Troubleshooting

- **`gh` or `git` "not recognized" right after install** — close PowerShell,
  reopen it, and re-run the script. `PATH` only refreshes in new windows.
- **`winget` not found** — you're on an older Windows build. Install Git from
  <https://git-scm.com/download/win> and the GitHub CLI from
  <https://cli.github.com/>, then re-run the script.
- **"running scripts is disabled"** — you skipped the `Set-ExecutionPolicy`
  line above. Run it, then re-run the script.
- **Re-check your login anytime**: `gh auth status`
