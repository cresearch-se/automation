# Create Branch Skill

Help the user create a new feature branch off the latest `main`.

## Step 1 — Ask for the branch name

Ask the user: "What should the new branch be called?"

Wait for their answer before continuing.

## Step 2 — Check for uncommitted work

Run `git status` before switching branches. If there is anything uncommitted:
- Tell the user what's there
- Ask whether to stash it or leave it — do not proceed until they decide
- If stashing: `git stash -u`

## Step 3 — Create the branch off latest main

Run this exact sequence — no shortcuts:

```bash
git fetch origin               # pull remote changes without switching branches
git checkout main              # switch to local main
git pull                       # bring local main up to date with origin/main
git checkout -b <branch-name>  # create new branch from fresh main
```

## Step 4 — Confirm and report back

After the branch is created, print a short summary:

```
Branch created: <branch-name>
Based on:       main @ <short commit hash> — <commit message>
Status:         local only (not yet pushed to origin)
```

## Rules

- Always fetch + pull main before branching — never create off a stale main
- Never push the new branch unless the user explicitly asks
- Never create a branch off the current feature branch — always go back to main first
- If the user gives a branch name with spaces, replace them with hyphens and tell them
