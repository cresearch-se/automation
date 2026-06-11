# Git Commit Skill

Help the user commit and push their changes to the right branch.

## Step 1 — Gather context

Run these commands to understand the current state:
- `git status` — what files changed
- `git branch` — current branch name
- `git log --oneline -5` — last 5 commits on this branch
- `git status -sb` — check if branch is ahead/behind origin

## Step 2 — Show the user a summary

Print a clear summary:
```
Branch:   <current branch>
Status:   <ahead by N commits / in sync with origin / no upstream set>
Last commit: <last commit message>

Changed files:
  Modified:  <list>
  New files: <list>
  (Excluded: output/, fixtures/ — generated files not staged)
```

## Step 3 — Decide what to do

Apply this logic and ask the user ONE clear question:

**If on `main` or `master`:**
→ Never commit directly. Tell the user and ask: "What should the new branch be called?"

**If on a feature branch AND branch is in sync with origin (all changes pushed):**
→ Default suggestion is a NEW branch (likely new work).
→ Ask: "You're on `<branch>` which is fully pushed. Create a new branch or continue on this one?"

**If on a feature branch AND branch has unpushed commits (ahead of origin):**
→ Default suggestion is SAME branch (likely continuing work).
→ Ask: "You're on `<branch>` with unpushed commits. Continue on this branch or create a new one?"

**If on a feature branch with no upstream set yet:**
→ Continue on current branch (not yet pushed at all).
→ Ask: "You're on `<branch>` (not yet pushed). Commit here or create a new branch?"

## Step 4 — Stage files

Stage only relevant files — NEVER stage:
- `tests/TeamworkDB/output/` — generated test output files
- `tests/TeamworkDB/fixtures/*.xlsx` — Excel fixture files  
- `*.html`, `mismatches_utilization.txt` — generated reports
- Any file the user says to skip

Show the user exactly what will be staged and confirm before proceeding.

## Step 5 — Commit

Write a clear, concise commit message based on what changed. Format:
- One line summary (what and why, not just "update files")
- If the user gives you a message, use that exactly

## Step 6 — Push

Ask: "Push to origin now?"
- If yes → `git push -u origin <branch>`
- If no → stop here, leave it as a local commit

## Rules
- Never use `git add .` or `git add -A` — always add files explicitly by name
- Never force push
- Never commit to `main` or `master` directly
- Always show the user what will be staged before doing it
