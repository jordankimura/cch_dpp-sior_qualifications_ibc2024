# Getting this onto GitHub

This folder is already a git repository with one commit. You do not need to run
`git init`.

## 1. Create an EMPTY private repo on GitHub

github.com → New repository

- Name: `si-qualification-2024` (or whatever you like)
- **Private** — see `LICENSING.md` for why this matters
- Do **not** tick "Add a README", ".gitignore" or a licence. The repo must be empty
  or the first push will conflict.

## 2. Push

From inside this folder:

```bash
git remote add origin https://github.com/YOUR-USERNAME/si-qualification-2024.git
git push -u origin main
```

If you have the GitHub CLI, this does both steps at once:

```bash
gh repo create si-qualification-2024 --private --source=. --push
```

## 3. Set your identity if git complains

The commit was authored with your name and email already. If git objects on a later
commit:

```bash
git config user.name  "Jordan Kimura"
git config user.email "jordank2@hawaii.edu"
```

## 4. Watch the checks run

Pushing triggers `.github/workflows/verify.yml`, which installs Python, runs the
harness, and uploads the report as a build artifact. The Actions tab will show a green
tick or a red cross within about a minute.

Then update the badge at the top of `README.md`, replacing `OWNER/REPO` with your
actual path. That badge is the point of the whole exercise: it says the checks pass on
the current commit, and anyone can click it to see the run.

## Working in it after that

```bash
git add -A
git commit -m "what changed and why"
git push
```

Every push re-runs the 23 checks. If a change breaks one, the cross appears before the
file reaches anybody.

## When DPP answers the Decisions tab

Do it as one commit per decision, with the decision number in the message —
`Decision 1: renumber ROH 1705.19 and 1705.20`. Then `git log` becomes the record of
who decided what and when, which is the thing a department usually cannot reconstruct
two code cycles later.
