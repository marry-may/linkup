# LinkUp

Event networking web app: join an event via link, create a quick profile, see who's here, get AI suggestions for who to meet, connect on LinkedIn.
Workshop MVP built in ~2 hours at the Vibe Coding Workshop, Alicante.

- Plan: `docs/PLAN.md`
- Tracker: `docs/project_tracker.xlsx` (managed only via `scripts/tracker.py`)
- Stack: Next.js (App Router, TS) · Tailwind · shadcn/ui · Supabase (Postgres, anonymous auth, RLS, Realtime) · Claude Haiku 4.5 · Vercel
- Repo: https://github.com/marry-may/linkup (branch `main`, auto-deploys to Vercel)

## Tracker commands

Tracker scripts run in the project venv (`python` only exists inside it on this machine):

```bash
python3 -m venv .venv && .venv/bin/pip install -r scripts/requirements.txt   # once
source .venv/bin/activate
python scripts/tracker.py next                       # next ready task
python scripts/tracker.py show P1-T02                # full task details
python scripts/tracker.py list --status "Not Started" [--phase P2]
python scripts/tracker.py update P1-T03 --status "In Progress"
python scripts/tracker.py update P1-T03 --status Done --note "tests passing"
python scripts/tracker.py decision "Use X" --options "X, Y" --rationale "..."
python scripts/build_tracker.py                      # regenerate workbook (keeps progress)
```

## Session workflow rules

1. At the beginning of every development session, run `python scripts/tracker.py next`.
2. Confirm the selected task with the user before implementation.
3. Work on one task at a time. Set it to "In Progress" when starting.
4. Do not start tasks whose dependencies are incomplete (`tracker.py` refuses without `--force`).
5. Mark tasks Done only after every acceptance criterion is met and tests/build pass.
6. Update the tracker through `scripts/tracker.py`, never by editing the xlsx manually.
7. Record architecture changes with `tracker.py decision` (Decisions sheet), and update `docs/PLAN.md` if the plan changes.
8. Never commit secrets.
9. Check `git status` before every commit.
10. Do not deploy the event link publicly until the pre-launch security check (P5-T02) is Done. Security acceptance criteria inside each task are not optional.

## Git workflow rules

- Default branch is `main`; remote `origin` = https://github.com/marry-may/linkup.git.
- After each meaningful, completed and tested feature (normally one tracker task), create a descriptive commit and push to `main`.
  - Message style: `P2-T04: Join flow with anonymous session and participant creation`.
  - Include the tracker update (`docs/project_tracker.xlsx`) in the same commit.
- Always run `git status` (and review `git diff --staged`) before committing.
- Never commit broken code: `npm run build` (and `npm test` once tests exist) must pass before a commit.
- Never commit secrets: `.env`, `.env.local` and any `.env.*` except `.env.example` stay ignored. If a secret is ever committed, stop, rotate the key, and tell the user.
- Stage files explicitly or review `git status` output first; don't blindly `git add -A` without checking for unexpected files.
- Pushing to `main` deploys to production on Vercel — only push working code.

## Security essentials (MVP)

- Supabase service-role key is not used by the app. All writes go through the user's session + RLS.
- `ANTHROPIC_API_KEY` is server-only: import the SDK only in files with `import 'server-only'`; never prefix with `NEXT_PUBLIC_`.
- `user_id` always comes from `supabase.auth.getUser()` on the server, never from form data.
- Validate all input server-side with the shared Zod schema in `src/lib/validation/profile.ts`.
- External links: `rel="noopener noreferrer nofollow"`; never use `dangerouslySetInnerHTML`.
- Send only minimal professional fields to the LLM; treat profile text as untrusted data.
- Return generic user-facing errors; never log profile contents.
- Load the `claude-api` skill before writing Anthropic SDK code.
