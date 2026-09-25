# LinkUp — Workshop MVP Plan

> The task-level roadmap lives in [`project_tracker.xlsx`](project_tracker.xlsx). Manage it with `scripts/tracker.py`.
> Items marked **⚑ Assumption** were not stated in the brief.

## 1. Product goal

LinkUp helps people at an event see who is attending, find relevant professional connections and connect on LinkedIn.
The first real use is the **Vibe Coding Workshop in Alicante**.

The goal today is **not** a production networking platform. It is to show how fast an idea can become a real, deployed product that people in the room actually use. We will build it in about 2 hours with AI-assisted development.

**Success means:** the event exists, one URL is shared with the room, people join from their phones in under 60 seconds, see and open each other's profiles, get AI suggestions for who to meet, and connect on LinkedIn. All of this runs on a public deployment.

## 2. MVP user journey

```
Event link / QR
  → /event/vibe-coding-alicante        landing: event info, "N people here", [Join event]
  → /event/…/join                      one-screen profile form (chips, minimal typing)
  → /event/…/people                    card grid, search, category chips, live updates
  → /event/…/p/[id]                    full profile, [Connect on LinkedIn]
  → /event/…/matches                   [Find people I should meet] → 1–3 AI matches
                                         with reason and conversation starter
  → /event/…/me/edit                   edit my profile
```

## 3. Architecture

| Layer | Choice | Why |
|---|---|---|
| Frontend | Next.js (App Router) + React + TypeScript, Tailwind CSS, shadcn/ui | Fast to build with AI, and the pages are mobile-first |
| Backend | Next.js **server actions**, no separate API | Least code; runs on Vercel |
| Database | Supabase Postgres with **Row Level Security** | Ownership rules are enforced in the database |
| Auth | **Supabase Anonymous Sign-ins** | No password or email, but still a real `auth.uid()` for RLS |
| Live updates | Supabase Realtime (`postgres_changes` on `participants`) | Built in; no polling fallback in the MVP |
| AI | Anthropic **Claude Haiku 4.5**, one call per request, rule-based fallback | Cheap, fast, and always returns a result |
| Hosting | Vercel (auto-deploys from `main`) | Pushing to main deploys straight to production |

The app **never uses the Supabase service-role key**. Every write goes through the user's own session and RLS. The two operations that need elevated rights are small `SECURITY DEFINER` SQL functions (see §4).

**⚑ Assumption:** the Supabase project is hosted in the EU. Identity is tied to the device: clearing cookies or switching browser means rejoining with a new profile.

## 4. Data model

```
events
  id uuid pk, slug text unique, name, description, location, starts_at timestamptz, created_at

participants
  id uuid pk, event_id fk → events, user_id uuid (= auth.uid()),
  name, role (job title), role_category (Designer|Developer|Founder|Product|Marketing|AI|Other),
  bio, skills text[], can_help_with text[], looking_for text[], ask_me_about,
  linkedin_url, instagram_url, website_url,
  last_matched_at timestamptz, created_at
  unique (event_id, user_id)
```

- The brief's `EventParticipant` table is merged into `participants`: each profile belongs to one event, since profiles across events are out of scope.
- The brief's `Match` table is dropped. Matches are generated on demand and not stored. `last_matched_at` is enough for the cooldown.
- **RLS rules:**
  - Anyone can read `events`.
  - Signed-in users, including anonymous ones, can read `participants`.
  - Users can only insert, update or delete a row where `user_id = auth.uid()`.
  - Column grants stop clients from changing `user_id`, `event_id` or `last_matched_at`.
- **Functions:**
  - `event_participant_count(event_id)` is readable by anonymous visitors, for the landing page count.
  - `claim_match_slot()` is an atomic 60-second AI cooldown.
- The Vibe Coding Alicante event is created by a **seed SQL script**. There is no organizer UI.

## 5. AI matching approach

1. **Pre-filter in code.** Score every other participant by tag overlap: my *looking for* against their *can help with*, *skills* and *category*, and the reverse. Keep the top 10.
2. **Send only the minimum data.** For each candidate, send first name, job title, category, skills, can-help-with, looking-for and ask-me-about. Candidates are keyed `c1…c10`. No URLs, bios or IDs are sent.
3. **One LLM call** to `claude-haiku-4-5` with a low `max_tokens` and a 15-second timeout. It returns JSON with up to 3 matches, each `{key, reason, starter}`.
4. **Validate the output.** Check it with Zod, drop any key not in the candidate set, then map keys back to participant IDs.
5. **Fall back** to the top 3 rule-based matches with template text on any error.
6. **Guard against abuse** with a 60-second cooldown per user (`claim_match_slot()`), a `max_tokens` cap and a monthly spend limit in the provider console.

Profile text is **untrusted**: it is wrapped in delimiters, and the prompt tells the model to treat it as data only. This addresses prompt injection (OWASP LLM01).

## 6. Development phases (~2–2.5 h)

| Phase | Tasks | Time | Exit criteria |
|---|---|---|---|
| **P1 Setup** | Scaffold · Supabase schema, RLS, seed · Design system · First deploy | ~35 min | Vercel URL live; DB and RLS in place; no secrets in git |
| **P2 Join & Profile** | Landing · Zod schema · Onboarding form · Join action · Edit my profile | ~35 min | A phone joins and edits its own profile; bad input rejected on the server |
| **P3 People** | Directory · Search and filters · Profile + LinkedIn · Realtime | ~25 min | Two devices see each other live; LinkedIn CTA works |
| **P4 AI Matching** | Match action · Abuse guard · Matches UI | ~25 min | 1–3 matches with reason and starter; cooldown enforced |
| **P5 Polish & Launch** | UI polish · Security check · Launch and QR | ~20 min | Security check Done; room is using the product URL |
| **P6 Backlog** | 12 post-workshop items | — | Does not block the workshop |

That makes **19 MVP tasks**, 8 of them security-related, plus 12 backlog tasks. If time runs short, cut **P3-T04 (Realtime)** first.

## 7. Deployment approach

- GitHub `marry-may/linkup`, branch `main`, connected to Vercel. Every push to `main` deploys to production.
- We deploy in P1 so the rest of the workshop ships continuously.
- Env vars live in Vercel and in `.env.local` locally, and are never committed. `.env.example` holds placeholders only.
- Public: `NEXT_PUBLIC_SUPABASE_URL`, `NEXT_PUBLIC_SUPABASE_ANON_KEY` (safe because of RLS).
- Server-only: `ANTHROPIC_API_KEY`.
- The launch URL is `https://<vercel-domain>/event/vibe-coding-alicante`, shared as a link and a QR code.

## 8. Security considerations

The essentials for a public MVP are built into the tasks themselves. One **pre-launch security check (P5-T02)** replaces per-phase gates for the workshop (see the Decisions sheet).

| Risk | Control | Task | OWASP |
|---|---|---|---|
| Secret leakage | `.env*` git-ignored; key only on the server; no service-role key | P1-T01, P1-T02, P4-T01 | A02, A04 |
| Vulnerable dependencies | Exact-pinned versions and lockfile; `npm audit` before launch | P1-T01, P5-T02 | A03 |
| Editing someone else's profile | RLS owner-only writes; `user_id` taken from the session only | P1-T02, P2-T04, P2-T05 | A01, A07 |
| Malicious input / XSS / bad links | Shared Zod schema checked on the server; https and host allowlist for URLs; React escaping; `rel="noopener noreferrer nofollow"` | P2-T02, P3-T03 | A05 |
| Leaking error details | Generic user-facing errors; AI fallback | P2-T04, P4-T01 | A10 |
| AI abuse and cost | Cooldown, `max_tokens`, spend limit | P4-T02 | LLM10, A06 |
| Prompt injection / PII sent to AI | Delimited data, validated output, minimal fields | P4-T01 | LLM01, LLM02 |
| Personal data in logs | Log IDs and metrics only | P4-T02 | A09 |

References are to the OWASP Top 10:2025 and the OWASP Top 10 for LLM Applications 2025.

**Deferred to backlog:**
- Members-only visibility per event
- CAPTCHA and join rate limits
- CSP and security headers
- CI and Dependabot
- Automated RLS tests
- Moderation tools
- Delete profile and data retention

## 9. Key risks

| Risk | Mitigation |
|---|---|
| 2-hour build overruns | Strict P0 order; deploy early; cut Realtime first; AI fallback |
| Spam or fake profiles through the public link | Supabase anonymous sign-in rate limits; one profile per device; delete in dashboard; CAPTCHA in backlog |
| AI cost or abuse | Cooldown, token cap, spend limit |
| Venue Wi-Fi unreliable | Pages work with a refresh; mobile data |
| In-app QR browsers lose the session | Test QR with native camera apps; rejoin is cheap |
| Personal data (GDPR, event in Spain) | Professional data only; privacy note; data deleted after the event (backlog) |

The full register is in the tracker's **Risks** sheet.

## 10. Assumptions

- **⚑** The LLM is Anthropic Claude Haiku 4.5; the brief only said "an LLM API". It's easy to swap.
- **⚑** Expect about 20–60 attendees. The UI is in English, and free tiers are enough.
- **⚑** Any signed-in user can read participants. Per-event isolation is in the backlog, and it's fine for now because there's only one event.
- **⚑** LinkedIn is optional but prominent. People without it can still join.
- **⚑** Profile photos are replaced by generated initials avatars.
- **⚑** Matches are not stored.
- **⚑** The QR code is a PNG generated once at launch, not a feature in the app.
- **⚑** Next.js 16+ renames `middleware.ts` to `proxy.ts`. Use whichever the installed version expects.
- The tracker scripts need Python 3.9+ and `openpyxl` in `.venv` (`scripts/requirements.txt`).
