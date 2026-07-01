# TestSheet SaaS — MVP Product Spec

**Version:** 1.0  
**Target launch:** 6–8 weeks from start of build  
**First buyer:** Finance / FP&A analyst (non-technical, credit card purchase, team of 2–10)

---

## The one-sentence pitch

> Upload two versions of your Excel model and instantly see every cell that changed — shared with your whole team in seconds.

---

## Problem being solved

Finance teams version-control Excel models by saving files as `budget_v3_FINAL_JIM.xlsx`. When something looks wrong in a presentation, they spend hours manually hunting for which cell changed and why. There is no automated way to diff two Excel workbooks and get a clean, shareable report.

---

## What the MVP does (and only this)

1. A team creates an account and a shared workspace
2. Any team member uploads an Excel workbook → TestSheet captures it as the baseline
3. When a new version of the model is ready, they upload it → TestSheet diffs it cell-by-cell
4. They see a clean drift report in the browser: which cells changed, from what value to what
5. They share a link to the report with anyone (view-only, no login required)
6. History: the last 10 runs per workbook are stored and browsable

---

## What the MVP does NOT do

These are explicitly cut from v1. They go on the v2 backlog.

| Feature | Reason cut |
|---|---|
| Invariant rules (range_bound, no_error, etc.) | Adds complexity; validate core drift value first |
| Excel add-in | Requires Microsoft Store approval; build after web app is proven |
| CI/CD webhook integration | Developer feature; primary buyer is non-technical |
| Email / Slack notifications on drift | Post-MVP engagement feature |
| Formula-level diff | MVP shows value drift; formula diff is a power-user feature |
| Mobile app | Desktop-first; FP&A analysts work at desks |
| SSO / SAML / enterprise auth | Enterprise tier; comes after team tier is validated |
| API access | Developer feature; post-MVP |
| Version branching / tagging | Git-like features; validate simpler history first |

---

## User stories (MVP)

### Auth & team

- **US-1** As a user I can sign up with Google or email/password
- **US-2** As a user I can create a team workspace with a name
- **US-3** As a team owner I can invite members by email
- **US-4** As an invited member I can accept an invite and join the workspace
- **US-5** As a team owner I can remove a member

### Workbooks (projects)

- **US-6** As a team member I can create a "workbook project" with a name and optional description
- **US-7** As a team member I can upload a .xlsx file to capture the baseline for a project
- **US-8** As a team member I can see all workbook projects in my workspace
- **US-9** As a team member I can delete a workbook project

### Running checks

- **US-10** As a team member I can upload a new version of a workbook and trigger a diff run
- **US-11** As a team member I can see the drift report for any run: cells that changed, their sheet, address, old value, new value, and drift kind
- **US-12** As a team member I can see a summary: total cells checked, cells drifted, pass/fail status
- **US-13** As a team member I can see the last 10 runs for a project and click into any of them
- **US-14** As a team member I can download the drift report as HTML

### Sharing

- **US-15** As a team member I can generate a shareable link to a specific run report
- **US-16** As anyone with a link I can view a read-only drift report without logging in

### Billing

- **US-17** As a team owner I can enter a credit card and subscribe to a paid plan
- **US-18** As a team owner I can see my current plan and usage
- **US-19** As a team owner I can cancel my subscription

---

## Screens

### Public / auth
1. **Landing page** — headline, demo GIF, pricing table, sign-up CTA
2. **Sign up** — email/password + Google OAuth
3. **Log in** — email/password + Google OAuth
4. **Accept invite** — team name shown, one-click join

### App (authenticated)
5. **Dashboard** — list of workbook projects in the workspace; last run status badge per project
6. **New project** — name + description + first .xlsx upload (captures baseline)
7. **Project detail** — list of runs (date, pass/fail, cells drifted); Upload new version button
8. **Run report** — summary bar + filterable drift table (filter by sheet, drift kind); Download HTML; Share link button
9. **Team settings** — member list; invite by email; remove member
10. **Billing** — current plan; upgrade/downgrade; Stripe customer portal link

### Shared (no login)
11. **Shared run report** — read-only version of screen 8, no nav, branding watermark on free tier

---

## Data model

```
Workspace
  id, name, created_at
  owner_id → User

WorkspaceMember
  workspace_id, user_id, role (owner | member), joined_at

WorkbookProject
  id, workspace_id, name, description, created_at
  baseline_run_id → Run (nullable, set after first upload)

Run
  id, project_id, created_at, status (pending | processing | complete | error)
  workbook_filename, workbook_storage_key
  cell_count, drift_count, passed (bool)
  share_token (uuid, nullable — set when shared)

DriftCell
  id, run_id, sheet, address, kind, baseline_value, current_value
  (index on run_id for fast lookup)

Subscription
  workspace_id, stripe_customer_id, stripe_subscription_id
  plan (free | pro | team), status, current_period_end
```

---

## Tech stack

| Layer | Choice | Reason |
|---|---|---|
| Frontend | Next.js 14 (App Router) + Tailwind | Fast to build, great SaaS component ecosystem |
| UI components | shadcn/ui | Pre-built, accessible, easy to customise |
| Backend API | FastAPI (Python) | Reuses all TestSheet engine code directly |
| Database | PostgreSQL via Supabase | Managed, free tier, built-in auth option |
| File storage | Supabase Storage (S3-compatible) | Workbook uploads; auto-expires old files |
| Auth | Supabase Auth | Email/password + Google OAuth; JWT; team invites |
| Payments | Stripe | Industry standard; Stripe Billing handles subscriptions |
| Background jobs | FastAPI + asyncio (simple queue) | Diff runs are fast enough (<30s); no Celery needed for MVP |
| Frontend hosting | Vercel | Zero-config Next.js deploy |
| Backend hosting | Railway | Simple Python deploy; scales easily |
| Email | Resend | Team invites + transactional email |

---

## Pricing

| Plan | Price | Limits |
|---|---|---|
| **Free** | $0 | 1 workbook project, 5 runs/month, 7-day run history, "Powered by TestSheet" on shared reports |
| **Pro** | $29/mo per seat | Unlimited projects, unlimited runs, 90-day history, no branding on shared reports |
| **Team** | $79/mo (up to 5 seats) | Everything in Pro + team workspace, member management, shared history |

Start with annual discount (2 months free) to improve cash flow.

---

## Build sequence (6–8 weeks)

### Week 1–2: Backend API + engine integration
- FastAPI app scaffold with Supabase connection
- File upload endpoint → triggers diff job
- Diff job: wraps existing TestSheet Python engine
- REST endpoints: projects CRUD, runs CRUD, drift cells read
- Auth middleware (Supabase JWT verification)

### Week 3–4: Frontend core
- Next.js scaffold + Supabase Auth (login, signup, Google OAuth)
- Dashboard, project detail, run report screens
- Drift table component (filterable by sheet + kind)
- File upload UI (drag-and-drop)

### Week 5: Teams + sharing
- Invite flow (email via Resend)
- Team settings screen
- Share token generation + public run report screen

### Week 6: Billing + polish
- Stripe integration (checkout, customer portal, webhooks)
- Free tier enforcement (run limits)
- Landing page
- Error states, loading states, empty states

### Week 7–8: Beta + fixes
- 5 beta users from target market (finance/FP&A contacts)
- Fix top issues from beta feedback
- Set up monitoring (Sentry, Vercel Analytics)
- Launch on Product Hunt

---

## Success metrics for MVP

| Metric | Target at 8 weeks |
|---|---|
| Beta users | 10+ using the product weekly |
| Paying customers | 3+ paying before public launch |
| Weekly active usage | >2 runs/week per active workspace |
| NPS | >30 from beta users |

---

## Biggest risks

**Risk 1: File size.** Large FP&A models can be 10–50 MB with thousands of formula cells. The `formulas` evaluator may be slow (>30s). Mitigation: add a processing timeout, show a progress indicator, and test with real large models in week 1.

**Risk 2: Formula coverage gaps.** The `formulas` library doesn't cover 100% of Excel functions. Some customer models may error. Mitigation: catch evaluator errors gracefully, show which cells couldn't be computed, fall back to cached values.

**Risk 3: Non-technical buyer friction.** FP&A analysts are comfortable with Excel but not with "uploading workbooks to a web app." Mitigation: extremely simple onboarding (3 clicks to first report), good empty states, and a demo mode with a pre-loaded sample workbook.

**Risk 4: Pricing too high to validate fast.** Mitigation: offer a 14-day free trial of the Team plan, no credit card required, to get teams in the door.
