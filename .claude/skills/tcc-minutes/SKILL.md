---
name: tcc-minutes
description: Generate formal Trustee Committee (TCC)-format meeting minutes as a Word (.docx) document from a Granola meeting. Use when the user wants to turn a Granola meeting's transcript/notes into minutes, draft TCC minutes, or produce board/committee minutes in the Straight Bat house style. Drives the full routine — connect Granola, select a meeting, pull the transcript, map it to the TCC template, and render the .docx.
---

# TCC Meeting Minutes

Produce a polished, formal set of minutes in the **Trustee Committee (TCC) house style**
from a Granola meeting, delivered as a Word `.docx`. The reference template is
`TCC_Minutes_30Mar2026.docx` (SharePoint, ComplianceTrustee site).

## Workflow

### 1. Identify the Granola account & meeting
- Confirm the connected account with `mcp__Granola__get_account_info`.
- If the user named a meeting, find it with `mcp__Granola__list_meetings` (widen
  `time_range` to `last_30_days` or `custom` as needed).
- If they did **not** name one, list recent meetings and ask which to use (prefer
  `AskUserQuestion`). Surface the meetings whose attendees match the relevant committee.

### 2. Pull the source material
- `mcp__Granola__get_meetings` → AI summary, attendees, metadata.
- `mcp__Granola__get_meeting_transcript` → verbatim transcript (may be large; read the
  saved tool-result file fully before drafting).
- Base the minutes on the **transcript** for substance; use the summary to cross-check
  structure and action items.

### 3. Draft the minutes content
Write a JSON spec (see `spec.schema.md`) that maps the discussion into the TCC structure:
- **Title block**: title, subtitle (`Entity | DD Month YYYY`), subheading (meeting name),
  status (`CONFIDENTIAL — DRAFT` until signed off).
- **Meeting Details** table: Entity, Meeting Type, Date, Time, Location, Chair,
  Minutes Prepared By, Date of Confirmation.
- **Attendees** table: Name, Title/Role, Status (Present / Apologies / via Teams).
- **Numbered sections** (`1. Opening & Quorum`, …, final `Close & Next Meeting`). Each has
  a presenter/time-allocated line and blocks: `sub` (sub-heading), `para`, `bullet`,
  `resolution` (rendered bold, prefixed `RESOLUTION:`).
- **Action Register**: Ref (A1, A2…), Action Item, Owner, Due By — one row per commitment.
- **Confirmation of Minutes**: true-and-accurate statement + signatory blocks.
- **Footnote**: note the source (Granola transcript), and that figures/names should be
  verified and `[bracketed]` items completed before sign-off.

Drafting rules:
- Formal, third-person, past tense ("The Committee discussed…", "It was agreed…").
- Convert dialogue into decisions, discussion points, and actions — never transcribe chatter.
- Capture every concrete commitment in the Action Register with a named owner.
- Preserve figures, dates, clause numbers and names exactly; flag anything uncertain with
  `[square brackets]` rather than inventing it.
- Infer Chair/minute-taker from the transcript; if genuinely unclear, ask.

### 4. Render the .docx
```bash
pip install python-docx --quiet   # first run only
python3 .claude/skills/tcc-minutes/render_minutes.py \
    --input /tmp/minutes_spec.json \
    --output "YY'MM'DD <Meeting Name> - Minutes (DRAFT).docx"
```
File-naming convention: `YY'MM'DD <Meeting Name> - Minutes (DRAFT).docx`.

### 5. Deliver
- Send the `.docx` to the user (SendUserFile).
- Note it is a DRAFT with `[bracketed]` placeholders, and that Microsoft 365 access is
  read-only here — the user uploads to SharePoint themselves (offer to adjust first).

## Compliance note
These meetings can touch regulated/governance content. Apply the org AI-use policy: do not
echo investor bank details, TFNs, employee personal data, SRNs, or AFSL-regulated/trading
data into the minutes — summarise around them and flag for manual completion.
