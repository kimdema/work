#!/usr/bin/env python3
"""Build a Word (.docx) version of the TCC-format minutes."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

NAVY = RGBColor(0x1F, 0x38, 0x64)
GREY = RGBColor(0x60, 0x60, 0x60)

doc = Document()

# Base style
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10.5)

def heading(text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = NAVY
    return h

def para(text=None, bold=False, italic=False, size=None, color=None, align=None, space_after=6):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    if text:
        r = p.add_run(text)
        r.bold = bold
        r.italic = italic
        if size:
            r.font.size = Pt(size)
        if color:
            r.font.color.rgb = color
    return p

def bullet(text):
    p = doc.add_paragraph(style='List Bullet')
    p.add_run(text)
    p.paragraph_format.space_after = Pt(3)
    return p

def kv_table(rows):
    t = doc.add_table(rows=0, cols=2)
    t.style = 'Light Grid Accent 1'
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    for k, v in rows:
        cells = t.add_row().cells
        cells[0].width = Inches(1.8)
        cells[1].width = Inches(4.7)
        rk = cells[0].paragraphs[0].add_run(k)
        rk.bold = True
        cells[1].paragraphs[0].add_run(v)
    return t

def grid_table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Light Grid Accent 1'
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        run = hdr[i].paragraphs[0].add_run(h)
        run.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].paragraphs[0].add_run(val)
            if widths:
                cells[i].width = widths[i]
    return t

# ---- Title block ----
para('TRUSTEE COMMITTEE MEETING — MINUTES OF MEETING', bold=True, size=15,
     color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
para('Straight Bat Private Equity Pty Ltd | 9 June 2026', bold=True,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
para('Governance Review — Implementation Kick-off', italic=True,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
para('CONFIDENTIAL — DRAFT', bold=True, color=RGBColor(0xB0, 0x00, 0x00),
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

# ---- Meeting details ----
heading('Meeting Details', 2)
kv_table([
    ('Entity', 'Straight Bat Private Equity Pty Ltd'),
    ('Meeting Type', 'Trustee Committee Meeting — Governance Review Implementation Kick-off'),
    ('Date', '9 June 2026'),
    ('Time', '11:00 AM – 12:00 PM'),
    ('Location', 'Matthews Meeting Room, Gate 8, 88 Jolimont St, East Melbourne VIC 3002 & TEAMS'),
    ('Chair', 'Stephen Gledden (Responsible Manager)'),
    ('Minutes Prepared By', 'Kim Dema (Company Secretary)'),
    ('Date of Confirmation', '[To be completed upon sign-off]'),
])

# ---- Attendees ----
heading('Attendees', 2)
grid_table(
    ['Name', 'Title / Role', 'Status'],
    [
        ['Stephen Gledden', 'Responsible Manager (RM) (Chair for this session)', 'Present'],
        ['Rob Nicholls', 'Responsible Manager (RM)', 'Present'],
        ['Kim Dema', 'Financial Responsible Manager (FRP) / Company Secretary / Fund CFO', 'Present'],
        ['Becky Barr', 'Compliance Manager / Fund Operations Manager', 'Present (via Teams)'],
    ],
    widths=[Inches(1.5), Inches(3.8), Inches(1.2)],
)

# ---- Sections ----
def section(num_title, presenter):
    heading(num_title, 1)
    para(presenter, italic=True, color=GREY, space_after=6)

def sub(title):
    para(title, bold=True, space_after=3)

section('1. Opening & Purpose', 'Chair: Stephen Gledden | Time allocated: 5 minutes')
sub('1.1 Purpose of Session')
para('The Chair opened the meeting and set out the purpose: to walk through the Governance '
     'Framework document section by section, agree the approach to each, and confirm '
     'accountabilities for the implementation work arising from the governance review. It was '
     'noted that the framework comprises six themes — independence, conflict management, '
     'documentation standardisation, trustee structure clarification, regulatory alignment, and '
     'governance culture.')
sub('1.2 Conflicts of Interest')
para('Consistent with the practice to be formalised (see section 6), the standing conflicts of '
     'the senior individuals were acknowledged as a recurring agenda item. No new transactional '
     'conflicts were declared in respect of the matters discussed.')

section('2. Governance Framework — Entity Structure & Delegation Model',
        'Presenter: Stephen Gledden | Time allocated: 15 minutes')
sub('2.1 Core Principle — Delegation of Authority')
para('The Committee discussed and endorsed the foundational principle of the framework: the '
     'board of Straight Bat Private Equity Holdings Pty Ltd holds ultimate decision-making '
     'authority for the funds management business and delegates that authority, for two distinct '
     'purposes, to:')
bullet('The Investment Manager (Straight Bat Private Equity Pty Ltd) — operates under delegation '
       'to run the funds and reports to the board on business performance and to the Trustee on '
       'fund performance (NAV, unit price, distributions); and')
bullet('The Corporate Trustee — a protected corporate entity that sets policy, acts strictly in '
       "investors' interests, and delegates execution to the Investment Manager under an "
       'Investment Management Agreement (IMA), while providing independent oversight including '
       'periodic third-party review.')
para('The Committee agreed that applying this distinction consistently across all governance '
     'documents is expected to resolve a significant number of the inconsistencies identified in '
     'the review.')
sub('2.2 Standardisation of Language')
para('It was agreed that documents should standardise terminology: "Board" refers to Holdings; '
     '"Trustee" refers to the corporate trustee; "Investment Manager" refers to Straight Bat '
     'Private Equity Pty Ltd. The Committee noted that this precision will reassure stakeholders '
     '(including Farrel, the broader team, and prospective institutional investors) who are less '
     'comfortable with functions being "blended into one pot".')
para('RESOLUTION: The Committee endorsed the entity/delegation model and the standardised '
     'terminology as the basis for all governance documentation going forward.', bold=True)

section('3. Independent Director Appointment — Trustee Board',
        'Presenter: Stephen Gledden | Time allocated: 15 minutes')
sub('3.1 Candidate Discussion')
para('The Committee discussed the appointment of an independent Non-Executive Director (NED) to '
     'the Trustee board. Meetings are scheduled this week (Thursday) with Tony McVeigh (2:00 PM) '
     'and Randall (1:00 PM / 9:00 AM). The Chair indicated Tony McVeigh is his preferred candidate '
     '— a recently retired Managing Partner of Hall & Wilcox (retiring 30 June), with 20+ years’ '
     'experience, recommended by Andrew O’Brien and known to Farrel.')
sub('3.2 Role Expectations')
para('Following discussion (Kim Dema and Becky Barr contributing), the Committee clarified that '
     'the role is intended to be more active than a typical NED — scheduled quarterly board '
     'meetings plus provision for ad-hoc virtual meetings to sign off proper instructions '
     '(e.g. shareholder loans for portfolio companies). Key attributes sought:')
bullet("Independence from the funds, making independent decisions in investors' interests;")
bullet('A legal background (essential for ASIC / Corporations Act compliance), given Rob Nicholls '
       'and Stephen Gledden are not deeply technical in this area;')
bullet("Knowledge of the Investment Manager's operations and the portfolio companies (attends "
       'events, reads scorecards, monitors businesses);')
bullet('Collaborative and constructive in challenge.')
sub('3.3 Independence Consideration')
para('A potential concern was noted that Tony McVeigh is from the same firm (Hall & Wilcox) as '
     "the Committee's adviser, Vanessa, which may limit independent challenge of her advice. The "
     "Committee confirmed comfort in Vanessa's competence as adviser, while agreeing to obtain a "
     'point of comparison (Randall) and to keep the NED and adviser roles distinct — the NED is '
     'not an adviser substitute. Two candidates previously suggested were ruled out: "Annette" '
     '(conflicted — EY engaged as tax adviser) and "David".')
sub('3.4 Board Process — Pre-Meeting Checklist')
para('The Committee agreed (Stephen Gledden / Kim Dema) that a standardised checklist must '
     'precede any Trustee board sign-off — proper instructions, related-party check, confirmation '
     'that supporting documents (e.g. loan documentation) are correctly executed — and that '
     'nothing is escalated to the NED / no board meeting is called until the checklist is '
     'complete. This checklist is to be written into the relevant charter.')
para('RESOLUTION: Rob Nicholls to lead the search for an independent NED to the Trustee board. '
     'A standardised pre-meeting checklist for Trustee board sign-offs to be developed and '
     'incorporated into the charter.', bold=True)

section('4. Roles & Responsibilities — Segregation of Duties',
        'Presenter: Stephen Gledden / Kim Dema | Time allocated: 15 minutes')
sub('4.1 Moving Away from "15 Hats"')
para('The Committee agreed to move from an undefined "15 hats" approach to precise role '
     'definitions, eliminating ambiguity ("grey") wherever possible by clarifying which '
     'activities relate to which entity.')
sub('4.2 Becky Barr — Roles')
bullet('Fund Operations Manager (for the Investment Manager) — administering Atomic, registry, '
       'distributions, KYC/AML; with inherent compliance requirements embedded in fund '
       'operations; and')
bullet('Compliance Officer — managing the Trustee Compliance Committee and undertaking the work '
       'required to maintain the AFSL (reporting, training, etc.), under delegation from the '
       'Trustee.')
para('It was acknowledged that compliance within fund operations is an inherent function (Elwood, '
     'as an operator, must operate the fund compliantly), and is distinct from the separate '
     'Trustee compliance function.')
sub('4.3 Kim Dema — Roles')
bullet('Company Secretary for board meetings (preparing agenda and papers);')
bullet('Compliance for the Investment Manager and Holding Company (risk management / governance, '
       'ASIC / Corporations Act matters for the trustee company not related to the AFSL);')
bullet('Cross-training with Becky Barr and developing the broader team (Elwood) for redundancy.')
sub('4.4 Key Person Risk vs. Segregation')
para('Kim Dema challenged a strict "complete segregation on paper" position on the basis that it '
     'could create key-person risk, noting the deliberate effort to share knowledge (Elwood, '
     'Mark, Al) to protect the company. The Committee reconciled this by confirming the '
     'operating-model distinction between accountability (a named individual) and responsibility '
     '(the team). Documentation should reflect actual practice as closely as possible: Becky Barr '
     'is accountable as Compliance Officer, but the compliance team includes Kim Dema and Elwood, '
     "who periodically review each other's work with a critical eye.")
sub('4.5 Trustee Compliance Committee vs. Board Meeting')
para('The Committee agreed it will not move to a separate audit & risk committee structure at '
     'this stage. Instead it will be more explicit that the Trustee Compliance Committee is the '
     'internal operational meeting (where Investment Manager personnel — principally Kim Dema and '
     'Becky Barr — perform administrative work for the Trustee under delegation), distinct from '
     'the board meeting (where the board and Trustee directors confirm work has been properly '
     'done). Accountability split agreed: Kim Dema prepares the agenda and papers for board '
     'meetings; Becky Barr runs the agenda and checks for the Trustee Compliance Committee.')
para('RESOLUTION: The Committee adopted the two-hat model for Becky Barr (Fund Operations '
     'Manager; Compliance Officer) and the multi-hat clarification for Kim Dema, with job '
     'descriptions and titles to make accountabilities explicit while preserving team-based '
     'redundancy.', bold=True)

section('5. Valuation Committee Independence & External Valuation / Audit Costs',
        'Presenter: Stephen Gledden | Time allocated: 10 minutes')
sub('5.1 Independent Chair — Danny van Aswegen')
para('The Committee discussed the proposed appointment of Danny van Aswegen (former PwC valuation '
     'partner, now of Marketline; recommended by Farrel) as independent chair of the Valuation '
     'Committee. He is regarded as aligned with the Committee’s "golden thread" approach to '
     'valuation (≈60% consistency of investment storyline, 40% mechanics) and is familiar '
     'with comparable portfolios. His revised proposal: $20k for a review, $50k for a portfolio '
     'valuation review, and an as-yet unquoted fee for the independent role (estimated ≈$100k '
     'p.a. to do well).')
sub('5.2 Cost Analysis & Potential Duplication')
para('Concerns were raised regarding total governance cost and possible duplication:')
bullet('Current Findex spend ≈ $15k per quarter, increasing (≈$25k budgeted) as scope expands;')
bullet('Proposed EY engagement: ≈$200k audit plus $60–80k for financial statements (noting '
       'Straight Bat is a public trading trust already preparing financial statements; conversion '
       'to AASB standards estimated at ~3 months);')
bullet('Combined governance costs potentially approaching ~$500k annually;')
bullet('Potential double-up between Danny van Aswegen and Findex — to be resolved. If EY prepare '
       'the audit and perform the valuations, the Findex opinion may be unnecessary (potential '
       '~$100k saving).')
sub('5.3 Items to Clarify')
bullet('Whether Danny van Aswegen is a registered valuer (ATO-accepted) — noted as a key question '
       '(he is CFA/CA qualified; registration to be confirmed, possibly held at company level);')
bullet('Scope differentiation between providers (valuation vs. review/opinion);')
bullet("Integration with audit requirements (EY as auditor's view to be obtained);")
bullet('Note: Anthony Cohen / Findex referenced; NCF group likely to use Anthony.')
para('RESOLUTION: Stephen Gledden to lead engagement with Yash and Kim Dema on valuation provider '
     'relationships, scope and budget analysis, to engage with the auditors (EY), and to revert '
     'to Danny van Aswegen — with the objective of a clear roadmap and avoiding duplication ahead '
     'of Thursday’s discussion. No cost commitment to be made until scope and end-goal are '
     'confirmed.', bold=True)

section('6. Conflict Management & Standing Conflicts Register',
        'Presenter: Stephen Gledden / Kim Dema | Time allocated: 10 minutes')
sub('6.1 Standing Conflicts')
para("The Committee discussed developing a standing conflicts register for key individuals. "
     "Stephen Gledden's conflicts were used as an example (Managing Partner; board member; "
     'Trustee NED; Responsible Manager; partner committee; investment committee; valuation '
     'committee; remuneration committee). The agreed position: conflicts are recognised and '
     'managed by independent challenge from others on the relevant bodies, who recognise the '
     'value of that challenge. Stephen Gledden will step off the Valuation Committee and the '
     'Trustee board by year-end; succession to be discussed.')
sub('6.2 Becky Barr — Structural Conflict')
para('It was noted that, at current scale, a single team delivers compliance across both the '
     'Investment Manager and the Trustee. This carries a structural limitation: Becky Barr is '
     'employed by the Investment Manager and holds B-class shares in the Manager, which may '
     'inhibit independent compliance oversight of the Investment Manager on the Trustee’s '
     'behalf. The Committee will identify this in the standing conflicts register and discuss at '
     "board level; a possible outcome is a change to Becky Barr's short-term incentive "
     'remuneration — potentially retaining B-class, or an alternative incentive of equal '
     'magnitude but differently structured.')
sub('6.3 Conflicts & Related-Party Policies')
para('The Committee agreed to read the Conflicts Policy and the Related-Party Transactions Policy '
     'together to ensure familiarity. It was noted these were adapted from Hall & Wilcox '
     'templates; a query regarding "consolidation into a single policy" was clarified as a '
     'branding difference only (no substantive change). Conflicts, standing conflicts and '
     'related-party transactions are to be a standard early agenda item for both Holdings and '
     'Trustee board meetings.')
para('RESOLUTION: Kim Dema to split the standing conflicts out of the existing risk-management '
     'dashboard, add the Trustee conflicts, structure by key individual with oversight mapped '
     'against each, and prepare a draft for board review.', bold=True)

section('7. Documentation & Charters Work Plan',
        'Presenter: Stephen Gledden / Kim Dema | Time allocated: 5 minutes (high-level)')
sub('7.1 Two Workstreams')
para('The Committee agreed documentation work falls into two parts:')
bullet('Policies — "weeding the garden": tightening dates, cross-references and revision dates as '
       'part of Becky Barr’s annual review. A list of documents requiring formal revision '
       'dates may need to be enacted by board minute.')
bullet('Charters — standardising committee charters (Investment Committee, Selection Committee, '
       'Valuation Committee, etc.), which currently exist in multiple inconsistent formats '
       '(one-pagers, PowerPoint, terms of reference). The approach: gather existing documents, '
       'identify duplicates/gaps, agree the correct standard format, and prepare a work plan — '
       'agreeing structure, language and key points before drafting.')
sub('7.2 SharePoint Structure')
para('It was noted the SharePoint restructure is on hold pending a decision to give each '
     'committee (Valuation, Board, etc.) its own site with relevant documents. The Committee '
     'agreed to review SharePoint together to identify existing charters per committee.')

section('8. Close & Next Meeting', 'Chair: Stephen Gledden | Time allocated: 5 minutes')
para('8.1 Three immediate accountabilities were confirmed for the period to the next meeting '
     '(see Action Register), with the broader governance framework sections to be completed when '
     'the Committee reconvenes.')
para('8.2 With a partner meeting due to commence at 12:00 PM, the Chair proposed reconvening '
     'within a couple of days to complete the remaining chapters. Becky Barr to find a suitable '
     'time.')
para('8.3 There being no further business, the meeting closed at approximately 12:00 PM.')

# ---- Action Register ----
heading('Action Register', 1)
grid_table(
    ['Ref', 'Action Item', 'Owner', 'Due By'],
    [
        ['A1', 'Split the standing conflicts out of the risk-management dashboard; add Trustee '
               'conflicts; structure by key individual with oversight mapped against each; prepare '
               'draft for board review.', 'Kim Dema', 'Next Board Mtg'],
        ['A2', 'Lead the search for an independent NED to the Trustee board (including meetings '
               'with Tony McVeigh and Randall this week; follow up Andrew O’Brien re David '
               'Nicholl introduction).', 'Rob Nicholls', 'Ongoing'],
        ['A3', 'Lead engagement with Yash and Kim Dema on valuation provider relationships, scope '
               'and budget; engage EY (auditors) re duplication; revert to Danny van Aswegen with '
               'scope/roadmap. Confirm Danny’s registered-valuer (ATO) status.',
               'Stephen Gledden', 'Before Thu mtg'],
        ['A4', 'Continue annual review of policies ("weeding"), using the governance framework as '
               'guidance; do not commence new documents pending agreed structure/language.',
               'Becky Barr', 'Ongoing'],
        ['A5', 'Develop a standardised pre-meeting checklist for Trustee board sign-offs (proper '
               'instructions, related-party check, executed documentation) and incorporate into '
               'the relevant charter.', 'Kim Dema / Stephen Gledden', 'Next Meeting'],
        ['A6', 'Identify Becky Barr’s B-class share / remuneration structural conflict in the '
               'standing conflicts register for board discussion; consider equal-magnitude '
               'alternative incentive structure.', 'Stephen Gledden / Kim Dema', 'Next Board Mtg'],
        ['A7', 'Committee to read the Conflicts Policy and Related-Party Transactions Policy '
               'together; confirm conflicts / standing conflicts / related-party transactions as a '
               'standing early agenda item.', 'All', 'Next Meeting'],
        ['A8', 'Review SharePoint together to identify existing charters per committee; prepare a '
               'charter standardisation work plan (agree format, language and key points before '
               'drafting).', 'Stephen Gledden / Kim Dema', 'Next Meeting'],
        ['A9', 'Schedule the follow-up meeting to complete the remaining governance framework '
               'sections.', 'Becky Barr', 'Immediately'],
    ],
    widths=[Inches(0.5), Inches(4.3), Inches(1.4), Inches(1.0)],
)

# ---- Confirmation ----
heading('Confirmation of Minutes', 1)
para('These minutes are a true and accurate record of the Trustee Committee Meeting (Governance '
     'Review Implementation Kick-off) of Straight Bat Private Equity Pty Ltd held on 9 June 2026.')
para('Chair', bold=True, space_after=0)
para('Stephen Gledden | Responsible Manager', space_after=0)
para('Date: ________________________', space_after=12)
para('Responsible Manager', bold=True, space_after=0)
para('Rob Nicholls | Responsible Manager (ORM)', space_after=0)
para('Date: ________________________', space_after=12)

para('Draft prepared from the Granola meeting transcript ("Governance Review - implementation '
     'kick-off [Matthew MR]", 9 June 2026). Figures and names captured from discussion and should '
     'be verified against source documents prior to sign-off. Square-bracketed items to be '
     'completed on finalisation.', italic=True, size=8.5, color=GREY)

out = "26'06'09 Governance Review Implementation Kick-off - Minutes (DRAFT).docx"
doc.save(out)
print("Saved:", out)
