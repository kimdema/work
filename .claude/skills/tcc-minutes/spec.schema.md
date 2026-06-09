# Minutes JSON spec

`render_minutes.py` consumes a single JSON object. All keys are optional except where a
section is wanted — omit a key to skip that block.

```jsonc
{
  "title":      "TRUSTEE COMMITTEE MEETING — MINUTES OF MEETING",  // optional; sensible default
  "subtitle":   "Straight Bat Private Equity Pty Ltd | 9 June 2026",
  "subheading": "Governance Review — Implementation Kick-off",     // meeting name, italic
  "status":     "CONFIDENTIAL — DRAFT",                            // red, centred

  "details": [                       // 2-col key/value table
    ["Entity", "Straight Bat Private Equity Pty Ltd"],
    ["Meeting Type", "Trustee Committee Meeting"],
    ["Date", "9 June 2026"],
    ["Time", "11:00 AM – 12:00 PM"],
    ["Location", "Gate 8, 88 Jolimont St, East Melbourne VIC & TEAMS"],
    ["Chair", "Becky Barr (Compliance Manager)"],
    ["Minutes Prepared By", "Kim Dema (Company Secretary)"],
    ["Date of Confirmation", "[To be completed upon sign-off]"]
  ],

  "attendees": [                     // 3-col table
    {"name": "Becky Barr", "role": "Compliance Manager (Chair)", "status": "Present"},
    {"name": "Stephen Gledden", "role": "Responsible Manager (RM)", "status": "Apologies"}
  ],

  "sections": [
    {
      "heading":   "1. Opening & Quorum Confirmation",
      "presenter": "Chair: Becky Barr | Time allocated: 5 minutes",
      "blocks": [
        {"type": "sub",        "text": "1.1 Quorum"},
        {"type": "para",       "text": "The Chair confirmed a quorum ..."},
        {"type": "bullet",     "text": "A weekly liquidity monitoring process ..."},
        {"type": "resolution", "text": "The Committee noted the Legal and Risk Register ..."}
      ]
    }
  ],

  "actions": [                       // 4-col Action Register
    {"ref": "A1", "action": "Amend the legal and risk register ...", "owner": "Kim Dema", "due": "Next Board Mtg"}
  ],

  "confirmation": {
    "statement": "These minutes are a true and accurate record of the Trustee Committee Meeting ... held on 9 June 2026.",
    "signatories": [
      {"role": "Chair", "name": "Becky Barr", "title": "Compliance Manager"},
      {"role": "Responsible Manager", "name": "Rob Nicholls", "title": "Responsible Manager (ORM)"}
    ]
  },

  "footnote": "Draft prepared from the Granola meeting transcript ... Verify figures and names before sign-off."
}
```

## Block types (inside `sections[].blocks`)
| `type`       | Rendered as                                            |
|--------------|--------------------------------------------------------|
| `sub`        | Bold sub-heading (e.g. `2.1 Quorum`)                   |
| `para`       | Body paragraph                                         |
| `bullet`     | Bulleted list item                                     |
| `resolution` | Bold paragraph, auto-prefixed `RESOLUTION: ` (override with `"prefix"`) |

The example at `examples/2026-06-09_governance_kickoff.json` is a complete, real spec.
