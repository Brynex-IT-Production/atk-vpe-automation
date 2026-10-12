# Grammarian Manual Flow — Locked Specification

Status: LOCKED for DEV implementation
AI dependency: PARKED / OPTIONAL FUTURE ENHANCEMENT

## Canonical WhatsApp flow

### Word of the Day — Message 1
Thank you for being the Grammarian for ATK Meeting {{meeting_no}}. 😊

*WORD OF THE DAY*

Please copy the next message, fill in the details, and send it back.

### Word of the Day — Message 2
Word:
Meaning:
Usage:

### Idiom of the Day — Message 1
Thank you! Your Word of the Day is received. ✅

*IDIOM OF THE DAY*

Please copy the next message, fill in the details, and send it back.

### Idiom of the Day — Message 2
Idiom:
Meaning:
Usage:

## Core strategy

1. Manual-first. Word/Idiom, Meaning and Usage are supplied by the member.
2. AI must never be required for completion. It may be added later only as an optional suggestion feature.
3. The parser accepts:
   - Word:/Idiom:, Meaning:, Usage:
   - capitalization differences
   - extra spaces
   - copied prompt text before fields
   - multiline field values
   - Example / Example usage as Usage aliases
   - a plain standalone Word/Idiom, then requests the missing fields
4. Partial input is preserved. If one or more fields are missing, request only the missing fields.
5. Duplicate inbound WhatsApp messages are protected by source Message ID.
6. Before saving, verify the member is still assigned as Grammarian for that meeting.
7. Real meeting assignment state takes priority over stale DEV/test conversation state.
8. Word completion advances to Idiom; partial data must not force a restart.
9. Final canonical values are stored in WA Submissions as GRAMMAR_WORD and GRAMMAR_IDIOM.
10. After both stages are complete, show one final review with Confirm / Edit Word / Edit Idiom.
11. The Grammarian task is completed and handed off to TMOD/VPE only after Confirm.
12. New Meta template submission is never automatic and requires explicit VPE approval.

## State model

WORD_PENDING
→ WORD_PARTIAL (optional)
→ WORD_COMPLETE
→ IDIOM_PENDING
→ IDIOM_PARTIAL (optional)
→ IDIOM_COMPLETE
→ REVIEW_PENDING
→ CONFIRMED
→ COMPLETED

From REVIEW_PENDING the member may choose:
- Confirm
- Edit Word
- Edit Idiom

Edits return to REVIEW_PENDING before completion. Final handoff must not occur until Confirm is selected.

## Recovery rules

- Only Word/Idiom received: save it in conversation state; request Meaning + Usage.
- Word/Idiom + Meaning received: preserve both; request Usage only.
- Meaning/Usage arrives later: merge with preserved partial data.
- Opposite-stage label received: do not silently misclassify it.
- Assignment changed: do not save; close the stale conversation.
- Unrecognizable input: resend only the small format needed for the current stage.
- AI/API failure: irrelevant to the manual workflow; collection continues normally.
