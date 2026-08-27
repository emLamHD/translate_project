# Translation provenance policy

| Provenance | NO-AI automatic use |
|---|---:|
| `human_approved`, `approved=true` | Allowed |
| `owner_manual`, `approved=true` | Allowed |
| `claude_silver_reference` | Rejected |
| `google_machine_draft` | Rejected |
| `unknown` | Rejected |
| Any entry with `approved=false` | Rejected |

Legacy facts for this work item:

- The two private bilingual references were produced with Claude and remain silver references only.
- The previous blind bilingual draft used Google Translate no-key and is classified `AI_DERIVED_LEGACY_DRAFT`.
- The previous flat cache has no per-entry provenance. Migration must mark its entries unapproved and cannot make them eligible for final translation.

The repository does not store private references, migrated cache, human TM, or reports containing customer content.
