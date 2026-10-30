# Implementation Plan

1. Add a host-restricted resolver for Wiki, Sheets, Base, and Docx URLs. Wiki nodes use `get_node` to obtain `obj_type` and `obj_token`.
2. Add a TableService for worksheet and Base metadata, bounded reads, single-cell writes, row append with `INSERT_ROWS`, and Base record create/update. Resolve ambiguity before proposing a write.
3. Extend the office action schema, transport, Mastra tools, gateway routing, OAuth scope storage, and approval summaries. Recheck target values before execution.
4. Update deployment documentation and environment examples. Test routing, permissions, OAuth, approval idempotency, data drift, and API failures.

Default: application identity first; user identity only after an explicit choice and OAuth. No automatic write retry after an uncertain response.
