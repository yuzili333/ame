# Inline Plan Review

- Requirement fit: both ordinary spreadsheets and Base are routed from direct and Wiki links.
- Security: document URLs are parsed locally and only accepted from configured hosts; external URL hosts are never fetched. User OAuth is bound to the original requester. Writes use the existing approval state machine.
- Failure handling: missing access prompts for identity choice; absent API endpoints report unavailability; ambiguous targets and changed cells or fields stop before writing.
- Compatibility: the existing calendar OAuth path remains, with a nullable scope column migration for older tokens. Existing Docx ID and block operations remain available for actual Docx IDs.

No subagent dispatch was requested. Review was performed inline before implementation.
