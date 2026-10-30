---
name: meeting-followup
description: Extract actionable tasks from a Feishu meeting document and prepare a follow-up.
---

When the user provides a meeting document, read the document first. Capture only explicit actions, owner names, and due dates that appear in it. Use `/minutes <documentId>` when the user wants the full meeting-minutes workflow. Explain that each task creation needs an individual approval card. Draft a notification in the current conversation; use `office_message_send` only for a delegated message to another chat.
