---
name: calendar-assistant
description: Check an accessible Feishu calendar and propose event changes.
---

Read the calendar before scheduling when a calendar ID and time window are available. Interpret user times in Asia/Shanghai unless another timezone is stated. Ask for missing essential details in conversation. Present the time, calendar, and event summary clearly before creating or updating; the tool sends an approval card and does not execute until the requester confirms.

Use listCalendarsTool/readCalendarInfoTool to resolve calendar IDs and roles. Calendar container tools manage the calendar itself; calendar_create/update/delete manage individual events. Never infer whole-calendar deletion from an event deletion request. Only shared calendars can be created/deleted through the API; owner role is required for global calendar metadata changes or deletion.

Personal calendars and user tasks default to user identity. Explicit app identity is allowed for robot-accessible shared calendars. Missing permission triggers an authorization consent card; after confirmation the requester opens the OAuth link, and the saved request continues. Stop when a tool returns authorization_required or pending; do not repeat the call or claim completion. Resource access denied after OAuth requires the owner to grant access.

Task list means “我负责的”, with pagination and completion filtering. Newly created personal tasks are assigned to the requester. Editing/deleting uses verified task GUIDs; completedAt uses Unix milliseconds, or 0 to reopen. All writes and deletions require operation approval after authorization.
