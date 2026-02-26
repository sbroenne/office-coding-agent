---
name: outlook-calendar
description: >
  Specialized skill for creating appointments, scheduling meetings,
  and managing calendar-related tasks from within Outlook.
version: 1.0.0
license: MIT
hosts: [outlook]
---

# Calendar & Appointments Skill

Activate this skill when the user wants to create appointments, schedule meetings, or manage calendar items.

## Appointment Creation Workflow

1. Gather details from user or email context:
   - Subject / title
   - Date and time (start + end or duration)
   - Location (if any)
   - Attendees (if a meeting)
   - Description / agenda
2. `display_new_appointment` → open appointment form with pre-filled details
3. Confirm what was created

## Extracting Meeting Details from Emails

When the user wants to create a meeting from an email:

1. `get_mail_item` → get sender and subject
2. `get_mail_body` → extract proposed dates, times, locations, agenda
3. Parse the key details:
   - **When**: Look for dates, times, "next Tuesday", "3pm", etc.
   - **Where**: Look for room names, links, addresses
   - **Who**: Sender + any mentioned attendees
   - **What**: Meeting topic / agenda items
4. `display_new_appointment` → create with extracted details
5. Summarize what was scheduled

## Best Practices

- **Always confirm ambiguous dates** — "next Friday" could mean two different days
- **Include timezone context** if participants are in different zones
- **Set reasonable defaults** — 30min for quick meetings, 60min for discussions
- **Add email sender as attendee** when creating from an email
- **Include email subject in appointment subject** for context

## Output Guidelines

- Confirm the appointment details before creating
- After creation, summarize: subject, date/time, location, attendees
- Suggest adding an agenda if none was provided
