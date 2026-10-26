---
name: outlook-email-reader
description: Read and interpret a folder of Outlook e-mails for a trip. Use when the user asks you to pull, read, inspect, digest, or use travel correspondence from an Outlook folder.
---

# Outlook E-mail Reader

Use Outlook in two stages. The tools are deliberately generic; this skill supplies the judgment needed to decide which correspondence is worth reading and to turn it into trustworthy travel facts.

## Retrieval method

1. Always call `list_outlook_folder` first. It returns lightweight metadata with short integer IDs only; do **not** try to retrieve every message body in a large folder.
2. Inspect subjects, senders, dates, and attachment flags and identify messages likely to contain operational trip information.
3. Prioritize:
   - airline, rail, cruise, or other transportation confirmations
   - schedule changes, rebookings, and cancellations
   - hotel/lodging confirmations and modifications
   - booked excursions, attractions, tours, and restaurant reservations
   - airport transfers, rental cars, and other ground transportation
   - ticket or reservation messages with dates, times, confirmation numbers, or record locators
4. Usually ignore obvious newsletters, destination marketing, loyalty promotions, surveys, sales offers, generic reminders with no new booking details, and unrelated mail.
5. Call `get_outlook_messages` only for likely-relevant messages, in batches of **no more than 25** short integer IDs. Use the short integer IDs exactly as returned. Start with the strongest candidates. A batch of 20–25 is fine when many messages look relevant; retrieve additional batches only when they are needed to complete or resolve the itinerary.
6. Each retrieved body is limited to its first 500 words. If important details appear to be missing because a body was truncated or because the message has attachments, flag the limitation instead of inventing information.
7. Do not retrieve every message simply because it might contain something useful. The goal is selective retrieval that keeps context small.

## Interpretation method

1. Consider relevant messages in chronological order so later messages can supersede earlier ones.
2. Separate actual bookings and operational travel messages from advertisements, newsletters, loyalty promotions, surveys, and generic "your trip is coming up" reminders.
3. Treat a later change notice as authoritative when it clearly modifies an earlier booking.
4. Treat a cancellation as canceling the corresponding booking unless a later message reinstates or replaces it.
5. Preserve confirmation numbers, record locators, ticket numbers, flight numbers, seat numbers, addresses, phone numbers, dates, and times exactly as written.
6. Do not infer that something is booked merely because an e-mail discusses, recommends, or advertises it.
7. When two messages conflict and the later one does not clearly supersede the earlier one, surface the conflict.

## Output for downstream itinerary work

When the user also wants an itinerary, do not dump e-mail bodies into the chat. Extract a concise set of confirmed travel facts and then use the `itinerary-builder` skill to organize them.

See `references/email-signals.md` for common signals that distinguish confirmations, changes, cancellations, and marketing messages.
