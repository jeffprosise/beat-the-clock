---
name: itinerary-builder
description: Assemble a chronological travel itinerary from confirmed travel facts such as flights, hotels, transfers, excursions, tickets, and reservations. Use when the user asks to build, organize, update, or review an itinerary.
---

# Itinerary Builder

Create a practical day-by-day itinerary from facts already established in the conversation or extracted by another skill. Do not re-fetch Outlook unless the needed correspondence has not yet been read.

## Method

1. Build one entry for every calendar day from the first travel day through the last when the dates are known.
2. Sort all confirmed events chronologically using the local date/time shown by the booking.
3. Keep transportation, plans/activities, and lodging conceptually separate.
4. For multi-leg travel, preserve every leg, flight/train number, route, and departure/arrival time.
5. Carry lodging forward to subsequent nights when the stay clearly spans them, but do not invent a checkout date if it is unknown.
6. Include ground transfers and pickup instructions when known.
7. Preserve confirmation numbers, record locators, ticket numbers, seats, addresses, and useful contact details exactly.
8. Distinguish booked activities from ideas or suggestions. Never promote a suggestion into a reservation.
9. Flag apparent timing conflicts, impossible connections, missing lodging, or other material gaps.
10. Never invent missing details. Use "not provided" only when the omission matters; otherwise leave the field out.

## Chat format

For normal chat, use a compact day-by-day format. If the user asks for a Word document, use the `itinerary-publisher` skill instead of trying to format the document in chat. The publisher intentionally omits pricing information even when prices appear in source e-mails.

## Structured handoff to the publisher

The publisher works best when each day can be represented with three visual columns:

- **transportation**: one or more flight/train/transfer items
- **plans**: travel purpose, excursions, booked activities, notes, confirmation details
- **lodging**: hotel/property details

See `references/itinerary-checklist.md` before finalizing a complex itinerary.
