---
name: itinerary-publisher
description: Create an attractive editable Word DOCX version of an established travel itinerary, using a compact three-column day-by-day layout with color-coded transportation, activities, and lodging. Use when the user asks to print, publish, export, download, or create a Word/DOCX itinerary.
---

# Itinerary Publisher

Publish the itinerary already established in the conversation. Do not re-read Outlook or re-plan the trip unless essential information is missing.

The visual design intentionally resembles a practical hand-built travel itinerary: a full-width date band for each day, transportation on the left, plans in the middle, and lodging on the right. See `references/document-style-guide.md`.

## Content discipline

The Word document is a concise travel reference, not an archive of every fact found in the correspondence.

- Include operationally useful information: dates, routes, times, flight/train numbers, lodging, addresses, confirmation/record locators, seats, transfers, booked activities, and important meeting/pickup instructions.
- **Do not include prices, fares, room rates, taxes, fees, totals, amounts paid, balances, currency amounts, or other pricing information**, even if those details appeared in the e-mails or earlier conversation.
- Do not add promotional text, loyalty-program marketing, cancellation policies, boilerplate terms, or verbose booking-provider language unless the user specifically asks for it.
- Do not add destination recommendations or decorative prose that were not part of the established itinerary.

## Rendering workflow

1. Convert the established itinerary to the JSON shape below, applying the content discipline above.
2. Run `scripts/render_itinerary.py` with that JSON as its single command-line argument.
3. The script prints JSON containing `filename`.
4. Tell the user the Word document is ready and include a markdown link exactly like:
   `[Download the itinerary](/itineraries/<filename>)`

The link should point directly to the generated DOCX. Do not add a preview link or attempt to trigger the download automatically.

## JSON shape

```json
{
  "title": "Optional trip title",
  "days": [
    {
      "date": "Monday, Sept. 21",
      "transportation": [
        {
          "title": "Cathay Pacific 603",
          "mode": "flight",
          "lines": ["HKG - KTM", "6:55 p.m. - 9:45 p.m."]
        }
      ],
      "plans": [
        {
          "title": "Fly to Kathmandu (A330)",
          "lines": ["R/L #EDA3TW", "Seats 14A and 14C", "Driver will be waiting to take us to the hotel"],
          "kind": "travel"
        }
      ],
      "lodging": {
        "title": "Dalai-La Boutique Hotel",
        "lines": ["Chaksibari Marg, Thamel", "Kathmandu 44600", "Expedia #72076541213989"],
        "full_details": true
      }
    }
  ]
}
```

## Column assignment rules

The three columns have strict semantic meanings.

### Left column: transportation

Use the day's `transportation` array only for **scheduled or booked transportation events that the traveler must catch, board, meet, pick up, or return on that date**. The left column is for transportation with a real schedule, reservation, pickup/return event, or similar operational commitment. This includes:

- Flights departing or arriving that day
- Trains departing or arriving that day
- Rental-car pickup or return
- Timed/booked transfers, taxis, shuttles, or private-driver pickups
- Scheduled ferries or boats
- Other booked transportation events with a specific departure, arrival, pickup, or return

Do **not** put ordinary self-directed travel instructions in `transportation`. Instructions such as `Drive to Bucharest`, `Drive to Bran`, `Continue to the hotel`, or `Head to the airport` belong in the middle-column `plans` array, normally with `kind: "travel"`.

Do **not** add entries merely to indicate that transportation remains available or continues from a prior day. Examples that MUST NOT appear in `transportation` include `Avis car rental continues`, `Rental car in use`, `Continue with rental car`, `Car available`, or similar status-only text.

If there is no scheduled/booked transportation event on a day, emit an empty array:

```json
"transportation": []
```

For example, if a rental car was picked up on Monday and the traveler drives to Bucharest on Tuesday, Monday may contain a `car` item for the rental-car pickup. Tuesday's `transportation` array should be empty, and `Drive to Bucharest` should appear in `plans` with `kind: "travel"`.

Every transportation item MUST include a `mode` field. Use one of these values:

- `flight` for airline flights
- `train` for rail travel
- `car` for rental cars or other self-drive segments
- `transfer` for taxis, shuttles, private drivers, or other ground transfers
- `ferry` for ferry or boat transportation
- `other` for any other transportation mode

The renderer uses `mode` to color the entire transportation cell:

- If the cell contains one or more items with `mode: "flight"`, the cell is light pink/red.
- If the cell contains transportation but no flights, the cell is light purple.

Flight numbers, routes, and departure/arrival times MUST NEVER be placed only in `plans`. Every flight segment must appear in `days[].transportation`, even when related travel notes also appear in `plans`.

For example:

```json
{
  "date": "Friday, October 9",
  "transportation": [
    {
      "title": "Delta 724",
      "mode": "flight",
      "lines": [
        "TYS - ATL",
        "2:38 p.m. - 3:45 p.m."
      ]
    },
    {
      "title": "Delta 14",
      "mode": "flight",
      "lines": [
        "ATL - FRA",
        "6:05 p.m. - 8:55 a.m."
      ]
    }
  ],
  "plans": [
    {
      "title": "Fly to Bucharest",
      "lines": ["R/L #HO3CHT"],
      "kind": "travel"
    }
  ]
}
```

A train example:

```json
{
  "transportation": [
    {
      "title": "ICE 612",
      "mode": "train",
      "lines": [
        "Frankfurt - Munich",
        "10:05 a.m. - 1:22 p.m."
      ]
    }
  ]
}
```

IMPORTANT: The renderer does not infer transportation mode from ordinary text. Cell coloring depends on `transportation[].mode`, so always supply the correct mode. Never omit a flight from `transportation` merely because related travel information also appears in `plans`.

IMPORTANT: Do not use `transportation` to describe an ongoing transportation state or an ordinary driving instruction. A rental car that was already picked up does not get repeated on subsequent days merely because it is still being used. `Drive to Bucharest` belongs in `plans`, not `transportation`. The left column should be blank on days with no scheduled/booked transportation event.

### Middle column: plans

Use `plans` for scheduled events, itinerary notes, sightseeing, and related travel notes that are not themselves scheduled/booked transportation events. Self-directed movement instructions such as `Drive to Bucharest`, `Drive to the hotel`, or `Head to the airport` belong here with `kind: "travel"`.

The pale-yellow highlight is reserved **only for scheduled events**. Use `kind: "scheduled_event"` only when the item represents a specific event the traveler is expected to attend or participate in on that date, such as:

- Conference talks or sessions
- Workshops or training classes
- Business meetings or appointments
- Performances or ticketed events
- Booked tours or excursions with a scheduled date/time

Do **not** use `scheduled_event` for generic travel instructions, transit notes, driving directions, sightseeing suggestions, meal suggestions, reminders, or unscheduled activities. Examples that must remain unshaded include `Drive to Bucharest`, `Visit Bran Castle`, `Check in at hotel`, and `Explore Old Town`. Use `kind: "travel"`, `kind: "note"`, or omit `kind` for those items.

Examples:

```json
{
  "plans": [
    {
      "title": "AI Workshop Day 1",
      "lines": [],
      "kind": "scheduled_event"
    }
  ]
}
```

The example above is highlighted pale yellow. This one is not:

```json
{
  "plans": [
    {
      "title": "Drive to Bucharest",
      "lines": [],
      "kind": "travel"
    }
  ]
}
```

IMPORTANT: Yellow highlighting is controlled entirely by `plans[].kind`. Never label an ordinary instruction or unscheduled activity as `scheduled_event` merely to emphasize it. If a middle-column cell contains a mixture of scheduled events and ordinary notes, keep the cell unshaded.

### Empty days

If there are **no confirmed plans or activities for a day**, leave the middle column blank by emitting an empty array:

```json
"plans": []
```

Do **not** invent placeholder entries such as `Open day`, `Free day`, `Nothing scheduled`, `No confirmed activities found`, or similar explanatory text. An empty middle-column cell is the intended presentation for a day with no confirmed activities.

### Right column: lodging

Use `lodging` for the hotel or other overnight accommodation. Set `lodging.full_details` true on the first day a property's address/confirmation details are shown; later days may carry only the hotel name with false.

Keep the document faithful to known facts and omit pricing information.
