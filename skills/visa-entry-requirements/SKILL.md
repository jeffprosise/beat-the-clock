---
name: visa-entry-requirements
description: Research current visa, electronic travel authorization, arrival form, transit, passport-validity, and other entry requirements for the countries in an itinerary. Assume the traveler is a U.S. citizen traveling on a U.S. passport. Use whenever the user asks what visas, entry permits, travel authorizations, arrival forms, or passport requirements are needed for a trip.
---

# Visa & Entry Requirements

Assume the traveler is a U.S. citizen traveling on an ordinary U.S. passport unless the user explicitly says otherwise.

Visa and entry rules change. Do not answer from memory. Use `web_search` to research the current requirements for every country the traveler will enter and any transit country where a transit requirement could apply.

## Method

1. Start from the established itinerary and enumerate every country entered. Distinguish a connection/transit from an actual entry when the itinerary makes that clear.
2. Search the Web separately for each country. Prefer authoritative sources in this order:
   - the destination country's official immigration, border, foreign-affairs, embassy, or consular site;
   - the U.S. Department of State country-information page;
   - official government travel-authorization or arrival-form sites.
3. Search specifically for all of the following when applicable:
   - visa or visa-free status and permitted stay;
   - electronic travel authorization/eVisa/ETA requirements;
   - mandatory online arrival or departure forms;
   - passport validity and blank-page requirements;
   - transit-visa requirements;
   - proof of onward travel or other material entry documentation.
4. Do not treat commercial visa-service sites, travel blogs, airline summaries, or search-result snippets as authoritative when an official source is available.
5. If official sources conflict or the answer depends on an itinerary detail that is unknown, say so rather than guessing.
6. Include the date the research was performed and make clear that requirements can change before departure.
7. Provide source links for the authoritative pages used.

## Output

Use a compact country-by-country format. For each country, state only requirements that matter to this trip. Clearly distinguish:

- **Required before travel**
- **Required/checked on arrival**
- **Not required** when that fact is useful (for example, "No visa required for stays up to X days")
- **Transit-only considerations**, if any

End with a short pre-departure checklist containing only actionable items.

See `references/research-checklist.md` before finalizing the answer.
