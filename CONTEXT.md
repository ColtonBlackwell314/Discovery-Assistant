# D365 discovery hub — project context

This file captures the full background, decisions, and reasoning behind the D365 discovery assistant built in this repo, so the history isn't locked inside a chat transcript. It covers the goal, how the design evolved, what got built, and what's still open.

## Goal

Colton (CRM consultant, Two Circles) wanted a tool to make D365 discovery work easier for CRM consultants: capturing client notes during discovery workshops, tracking team structure, mapping as-is/to-be process flows, and logging requirements — with no paid APIs, and a design that could later have AI layered on top without a rebuild.

## How it evolved

1. **First pass — single-purpose checklist tool.** A per-module (Sales, Marketing, Customer Service, Field Service, Project Operations, Power Platform, Data & Security, Reporting) question bank with notes capture, a requirements log, and CSV export. Fully offline, localStorage-backed, one file.

2. **Reframed as a centralized hub.** Colton clarified he actually wanted a bigger picture: a web app that centralizes discovery notes, team structure, and process flows (as-is/to-be) in one place, with everything cross-linked — not just a notes checklist. Wireframes were sketched first (dashboard/overview, session notes by module, team structure org view, process flow as-is/to-be) before any code changed, per his request to brainstorm before building.

3. **Implemented the hub.** The single HTML file was rebuilt with six sections: Overview (metrics + recent activity across everything), Discovery notes (module questions), Team structure (stakeholders with role/decision authority), Process flows (as-is/to-be steps per flow), Requirements log, and Search context. A shared `tags` field on notes, team members, flows, and requirements is what makes cross-linking and search work — tag the same topic in two places and both surface together.

4. **AI research and safety note.** Colton asked about free AI API options. Researched Google's Gemini API free tier (Flash-Lite: 1,000 requests/day, no card required, but Google can use free-tier prompts/responses to improve their products — a real consideration for client data). He then shared a screenshot containing a live Gemini API key, which was flagged as sensitive and should never be committed to a repo, screenshotted, or embedded client-side.

5. **AI wireframe, then AI implementation (as placeholder logic).** Sketched where AI would slot in: summarizing module notes, suggesting team structure from notes, drafting to-be process steps, and auto-tagging/drafting a functional design doc from the requirements log. These were then implemented as real, working features in the app — but using local keyword heuristics only, no external API calls, clearly labeled with a purple "AI" badge. The intent: the shape of the feature is real and wired up now; swapping in a real model call later (e.g. Gemini) only touches those specific functions, and the API key itself should never live in client-side HTML — it would need a server or serverless proxy in front of it.

6. **Two primary data-entry paths + responsive rebuild.** Colton wanted exactly two ways data enters the app: typed notes, and Word (.docx) transcriptions. Also wanted discovery notes collapsed to one page (instead of clicking through separate module tabs) with the ability to add/remove custom note pages, a global search bar at the top that searches across all pages (not just a dedicated search page), and the whole thing rebuilt to be responsive using Bootstrap.
   - **Docx import** was implemented with zero external libraries: a hand-written zip-file reader plus the browser's built-in `DecompressionStream('deflate-raw')` API pulls `word/document.xml` out of the uploaded `.docx`, inflates it, and strips it to plain text — entirely client-side, nothing uploaded anywhere.
   - **Bootstrap** is loaded from its CDN (`cdn.jsdelivr.net`) in the HTML `<head>`/before `</body>`. Note: when previewed inside Cowork's own sandboxed artifact viewer, its content security policy blocks that CDN request (only Chart.js/Grid.js/Mermaid are allowed there), so the in-chat preview shows unstyled HTML. Opening the actual saved file in a normal browser works fully — this is a preview-sandbox quirk, not a bug in the file.

7. **Project switcher.** Replaced the small dropdown `<select>` for choosing a project with a proper button in the navbar (📁 project name ▾) that opens a menu listing every saved project with its requirement count, a rename field for the current one, per-project delete, and "+ New project". Each project's data — notes, docx-derived text, team, flows, requirements — stays fully isolated by its own key in `localStorage`; switching just repoints the whole UI at a different key.

## What's built today

Single self-contained file: `d365-discovery-assistant.html` (also mirrored as a live Cowork artifact for quick access). Sections:

- **Overview** — metric cards (notes captured, requirements logged, open decisions) and a merged recent-activity feed across all sections.
- **Discovery notes** — one scrollable page. All eight standard D365 modules as an accordion, each with: a free-form notes textarea, a ".docx" transcript upload that appends extracted text into that same textarea, the original itemized question list (typed notes + tags + priority/type + "add to requirements"), and an "AI: summarize this module's notes" button. Below the accordion, an open-ended list of custom note pages you add/remove yourself, each with the same two entry paths.
- **Team structure** — add stakeholders (name, role, decision authority, tags) by hand, or click "AI: refresh team structure from discovery notes" to get suggested names pulled from existing notes (confirm or dismiss each).
- **Process flows** — per flow, as-is and to-be steps side by side, plus "AI: draft to-be steps from as-is" which proposes steps you can accept, edit, or remove.
- **Requirements log** — everything logged from discovery notes, with an "AI: draft functional design doc" button that assembles a grouped text draft.
- **Global search** — a search box permanently in the top navbar (not a separate page) that live-searches notes, requirements, team members, and flows by keyword or tag, with click-to-jump results.
- **Project switcher** — navbar button + dropdown for creating, renaming, switching between, and deleting projects, each with isolated data.

Everything persists in the browser's `localStorage`. No network calls, no API keys, no paid services anywhere in the app logic itself.

## Open items / next steps

- If/when real AI gets wired in: the four `ai*` functions in the script (`aiSummarizeModuleNotes`, `aiSuggestStakeholders`, `aiSuggestToBeSteps`, `aiDraftFDD`) are the only places that would change. The API key must not live in this HTML file — it needs a small backend or serverless function in front of it so it isn't exposed to anyone who views the page source.
- Google's Gemini free tier (Flash-Lite: 1,000 req/day, no card) is a reasonable starting point for prototyping, but its terms allow Google to use free-tier prompts/responses to improve their products — worth checking with compliance before pointing it at real client discovery data.
- Bootstrap is CDN-loaded; if this ever needs to run fully offline with zero network dependency (even for styling), Bootstrap's CSS/JS would need to be vendored into the repo instead of pulled from `cdn.jsdelivr.net`.
- No automated tests exist yet; verification so far has been manual (opening the file, exercising each feature).
