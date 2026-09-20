
# Specification: Troop 303 Court of Honor Coordination Tool

## 1. Project Overview
A specialized React application designed to streamline the coordination of Scouting Courts of Honor (CoH). This tool manages the timeline, parses advancement data, and generates pre-formatted emails to key stakeholders.

### Core Constraints (Non-Negotiable)
* **Local-First Architecture:** No backend server or database.
* **Privacy:** All processing (XLSX parsing and Doc generation) must happen in the browser. 
* **Persistence:** Use `localStorage` to save contact info, Google Doc links, and milestone progress. 

---

## 3. Persistent Data (Settings)
The user should be able to configure and save the following in a "Settings" panel:
* **Advancement Chairs:** Denise Kim, Katherine Loya Dial.
* **Leadership:** Bruce McGurk (Solo contact), Michael Finegan, Dennis Sinclitico.
* **Youth Leaders:** Placeholder for current SPLs (e.g., Ira, Jordan).
* **Troop List:** `troop303-orinda@googlegroups.com`.

---

## 4. The Coordination Workflow (The "Milestone Tracker")

The UI should be a vertical stepper. Each step contains a "Generate Email" button that triggers a `mailto:` link.

### Milestone 1: Advancement Data Request (3 Weeks Out)
* **To:** Denise Kim, Katherine Loya Dial.
* **Subject:** Advancement and Court of Honor [CoH Date].
* **Body:** > Hey Denise and Katherine, we have a Court of Honor on Monday, [Date]. Where are we on the data for that? If I can get the report by Saturday morning, I can make the script for the MC's so they have time to look at it over the weekend.

### Milestone 2: Bruce’s Separate Reminder
* **To:** Bruce McGurk (**Note:** Bruce must never be CC'd on other coordination threads).
* **Subject:** Troop 303 Court of Honor - [CoH Date].
* **Body:** > Hey Bruce, will you be able to make it to the Court of Honor on Monday, [Date]? It's always nice to have you there to officially open and close the ceremony.

### Milestone 3: MC & Color Guard Recruitment
* **To:** Senior Patrol Leaders, Michael Finegan, Dennis Sinclitico.
* **Input Field:** The UI must provide text boxes to enter the names of the recruited MCs and Color Guard.
* **Body:** > Hi [SPL Names], I’m starting to build out the team for our upcoming CoH on Monday, [Date]. Who will be serving as our MCs and who is assigned to the Color Guard? I’d like to get their names into the script as soon as possible. As a reminder, they need to arrive at 6:00 PM for rehearsal.

### Milestone 4: Recognition Preview (Parent Email)
* **Requirement:** Requires `.xlsx` upload first.
* **To:** Troop Distribution List (Parents).
* **Body:** > The following scouts will be recognized at the Court of Honor on [Date]:
    > [BULLETED LIST OF NAMES FROM PARSED FILE]
    > Please let me know if any corrections are needed.

### Milestone 5: Achievement Coordinator Review
* **Input Field:** The UI must provide a text box for the "Google Doc Script Link."
* **To:** Denise Kim, Katherine Loya Dial.
* **Body:** > The script is ready. Here is the link to the google doc: [DOC_LINK]. Please correct any mistakes you see before I send it out to the scouts.

### Milestone 6: Final Script & Rehearsal Instructions (1 Week Out)
* **To:** MCs, Color Guard, SPLs.
* **Body:** > [MC Names], thank-you for being our MC's. [Color Guard Names], thank-you for being the Color Guard. Please arrive at 6:00 PM (one hour early) for rehearsal. Here is the link to the script: [DOC_LINK]. MCs, please be prepared to pronounce all scout names correctly.

---

## 5. Script Generation Logic
The tool must include a button to **"Download Master Script (.docx)"**.
* **Structure:** Standard Troop 303 ceremony flow (Opening, Rank Advancements, Merit Badges, Closing).
* **Variables:** Auto-populate the document with the MC names, Color Guard names, and the full list of Scout awards parsed from the Excel file.

