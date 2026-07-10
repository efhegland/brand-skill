---
description: "Builds, regenerates, or modifies an M1-branded PowerPoint deck. Use when the user mentions: PowerPoint, ppt, pptx, presentation, slide deck, slides, deck, brand deck, M1 deck, building a deck, regenerating a deck, slide outline, or any request to produce slides in the M1 brand."
disable-model-invocation: false
---

### Presentation tasks

Building a PowerPoint deck, slide layout selection, or presentation from an outline.

- Follow the complete workflow below.

### Main directives

These rules apply to EVERY build, no exceptions:

1. **No placeholder data in the final deck.** Charts must contain the actual data from the outline (not Q1–Q4 / Series 1–3 template defaults). Tables must contain real cells (not sample rows). Textboxes must contain real content (not "Subhead | Any key takeaways from this section."). **Never ship a slide with visible template placeholder content** unless the user has explicitly asked for it (e.g., "leave the chart blank so I can fill it in").
2. **No empty or incomplete slides.** A slide with a title and no body content is a defect, not a "minimal" slide. An empty body placeholder (PowerPoint's "Click to add text" prompt) is worse than a placeholder data slide — it ships visibly broken. If a slide doesn't have enough real content to stand on its own, **do not include it in the deck**. Emit a routing warning naming the missing content so the outline can be extended, but do NOT build the slide.
3. **Decision tree when outline lacks data for an intended slide:**
   - Can the helper populate it from outline data that exists? → Enhance the helper.
   - Can the slide be replaced with a different format that uses the data you DO have? → Replace.
   - Is the outline simply missing the data? → **Skip the slide entirely.** Emit a routing warning. Do not ship a title-only / empty body / placeholder-stub slide.
4. **Verify visually before declaring a deck finished.** Render every chart, graphic, table, and any slide that's been demoted or special-cased via the qlmanage loop and read the PNGs. Spot-check for placeholder leak and empty placeholders.
5. **When a directive is violated, the build is a defect — fix and re-verify.** Don't ship a deck and ask the user to clean up.

---

## Presentation workflow (PowerPoint)

### Setup — guided steps with progress breadcrumb

Before building, walk the user through setup **one step at a time** using tasks as a visual progress breadcrumb. The task bar at the top shows all steps, with the current step marked in-progress (spinner). This gives the user a clear sense of where they are.

#### Tool prerequisites

Both `AskUserQuestion` and `TaskCreate`/`TaskUpdate` are deferred tools. Before the first setup step, run `ToolSearch` with query `"select:AskUserQuestion,TaskCreate,TaskUpdate"` to load their schemas. This only needs to happen once per session.

#### How to use `AskUserQuestion`

**Use `AskUserQuestion` with selectable options** for every question that has a finite set of answers (yes/no, multiple choice). This renders as a tab/arrow UI that the user can click or arrow-key through. Only fall back to free-text input when the answer is open-ended (e.g., file path, pasted content).

Every call must include a `questions` array. Each question object requires:

- `question` (string) — the question text
- `header` (string) — short label, max 12 chars (e.g., "Font", "Notes", "Save as")
- `options` (array of 2-4 objects) — each with `label` (string, 1-5 words) and `description` (string)
- `multiSelect` (boolean) — `false` for single-choice

Example:

```json
{
  "questions": [{
    "question": "Is the Inter Variable font installed on your system?",
    "header": "Font",
    "options": [
      {"label": "Yes", "description": "Inter Variable is already installed"},
      {"label": "No", "description": "I need to install Inter Variable first"}
    ],
    "multiSelect": false
  }]
}
```

#### Step creation — all at once, up front

At the start of the setup flow, create **all** setup tasks in a single batch so the full breadcrumb is visible immediately. Skip any step whose answer is already known from context (e.g., content was provided with the request, no image tags found). Only create tasks for steps that will actually run.

The steps and their task definitions are:

| Step   | Task subject       | activeForm                  | Condition                                              |
| ------ | ------------------ | --------------------------- | ------------------------------------------------------ |
| Q1     | Font check         | Checking font               | Always                                                 |
| Q1b    | OneDrive check     | Checking OneDrive           | Always                                                 |
| Q2     | Content            | Getting content             | Skip if content already provided                       |
| Q2b    | Copy polish        | Polishing copy              | Always (after Q2 content is available)                 |
| Q2c    | Visuals scan       | Scanning for visuals        | Always (after Q2b outline is final)                    |
| Q2d    | Section break mode | Choosing section break mode | Only if the outline has ≥1 section divider             |
| Q2e    | Content bullet mode| Choosing bullet style       | Always (default Yes — body content is near-universal)  |
| Q3     | Speaker notes      | Setting up notes            | Always                                                 |
| Q4     | File location      | Setting save location       | Always                                                 |
| Q5     | Images             | Collecting images           | Only if content has image tags                         |
| Review | Slide plan         | Reviewing slide plan        | Always                                                 |
| Build  | Build deck         | Building presentation       | Always (created with setup tasks)                      |
| Verify | Visual check       | Verifying rendered cover    | **Always — mandatory before declaring delivery**       |

Create all applicable tasks at once, then immediately set Q1 (or the first applicable step) to `in_progress`.

#### Step execution — one at a time

For each step: mark it `in_progress` → ask the question via `AskUserQuestion` → process the answer → mark it `completed` → mark the next step `in_progress` → repeat.

**Q1 — Font check:** Silently check whether Inter Variable is installed. On macOS, `fc-list` is **not** available — check the font file directly with `test -f ~/Library/Fonts/InterVariable.ttf` (or `system_profiler SPFontsDataType | grep -i "inter variable"`). On Windows, check the system font directory. If the font is found, mark this task completed automatically — no question needed, just briefly note "Inter Variable detected" in your output. If it is NOT found, ask the user: "Inter Variable isn't installed on your system. Would you like to install it now?" with options "Yes, install it" / "Skip for now". If yes, copy `~/.claude/skills/m1-brand/assets/InterVariable.ttf` to `~/Library/Fonts/` (macOS) or offer to open it for double-click install. See `brand-font.md` for full detection/install detail. Wait for confirmation before continuing.

**Q1b — OneDrive connectivity check:** Silently test whether the M1 OneDrive image catalog is reachable before any build step runs. OneDrive mounts the M1 org folder at platform-specific locations — auto-detect across the candidates below using `os.path.isdir(os.path.expanduser(...))` (the `~` form is portable across macOS and Windows):

| Platform | Candidate path |
|---|---|
| macOS | `~/Library/CloudStorage/OneDrive-M1Finance/_Creative Assets/_Creative Asset Hub/Images-ArticlesStockImages/in/img-OrigStockDownloads/` |
| Windows | `~/OneDrive - M1Finance/_Creative Assets/_Creative Asset Hub/Images-ArticlesStockImages/in/img-OrigStockDownloads/` |
| Windows (alt) | `~/OneDrive-M1Finance/_Creative Assets/_Creative Asset Hub/Images-ArticlesStockImages/in/img-OrigStockDownloads/` |

**This candidate list is universal across all M1 co-workers** — `~` resolves to each user's own home directory and the org-folder names are M1's standard OneDrive mount names (macOS uses the `CloudStorage/OneDrive-M1Finance` form; Windows typically uses `OneDrive - M1Finance` with spaces, occasionally without). Never substitute an absolute path containing a specific user's home (e.g., `/Users/<someone>/...` or `C:\Users\<someone>\...`) and never reference personal credentials.

If any candidate resolves, store that path as the session's OneDrive root for the rest of the build and mark Q1b completed silently (just note "OneDrive catalog detected" in your output). If NONE of the candidates resolve, prompt the user: "I can't find the M1 OneDrive image catalog at any of the standard locations. Open the OneDrive app and let me know when the `_Creative Assets` folder has synced — or paste the absolute path to your local `_Creative Assets` folder and I'll use that for this build." Options:

- **"I'm connected — re-check"** — re-runs the auto-detection across all candidates above
- **"Here's the path"** — user pastes an absolute path; validate with `os.path.isdir()`, and on success store as the session's OneDrive root
- **"Skip OneDrive for this build"** — go straight to no-OneDrive mode (see below)

When a user pastes a path, **use it for this build session only**. Do not write it into the skill, into any script, or into a config file — it's per-session state. Hardcoding a co-worker's local path would break every other co-worker's build.

**Retry up to 3 total attempts** across all approaches (re-check, path paste, etc.). If the catalog is still unreachable after the 3rd attempt — or the user selects "Skip OneDrive for this build" at any point — mark the deck as **no-OneDrive mode** and proceed to Q2. In no-OneDrive mode:

- Every image slot (S9/S10/S11 auto-imagery, named-file `<img name>` tags, and bare-tag `<img>` photo sourcing) builds with a **grey-box placeholder** in the image area
- Each placeholder includes a visible on-slide "Insert image here" marker so the user can see exactly where to drop the file
- The final build summary lists every slide that needs a user-supplied image and reminds the user to insert images before sharing the deck

**Q2 — Content:** "What content should go into the deck? You can paste text, share a file path, or describe what you need." Accept any text format: markdown, .txt, .md, pasted text, outlines, briefs, or conversational descriptions. If the content is sparse, follow up for: topic/objective, audience, key points or outline, any data for charts. If the user provides a brief: interpret it per brand rules — understand the objective, extract core ideas, rewrite in M1 voice. Never regurgitate a brief.

**After receiving content, validate the outline format** (see "Expected outline format" below). If the content does not match the expected format, pause and educate the user on the correct structure before proceeding. Do not attempt to guess the mapping — get it right first.

**Q2b — Copy polish:** "Would you like Emilybot to polish the copy before building? She'll rewrite the text in M1 voice while keeping your outline structure intact." with options "Yes, polish it" / "No, use as-is". If yes: load `brand-copywriting.md`, then `copywriting/copywriting-philosophy.md` and `copywriting/copywriting-voice-tone.md` per its routing, then rewrite the outline content following M1 brand voice rules. **Preserve the exact outline hierarchy** — `Title:`, Roman numerals, capital letters, Arabic numerals, dashes/bullets, and image tags must remain in place and unchanged. Only rewrite the text within each level: tighten wording, apply conversational clarity, cut adverbs, use benefit-first language, and ensure sentence case. Do not add or remove outline levels, merge slides, or restructure sections. Show the user the polished outline and ask: "Does this look good?" with options "Looks good" / "I have changes". If they have changes, adjust and re-present. Once approved, use the polished version for all subsequent steps.

**Q2c — Visuals scan:** Runs after the outline is finalized. Goal: ensure ~40–50% of content slides have a visual. This step is additive — it works on top of whatever the user already tagged.

**Step A — Auto-assign chart/graphic slides (no question needed):**

This step assumes the pre-analysis pass described in **Step 3** has already categorized each `A.` subsection as Graphic, Chart, or plain content. Walk the categorized subsections and substitute the appropriate template slide. Show the substitution in the slide plan table — user can override at Review.

**Reference the Graphics vs Charts framework above** for the boundary between guiding-number visuals and spreadsheet-data visuals. The substitution table below is organized by category to make that boundary explicit:

| Category | Trigger in outline | Substitute | Clone from |
|---|---|---|---|
| **Graphic** | Time-sequence / phases with NO per-step numeric data ("Q1: design, Q2: build, Q3: launch", "step 1, 2, 3") | S14 timeline graphic | `prs.slides[13]` |
| **Graphic** | Concepts that build upon or nest inside one another ("core / utility / desire", "layers built on") | S15 target graphic | `prs.slides[14]` |
| **Graphic** | Small fixed count of topics / takeaways with key learnings under each ("3 pillars", "5 priorities") | S16 takeaways graphic | `prs.slides[15]` |
| **Chart** | Spreadsheet data with groupings — multi-series numeric values, before/after comparisons, trend lines | S12 bar/line chart | `prs.slides[11]` |
| **Chart** | A single category of items that add up to 100% (allocation %, "X out of Y", portfolio splits) | S13 pie chart | `prs.slides[12]` |
| **Chart** | Larger spreadsheet data where presenting many values is the focus (matrix, specs, feature list) | S17 table | `prs.slides[16]` |

**When the boundary feels close, read the speaker note on each candidate template slide** (`prs.slides[N].notes_slide.notes_text_frame.text`). Those notes are written specifically to answer "best for X". A time-series with per-period dollar figures → Chart (S12 bar/line). A time-series with only step names → Graphic (S14 timeline). The presence of per-item numeric data is the boundary.

**Title number-keyword hint — graphic first, image fallback:** If a subsection's TITLE contains a count word ("Two", "Three", "Four", "Five", "Six") or a small digit (2–10), that is a strong hint the content fits a graphic shape. Try the structural match first (sequential → S14, nested/concentric → S15, discrete-with-takeaways → S16). If the content doesn't cleanly fit any graphic, fall through to plain content — `set_content_series` will then catch the sparse case and promote to an image variant. See the S9/S10/S11 polish rule below for the sparse-promotion threshold (≤3 L1 bullets, ≤60 words, no L3 nested).

**Build-time routing audit (`routing_warnings`):** `set_content_series` automatically flags any subsection whose title has a number keyword AND whose L1 bullet count matches that number (e.g. title "Two models" + 2 L1 bullets, title "Three stages" + 3 L1 bullets). The warning surfaces at the end of the build alongside imagery warnings. Treat it as a checklist item: open the slide, judge whether a graphic would communicate better. If yes, re-route in the build script; if no (e.g. image-fallback was intentional like the I.C "Two models" case), the warning is informational — the build still succeeded. The audit is implemented in `_title_number_keyword()` + `_count_l1_bullets()` and only fires when BOTH signals match (a title with "two" but 5 L1 bullets is NOT flagged — likely just rhetorical use of the word).

For chart/graphic substitutions: clone the template slide, set the title via `set_content_title`, then **populate every brand-template text slot from the outline content** — don't leave the template's `"Subhead | Any key takeaways from this section."` scaffolding visible in the output.

- **S14 / S15** helpers accept `(label, description)` tuples (or plain strings, label-only). The label fills the topic/ring/step rectangle; the description fills the per-topic subhead/takeaway slot, preserving the template's bold-subhead / regular-takeaway styling.
- **S14 also accepts an optional `takeaway=` summary** that fills the bottom key-takeaway textbox. If your outline doesn't have one, omit it — the helper deletes the textbox rather than leaving the placeholder string.
- **S16 supports a flex row of 1–4 proof points per topic.** Each entry in `topics` is one of:
    - `"Topic"`                          — topic pill only, no proofs
    - `("Topic", "takeaway sentence")`   — one WIDE proof spanning the full right-side width
    - `("Topic", [proof, …])`            — 1–4 proofs laid out as a flex row of equal columns with even gaps
  Each `proof` is a `(subhead, takeaway)` tuple OR a `"subhead: takeaway"` string. The helper clones the per-topic source textbox as many times as needed and resizes them to: 1 → full row, 2 → halves, 3 → thirds, 4 → quarters. Default (1 proof) is wide — that is the norm. Use multi-proof rows only when the outline supplies sub-bullets under that pillar.
- **Outline syntax for S16 multi-proof:** put each proof point as a sub-bullet (`-`) under the numbered pillar in `test-01.md` (or Emilybot's outline). Each sub-bullet is `Subhead: takeaway` — the colon splits the bold subhead from the regular takeaway. No sub-bullets → 1 wide proof using the pillar's own colon-separated description.
- **Unused template scaffolding is removed, not left blank.** If S14 only has 3 stages, the 4th and 5th ovals + step-labels are deleted. The S16 helper always deletes the template's prebuilt 4-column Topic-3 scaffolding (TextBox 31/32/33) up front and re-builds dynamically based on the actual proof count.
- **Why:** the brand-correct scaffolding rule applies to *non-text* shapes only (chart geometry, decorative rules, layout containers). Any visible placeholder *text* in the rendered output is a build defect — the reader sees "Subhead | Any key takeaways…" and the deck looks unfinished.
- **What still requires Emilybot copy:** content the outline genuinely does not contain. Today that's (a) the S14 bottom takeaway summary, written as `Key takeaway: …` after the step list, and (b) the S17 comparison-table cell data.

**Step B — Auto-imagery (no user prompt):** `set_content_slides` inserts one image slide per ~4 content slides, picking the best brand-curated match from the OneDrive catalog. Full rules (cadence, source path, matching, fallbacks, sparse-text promotion) live in the **S9 / S10 / S11** polish rule below — do not duplicate them here. User-supplied images dropped in chat are handled via Q5 below.

**Q2d — Section break mode (only if the outline has ≥1 section divider):** "How would you like the section break slides styled?" with options:
1. **Light mode (S3)** — clean white-on-light treatment
2. **Dark mode (S4)** — high-contrast white-on-dark treatment
3. **Alternate** — alternate between light and dark for visual rhythm (first divider = light, second = dark, …)

S3 and S4 are visually consistent in structure (title + optional supporting text) — they differ only in colour treatment. Apply the chosen mode to every section divider in the deck. If "Alternate" is chosen, track parity across the build and flip mode at each divider.

**Q2e — Content bullet mode:** "How should the subhead and body text on content slides be styled?" with options:
1. **Default — text and body** *(recommended)* — Level 1 (subhead) and Level 2 (body) are plain text with no bullets. Bullets only appear at Levels 3–5.
2. **Text with bullets** — Level 1 stays plain text; Level 2 gets a `•` bullet (so body text reads as a bulleted list under the subhead).
3. **Bullets with bullets** *(not recommended)* — Both Level 1 and Level 2 get `•` bullets. Effectively everything in the body is a bulleted list. Rarely matches the M1 voice.

This choice is global — apply to every S7 content slide in the deck. Pass the chosen mode to `set_content_bullets(slide, ph_idx, items, bullet_mode=...)`. Levels 3–5 always carry their template bullets regardless of mode.

**Q3 — Speaker notes:** "Do you want speaker notes generated for each slide?" with options Yes / No. If yes, add notes to each slide's notes placeholder with talking points that expand on the slide content — what to say, not what's on screen. Write notes in conversational tone, 2-4 bullet points per slide.

**Q4 — File name and location:** Suggest a default name: `[Topic]-Presentation-[YYYY-MM-DD].pptx` on the Desktop. "Where should I save the file?" with options "Use default" / "Let me specify". If they choose default, proceed with the suggested name. If they choose to specify, ask for the path as free text.

**Q5 — Images (only if the content references image tags):** If the source content contains any image tag — `<img>`, `<img anything>`, `<image>`, `(image)`, `(img)`, or any similar pattern the user writes — treat it as a request for an image on that slide. Accept whatever the user wrote; don't require a specific format. For tags that include a filename or label (e.g. `<img hero-screenshot>`), ask the user to drop the file into chat and match by filename similarity. For bare tags with no filename (e.g. `<img>` or `<img >`), route that slide through the Q2c photo sourcing flow instead — offer the 4-option picker (Claude's pick / OneDrive / Unsplash / Skip). Show a match table for any named-file tags. Proceed directly to building if every tag is resolved; only ask for confirmation when a match is ambiguous.

**Review — Slide plan:** After all setup questions are answered, show the user a slide plan table before building. The table must include: slide number, slide type (e.g., S1 Cover, S4 Section break, S7 Content, Content + Image), title, a brief content summary, and a **Visual** column showing what graphic treatment was assigned (e.g., `bar chart (auto)`, `photo — Claude's pick`, `timeline (auto)`, `—` for text-only). Ask: "Does this slide plan look right?" with options "Looks good" / "I have changes". If the user has changes, adjust the plan and re-present. Only after approval, mark the **Review** task as `completed`, mark the **Build deck** task as `in_progress`, and proceed to the build steps below.

**Slide feedback format:** ALL slide-related feedback — slide plans, build summaries, revision notes, delivery recaps — must be presented in **table format**. Never use plain bullet lists or numbered lists to describe slides. Always use a markdown table with columns for slide number, type, title, and relevant details.

**Section divider supporting text (S3 / S4):** The master no longer fills `ph_idx=11` with default text. If the outline provides supporting text for a section, populate `ph_idx=11`. If no supporting text is given, **delete the `ph_idx=11` shape entirely** (don't just set it to empty) — leaving an empty placeholder is annoying to the user. See the S3/S4 polish rule below for the helper.

### Expected outline format

The outline must follow this hierarchical structure. Each level maps directly to a slide type:

```
Title: Presentation title here

I. Section name
    A. Slide title
        1. Text entry or talking point
        2. Another text entry
            - Bullet item
            - Another bullet
            • Also valid bullet syntax
    B. Another slide title
        1. Text entry
II. Next section
    A. Slide title
        1. Text entry
```

#### Level-to-slide mapping

| Outline level  | Example                     | Maps to             | Slide type                                                                                                                     |
| -------------- | --------------------------- | ------------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| `Title:`       | `Title: AI enablement week` | Deck title          | S1 (Cover) — appears once at the top of the outline                                                                            |
| Roman numeral  | `I.`, `II.`, `III.`         | Section breaker     | S3 or S4 (Section divider) — creates a breaker slide between sections                                                          |
| Capital letter | `A.`, `B.`, `C.`            | Content slide title | S7 (Content Text) or best-fit slide type based on content                                                                      |
| Arabic numeral | `1.`, `2.`, `3.`            | Slide text entries  | Text lines or top-level bullets on the content slide                                                                           |
| Dash or bullet | `-`, `•`                    | Bullet items        | Indented sub-bullets on the content slide at `lvl=2`. **Always use `set_ph_grouped_bullets()`** — groups Arabic numeral items as parents with their dash/bullet children. Master handles bullet char, teal color, and indent. |

#### Rules

- **One title only:** The outline has exactly one `Title:` line at the top. This becomes the S1 cover slide.
- **Roman numerals = breaker slides:** Every `I.`, `II.`, `III.` etc. produces a section divider slide (S3 or S4). These are visual breakers, not content slides. Use S3 when the Roman numeral has sub-text that provides context; use S4 for bold/minimal breakers with just the section name.
- **Capital letters = content slides:** Every `A.`, `B.`, `C.` under a Roman numeral becomes a content slide. The letter text is the slide title.
- **Arabic numerals = text on the slide:** `1.`, `2.`, `3.` entries under a capital letter are the text content for that slide. They become bullet points or text blocks.
- **Dashes/bullets = grouped sub-bullets:** `-` or `•` items nested under an Arabic numeral become indented sub-bullets on the same slide. **You must use `set_ph_grouped_bullets(slide, ph_idx, groups)` for any slide that has dash/bullet children.** Pass Arabic numeral items as parent strings and their dash/bullet children as the child list. This renders parents at `lvl=0` (cloned from template) and children at `lvl=2` with master-inherited bullet character, teal color, and indent — plus automatic spacing between groups. Never use `set_ph_bullets()` on slides with mixed parent/child content.
- **Arabic numerals directly under a Roman numeral (no capital letter):** When `1.`, `2.`, `3.` items appear directly under a Roman numeral without an intervening `A.`, `B.` level, create a single content slide using the Roman numeral's name as the slide title. The arabic numerals become the bullet content on that slide. This is a shorthand for sections with only one content slide.
- **Agenda is auto-generated:** The agenda slide (S2) is built automatically from the Roman numeral sections — do not expect it in the outline.
- **Closing is auto-generated:** The closing slide (S18) is always added as the last slide — do not expect it in the outline.
- **Image tags** (`<img>`, `<image>`, `(image)`, `(img)`) can appear at any level and are collected during Q5. Tags support an optional sizing mode: `<img name, fit>` (default — full image, no cropping) or `<img name, crop>` (fill placeholder, crop edges).

#### Format validation

After receiving content (Q2), parse it against this structure. If the outline does not match, **do not proceed**. Instead:

1. Show the user what you found (e.g., "I see flat bullet points without the Roman numeral / capital letter hierarchy").
2. Explain the expected format with a short example.
3. Ask them to revise the outline or confirm they'd like you to interpret it as-is and propose a mapping for their approval.

This prevents mismatched slides and ensures breaker slides are used correctly.

### Prerequisites (silent — do not ask)

These are handled automatically without prompting the user:

1. **python-pptx + Pillow:** Ensure both libraries are available. Run `pip3 install python-pptx Pillow` if needed. (Pillow is required by the `insert_picture()` helper for image dimension calculations.)
2. **Template:** The source template is `assets/M1-Presentation-Template_2026.pptx`. Never modify this file.
3. **Copy the template** to the user's chosen location from Q4.

## Graphics vs Charts

A slide is most useful as a visual when it conveys either a **small structural idea** ("there are 3 phases", "the four quarters of this year") OR an **actual data set**. These are two distinct use cases — different template slides for each:

- **Graphic** — used when a slide's content centers on a **SMALL GUIDING NUMBER** with no per-item numeric data. Examples: "Three pillars of stewardship", "Four quarters of the roadmap", "Five layers of the architecture". The visual communicates the count + structure; the words on the slide describe what each item IS, not what its value is.
- **Chart** — used when the slide carries a **SPREADSHEET'S WORTH** of values. A percentage breakdown that adds to 100% → pie. A multi-row table of metrics → table. A bar/line chart for a series of numeric values → bar/line.

The template provides three graphics and three charts:

| Slide | Category | Template title         | Best for (from the speaker note on the template slide) |
|-------|----------|------------------------|--------------------------------------------------------|
| S12   | Chart    | Bar Chart Sample       | Spreadsheet data that has groupings |
| S13   | Chart    | Pie Chart Sample       | A single category of items that add up to 100% |
| S17   | Chart    | Table Sample           | Larger spreadsheet data where presenting a large amount of data is the focus |
| S14   | Graphic  | Timeline Graphic Sample | Timelines or things that go into a clear sequential order |
| S15   | Graphic  | Target Graphic Sample  | Concepts that live inside one another or build upon one another |
| S16   | Graphic  | Takeaways Graphic Sample | Summaries that have clear categories and can have key takeaways |

Each template slide's speaker note (`prs.slides[N].notes_slide.notes_text_frame.text`) is the **authoritative "best for X" guidance** when picking a slide type — read it during the pre-analysis pass (below) whenever the choice is ambiguous.

---

### Step 3 — Map content to slides

The template has 19 example slides. Each is a reusable stamp. For each section of content, pick the best-fit slide and clone it.

**Pre-analysis pass (do this BEFORE rendering anything):** scan every L1 subsection of the outline (every `A.` heading) and decide its slide type up front. For each subsection, read the title and the full body. If the content centers on:

- A **small guiding number** ("3 phases", "4 layers", "5 quarters") with no per-item numeric data → **Graphic** (S14 timeline / S15 target / S16 takeaways)
- **Spreadsheet-style values** (per-item %, per-item counts, table of metrics) → **Chart** (S12 bar/line / S13 pie / S17 table)
- A **time-sequence with actual per-period numeric values** ("Q1 2024 $5M, Q2 $6M, …") → **Chart** S12 bar/line (numbers per step → chart, not graphic)
- Just a **time-sequence with no per-step numbers** ("Q1: design, Q2: build, Q3: launch") → **Graphic** S14 timeline
- Anything else → plain S7 or S8 content (single- or multi-line title)

When the choice is genuinely ambiguous, read each candidate template slide's speaker note via `prs.slides[N].notes_slide.notes_text_frame.text` — those notes were written specifically to answer "is this slide a fit?".

| Content need                    | Template slide | Layout          | When to use                                                                                                                                                                                                               |
| ------------------------------- | -------------- | --------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Cover / title page              | S1             | Title Slide     | Always first. Title + optional subtitle.                                                                                                                                                                                  |
| Agenda or overview              | S2             | Agenda          | Half-image + title + bullet list of topics.                                                                                                                                                                               |
| Section divider — light mode    | S3             | 1_Title Slide   | Light-mode section break. Title + optional supporting text. Used when Q2d = Light, or every other divider when Q2d = Alternate.                                                                                           |
| Section divider — dark mode     | S4             | SectionBreak    | Dark-mode section break. Title + optional supporting text. Used when Q2d = Dark, or every other divider when Q2d = Alternate. Same structure as S3 — only colour treatment differs.                                       |
| Key statement or thesis         | S5             | Statement Slide | Two-column: bold statement left, elaboration right. For "big idea" moments.                                                                                                                                               |
| Quote                           | S6             | 1_SectionBreak (used for quotes) | A quote slide. Only use when the content is an actual quote, or ask the user if they'd like to include a quote that supports their point. Never use as a generic section divider. Use `set_quote()` — see Polish rules.   |
| Bullet-point content (1-line title)  | S7        | Content Text    | Title + structured bullets with indent levels (1-5). The workhorse slide. Use when the title fits on **one line** at 42pt.                                                                                                |
| Bullet-point content (multi-line title) | S8     | Content Text    | Same body structure as S7 but title font drops to 36pt and the body shifts down to accommodate a **2-line title**. Use `pick_content_slide_index(title)` to auto-select S7 vs S8.                                          |
| Content + image (right, 1-line title)   | S9     | Content + Image | Auto-injected by `set_content_slides` for ~1 in 4 content pages (default image variant). Body width ≈ 420pt — narrower than S7.                                                                                            |
| Content + image (left, 1-line title)    | S10    | Content + Image | Auto-injected for every 3rd image slide (visual rhythm). Body width ≈ 439pt. Same content structure as S9.                                                                                                                |
| Content + image (multi-line title)      | S11    | Content + Image | Used in place of S9 / S10 when the title needs to wrap (36pt, image right). No image-left multi-line variant exists in the template.                                                                                      |
| Chart (bar, line)               | S12            | Content-Other   | Single chart with title. Map real data to the chart object.                                                                                                                                                               |
| Pie chart with breakdown        | S13            | Content 2 col   | Pie chart left + component legend right (up to 6 items).                                                                                                                                                                  |
| Timeline or milestones          | S14            | Content-Other   | Sequential events across a horizontal timeline (up to 5 points).                                                                                                                                                          |
| Conceptual hierarchy            | S15            | Content-Other   | Concentric circles (4 layers) + descriptions. For nested relationships or priorities.                                                                                                                                     |
| Multi-topic overview            | S16            | Content-Other   | Topic grid: 3 main topics with sub-sections. For capabilities, pillars, or comparisons.                                                                                                                                   |
| Data table                      | S17            | Content-Other   | Full table. For structured comparisons or detailed data.                                                                                                                                                                  |
| Closing / thank you             | S18            | Blank           | Always last. "Thank you" or closing message.                                                                                                                                                                              |

**Rules:**

- **Follow the outline hierarchy:** When the content uses the expected outline format, mapping is deterministic — `Title:` → S1, Roman numerals → S3/S4 breaker slides, capital letters → content slides (S7/S8 chosen by title length, or auto-promoted to S9/S10/S11 by the imagery rule), agenda auto-generated from Roman numerals → S2, closing → S18. Do not skip breaker slides or collapse sections.
- **Content fidelity:** Preserve the source outline's wording and structure as closely as possible. Do not heavily rewrite, summarize, or paraphrase the user's text. Use the outline's own words for bullet points and titles — only adjust for grammar, sentence case, and brand voice consistency (e.g., curly quotes, removing jargon). The outline is the content; your job is to place it on slides, not rewrite it.
- If the same slide type is needed more than once (e.g., three bullet-point content slides), clone it multiple times.
- Only create a new layout from scratch if none of the 15 templates fit. This should be rare.
- Prefer minimal adaptation — use the template slide as-is and fill in content rather than rearranging shapes.
- **Quote slide (S6):** Never use as a filler or section divider. Only clone S6 when the content contains an actual quote. If no quote exists in the source material, ask the user: "Would you like to include a quote that supports this point?" If they accept, search for a relevant quote — check [100+ Most Famous Quotes Of All Time (Updated 2026)](https://wisdomquotes.com/famous-quotes/) first, then fall back to a general web search. The quote must be relevant to the slide's content. If no fitting quote is found, let the user know and assume they'll source their own. If they decline, skip the quote slide entirely.
- **Agenda slide (S2):** Always populate every agenda item from the source content. The agenda should reflect the full structure of the presentation — every major section gets a line item.
  - **Numbered list (ordered):** Agenda items must be an ordered list (1, 2, 3…), not bullets. Use `<a:buAutoNum type="arabicPeriod"/>` in `<a:pPr>` instead of `<a:buChar>`. Each major section of the presentation gets a numbered item.
  - **Subsections (not default):** Do NOT add sub-bullets unless the source material explicitly contains sub-items under a section. When sub-items are present, add them as bullet items (unordered, `<a:buChar char="\u2022"/>`) indented under the numbered parent (`lvl="1"`). Only if the text fits the placeholder box.
  - **Image placeholder (`ph_idx=10`):** Do NOT touch this placeholder during the build. Leave the default image exactly as it is in the template. Only replace the image after the deck is built AND the user has provided a replacement file. After delivering the deck, ask: "The agenda uses a default image — would you like to choose a different one?" If yes, instruct the user: "Drag an image into this chat from one of these sources and I'll swap it in:"
    1. **OneDrive Images:** `~/Library/CloudStorage/OneDrive-M1Finance/_Creative Assets/_Creative Asset Hub/Images-ArticlesStockImages/in/img-OrigStockDownloads/`
    2. **Unsplash:** [https://unsplash.com](https://unsplash.com) — download the image first, then drag the file into chat
  - Before suggesting the OneDrive path, silently check whether the directory exists (e.g., `os.path.isdir()` or equivalent). If it does not exist, skip the OneDrive option entirely and tell the user: "It looks like you're not connected to OneDrive right now. You can grab an image from Unsplash instead:" then link to [https://unsplash.com](https://unsplash.com) and instruct them to download and drag the file into chat.
  - Only after receiving the image file from the user, use `placeholder.insert_picture(image_path)` to replace the agenda image in the saved deck.
  - **Body placeholder (`ph_idx=1`) — empty paragraph fix:** The agenda template's body placeholder starts with an empty paragraph (no runs). When adding items, do NOT clone the empty first paragraph. Instead, clear the `<a:txBody>` and build new `<a:p>` elements with explicit `<a:r>` and `<a:t>` children. Copy the `<a:pPr>` (paragraph properties) and `<a:rPr>` (run properties) from the layout's default if available, or set font to Inter at 18pt minimum. This prevents silent failures where paragraphs are added but contain no visible text.
- **Image references in content:** Treat any image-like tag as a request for a visual on that slide — `<img>`, `<img anything>`, `<image>`, `(image)`, `(img)`, or any similar pattern the user writes. Do NOT trigger on the bare word "image" in normal sentence text. When a tag is detected, always use the **Content + Image** layout (not S7) for that slide.
  - **Tag with filename** (e.g. `<img hero-screenshot>`, `<img hero-screenshot, crop>`): collect via Q5 — user drops the file and it's matched by filename similarity. Supports optional sizing mode after a comma: `fit` (default, full image visible, no cropping) or `crop` (fill placeholder, crop edges). Build with `add_slide_from_layout(prs, "Content + Image")`, fill `ph_idx=0` + `ph_idx=1`, insert via `insert_picture(slide, 12, path, mode)`.
  - **Bare tag** (e.g. `<img>`, `<img >`): route through the Q2c photo sourcing flow — offer the 4-option picker (Claude's pick / OneDrive / Unsplash / Skip).
  - **Fallback:** If no image is provided for any tagged slide, build it as S7 text-only and note it in the output.

### Step 4 — Build the deck

Using `python-pptx` with package-level slide cloning:

#### Helper functions

All slide cloning, deletion, text injection, and cleanup functions live in `/scripts/pptx_helpers.py`. Read that file and use its functions directly. Key functions:

**`set_ph_grouped_bullets(slide, ph_idx, groups)`** — **Required for any slide with dash/bullet children.** Pass a list where each item is a plain string (no children) or a `(parent_text, [child_text, ...])` tuple.

Operates in two passes:
1. **Style pass** — clones template paragraphs at the correct level so master slide formatting is preserved exactly (no manual overrides):
   - Parents → clone `lvl=0` template paragraph (navy bold subhead, no bullet)
   - Children → clone `lvl=2` template paragraph (teal 14pt, master bullet char + indent)
2. **Spacing pass** — applies `spcBef` on each group's parent after the first.

**Critical level note** — the Content Text layout explicitly defines:
- `lvl=0`: subhead (navy bold, `buNone`)
- `lvl=1`: secondary text (teal 16pt, `buNone` — **no bullet**, do NOT use for bulleted children)
- `lvl=2`: bullet level (teal 14pt, inherits bullet char + indent from master) ← use this for children

Import: `from pptx_helpers import set_ph_grouped_bullets`.

> **Important:** When iterating `slide.placeholders`, never access `.text` directly — `PlaceholderPicture` objects (e.g. the agenda image at `ph_idx=10`) do not have a `.text` attribute and will raise `AttributeError`. Always use `get_ph_text(ph)` from `pptx_helpers` instead, which returns a safe fallback string for picture placeholders.

---

### Step 5 — Verify the rendered output (MANDATORY)

**Do NOT skip this step.** Every text-fitting slide type must be visually verified against PowerPoint's actual render — not against the algorithm's predicted layout. Verification uses `scripts/render_for_review.py` (driven by macOS Quick Look's `qlmanage`).

**Which slides to verify:**

- **Always** — cover (S1)
- **If the deck has section breaks** — every S3 / S4 section divider (titles can exceed the narrower S3 placeholder width)
- **Always** — every S6 quote (multi-line layout most vulnerable to per-line wrap drift)
- **If the deck has multi-line content titles** — every S8 / S11 slide

The build script (`scripts/build_demo.py` or equivalent) auto-renders these slides after `prs.save()` and prints one path per slide:

```
Verify slide1 (Title Slide): /…/testing/_review/M1-digital-private-bank.pptx.png
Verify slide3 (1_Title Slide): /…/testing/_review/slide-3.png
Verify slide14 (1_SectionBreak): /…/testing/_review/slide-14.png
```

**The required action:** read each PNG with the Read tool and confirm:

- Title line count matches the algorithm's prediction
- No widows — no single word stranded as the LAST line of a paragraph. (Short opener like `M1:` is OK; short closer is NOT.)
- No overflow past the footer line

Use Quick Look **only** for layout: line count, wrapping, widows, and overflow.
Do **not** judge the *font* from the QL render (see the limitation note below).

**If anything is wrong:** identify which constant to adjust:
- Cover / section / content-title widows → bump `TITLE_SAFETY_PT` in `pptx_helpers.py`
- Quote widows → bump `QUOTE_SAFETY_PT` in `pptx_helpers.py`

Rebuild, re-render, re-read. Iterate until every flagged slide is clean. Only then mark the Verify task `completed` and deliver.

**Why this step is non-optional:** fontTools-measured `Inter Variable Thin` widths drift from PowerPoint's actual render by enough to produce widows the algorithm can't predict. The qlmanage render is the source of truth **for layout** — it's the wrapping the user's PowerPoint will display. Multiple prior cycles delivered "fixed" decks that still had widows; the only way to break that pattern is to render and look before claiming done.

**Known qlmanage limitation — variable-font named instances:** macOS Quick Look does NOT reliably render the `Inter Variable *` named instances. It can produce inconsistent **serif** renders even when the font resolves fine in PowerPoint — so a serif look in a QL PNG is *not by itself* proof of a problem. But it is also not proof of correctness: a serif render can equally mean PowerPoint will genuinely fail to find the name. **Do not judge the font from QL either way.** Instead validate font-NAME resolution directly: (1) confirm the deck's run typefaces are only the four brand names (the `BRAND_FONT_*` constants); (2) confirm each of those strings matches an installed **Full Name** via `system_profiler SPFontsDataType | grep -i "inter variable"` (the installed Full Names are `Inter Variable Thin`, `Inter Variable`, `Inter Variable SemiBold`, `Inter Variable Medium` — if a requested name isn't in that list, PowerPoint will serif-fallback); (3) open the real `.pptx` in actual PowerPoint — the authority for how fonts render. Use QL for layout/wrapping/widows only.

**Known qlmanage limitation — SVG graphics:** macOS Quick Look does NOT render SVG content embedded inside .pptx files. Any slide that uses SVG decorations (currently the S6 quote — its decorative `“` and `”` marks are SVG pictures) will show those graphics as empty rectangles in the verification PNG. **This is a Quick Look limitation, not a build bug.** Inspect the underlying .pptx XML (`<p:pic>` elements + `ppt/media/*.svg`) to confirm SVGs are embedded, then open the file in actual PowerPoint to confirm the decorative graphics render correctly. qlmanage still catches title/text wrapping and widow issues on these slides — only the SVG layer is invisible.

**Helper invocations:**
```python
from render_for_review import render_cover_to_png, render_slide_to_png

# Cover (slide 1) — fast path, qlmanage already renders slide 1
cover_png = render_cover_to_png(out_pptx_path)

# Any other slide — temp-pptx with only the target slide, then qlmanage
quote_png = render_slide_to_png(out_pptx_path, slide_index=13)
```

Print every PNG path so the model can `Read` them in sequence before declaring done.

---

## Polish rules

Polish rules are applied after content is injected into each slide. They correct layout, sizing, and typographic issues that the raw injection does not handle. Each rule is permanent — apply it every time you build that slide or element type.

---

### Global — visual self-verification (MANDATORY)

See **Step 5 — Verify the rendered output** above for the full verification loop (which slides, what to check, helper invocations, SVG caveat). This rule is the same content stated as a permanent polish requirement: every build MUST be visually verified, no exceptions.

**Tunable constant:** `TITLE_SAFETY_PT` in `pptx_helpers.py` is the cover/section/content-title width-margin in points. Empirically tuned ("PowerPoint demanded this much slack"), not theoretical. Re-tune if Inter Variable Thin metrics, template placeholder dimensions, or title font weight changes.

---

### Global — titles always use Inter Variable Thin

**Rule:** Every slide title — cover (S1), section break (S3/S4), quote (S6), and content (S7/S8) — is rendered in **Inter Variable Thin**, no exceptions. The brand voice + visual language requires this consistency. The polish helpers (`set_cover_title`, `set_section_break`, `set_quote`, `set_content_title`) all stamp `typeface="Inter Variable Thin"` on the title runs.

If you find yourself calling `set_ph()` directly on a title placeholder, that's the wrong helper — switch to the slide-specific title helper instead, or the title will inherit the master's typeface (Inter Variable) and look heavier than the rest of the deck.

---

### Global — every multi-line title gets the EXACT S8 treatment

**Rule:** Whenever a slide's title runs to more than one line, three things MUST happen — every one of them, no exceptions, regardless of slide type (S7, S8, S9, S10, S11, S12, S13, S14, S15, S16, S17):

| Property | Single-line title (S7 path) | **Multi-line title (S8 path)** |
|---|---|---|
| **Title font size** | 42pt (layout default) | **36pt** (sz="3600") |
| **Title placeholder y-position** | inherits layout (typically 0.000") | **0.269"** (offset y = **246165 EMU** — the exact value extracted from the S8 template) |
| **bodyPr anchor** | bottom (`anchor="b"`) | bottom (`anchor="b"`) |
| **Typeface** | Inter Variable Thin | Inter Variable Thin |
| **Line-breaking** | n/a (1 line) | phrase-aware (NATURAL_PUNCT / BREAK_BEFORE), capitalize after INTRO_PUNCT, **NO widows** |

The y-shift is the **critical** part that's easy to miss. Without it, a 2-line title at 36pt on an S7-cloned slide (which has the title placeholder at y=0) renders the top line right at the slide edge — visually crowded. The S8 template solves this by placing its title at y=0.269", giving 2 lines of 36pt Inter Variable Thin breathing room above the body anchor.

**`set_content_title()` is the canonical implementation.** It checks if the title fits one line at 42pt; if so, leave the placeholder at its default y. If multi-line: shift y to `246165 EMU` (constant `S8_TITLE_Y_OFFSET_EMU`) AND render at 36pt with `_layout_phrase_lines(min_pt=28, max_pt=36)`. The 28pt floor is only hit by exceptionally long titles — most 2-line titles land at exactly 36pt.

**How it's enforced:** every title placeholder must be populated via one of these polish helpers — never bare `set_ph(slide, 0, title)`:

| Slide type | Title helper | Title placeholder width | Default → multi-line size |
|---|---|---|---|
| S1 cover | `set_cover_title(slide, title)` | 741pt | 72pt → 28pt (sized to fit, 2-line preferred) |
| S3 / S4 section break | `set_section_break(slide, title, supporting_text=None)` | 594pt / 809pt | 42pt → 24pt |
| S6 quote | `set_quote(slide, quote, attribution)` | 508pt (quote-specific layout) | n/a (uses `_layout_quote_lines`) |
| S7 / S8 content (text)  | `set_content_title(slide, title)` | 864pt | 42pt → 36pt → 28pt |
| S9 / S10 / S11 image-content | `set_content_title(slide, title)` via `set_content_slides`/`set_content_series` | 864pt | 42pt → 36pt → 28pt |
| S12 bar / S13 pie / S17 table — **charts** | `set_content_title(slide, title)` | 864pt | 42pt → 36pt → 28pt |
| S14 timeline / S15 target / S16 takeaways — **graphics** | `set_content_title(slide, title)` (invoked inside `set_s14_timeline` / `set_s15_target` / `set_s16_takeaways`) | 864pt | 42pt → 36pt → 28pt |

The chart/graphic helpers (`set_s14_timeline`, `set_s15_target`, `set_s16_takeaways`) all delegate the title to `set_content_title` internally. The result: a multi-line title on a chart or graphic slide gets the exact same wrap, sizing, and capitalization as a multi-line title on S8.

**Why it matters:** without this rule, a chart or graphic slide with a long title renders at 42pt with PowerPoint's default wrap — title overflows the top of the slide. S8 already solved this problem for content slides; the rule extends that solution everywhere a title placeholder lives.

---

### Global — avoid consecutive same-format slides

**Rule:** Don't use the same chart/graphic slide-type for two adjacent slides — UNLESS the second is an explicit continuation of the first (e.g., a content overflow with a `" (continued)"` title). Repeated format = visual monotony; the reader stops noticing what the format is supposed to communicate.

**How to apply** during the Q2c pre-analysis pass: walk the planned slide order. If two consecutive slides both want the same chart/graphic template (e.g. both S16 takeaways, or both S14 timeline), demote the WEAKER candidate to plain content (S7/S8) and keep the STRONGER candidate as the chart/graphic. "Stronger" means the slide whose content most cleanly matches the template's "best for X" speaker note. The demoted slide still communicates the content as a regular bulleted slide — just without the visual template.

**Continuation exception:** when `set_content_series` or `set_content_slides` produces a multi-page overflow with a `" (continued)"` title, those consecutive same-layout slides are fine — the reader understands they're reading one logical section split across pages. The continuation title makes the relationship explicit.

**Example from this deck's build:** test-01.md has 4 subsections that structurally fit S16 takeaways (II.A relationship model, II.B three pillars, V.B rates, V.C what others say). Naively routing all 4 to S16 produces two adjacent-S16 pairs. The build script splits them: II.A and V.C demote to plain content; II.B and V.B keep the S16 treatment because their content is the canonical "guiding number with key takeaways" use case.

---

### Global — never overflow the footer

**Rule:** Content must never run into the slide's footer area. Leave at least **0.2"** of breathing room between the bottom of body content and the footer line. Any content that would push past that boundary spills onto a new continuation slide.

**Continuation slide title:** `"{original title} (continued)"`. Always just that — no page counter, no "(continued) (continued)" stacking.

**L1 grouping is sacred — never split an L1 across pages.** If a Level 1 subhead's descendants (L2 body, L3 bullets, L4/L5 sub-bullets) would overflow the page, the **entire L1 group** moves to the next slide together. Don't leave the L1 stranded on one slide with its descendants on the next.

**Continuation L1 spacing:** The first L1 on a continuation slide always gets the **+18pt L1↔L1 spacing** as if a prior L1 had been on the same page — preserves the visual rhythm of the outline across the page boundary.

**Last-resort split:** If a single L1 group is itself taller than one page (very rare with normal content), the group is split anyway — the rule is broken rather than refuse to build. A warning should be surfaced in this case so the source content can be shortened.

**How it's enforced:** `set_content_slides(prs, title, items, bullet_mode)` in `pptx_helpers.py` is the only correct entry point for building content slides — it measures every paragraph using per-level Inter font metrics, paginates by L1-group boundary, and stamps continuation titles automatically. Calling `set_content_bullets` directly bypasses pagination and risks footer overflow.

---

### Global — no widows (typographic definition)

**Rule:** In any multi-line text layout (cover title, section break title, content title, quote, etc.), **never leave a single word stranded as the FINAL line** of the paragraph. A short *closing* line is a widow and is visually undesirable.

A short *opening* line is fine. Titles like `"M1: the digital private bank"` are expected to render as:

```
M1:                         ← single-word opener — OK
The digital private bank
```

That short first line is an intentional pivot (often after a colon or em-dash) — it's not a widow.

**How it's enforced:** Both layout helpers (`_layout_phrase_lines` and `_layout_quote_lines` in `pptx_helpers.py`) reject any candidate split where any line **after the first** contains fewer than 2 words. The first line is never widow-checked. If no widow-free split fits, the algorithm sizes down or tries a different line count until a valid layout is found. The greedy fallback (used only as the absolute last resort) is the one place this rule may be violated — if you see a widow in output, the text is genuinely too long for the placeholder and the source content should be shortened or the placeholder enlarged.

---

### S1 — cover slide (title)

**Rule:** Never inject the cover title with a plain `set_ph()`. Always use `set_cover_title()` from `pptx_helpers`.

**Font:** Always Inter Variable Thin (`typeface="Inter Variable Thin"`). No exceptions.

**Algorithm — 4-phase priority:**
1. **Natural 2-line (preferred):** Sweep 72→42pt in 2pt steps. At each size, find the best split at a natural break point (BREAK_BEFORE word or after NATURAL_PUNCT). Stop at the largest font where a natural 2-line split fits. This is the most common result.
2. **Any 2-line:** Same sweep, but accept non-natural breaks. Used only when the title has no prepositions or punctuation to break on.
3. **3-line fallback:** Same sweep, up to 3 lines. Only reached if no 2-line split exists at any size.
4. **Greedy fallback at 42pt:** Absolute floor.

**Key behaviours:**
- Prefers 2 lines at a smaller font over an awkward break at a larger font
- Capitalizes the first letter of a new line **only** when the preceding line ends with introductory punctuation (`:`, `—`, `–`, `;`) — not on prepositional or conjunctive breaks
- Never goes below 42pt

**Implementation:**
```python
from pptx_helpers import set_cover_title
size, lines = set_cover_title(slide, title)
print(f"Cover  {size}pt  {lines}")
```

**Examples:**
- `"M1: the digital private bank"` → `["M1:", "the digital private bank"]` at 72pt
- `"M1 Intelligence: your always-on financial advisor"` → `["M1 Intelligence:", "Your always-on financial advisor"]` at 50pt (natural break after colon, second line capitalised)
- `"Built for investors who think in decades"` → `["Built for investors", "who think in decades"]` at 72pt (break before relative pronoun "who", no cap)
- `"Building wealth for the long term"` → `["Building wealth", "for the long term"]` at 72pt (no cap — break is prepositional)

---

### S3 / S4 — section break slides

**Rule:** Never inject section-break titles with plain `set_ph()`. Always use `set_section_break()` from `pptx_helpers`. Same helper handles both S3 (light mode) and S4 (dark mode) — the only difference is which template slide you clone (`prs.slides[2]` for S3, `prs.slides[3]` for S4), chosen per the user's Q2d answer.

**Font:** Always Inter Variable Thin. Phrase-aware breaks, fontTools-measured widths, **height-constrained** layout (title placeholder is only ~76pt tall, so 2-line layouts above 30pt are rejected).

**Supporting text (the key polish rule):** The supporting-text shape (`ph_idx=11` on S3, demoted to a regular shape on S4) is automatically handled:
- If `supporting_text` arg is provided and non-empty → populate the shape
- If `supporting_text` is omitted or empty → **the shape is deleted entirely**. Setting it to `""` would leave the default "Optional text…" showing through; full deletion is the only way to suppress it.

**Why deletion (not clearing):** The master no longer fills `ph_idx=11` with default text, but the slide-level shape on each cloned slide still has it. Only removing the shape removes the placeholder text — empty string injection still inherits the default.

**Implementation:**
```python
from pptx_helpers import set_section_break

# No supporting text — shape gets deleted
set_section_break(slide, "Why personal finance is broken")

# With supporting text — shape gets populated
set_section_break(slide, "M1 is the third option",
                  "An integrated platform with intelligent automation — "
                  "what private banks provide, for everyone.")
```

**Examples:**
- `"Why personal finance is broken"` on S3 → 1 line at 40pt (drops from 42pt because S3 title width is only 594pt)
- `"Platform traction"` on S4 → 1 line at 42pt (S4 title width is 809pt — fits comfortably)
- `"The M1 advantage"` on S4 + supporting text → title at 42pt, supporting text populated at 16pt (template default)

**Known caveat:** S3 and S4 title placeholders have different widths (594pt vs 809pt), so the same long title may render 1-line on S4 and 2-line on S3. Visual consistency between light/dark modes requires aligning these widths in the template (out of scope for the polish helper).

---

### S7 / S8 — content slide (5-level body)

**Rule:** Use `set_content_slides(prs, title, items, bullet_mode=...)` for any content slide — it is the only function that handles pagination, continuation titles, and L1-group preservation correctly. It internally:
1. Picks S7 (single-line title) or S8 (multi-line title) via `pick_content_slide_index`
2. Polishes the title via `set_content_title` (Inter Variable Thin, phrase-aware breaks, no widows)
3. Paginates items by L1 boundary and adds continuation slides as needed
4. Applies L1↔L1 (+18pt) and L2↔L2 (+12pt) group-separator spacing — including the continuation L1 case

Never call `set_ph()` on the title (bypasses Inter Variable Thin + phrase-break). Never call `set_content_bullets` directly unless you're sure the content fits on one slide (it skips pagination).

```python
from pptx_helpers import set_content_slides
slides = set_content_slides(prs, title, items, bullet_mode="default")
# slides is a list — 1 entry if everything fit, 2+ if paginated.
```

**Title behavior:**
- S7 path — title fits in 1 line at 42pt → render at 42pt, 1 line, Inter Variable Thin
- S8 path — title doesn't fit at 42pt → drop to 36pt, run `_layout_phrase_lines` for natural-break layout. Sizes down to 28pt if needed. No widows.

**Body grouping spacing** (applied automatically by `set_content_bullets`):
- Two consecutive L1 (subhead) paragraphs → +18pt space before the second
- Two consecutive L2 (body) paragraphs → +12pt space before the second
- All other transitions inherit the template's default spacing

This visually groups items that go together — paragraphs at the same structural level get a clear gap between them, while content within a group stays tight.

The template defines 5 styled body levels (same on both S7 and S8):

| Outline marker | OOXML lvl | Visual                                |
|----------------|-----------|---------------------------------------|
| `A.`           | 0         | Subhead — navy bold                   |
| `1.`           | 1         | Body text — teal                      |
| `-` (dash)     | 2         | Bullet (`•`) — teal 14pt              |
| nested `-`     | 3         | Sub-bullet (`-`) — light blue         |
| further nested | 4         | Final sub-bullet (`›`) — light blue   |

**Input format:** `items` is a flat list of `(level, text)` tuples where level is 0–4. Multiple consecutive same-level entries are fine.

**Bullet mode (set globally from Q2e):**
- `"default"` — L1 and L2 are plain text; bullets only at L3+ (recommended)
- `"text_with_bullets"` — L2 gets `•` added; L1 stays plain
- `"bullets_with_bullets"` — L1 and L2 both get `•` added (not recommended)

L3–L5 always carry their template bullets regardless of mode.

**Implementation:**
**Full example:**
```python
from pptx_helpers import (
    duplicate_slide, pick_content_slide_index, set_content_title, set_content_bullets,
)

slide = duplicate_slide(prs, pick_content_slide_index("Slide title"))
set_content_title(slide, "Slide title")     # always Inter Variable Thin
set_content_bullets(slide, 1, [
    (0, "First subhead"),
    (1, "Body under the first subhead"),
    (2, "First bullet"),
    (2, "Second bullet"),
    (3, "Sub-bullet"),
    (4, "Final nested item"),
    (0, "Second subhead"),                  # ← gets +18pt space before
    (1, "Body under the second subhead"),
    (1, "Second body line"),                # ← gets +12pt space before
], bullet_mode="default")
```

**How it works internally:** `set_content_bullets` clones the existing template paragraph at each level (preserves all master-inherited styling — font, colour, size, indent). For modes that add bullets to L1/L2, the helper removes any `buNone` inheritance and writes an explicit `<a:buChar char="•"/>` to the paragraph's `pPr`, alongside hanging-indent `marL` and `indent` attributes so the bullet doesn't collide with the text (the layout's `lstStyle` sets `marL=0` on those levels by default).

---

### S9 / S10 / S11 — image + content slides (auto-imagery)

**Rule:** `set_content_slides()` automatically picks ~1 in every 4 content pages to become an image slide. Users do not see a question for this — no `AskUserQuestion` flow.

**Cadence:**
- One image per `IMAGE_SLIDE_CADENCE = 4` content pages (the 4 pages produced by text-only pagination form one "window")
- Within each window, the page with the **highest filename-token overlap** against the OneDrive image library wins the image slot
- Every **3rd** image slide of the deck uses **S10** (image left). Image slides 1, 2, 4, 5, 7, 8, … use **S9** (image right)
- If the chosen page's title needs to wrap to multiple lines, it uses **S11** (image right + 36pt title) since the template has no image-left multi-line variant. The S10 counter holds — S10 is used at the next eligible 3rd image slide

**Image source:** the M1 OneDrive folder at `~/Library/CloudStorage/OneDrive-M1Finance/_Creative Assets/_Creative Asset Hub/Images-ArticlesStockImages/in/img-OrigStockDownloads/` — already brand-curated and licensed. No external network calls. No API keys.

**Matching:** filename tokens (split on `-`, `_`, and camelCase, with stock-service slugs like `iStock`, `shutterstock`, photographer hash slugs, and pure-digit IDs stripped) are scored against the slide's content tokens (title + L1 subheads + L3 bullets, minus stopwords). Score = `|slide_tokens ∩ filename_tokens|`. Highest scorer wins, ties broken alphabetically.

**Fallbacks:**
- No keyword match in a window → designated generic image (`HowWeDoIt-AI-midjourney.png`) is used, and a line is added to the build's imagery warnings
- OneDrive folder unreachable (folder missing, M1_FORCE_NO_ONEDRIVE=1, wrong account) → image slides build with **grey-box placeholders** and a visible "Insert image here" marker per slide. Final build summary lists every slide that needs a user-supplied image so the user knows where to drop files before sharing the deck. (Q1b's pre-build connectivity check should catch this *before* build starts — see the Setup flow.)

**Cadence skips chart/graphic slides:** The 1-in-4 image cadence is computed only over **plain S7/S8 content pages**. Slides that the Q2c pre-analysis pass substituted to a chart (S12/S13/S17) or graphic (S14/S15/S16) are already visual coverage and are NOT eligible for the image-cadence slot. In practice this means: build the chart/graphic substitutions first, then pass only the remaining text-content subsections to `set_content_series` — the imagery loop inside `set_content_series` then sees only S7/S8 pages and the cadence math stays correct.

**Sparse-text promotion — sparse plain content always becomes an image slide:**

A plain S7/S8 page with only a few short bullets leaves obvious empty space underneath the text. Those slides are stronger as image variants — the image fills the empty real estate and gives the bullets a visual partner. `set_content_series` detects this in Phase 2a and promotes sparse pages to S9/S10/S11 **before** the 1-in-4 cadence pass, so every sparse page gets an image regardless of where it falls in the cadence window. The cadence then fills any remaining gaps among dense pages.

A page qualifies as sparse when **all three** are true (constants in `pptx_helpers.py`):
- L1 bullet count ≤ `SPARSE_L1_MAX` (3)
- Total word count across all bullets ≤ `SPARSE_WORD_MAX` (60)
- No L3-or-deeper nested bullets (deep nesting signals density — those slides are visually full already)

If you want a sparse subsection to stay plain text (rare — usually only when the subsection is sandwiched between two image slides and a third image would be too much), bump it above one of the thresholds in the outline or, in the rare case where the rule should be bypassed, call `set_content_slides` directly for that subsection.

**Q2c pre-analysis — number-keyword titles get graphic-first routing:**

When the subsection's TITLE contains a count word ("Two", "Three", "Four", "Five", "Six", or digits 2–10), it is a strong hint that the content fits a graphic shape (S14 timeline, S15 target, or S16 takeaways). During the Q2c pre-analysis pass, check whether the content structurally fits one of those graphics (sequential → S14, nested/concentric → S15, discrete-with-takeaways → S16). If yes, route there. If no graphic fits the structure, fall through to the standard plain-content path — sparse-text promotion will then catch the sparse case and convert it to an image variant. **Graphic first, image fallback.** This rule applies at planning time and the routing decision shows in the slide-plan review table so the user can override.

**Body width caveat:** S9/S11 body width is ~420pt, S10 body width is ~439pt — roughly half what S7/S8 provide. Image-slide content may overflow the narrower body and produce one or more S7/S8 text continuations (with " (continued)" titles), exactly like text-only pagination overflow.

**Implementation:**
```python
from pptx_helpers import set_content_slides, imagery_warnings

slides = set_content_slides(prs, title, items, bullet_mode="default")
# Some of `slides` may be S9 / S10 / S11 (image + content); subsequent
# overflows from those become S7 / S8 continuations automatically.

# After the build, check imagery_warnings for any fallback usage or
# OneDrive accessibility notes to surface in the final build summary.
```

**Deprecated:** `add_slide_from_layout(prs, "Content + Image")` — the old layout-only flow. Now that S9, S10, and S11 are example slides in the template, prefer direct cloning via `duplicate_slide`.

---

### S15 — target graphic (nested concepts)

**Rule:** Use `set_s15_target(slide, title, levels)` from `pptx_helpers`. The helper handles ring/marker/textbox positioning — never set those by hand.

**Input ordering:** `levels` is a list of 2–4 entries ordered **INNERMOST → OUTERMOST**.
- `levels[0]` → innermost ring (Core slot — darkest, smallest)
- `levels[-1]` → outermost ring (Desire slot — lightest, largest)

Each entry is `"Label"` (label only) or `("Label", "Takeaway sentence")`.

**Marker-color matching (CRITICAL):** The right-side subhead textboxes stack top→bottom matching ring colors: TOP marker = OUTERMOST (lightest), BOTTOM marker = INNERMOST (darkest). The helper writes each pair into the textbox whose marker matches its ring. Verify visually: top subhead's marker color must match the largest/lightest ring; bottom subhead's marker must match the smallest/darkest ring. Mismatched markers are a build defect.

**Ring label text fit (auto):** Ring interior labels (Foundation/Growth/Leverage/Legacy in the demo, or Core/Need/Utility/Desire in the template default) are auto-widened to 80% of their oval's width on every call. This prevents long labels like "Foundation" from wrapping inside the small inner oval (the template's default 657pt-wide label box was too narrow for words >7 characters at 15pt Inter). The widening is centered horizontally so the text stays inside the ring.

**Level count behavior:**
- **N=4 (canonical, recommended):** all 4 rings used. The template was designed for this case.
- **N=3 / N=2:** outermost rings + their labels + pair textboxes + marker stripes + ovals are removed. By default the remaining rings **auto-enlarge to fill the original 4-ring footprint** (`enlarge=True`), and the pair textboxes are redistributed evenly between the top and bottom positions so each pair stays vertically near its ring.
- **N>4:** call is clamped to first 4 levels and a `routing_warning` is emitted. The template ships 4 ovals; adding more rings requires manual template editing (clone an oval, add a lighter color, add label+marker+pair textbox). Not currently supported by the helper.
- **N<2:** `routing_warning` emitted (S15 needs at least 2 nested concepts to communicate "build upon").

**Style-by-cloning rule (project-wide):** When the helper needs to insert a NEW textbox into a brand slide (e.g., the summary block on S15), **never construct it from scratch with manual `font.name`, `font.size`, `font.color` settings.** Instead, clone an existing textbox on the same slide that already carries the desired brand style, then change its text. Cloning preserves the master's paragraph properties, run properties, and theme color references — and these stay correct even if the brand template is later re-themed. Manual font/color settings drift away from the template and produce slides that look "almost right but not quite." This rule applies to ANY new text element added to a brand slide, not just S15.

The S15 `summary=` block clones `TextBox 13` (the innermost pair textbox), keeps only the second paragraph (the regular takeaway style — not the bold subhead), sets the summary text, and repositions the cloned box above the rings. Implementation: `_add_s15_summary()` in `pptx_helpers.py`.

**Fill modes for deleted areas:**
- **Default — `enlarge=True`** (the helper's default when `summary` is None): scale up the kept rings so the outermost one matches the original outermost size. Redistribute pair textboxes evenly. The graphic fills the slide naturally — no visible empty space.
- **`summary="text..."`** (optional argument): keep the rings at their template positions (small, clustered at the bottom-left) and add a body-text block above the rings with `summary` content. Use this when the freed-up space should communicate a summary insight rather than a bigger graphic. The helper auto-sets `enlarge=False` in this mode unless the caller explicitly overrides.
- **Both (option c)** — `summary="..."` AND `enlarge=True`: caller asks for both modes. The graphic enlarges AND a summary text block is added. Use sparingly — both elements compete for vertical space.

**Why N=4 is the canonical case:** Each ring has a fixed brand color (dark navy → medium → light → lightest), and the visual is balanced for 4 layers. Removing rings leaves the smaller inner rings clustered at the bottom of the original visual frame; the title and right-side text still align, but the graphic looks less balanced. When the content has exactly 4 concepts that nest, use S15. When the content has 2–3, S15 still works but consider whether S16 (parallel discrete topics) or plain content might communicate better.

**Mapping (inner → outer)** — for reference when debugging:
```
ring labels:    TextBox 49 (Core)        → TextBox 2 (Need)        → TextBox 3 (Utility)    → TextBox 4 (Desire)
pair textboxes: TextBox 13 (bottom/dark) → TextBox 11              → TextBox 9              → TextBox 54 (top/light)
ovals:          Oval 25 (innermost)      → Oval 23                 → Oval 21                → Oval 20 (outermost)
marker stripes: Rounded Rectangle 12     → Rounded Rectangle 10    → Rounded Rectangle 8    → Rounded Rectangle 5
```

**Implementation:**
```python
from pptx_helpers import duplicate_slide, set_s15_target
s = duplicate_slide(prs, 14)  # S15 template index

# Canonical N=4 — no fill-mode args needed
set_s15_target(s, "The wealth journey, layer by layer", [
    ("Foundation", "Emergency cash that earns more than it sits idle"),  # innermost
    ("Growth",     "Invested portfolios that compound across decades"),
    ("Leverage",   "Borrowing power without forced selling"),
    ("Legacy",     "Stewardship that protects what you've built"),       # outermost
])

# N=2 enlarged (default) — rings fill the original 4-ring footprint
set_s15_target(s, "The two-layer trust model", [
    ("Custodial", "Your assets sit with the custodian — never on M1's books"),
    ("Insured",   "SIPC + FDIC coverage applied to the right account types"),
])

# N=2 with summary — rings stay small at bottom, summary body text on top
set_s15_target(s, "The two-layer trust model", [
    ("Custodial", "Your assets sit with the custodian — never on M1's books"),
    ("Insured",   "SIPC + FDIC coverage applied to the right account types"),
], summary="Your money never sits on M1's balance sheet — it stays with regulated custodians.")
```

**Dedicated regression:** `scripts/build_target_test.py` builds `testing/test-target-01.pptx` with N=4, N=3, and N=2 cases. Re-run any time `set_s15_target` is changed; verify each slide's marker-to-ring color match in the rendered PNGs.

---

### S6 — quote slide

**Rule:** Never inject a quote with plain `set_ph()`. Always use `set_quote()` from `pptx_helpers`. The S6 layout name is `1_SectionBreak` (misleadingly named in the template) but it IS the quote layout — quote text goes in `ph_idx=0`, attribution in `ph_idx=11`.

**Template note:** The opening (`“`) and closing (`”`) curly quote marks are decorative shapes positioned by the master — `set_quote()` does NOT include them in the text. Pass plain quote text in.

**Font:** Always Inter Variable Thin. Same fontTools-measured widths as `set_cover_title`.

**Layout algorithm (different from cover titles):** Cover titles prefer fewer lines; quotes prefer the largest font that fills the box. `_layout_quote_lines` iterates 72→32pt and at each size tries n_lines = 2, 3, …, height_max (constrained by box height × line-height factor 1.2). Returns the first valid fit, prioritising naturals.

**Validations applied:**
- **Sentence integrity** — no period/!/? mid-line. Splits like `"my strategy. It executes"` are rejected; the sentence-ending punctuation must be the last character of a line.
- **Balance penalty** — `max(width) − min(width)` is weighted into scoring, so orphan lines like `"Managing"` alone are discouraged.
- Recognises sentence boundaries (`.` `!` `?`), clause boundaries (relative pronouns: who/which/that), prepositions, and conjunctions as natural break points.
- Writes attribution to `ph_idx=11` prefixed with `— `. Pass the attribution without a leading dash; the function adds it.

**Implementation:**
```python
from pptx_helpers import set_quote
size, lines = set_quote(slide, quote_text, attribution)
print(f"Quote  {size}pt  {lines}")
```

**Examples:**
- `"Managing money used to feel overwhelming. With M1, it just runs." / "M1 client"` → 4 lines at 56pt, sentence 1 spans lines 1–3, sentence 2 stands alone on line 4
- `"I tell M1 my strategy. It executes — without me, without emotion." / "Jamie R., M1 client"` → 4 lines at 54pt, sentence 1 alone on line 1; em-dash break capitalises the next line

---

### Verification

To regression-test the polish rules after any change to `pptx_helpers.py`, run `python3 scripts/verify_polish.py` from the repo root. It builds one `.pptx` per case into `testing/` and prints width-based pass/fail. Add new cases by appending to the `CASES` list at the top of the script.
