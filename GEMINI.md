# Medical Flashcard Generation Assistant — Calibrated Style Rules & Workflow

This document persists the calibrated card-drafting guidelines, visual processing protocols, and workflow gates for MBBS flashcard generation.

---

## 1. Calibrated Deck Style (Derived from Active Anki Decks)

Analysis of existing cards from active decks (`MBBS Jassem A.` Year 1 and Year 2 Stage 2, `MBBS Year 1 AK`) establishes the following rules:

### A. Atomicity & Cloze Syntax
- **One Discrete Concept per Card**: Strictly test one fact, enzyme, rate-limiting step, physiological response, histological layer, or clinical sign per card. Avoid run-on multi-cloze cards that test entire paragraphs at once.
- **Cloze Bracket Precision**:
  - Enclose **only** the operative keyword, number, or structure name:
    - ✅ `The outermost muscle membrane is the {{c1::epimysium}}`
    - ❌ `{{c1::The outermost muscle membrane is the epimysium}}`
    - ✅ `CNV trigeminal is a {{c1::both motor and sensory}} nerve`
    - ✅ `Low {{c1::O2}} levels in the blood stimulate the {{c2::kidneys}} to produce {{c3::erythropoetin}}` *(only when testing related steps across cards)*
  - Never cloze fluff words, verbs of state (e.g., "is", "causes"), or articles ("the", "a").
- **Paired / Directional Clozes**:
  - When two items form a single reciprocal fact or direction, combine them under the same cloze identifier:
    - ✅ `DNA Polymerase synthesizes from {{c1::5}}' to {{c1::3}}'`
    - ✅ `PKU prevents conversion of {{c1::Phe}} -> {{c1::Tyr}}`
    - ✅ `In nicotinic receptors, {{c1::2}} Ach molecules bind to {{c1::2}} alpha subunits for activation`
- **In-Bracket Hints & Cues**:
  - Use concise directional or phonetic cues inside brackets when disambiguation prevents rote guessing:
    - `Kinesin transports cargo to the {{c1::(+) periphery (think p for plus and periphery)}}`

### B. Phrasing Conventions & Typography
- **Declarative & Direct**: Formulate concise, high-yield medical assertions rather than vague conversational prose.
- **Emphasis & Contrast**:
  - Use `<b>...</b>` and `<u>...</u>` deliberately to anchor opposing concepts:
    - `More resistances in <b>series</b> results in {{c1::greater}} total resistance`
    - `More resistances in <b>parallel</b> results in {{c1::lesser}} total resistance`
- **Categorization Lists**:
  - For discrete classifications (e.g. 2 amino acid precursors, 3 hormone classes, staging levels), format cleanly with `<ol><li>...</li></ol>`:
    ```html
    The source organ of each form of hypothyroidism is:
    <ol>
      <li>Tertiary -- {{c1::Hypothalamus}}</li>
      <li>Secondary -- {{c1::Pituitary}}</li>
      <li>Primary -- {{c1::Thyroid}}</li>
      <li>Resistance -- {{c1::Peripheral tissue}}</li>
    </ol>
    ```

### C. 'Extra' Field Conventions
The 'Extra' / 'Back Extra' field must **never** duplicate the front cloze text. It is reserved for high-yield value-adds:
1. **Mechanism / Physiological Rationale**:
   - Short explanation of why the fact is true:
     - `<i>(All that matters is vein pressure &lt; artery pressure within a tissue)</i>`
     - `(Diffuses across the PM and binds to intracellular receptors)`
2. **Mnemonics & Memory Hooks**:
   - Highlight phonetics or visual wordplay with bold/underline:
     - `<u><b>Per</b></u>imysium covers the fascicle which is basically a <u><b>pair</b></u> of muscle fibers`
     - `The Villain replaced the Good guy` *(Valine replaces Glutamate in HbS)*
     - `<i><b><u>A</u></b>nterior for <b><u>U</u></b>p-growth for <u><b>O</b></u>ral<br><b><u>P</u></b>osterior for <b><u>D</u></b>own-growth for <b><u>D</u></b>iencephalon</i>`
3. **Clinical Pearls / Consequences**:
   - Clinical context of deficiency/excess:
     - `(High iodine can cause or exacerbate hyperthyroidism via Jod-Basedow)`
4. **NO Lecture or Slide Attribution Tags**:
   - Do NOT append lecture source tags or slide numbers (e.g. NEVER append `[Lecture: ..., Slide ...]`). Keep answers and extra fields focused purely on the medical science.

### D. Note Type & Field Compatibility
- **`1. Basic`** (Primary Q&A Note Type):
  - Fields: `Front`, `Back`
  - Use for simple definitions, direct questions, and categorical lists.
  - Answers must be **simple, direct, and free of unnecessary academic jargon** (e.g. avoid terms like *hyperpharmacology* or *promiscuity*; state the mechanism in plain, punchy English).
- **`Cloze++`** / **`Cloze`** (Cloze Note Types):
  - Fields: `Text`, `Back Extra`
  - Use when testing directional relationships (e.g. 5' to 3'), specific numbers/percentages, or tightly coupled reciprocal facts.
- **Focus on Core Learning Outcomes**:
  - Prioritize the lecture's explicitly stated Learning Outcomes at the start of the slides.
  - Do not over-index on transient case illustrations or example drugs (e.g. propranolol) when testing high-level lecture principles.

### E. Strict Negative Constraints & Slide Filtering
- **Non-Examinable Slide Exclusions**:
  - Automatically skip Slide 1 (Title slide with module codes, lecture numbers, university logos).
  - Automatically skip Intended Learning Objectives (ILOs) slides (used strictly as an internal blueprint, NEVER tested as flashcards).
  - Automatically skip administrative slides, reading lists, textbook recommendations, Mentimeter polls, and closing slides.
- **Strict Content Exclusions (Never create cards for these)**:
  - ❌ **Administrative metadata**: Never ask what the module code is (e.g. 4MBBS102), lecturer name, lecture number, or session title.
  - ❌ **Learning objectives**: Never ask "What are the learning objectives of this lecture?".
  - ❌ **Historical trivia & awards**: Never ask who won a Nobel Prize (e.g. Robert Edwards, Steptoe), what year/decade a discovery was made, or biographical facts about scientists.
  - ❌ **Introductory justifications & 'Why study' questions**: Never ask "Why study human embryology?", "What clinical/scientific disciplines benefit from embryology?", general justifications of why the field matters, career overviews, or course overviews.
  - ❌ **Anatomy: Listing anatomical parts or structures**: Never create cards asking to list, enumerate, or identify anatomical parts/structures (e.g., *"What are the parts of the pancreas?"*, *"List the muscles of the anterior thigh"*, *"Name the branches of the facial nerve"*). Pure anatomical identification and part naming are exclusively tested via **Image Occlusion** in Anki.
- **Examinable Medical Focus**:
  - **Permissible Anatomy Focus**: For anatomical topics, strictly test high-yield spatial, functional, and clinical relationships rather than rote parts:
    - *Spatial & Anatomical Relations*: Anterior, posterior, medial, lateral, superior, inferior, deep, and superficial relations.
    - *Neurovascular Supply*: Motor/sensory innervation, nerve roots, arterial supply, venous drainage, and collateral circulation.
    - *Fascial Planes & Compartments*: Boundaries, contents, and fascial layers.
    - *Lymphatic Drainage*: Primary drainage basins and sentinel nodes.
    - *Functional Biomechanics*: Muscle actions, synergists, and antagonists.
    - *Clinical Correlates*: Entrapment neuropathies, fracture vulnerabilities, physical exam maneuvers, surgical landmarks, and referred pain.
  - Anatomical relations, neurovascular pathways, fascial planes, histology.
  - Physiological mechanisms, timelines, hormones, receptor interactions.
  - Embryological milestones, germ layers, cell lineage, compaction/gastrulation.
  - Clinical correlates, congenital anomalies, teratogens, diagnostic signs and clinical procedures (e.g. IVF stages, ICSI mechanism, PGT stage).
- **Dynamic Content-Driven Card Scaling**:
  - Do NOT artificially cap card count to an arbitrary number.
  - Thin slides yield 0–1 cards, average slides yield 2–3 cards, and content-heavy slides (e.g. dense mechanisms, germ layer tables, clinical embryology sequences) dynamically scale up to 4–5 cards to reflect actual lecture depth.

---

## 2. Lecture Content & Visual Processing Protocol

### A. Module & Block Architecture (King's College London MBBS Year 1)
Pre-clinical curriculum folders in `./Lectures_OneDrive/` are structured by Module and Block:
- **FBS — Foundations of Biomedical Sciences**:
  - Focuses on basic scientific foundations: molecular and cell biology, biochemical pathways, genetics, and foundational neuroscience.
  - Blocks:
    - `MCB`: Molecular and Cellular Biology
    - `NBSS`: Neural Basis of Systemic Structure / Neuroscience
    - `NAM`: Nerve and Muscle
    - `DH`: Development and Homeostasis
    - `IAI`: Infection, Adaptation and Immunity
    - `EAB`: Elimination, Absorption and Balance
    - `GEN`: Genetics
- **PAS — Physiology and Anatomy of Systems**:
  - Focuses on pre-clinical body systems: anatomical relationships, systemic physiology, and principles of pharmacology.
  - Blocks:
    - `FAH`: Fundamental Anatomy and Histology
    - `FPP`: Foundations of Pharmacological Principles
    - `ACRS`: Anatomy of Cardiovascular and Respiratory Systems
    - `PCRS`: Physiology of Cardiovascular and Respiratory Systems
    - `AGRS`: Anatomy of Gastrointestinal and Renal Systems
    - `PGRS`: Physiology of Gastrointestinal and Renal Systems
    - `AMS`: Anatomy of Musculoskeletal System
    - `Cranial Nerves`: Cranial Nerve Anatomy and Pathways
- **IVBCP — Introduction to Values-Based Clinical Practice**:
  - Clinical communication, ethics, law, and professionalism.
  - Blocks: `Clinical Skills`, `Human Values`.

### B. Block-Targeted AI Querying & Card Generation
When the user asks questions or requests card generation scoped to a specific block (e.g., MCB, FAH, NBSS, FPP):
1. Constrain lecture discovery and context extraction strictly to `./Lectures_OneDrive/Year 1/<Module>/<Block>/`.
2. Cross-reference slide content and citations from that block's folder (e.g. `[FAH01L Introduction to Anatomy and Dissection · Slide 12]`).

### C. Processing Protocol
When assigned a lecture topic or file:

1. **File Discovery**:
   - Recursively search `./Lectures_OneDrive/` for matching `.pptx` or `.pdf` lecture slide decks.
2. **Dual-Channel Content Extraction**:
   - **Text & Structure**: Slide headlines, bullet points, clinical breakout boxes, bolded instructor remarks, and tables.
   - **Visual Diagrams & Graphics**:
     - Actively parse flowcharts, biochemical pathways, histology microphotographs, anatomical drawings, and cycle schematics.
     - Extract mechanisms represented visually (e.g., feedback loops, rate-limiting enzymes, organ axes, diagnostic algorithms) and translate them into atomic Cloze cards.
3. **Notion Synchronization**:
   - Check connected Notion databases/pages for lecture objectives, syllabus outcomes, and lecturer-emphasized high-yield notes to weight card density appropriately.

### D. Lecture Notes Generation & Diagram Cropping Protocol
When generating lecture notes and syncing diagrams to Notion:
1. **Style Calibration**:
   - Always calibrate style and tone against `FPP-2`, `MCB-7`, and `MCB-13` (persisted in `./references/style-examples.md`).
   - Follow `my-lecture-notes.md`: `## 🎯 Learning Objectives` verbatim, `## 📝 Notes`, `<b><u>Major block</u></b>`, `<u>Topic:</u>`, step-by-step bullets for mechanisms, short cause-and-effect paragraphs for concepts.
   - **Complex Term Clarifications**: Whenever encountering complex, obscure, or hard-to-understand medical, biochemical, anatomical, or pharmacological words/phrases, provide a brief description ~1 to 2 sentences long in brackets `(...)` right after the term explaining what it means in plain English.
   - **Diagram-Driven Conciseness & Visual Priority**: When a diagram, graph, curve, flowchart, or anatomical schematic effectively explains the concept, let the visual do the heavy lifting! Always emit `(slide N: what the diagram shows)` on its own line followed by a 1-line user revision caption. Since the visual provides the bulk of the information, generate much less text about it: keep accompanying text lean and punchy (key takeaways, critical numbers, or clinical significance) rather than redundant walls of text describing what is already obvious in the visual.
   - **Mandatory Complete Sequential Slide Coverage**: Process the lecture in a unified single pass without splitting decks into 2 chunks. Step through every content slide sequentially from Slide 2 to the end of the presentation without skipping intermediate mechanism, compensation, or pathway slides.
2. **Diagram Cropping vs. Full Slide**:
   - **Never paste entire slide canvases** with slide titles, module headers, slide numbers, and blank outer borders.
   - Crop down tightly to the required diagram section (the bounding box of the diagram plus its callouts, arrows, and labels).
   - If a slide contains overlay text boxes, callout arrows, or braces, capture the cropped rendered diagram with all labels intact rather than an incomplete raw shape.
   - Constrain all images to max width 960px, max height 540px in standard RGB PNG format.

---

## 3. Card Formulation & Review Workflow Gate

### A. Preview Format
Every card proposal must be displayed directly in the chat inside a clear markdown preview table:

| # | Front / Cloze Text | Extra Field | Source (Slide # / Notion) |
|---|--------------------|-------------|---------------------------|
| 1 | `Text with {{c1::cloze}}` | Mechanism / mnemonic / clinical pearl | Slide X / Objective Y |

### B. STRICT REVIEW GATE
- **NO Automatic Pushing**: Under NO circumstances should `anki_create_note`, `anki_batch_create_notes`, or any write tool be executed without explicit user instruction.
- **Workflow Steps**:
  1. Draft and display the preview table.
  2. Ask for user review, corrections, additions, or discards.
  3. Wait for explicit approval (e.g., *"looks good, push them"*).
  4. Only upon explicit confirmation, push the notes to the appropriate target deck in Anki via MCP.

---

## 4. GitHub-Hosted Announcements System

<!-- Schema: [{"id": "unique-id", "title": "...", "message": "...", "type": "info", "active": true, "expires_at": "2026-12-31"}] -->

The announcements system reads remote announcements from `https://raw.githubusercontent.com/Synapse-Study/synapse/main/announcements.json`.
- **id** (`str`): Unique identifier for deduplication and dismissal tracking.
- **title** (`str`): Short headline for the announcement banner.
- **message** (`str`): Markdown-supported description or update message.
- **type** (`str`): Visual style badge (`info`, `warning`, `success`, `festive`).
- **active** (`bool`): Master toggle for displaying the banner.
- **expires_at** (`str`): Expiration date string (`YYYY-MM-DD`); banners expire automatically past this date.
- Dismissed announcement IDs are stored in `dismissed_announcements` within `config.json`.
