My lecture notes

Turn lecture slides into notes that read as if the user wrote them: lean, in their own words, built for revision rather than completeness. The reference pages are the user's real notes; the traits below were read from them. Skim references/style-examples.md before writing, since the voice is easier to match from examples than from rules.

Workflow
Read every slide. PPTX: read the pptx skill, then extract all text. PDF: read the pdf-reading skill. Read to the end; summaries sit on the last slides. Find the learning objectives slide. Slides that are only images give you a title and nothing else, so never invent what they show.
Find the Notion page. Load the Notion tools with tool_search, then notion-search the session code and title. The page lives in the Sessions database (path "MBBS Year 1 - Home / Sessions"). Watch for near-duplicates. notion-fetch it:
Blank page: write the notes.
Page with content: never replace it. Append below it, or ask.
No matching page: say so and ask before creating one, because Sessions has its own properties.
Check the layout once. Fetch notion://docs/notion-flavored-markdown-spec before the first write.
Write the notes in the style below and insert at the end of the page.
Reply briefly (see the end).

Never touch Status, Confidence, Reviewed, dates or other properties unless asked.

Page layout
## 🎯 Learning Objectives
(the slide's objectives, copied verbatim)

## 📝 Notes
(everything else)
Copy the objectives word for word, since exam questions follow the lecturer's wording. If the page already has them, leave them alone. If the page has tick boxes, keep tick boxes; otherwise plain bullets. If the lecture groups objectives under a heading, keep that heading in bold and underline.
Everything goes under one Notes heading. No callouts, no tables, no decorative emojis in body text, no self-test, no to-do list. The user's pages do not use them, and extra apparatus makes the page feel like someone else's. Headings always use the exact emojis: `## 🎯 Learning Objectives` and `## 📝 Notes`.
Follow the order of the lecture, not a reorganised order.
How the user writes

Two modes, chosen by content.

Mechanisms, pathways, sequences of steps become plain bullets that stay close to the slide's own wording, one step per bullet. (ER targeting via SRP, nuclear import, peroxisome targeting.)
Concepts and comparisons become short paragraphs in plain, slightly informal English, explaining cause and effect. (Why a competitive inhibitor raises Km and leaves Vmax alone, what a transition state does.)

Explain the why. Where a slide states a result, add the one or two sentences that make it make sense: what a high Km means for an enzyme sitting at low substrate, why increasing [S] overcomes a competitive inhibitor but not a non-competitive one. Put this in the user's voice without marking it as an addition. Keep it correct; if unsure, leave it out and put "(check)" instead.

Explain complex words and phrases in brackets. Whenever you use or encounter complex, obscure, or hard-to-understand medical, biochemical, anatomical, or pharmacological words or phrases, provide a brief description ~1 to 2 sentences long in brackets `(...)` right after the term explaining what it means in plain English (e.g., *promiscuous ligands (molecules that bind to multiple distinct, unintended receptors across different tissues)*, *pleiotropy (when one single gene mutation impacts multiple seemingly unrelated organ systems)*).

Describe diagrams in words. When a slide's point is a diagram (feedback inhibition, allosteric T/R states), narrate what happens step by step in a paragraph, as if explaining it to a friend. If you cannot read the diagram, do not guess; say "see slide N" and move on.

Memory hooks and links to earlier material. Where it is natural, add a short hook: a link to something already learned ("remember PEX from Zellweger syndrome?"), an analogy ("similar to ATP"), a one-line rule ("Lyases do not use water to break bonds"), or a nickname in italics (sorting office). One or two per lecture. Do not force them.

Extend or correct the slide briefly. If the slide is incomplete or looks wrong (a missing seventh enzyme class, a mislabelled term), say so in a short line: what the slide says, what is standard, and "check". Never silently fix it, and never invent a fact to fill a gap.

Pointers. The user finishes with pointers to things to follow up, such as "Check KEATS" for a clinical case discussed afterwards. Where a slide says details come later, write the same kind of pointer.

What to leave out

The user keeps the conclusion and drops the machinery, so do the same:

Worked numerical examples and data tables (keep only the take-home point).
Exact amino acid sequences and other pure-recall strings, unless the lecture clearly stresses one.
Historical names and dates, unless they are the learning objective.
Derivations (state the result, e.g. Km = [S] at half Vmax).
Slides that exist only for context: title, thank-you, contact details, decorative images.
Long lists of examples. Keep one or two the lecturer emphasised.

Learning objectives decide what stays. If a slide does not serve one, it needs a good reason to be in.

Formatting in Notion

Notion treats every line as its own block, and a line placed straight after a bullet can be swallowed into that bullet as an indented child. So spacing is part of the style, not decoration. The first MCB-11 page came out as one block of same-level bullets for exactly this reason.

Spacing and nesting rules

Put a blank line between every block: heading, paragraph, bullet list, screenshot marker, caption. Never put a heading or paragraph on the line directly after a bullet.
Before each major block (<b><u>…</u></b>), add an empty paragraph, written <p /> on its own line with a blank line either side, so sections have visible breathing space. Use a blank line (no <p />) before a smaller underlined topic heading.
Headings and explanatory paragraphs are always top level. Only a real sub-point is indented, with a tab under its parent bullet. Never indent a heading.
Keep each bullet list short and under its own heading, rather than one long list that runs through the whole page.
A paragraph that introduces a list sits above it with a blank line between them; a paragraph that follows a list gets a blank line before it.

Text styling

Topic headings inside Notes are underlined plain text (<u>Mitochondria:</u>), not ## headings. Use <b><u>…</u></b> for a major block such as a whole sub-topic.
Italics for asides, nicknames and side remarks.
Bold only occasionally, for a key formula or a one-off warning.
Bullets for steps; paragraphs for explanation.
Formulas as plain text, escaping literal brackets: V0 = Vmax\[S\] / Km + \[S\].
Do not repeat the page title at the top.
Escape these when literal: \ * ~  $ [ ] < > { } |. Avoid <and>` as arrows; write "then" or use a dash.

Shape to follow

<p />

<b><u>Major block</u></b>

Short intro paragraph in the user's voice.

<u>Sub-topic:</u>

- Step one
- Step two

(slide 6: what the diagram shows)

Caption in the user's style.

<u>Next sub-topic:</u>

Paragraph explaining cause and effect.

- Point
	- Sub-point (tab-indented under its parent only)

After writing, fetch the page once and check the structure: headings at top level, blank space between sections, nothing unexpectedly nested. Fix it with a targeted update before replying.

Screenshots

The user pastes slide screenshots and writes a one-line caption saying what to take from the picture ("Ignore the bottom diagram", "Know the example above", "Look at left hand side for clear diagrams"). You cannot paste slide images. Where a screenshot would normally go, write a line such as (slide 12: SRP and translocon diagram) followed by the caption the user would write. Use captions that guide revision: what to notice, what to ignore, what to learn by heart.

Reply

Keep it short. Name the page written to, then say only what the user needs before trusting it: which slides were image-only and so described by title only (with slide numbers), and any slide errors or "(check)" items you wrote in. No recap of the notes.

Fallbacks
No Notion access or no matching page: write the same notes as a Claude Doc (or markdown file if Docs is unavailable), same layout, and say where it is.
Several lectures in one request: do each page separately.
Very long page: insert in two calls, splitting at a topic boundary.