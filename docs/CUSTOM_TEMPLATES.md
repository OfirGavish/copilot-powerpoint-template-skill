# Custom templates

The builder is **profile-driven**, not an automatic visual-inference engine.
Replacing filenames alone does not make an arbitrary deck compatible. Back up
your skill and source decks before editing; originals must not be overwritten
by generation.

## Easiest route: preserve the NOVA compositions

1. Open copies of the four NOVA WORKS PPTX assets in PowerPoint.
2. Edit fonts, colors, logos, and backgrounds while preserving the reference
   compositions, mapped shape names, object types, and intended editable areas.
3. Save all four variants under `nova-deck-designer\assets\templates`.
4. Review the files as trusted, user-approved inputs, then refresh integrity
   hashes from the repository root:

   ```powershell
   python tools\update_manifest.py
   ```

5. Test every variant and inspect the resulting presentations before installing
   the updated folder.

Keep the four keys `en-dark`, `en-light`, `he-dark`, and `he-light`. Each plan's
`language` must be `en` or `he` and match its selected key. Changing languages
requires appropriate authored layouts and direction, not just translated text.

SHA-256 hashes identify the approved assets. Do not automatically regenerate
hashes to bypass an unexplained mismatch: first inspect the changed files and
obtain approval.

## Different decks: inspect and map explicitly

Place approved PPTX files in `nova-deck-designer\assets\templates`, then inspect
each one locally:

```powershell
python nova-deck-designer\scripts\inspect_template.py nova-deck-designer\assets\templates\NOVA-EN-LIGHT.pptx
```

To save the inventory, append `--output inventory.json`. Review reference slide
indices, shape names, and shape types. Choose the reference slides and bindings
in `nova-deck-designer\references\templates.json`, then run
`python tools\update_manifest.py` on the approved files and test every variant.

Mapped shape names **must be unique on each reference slide**. Use native
editable text, chart, and table objects for their corresponding field types;
a flattened image is not an editable replacement.

**Map or remove all non-branding sample content in replacement decks.** Bind
every retained sample-content object to a supported content field, or remove
it from the reference, layout, or master before use. Unmapped text, pictures,
charts, and tables are retained; they are not automatically discovered or
cleared. Rejecting unknown JSON fields does not detect unmapped sample content
in the source deck. Review retained media and embedded data as well as visible
slide text.

### Manifest schema v1

This abbreviated example shows the structure; it is not a complete profile:

```json
{
  "schema_version": 1,
  "templates": {
    "en-light": {
      "asset": "assets/templates/NOVA-EN-LIGHT.pptx",
      "sha256": "<approved asset SHA-256>",
      "language": "en",
      "direction": "ltr",
      "layouts": {
        "cover": {
          "reference_slide": 1,
          "fields": {
            "title": {
              "shape": "content.title",
              "type": "text",
              "required": true
            },
            "subtitle": {
              "shape": "content.subtitle",
              "type": "text",
              "required": false
            }
          }
        }
      }
    }
  }
}
```

Asset paths in JSON are relative to the **skill root**, not the `references`
directory. `reference_slide` is a one-based slide index. Field paths may address
nested array content, for example
`cards.0.heading`; supported field types are `text`, `chart`, and `table`.
New fields need profile and content-plan updates. New object kinds require
builder support as well; the builder does not infer a mapping from appearance.

## Content contract

The complete required fields and mappings are defined in
[`references/content-format.md`](../nova-deck-designer/references/content-format.md).
A plan starts with:

```json
{
  "language": "en",
  "slides": [
    {
      "layout": "cover",
      "title": "NOVA WORKS",
      "subtitle": "Pilot proposal",
      "meta": "Proposed discussion"
    }
  ]
}
```

Reference layout keys are `cover`, `section`, `three-cards`, `metrics`, `chart`,
`roadmap`, `table`, and `closing`. The pilot examples contain six slides; the
`all-layouts` examples exercise all eight.

- Unknown fields and missing required bindings fail rather than silently
  dropping content. Omitted optional text fields are cleared so sample copy
  does not leak into output.
- Chart payloads supply `categories`, `series` with `name` and `values`, and an
  explicit `number_format`. These replace the native chart data. Currently the
  renderer supports **category charts with a value axis**, such as the bundled
  column chart; unsupported chart families fail rather than being converted.
  Each repeated chart gets independent chart/workbook parts, so editing one
  does not change another.
- Table payloads must **exactly match** the source native table's row and column
  dimensions, including the header row. The bundled reference is **four rows
  by three columns**, and every cell value must be text. Split larger tables
  across repeated table slides, or author and bind a different reference.
- Source speaker instructions are removed unless supplied in the plan's
  `notes`. Review supplied notes before redistribution.
- Use approved facts only. NOVA sample figures are synthetic, not evidence of
  results. Do not present invented metrics or commitments as facts.

## Hebrew, fonts, and preserved design

Hebrew cards read right to left through the reference's **authored placement and
bindings**. The builder reverses logical row values into the native Hebrew
table. Keep logical content in the documented order; do not manually reverse
characters or double-reverse table values. Check mixed-direction text,
punctuation, numbers, and Latin words in PowerPoint. Preserve the actual logo
unmirrored.

The direct builder clones reference compositions and edits their editable
objects, preserving masters, layouts, themes, font parts, and retained media.
Default Aptos, Aptos Display, and Arial fonts are referenced; **no font binaries
are shipped**. Host font substitution can change
wrapping. Only embed fonts when their license permits it; the renderer retains
embedded font data rather than granting redistribution rights.

## Validate and install your update

From the repository root in a local development environment:

```powershell
python -m pip install -r requirements-dev.txt
python nova-deck-designer\scripts\build_presentation.py --probe
New-Item -ItemType Directory -Force outputs | Out-Null
python nova-deck-designer\scripts\build_presentation.py --template en-light --content examples\all-layouts-en.json --output outputs\all-layouts-en-light.pptx
python -m unittest discover -s tests -v
```

Actually run the read-only probe. Repeat generation for all four template keys,
using the English or Hebrew plan as appropriate and a new filename each time.
Existing output files are refused. Open the outputs and check all eight layouts,
branding, overflow, optional-field clearing, notes, editable chart/table data,
repeated-chart independence, and RTL reading order. Render slides when possible;
state which visual checks were not performed.

For optional Windows visual QA, [tools/render_windows.ps1](../tools/render_windows.ps1)
requires installed desktop PowerPoint. After generating decks in `outputs`, run:

```powershell
powershell -NoProfile -File tools\render_windows.ps1 -InputDirectory outputs -OutputDirectory outputs\rendered
```

It exports slide PNGs and a `text-overflow.json` report under the output
directory. It opens isolated read-only copies, does not save changes to the
source decks, and leaves existing user presentations open and untouched.
Text-height checks do not replace visual inspection for clipping, reading order,
or other layout issues. Rerunning replaces prior renders in the selected output
directory.

After local validation, back up the installed folder outside the active skills
folder, replace the whole `nova-deck-designer` folder, wait for sync, select
**Refresh**, and start a **new chat**. Test actual artifact delivery in the
target host; a locally generated file alone is not delivery. See
[installation and rollback](INSTALLATION.md#back-up-update-and-retest).

## Redistribution checklist

- Obtain customer approval for replacement logos, design, content, and
  sensitivity labels; retain required legal notices and handling restrictions.
- Clean author metadata, speaker notes, comments, and links before sharing.
  The builder's note handling is not a general privacy scrubber.
- Remove confidential content and external references that recipients should
  not receive. Inspect retained media and embedded workbooks too.
- Confirm licenses for replacement decks, artwork, and any embedded fonts.

The [custom NOVA Skill Attribution License](../LICENSE) covers the repository's
original code, documentation, and NOVA demo artwork. If you republish the skill
or reusable template package, retain the copyright/license and give reasonably
visible creator credit to **Ofir Gavish** in your README or primary project page.
See [ATTRIBUTION.md](../ATTRIBUTION.md).

Ordinary finished presentations do not require creator credit under the output
exception. That exception does not cover reusable skill/template redistribution.
Third-party fonts, Office, trademarks, and independently authored replacement
materials retain their own rights. These public examples contain only fictional
NOVA WORKS branding and synthetic data.
