---
name: nova-deck-designer
description: "Create an editable presentation from one of four bundled NOVA WORKS PowerPoint templates: English or Hebrew, dark or light. Use direct PPTX generation in a code-capable runtime instead of passing local asset paths to attachment-only template tools. Preserve the actual template design and use real Hebrew RTL."
license: LicenseRef-NOVA-Skill-Attribution-1.0
metadata:
  author: "Ofir Gavish"
  version: "1.0.2"
  generation-mode: "direct-pptx"
---
# NOVA deck designer

Created by Ofir Gavish. Redistribution terms are in LICENSE.txt.
Ordinary generated presentations do not require creator credit; do not add
such a credit to a presentation unless the user asks for it.

Use the actual selected PPTX as the design source. NOVA WORKS is a fictional
demo brand. Never present its sample metrics as real customer results.

## Select one reference

Ask for the language first (English or Hebrew), then the theme (dark or
light), if those choices have not already been answered explicitly in this
conversation. Do not infer the theme from the active presentation.

| Language | Theme | Template key |
|---|---|---|
| English | Dark | en-dark |
| English | Light | en-light |
| Hebrew | Dark | he-dark |
| Hebrew | Light | he-light |

Use `references/templates.json` to resolve the exact bundled asset and
bindings. Do not guess filenames, OneDrive URLs or sandbox paths.

## Runtime gate and generation

A skill-local file is NOT an uploaded chat attachment. Never pass its local
path or a `file://` URL to a template tool requiring an attachment URL.

1. Establish that this runtime provides Python code execution, the installed
   dependencies, and a way to publish generated files as downloadable artifacts.
2. Read `references/content-format.md`.
3. Actually run `python "<actual-skill-root>/scripts/build_presentation.py" --probe`.
   Require a real successful result, not a narrated or invented probe.
4. Build a UTF-8 content JSON plan matching the chosen profile. Use the user's
   facts, requested slide count and audience. Ask for the topic if missing.
5. Run the same script with `--template <key> --content <actual-plan.json>
   --output <new-host-artifact-path.pptx>`.
6. Check its actual receipt, then publish the generated file through the
   host's supported artifact/file-output mechanism.

The helper reuses complete reference slide compositions and native editable
objects. Repeated chart slides receive independent chart/workbook parts.
It preserves the original masters, layouts, themes, fonts and retained media.

A local output file is not proof of delivery. Return the actual file artifact
or download link, not a plain sandbox path. This creates a NEW presentation;
do not claim the open canvas changed unless a separate supported host commit
succeeds and is verified.

If code execution, a dependency or artifact publication is unavailable, state
the exact missing capability. Do not install dependencies in the host, invent
tool names/results, or silently approximate the template. An explicitly
described manual attachment/open-copy fallback is acceptable, but is not the
automatic bundled-template workflow.

## Content and design

- Never overwrite the bundled assets, an existing output or the user's source.
- Preserve the selected template's branding, geometry, fonts, palette, footer
  and legal notices. Do not redraw logos or substitute AI-generated branding.
- Use only supported bindings. Missing required fields, unbound content or
  unsupported objects must be reported, not silently dropped.
- Keep headings and body copy short. Split dense content across slides rather
  than shrinking the template's typography.
- Do not carry sample dates, values, names or speaker instructions into output.
  Supply every required field; omitted optional fields are cleared.
- Use existing facts only. Label proposed actions/measures and synthetic
  numbers as such. Never invent business results or commitments.
- An agenda can use the three-card layout. A six-slide story can use cover,
  three-card agenda, three priorities, roadmap, decision table and closing.

## Hebrew

Use genuine RTL paragraph direction and right alignment. Read cards and
process steps from right to left. Keep Latin product names, URLs, emails,
numbers and signs internally LTR using native run direction. Never reverse
characters or mirror the logo. The helper handles generated text direction;
do not run a separate XML rewrite that could damage namespaces or branding.
Preserve supplied characters exactly, including adjacent Hebrew/Latin text,
hashtags and product prefixes. Desired spaces must come from the input, not
automatic insertion. Keep Hebrew prefixes outside an email's LTR token.

## Check and hand off

Check the selected key, language/theme, slide count, editable content, chart
data, unchanged branding, no leftover sample content, no clipping/overflow,
and correct mixed Hebrew-English order. Inspect rendered slides when the host
supports rendering; distinguish performed checks from unverified ones.

For an existing presentation, preserve source facts, tables/charts, hyperlinks,
meaningful images and notes. Do not force unsupported source objects into this
JSON schema. Use a faithful compatible workflow or report the limitation.

Report the actual selected reference filename and the real delivered output.
Do not say complete unless generation and delivery both succeeded.
