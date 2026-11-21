# Portable bilingual deck builder

**Created by [Ofir Gavish](https://github.com/OfirGavish).**

The portable variant targets capable Cowork sessions, agents using the
Microsoft Copilot harness in Copilot Studio, and Scout. It ships the same four
English/Hebrew dark/light PPTX decks as the PowerPoint variant, not just
instructions for users to find their own templates.

## Build

From the repository root in an approved local developer environment with
`requirements-dev.txt` installed (including the runtime and frontmatter-validation
dependencies):

```powershell
python tools\package_skill.py --variant portable
```

To package both variants, use `--variant all`. With no argument, the existing
PowerPoint-only packaging behavior is retained.

The portable build produces:

| Output under `dist` | Purpose |
|---|---|
| `bilingual-multi-template-deck-builder.zip` | Self-contained install ZIP; SKILL.md is at the archive root |
| `bilingual-multi-template-deck-builder\` | Complete unpacked skill for folder-based installation |
| `bilingual-multi-template-deck-builder.build.json` | ZIP SHA-256, byte sizes, file count and output paths; not part of the skill |
| `cat\submissions\bilingual-multi-template-deck-builder\` | Separately licensed MIT CAT edition with metadata/README sidecars |

Builds replace these generated outputs. `dist` is intentionally ignored by Git.
The ZIP is reproducible from the same source bytes and Python/zlib toolchain.
Do not upload the build receipt as a Scout automation or install it as a skill.

## One runtime, two instruction sets

The canonical scripts, assets and binding/content references remain under
`nova-deck-designer`. The packager combines these with the portable authoring
overlay under `variants\bilingual-multi-template-deck-builder`, the root license
and requirements, and repository examples. There is no second renderer to
maintain. All four PPTX files and the binding manifest are copied byte-for-byte.

Do not install the overlay itself: it intentionally contains no duplicate
renderer or assets. Install the generated folder/ZIP instead. Make shared
runtime/template changes in the original skill and rebuild both variants.
The original PowerPoint instructions and assets are not modified by packaging.

The portable package adds:

- Capability-gated native/Python routes, host setup and actual file delivery.
- A self-contained requirements file and both languages' example plans.
- A CAT metadata sidecar in the submission output only, not the install ZIP.

The original packager's observed 2 MiB raw-ZIP / 5 MiB aggregate guards apply
only to the PowerPoint variant. No universal replacement limit is hard-coded
for portable hosts. Check the target client's current constraints; ZIP import
is not permission to run arbitrary code, install dependencies, or publish files.
The renderer still enforces its own supported expanded PPTX size and content
contract on both variants.

## Install and invoke

Follow the [packaged overview](../variants/bilingual-multi-template-deck-builder/README.md)
and [host-specific instructions](../variants/bilingual-multi-template-deck-builder/references/hosts.md).
The ZIP must preserve its root SKILL.md and every supporting file.
If a client requires a named folder, use the generated unpacked folder instead.
Neither package is an admin-center M365 plugin/app deployment ZIP.

The skill asks for language and theme, then chooses a supported native-template
handoff or the shared Python builder. Python is not a blanket prerequisite.
Only the Python route requires the real `--probe`. Both routes must use the
actual bundled PPTX, preserve editable design and deliver the actual file.
No bundled asset path is passed to an attachment-only template importer.
No existing presentation or open canvas is automatically updated.

Host support is conditional on actual bundled-file access and binary output.
The Python route also requires Python 3.10+ and the specified dependencies.
Instruction-viewing, native presentation tools and code execution can have
different file scopes. An actual host-managed OneDrive location for an imported
skill may provide a supported source reference; never guess one or bypass
an enforced scope restriction. The original renderer checks are reusable
evidence, not proof of live import/execution/delivery in every host.
Missing capabilities are explicit blockers, never successful fallbacks.

Initial 1.0.0 live testing established real Studio Hebrew/dark generation and
downloaded-artifact design preservation, plus Scout local/file-backed execution.
Cowork ZIP import and language/theme selection worked, but its direct
skill-resource handoff did not. A mixed Hebrew/English visual-spacing defect
was found through native PowerPoint rendering. These observations must remain
pinned to that package hash; a later build needs its own host evidence.

## Live-tested routes in 1.1.0

The 2026-11-17 portable 1.1.0 test ZIP has SHA-256
`36086c85ab6c9b26f74d2ff1cfab7570ee70e0e37dcf7566b086637ba1246fc6`.

| Host | Tested route and scope |
|---|---|
| Cowork | Full ZIP import; original bundle discovered in host-managed OneDrive storage and all 18 files staged through supported file tools; Python builder generated and returned all four eight-slide profiles without manual template attachments. The returned Hebrew deck opened in the native PowerPoint web viewer with correct table values and mixed-script spacing. |
| Copilot Studio | Full ZIP replaced in an unpublished Copilot-harness agent; fresh preview generated four actual eight-slide PPTX attachments. Independently retrieved binaries preserved every bound text field, native charts/workbooks/tables, and all 21 source master/layout/theme/font parts per profile. All 32 native PowerPoint renders matched Scout's corresponding renders pixel-for-pixel. |
| Scout | Supported file-backed registration loaded the full extracted ZIP; actual probe, four eight-slide builds and native PowerPoint rendering succeeded. This does not establish a ZIP-import UI or Co-create-provider binary delivery. |

These are scoped smoke tests of the direct-builder route, not a claim that all
native editing operations work or every client/tenant is certified. The focused
Hebrew/Latin separator fix restored native spacing without changing template
assets, fonts, or the other seven rendered slides in either Hebrew theme.

The native-only Cowork diagnostic did reach the actual bundled OneDrive template
and preserve its design, but its live text-replacement interface could not edit
table cells. That six-slide diagnostic retained sample table rows and **did not
pass**. Use the complete-bundle Python fallback for such unsupported objects;
do not deliver stale sample values or silently flatten/recreate the table.

Cowork duplicate imports can create an automatically suffixed skill name such
as `bilingual-multi-template-deck-builder-1`. Select the intended installed copy
explicitly during testing; back up or remove obsolete test copies before a
production rollout rather than relying on ambiguous automatic invocation.

Package and renderer versions are separate identifiers. The 1.1.0 package uses
the shared renderer's `1.0.0` receipt identifier. Its actual script SHA-256 is
`da9062b4728af636386b28d6a748945b1cbd328ac1c1ec4c9ed333b6c3452ce8`;
the build receipt records both versions and that hash. A different numeric
helper version alone does not establish a stale script. Studio displayed an
"incomplete skill" debug label despite returning independently verified outputs;
that UI diagnostic remains unclassified, not a reason to fabricate a failure
or claim universal host certification.

Capture delivery hashes **before** opening outputs in an editing-capable Office
viewer. In the Cowork test, simply opening the Hebrew/dark output in PowerPoint
web caused Office to add a sensitivity-label record and reserialize package
parts; its post-view bytes no longer matched the original build. Keep that
viewed file and its label, and retain a separate unopened audit copy from the
original build. Do not remove labels or mistake client-side reserialization
for the builder having changed its source template.

Detailed private run receipts and artifact/viewer evidence live under the
ignored `outputs\host-tests\2026-11-17` directory. They are intentionally not
part of the reusable ZIP or CAT submission.

## Final acceptance follow-up in 1.1.1

The corrected 1.1.1 ZIP has SHA-256
`6bdb14d7bf55202bcf142bcecabf115d0f43ad80150d516e5a9044b0203e12cb`.
Its shared renderer SHA-256 is
`45d44c856c67570c93503096fe54c0bb3b107c48a7a752d66b2efb8a7435e902`.
The shared renderer's internal receipt identifier remains `1.0.0`; package
metadata and exact script hashes distinguish the builds.

The separate clean Scout test discovered a defect that example percentages had
not exposed: template series-level percentage labels overrode a requested raw
number format. The fix applies explicit formatting to plot, series and point
labels, workbook values and numeric axes, and clears inherited axis tick units.
Raw `11, 22, 33`, fractional `0.11, 0.22, 0.33` displayed as percentages, and
whole-number percentage inputs are covered by regression and native PowerPoint
checks. All 51 repository tests pass.

| Acceptance item | Result and boundary |
|---|---|
| Clean Scout installation and separate-session generation | Complete canonical folder installed in Scout's actual local skill directory, replacing the old loader. A separate Scout automation session loaded the installed skill, authored its own task-specific plan, ran the actual probe/builder, and returned the requested local PPTX link. The final concise six-slide deck passed 51 native text-fit checks without source/output patches, plus independent content/design verification. |
| Fresh Scout layout quality | An earlier autonomous run chose a long card heading that failed a native text-height check; it was not treated as visual acceptance. A separate fresh run with explicit short headings passed. Character limits alone remain insufficient to guarantee visual fit for arbitrary content. |
| Studio warning | A minimal instruction-only echo skill returned its exact requested result while displaying the same "Skill incomplete" label. This reproduces the label independently of the NOVA package, code, templates and optional metadata. Its internal host mapping is not diagnosed, but it is not evidence of a NOVA-specific package failure. The temporary control skill was removed afterwards. |
| Native-only boundaries | Cowork accepted the actual bundled template but refused one bounded existing-table-cell edit and one chart-part text edit; its exposed chart operation only creates a new chart. Studio exposed no native presentation editing tools. Both correctly reported an expected capability blocker rather than claiming native success. The complete tested Python fallback is required on those surfaces. |
| Corrected Studio and Cowork build | Each freshly imported 1.1.1 and returned four eight-slide language/theme decks plus a two-slide raw/fractional-chart diagnostic. Studio's five actual downloaded binaries were independently checked for content, chart formats, native tables/workbooks and preserved design, opened in desktop PowerPoint and rendered. Cowork returned all five cloud artifacts and a persisted pre-view provenance JSON; verification there used its host checks and read-only JSON viewer, not local parsing of cloud Office documents. |
| CAT official commands | All three were attempted in a real CAT checkout, but failed before importer/build execution because dependencies could not be installed. The locked `mdast-util-to-hast@13.2.2` returned E404 from both configured and public registries. A baseline comparison without this submission failed identically. No CAT-generated bundle or official green validation is claimed. |

Scout's app can normalize optional SKILL.md frontmatter on restart; in this
test it retained the instruction body and all resource files, including the
full license and README creator credit. The clean test used directory-based
installation, not an asserted ZIP-import UI, and a requested local-file
destination, not Co-create-provider binary delivery. Separate test automations
were disabled one-shots with Teams notifications off and were deleted after
the runs; their session evidence remains.

## CAT handoff

CAT's current contribution format is an **unpacked**
`submissions\bilingual-multi-template-deck-builder` directory, not a ZIP.
The generated submission includes matching frontmatter/folder names, platform
tags, author information, a human-facing README and the full MIT LICENSE.txt
for the separately approved CAT edition.
The install ZIP omits catalog metadata, while keeping a human-readable credit
README. CAT strips the metadata and README from its generated download but uses
the README on the catalog page, preserving visible project-page attribution.

Do not hand-edit `src\content\skills` in CAT. Copy the generated submission into
a local CAT checkout and run its documented commands:

```powershell
npm run check:submissions
npm run import:submissions
npm run build
```

This repository does not open a PR. Initial attempts to run these commands
against CAT commit `189409b5be6caf7c83bfaac1857044e72b053dbe`
were blocked before validation/import/site generation by unavailable locked
dependencies. Subsequent local repair and successful commands are documented
below; the pristine upstream lock was not silently presented as working.
The initial custom-license contribution question was resolved only after the
author explicitly approved a separate MIT edition, not by stripping attribution
or changing the original project's license.

## CAT blocker resolution

The official commands subsequently passed on the same CAT source commit using
a documented **local compatibility repair**, not the pristine upstream lock:

- Pin `mdast-util-to-hast` to the available `13.2.1`, within its parent's
  `^13.0.0` requirement.
- Pin Astro's `rolldown` to the available `1.3.1`, within Astro's `^1.2.16`
  requirement; the locked `1.3.2` was also unavailable in the configured registry.
- Regenerate the local validation lockfile from the unchanged direct dependency
  ranges plus those explicit overrides. A targeted lock refresh left stale nested
  entries, so the complete resolution was necessary. It selected compatible
  patch updates including Astro `7.4.8` and Vite `8.3.7`.

The full patch and original lockfile are retained under the ignored
`outputs\host-tests\2026-11-17\blocker-resolution` directory. The repair belongs
to the local CAT checkout, not to the reusable skill or its runtime dependencies.
It has not been committed or submitted upstream. A pristine checkout still
needs the unavailable dependencies addressed before its CI can reproduce this.

With that environment, `check:submissions` validated all 161 submissions,
`import:submissions` imported all 161, and `build` generated 655 pages.
All 24 official importer regression tests also passed. The actual generated
download has 17 files, omitting only the catalog metadata and README sidecars;
all four templates, scripts, references, examples and the complete custom
license match the submitted bytes. Its real extracted builder probe and all
four eight-slide example builds succeeded.

The generated page and served ZIP were checked through a local HTTP preview.
The page visibly credits Ofir Gavish and labels the custom terms as
**See LICENSE**, not MIT. The packager now changes reference links only in
the CAT README sidecar to canonical submission-source URLs; otherwise those
relative links led to nonexistent catalog-site routes. The install ZIP keeps
working local reference links and remains byte-identical to the live-tested
1.1.1 ZIP. These GitHub source URLs become public when the submission is merged.
The unpublished slug also has no Giscus discussion yet; its discussion lookup
404 is not a missing skill asset, and no public discussion was created.

**Technical success does not resolve the license agreement.** CAT's
`CONTRIBUTING.md` still says contributions agree to the repository's MIT terms,
while `submissions/README.md` and the importer support non-MIT license files.
Retaining a custom LICENSE.txt and seeing the right page label does not itself
waive that contribution clause. That was the remaining approval boundary;
the author subsequently approved the separate MIT CAT edition described below.
No automatic relicensing of the original project is performed.

## Author-approved MIT CAT edition

On 2026-11-17 the author explicitly approved a **separate MIT-licensed CAT
distribution**, leaving the original custom-licensed project and tested ZIPs
unchanged. `python tools\package_skill.py --variant portable` now builds both:

| Distribution | Terms |
|---|---|
| `dist\bilingual-multi-template-deck-builder.zip` and its unpacked folder | Original NOVA Skill Attribution License 1.0; visible redistribution credit condition retained |
| `dist\cat\submissions\bilingual-multi-template-deck-builder` | Standard MIT License with Copyright (c) 2026 Ofir Gavish; no additional README/project-page credit condition |

The CAT edition changes only `LICENSE.txt`, the license declaration in SKILL.md,
and the human README's licensing section and catalog reference links. Metadata,
runtime scripts, binding references, examples and all four PPTX templates stay
the same. Edition resources live in `editions\cat-mit`; they are not copied into
the original ZIP. The build receipt records the original and CAT license
identifiers and the MIT license hash separately.

The standard MIT grant is unmodified apart from the correct original-creator
copyright line. CAT's current detector names only an exact match to Microsoft's
own MIT text, so an Ofir-copyright MIT file can still display **See LICENSE**.
The actual bundled terms and SKILL.md declare MIT; do not replace Ofir Gavish's
copyright with Microsoft Corporation merely to change that UI label.

CAT removes the human README from its download but uses it on the catalog page.
The page and bundled SKILL.md still identify Ofir Gavish as the creator, and the
download carries the complete MIT copyright and permission notice. That visible
credit is provenance, not a reinstated custom attribution condition.

After running the official CAT importer, its actual generated MIT download can
be saved separately as
`dist\cat\bilingual-multi-template-deck-builder-mit.zip`. It is not the ordinary
custom-licensed install ZIP, and its content differs only in the documented
license/notices plus CAT's usual removal of README and metadata sidecars.

This resolves the author's license-choice blocker without asserting maintainer
acceptance or publishing anything. The local CAT dependency compatibility repair
and its not-yet-landed upstream status remain as documented above.

## Review follow-up in 1.1.2

The portable skill's review follow-up keeps the existing templates and license
editions while tightening the content/runtime contracts:

- Select Python only when the COMPLETE bundle is accessible: builder, manifest
  and all four assets at their original relative paths. Native access to one
  chosen PPTX does not establish this full-bundle Python capability.
- Reject literal dots in input object keys before flattening; dotted binding
  strings in the manifest remain valid. Unsupported fields must error rather
  than vanish during lookup.
- Preserve the exact supplied text while assigning native run direction.
  Desired spaces come from the plan; do not rewrite mixed-script hashtags,
  adjacent Hebrew/Latin words or product prefixes.
- Keep adjacent Hebrew prefixes/suffixes outside common ASCII email tokens.
  ASCII email/URL cases and punctuation are covered separately from unverified
  internationalized-address behavior.
- Give added chart series distinct selected-theme accents while preserving the
  first series' authored styling and the original masters/themes/media. Native
  multi-series legends reserve space below the plot; the original high-value
  overlap case is retained in regression/render evidence, not avoided by
  changing test data.
- Keep human installation and first-use material in the root README. The
  agent-facing host reference now contains only resource, execution and delivery
  contracts.

The PowerPoint-specific 1.0.1 metadata correction was independently confirmed
to load by the user. PowerPoint skill 1.0.2 retains that correction while sharing
the reviewed renderer; this does not diagnose an unrelated package's failure.
The 1.1.1 live-host evidence above is historical, not a claim that every newer
build has been re-tested in every host or tenant.
