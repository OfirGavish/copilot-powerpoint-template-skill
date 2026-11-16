# Installation

## Before installing

Use a PowerPoint Copilot client that exposes custom skills. The
[Microsoft skills documentation](https://support.microsoft.com/en-us/powerpoint/copilot/copilot-in-powerpoint-skills)
describes the OneDrive folder workflow; availability and UI can vary by client.

This skill needs a runtime that already provides:

- Python with `python-pptx>=1.0.2,<2`, `defusedxml>=0.7.1,<1`, and `lxml>=4.9,<7`.
- Code execution and access to the installed skill's local assets.
- A real file-artifact tool or equivalent supported file-output capability.

Installing instructions does not grant execution, filesystem access, or file
publication. Do **not** ask Copilot to install packages inside its sandbox.
Local developer setup is separate; see the [README quickstart](../README.md#local-development-quickstart).

## Recommended: install the folder

1. Obtain a local copy of this repository:
   [OfirGavish/copilot-powerpoint-template-skill](https://github.com/OfirGavish/copilot-powerpoint-template-skill).
2. In PowerPoint Copilot, open **Settings (…) → Manage skills → Custom skills**.
3. Choose **Create OneDrive folder** if one does not exist, then **Open skills folder**.
4. Copy the entire repository subfolder **`nova-deck-designer`** into the opened
   skills folder, preserving its contents and relative paths.
5. Wait for OneDrive to finish syncing.
6. Select **Refresh** in Manage skills.
7. Start a **new chat** and use the [first prompt](../README.md#first-prompt).

The installed structure must begin with:

```text
nova-deck-designer\
  SKILL.md
  assets\templates\
  references\
  scripts\
```

Do not install the whole repository, nest an extra `nova-deck-designer` folder,
or copy only `SKILL.md`. The containing folder and frontmatter
`name: nova-deck-designer` must match.

## Optional: create a small skill ZIP

If your client offers custom-skill ZIP upload, package only the installable skill.
From the repository root in a local development environment, run:

```powershell
python tools\package_skill.py
```

Use the resulting **`dist\nova-deck-designer.zip`**. Never upload GitHub's
**Code → Download ZIP** archive or a ZIP of the complete repository as a skill.
Downloading a repository archive for local extraction is not skill installation.
The packager runs the read-only probe, checking dependencies and asset hashes.
Its local guards reject ZIPs over 2 MiB or aggregate extracted content over
5 MiB; these guards are not universal client guarantees.

Some observed raw ZIP uploader client paths enforce **2 MiB (2,097,152 bytes)**
while displaying a generic **5 MB** error message. This is an observed,
client-dependent behavior, **not a universal documented limit**. The compact
bundled NOVA assets are below that threshold; custom decks may not be.
Prefer the documented folder workflow if ZIP upload fails.

This repository includes no organization-admin deployment manifest or icons.
The skill folder/ZIP is not an admin-center deployment package.

## Verify generation and delivery

In the new chat, select the skill and provide language, theme, topic, and facts.
The skill must really execute its bundled builder with `--probe` before building.
For a local developer check, run this from the repository root:

```powershell
python nova-deck-designer\scripts\build_presentation.py --probe
```

The probe is read-only. A narrated success is not a probe, and a successful
local probe does not establish host artifact-publication support.

Confirm that the returned **new PPTX** opens in PowerPoint, has six slides for
the first prompt, retains the selected branding, and contains editable objects.
Check clipping, overflow, chart/table values, and Hebrew reading order where
applicable. Use a rendered preview when available and distinguish performed
checks from unverified ones.

A local path is not a delivered artifact. The host must actually publish the
file. The open PowerPoint canvas is not automatically replaced or updated.

## Creator credit and redistribution

The package includes a README crediting **Ofir Gavish** and the complete
[custom attribution license](../LICENSE). If you redistribute the skill or
a reusable template package, retain the copyright and license and include
reasonably visible creator credit in your README or primary project page.
Keeping the packaged README is one way to retain that credit.

Ordinary finished presentations generated with the skill do not require this
credit. See [ATTRIBUTION.md](../ATTRIBUTION.md) for a summary; the full license
controls. This is not standard MIT or an OSI-approved license.

## Troubleshooting

| Symptom | Action |
|---|---|
| Skill missing or old behavior persists | Check folder nesting and matching frontmatter, wait for sync, select Refresh, and start a new chat. |
| ZIP upload fails with a size message | Use the folder workflow; client-dependent raw upload limits may differ from the displayed message. |
| Python, dependencies, or code execution unavailable | Report the exact missing capability. Do not claim successful generation or install packages in the Copilot sandbox. |
| No artifact/file-output capability | State that delivery is blocked. An existing local output is not a download link. |
| Template import rejects a local skill path | Do not pass local paths or `file://` URLs to an attachment-only importer. Use the direct builder. |
| Integrity/hash mismatch | Stop and review the asset. Only refresh hashes for trusted, user-approved edits; see [customization](CUSTOM_TEMPLATES.md). |
| Output already exists | Choose a new filename; the builder refuses to overwrite an existing output. |

If the runtime cannot complete the direct-builder workflow, describe an
**explicit manual fallback**: open a bundled template locally and save a new
copy, or attach a user-approved copy in a chat that supports attachments and
use a compatible workflow. Attachment support alone does not guarantee editing
or generation. State any remaining limitations; do not call this fully automatic
bundled-template success. No cloud template URL or extra connector is required.

## Back up, update, and retest

1. Back up the installed skill and any approved custom decks outside the active
   skills folder so they are not discovered as duplicate skills.
2. Test the updated copy locally, including every language/theme you use and
   the applicable content examples. Use new output filenames.
3. Replace the installed **whole `nova-deck-designer` folder** with the approved
   version; do not leave stale scripts or profiles from an older copy.
4. Wait for OneDrive sync, select **Refresh**, and start a **new chat**.
5. Generate and open a fresh output in the intended host. If it fails, restore
   the backed-up folder, sync, Refresh, and start another new chat.

The approach was confirmed in one desktop PowerPoint setup. The public NOVA
adaptation is locally validated, not universally runtime-certified; verify
execution and delivery in your own client after updates.
