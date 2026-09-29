---
name: smart-ppt
description: >
  Create or edit PPTX presentations from complete user templates, user assets, or
  built-in styles. Preserve template layout; add relevant visuals, speaker notes,
  and a per-slide source report. Use for academic defenses, reports and teaching.
---

# smart-ppt

Locate this SKILL.md directory first. Use Python 3.10+, python-pptx, Pillow and the bundled scripts. A user may place complete PPTX files or nested template folders in `templates/`; each PPTX is an independent candidate. Do not publish user templates.

## Route the request

- Existing PPTX or a matching file in `templates/`: read [template-workflow.md](references/template-workflow.md). Scan with `python scripts/template_library.py list`, choose one complete deck, and edit/add/delete/reorder pages as the user asks. Preserve all unspecified pages by default.
- Brand assets without a PPTX: read [creator-workflow.md](references/creator-workflow.md) and build a new template.
- Neither template nor assets: read [style-workflow.md](references/style-workflow.md) and select or customize a built-in style.
- If intent is unclear, ask about an existing template or preferred style. Do not repeat questions answered in the session.

## Template invariants

- Inspect the selected deck and its layouts before editing. Show available candidates and explain the choice when the user has not chosen one.
- Keep the original PPTX read-only and write a new output. Full-deck mode retains the entire source unless the user requests a page change. Do not silently use a few sample pages as the whole deck.
- Fill existing text shapes through their text frames and replace pictures in their original shapes. Never overlay an unrelated textbox or image to hide template content. Do not move or resize existing placeholders; log any intentional geometry change.
- If content exceeds a page's capacity, edit it or add a suitable page. See [layout-safety.md](references/layout-safety.md).

## Visual and content standard

- Body text, chart labels and ordinary captions should generally be at least 14 pt. Simplify or split dense pages instead of shrinking the type. Smaller source notes may be used when necessary.
- Use black body text on light backgrounds and white body text on dark backgrounds. Gray is for supplementary information; headings may use theme colors.
- Add consistent icons or illustrations where they explain an idea. Review each template image against its new page text. If it is no longer relevant, use a suitable user asset, generate an original illustration, or find an image with usable rights; replace its picture data without changing the shape's position, size or crop. Record its origin and license. See [visual-standards.md](references/visual-standards.md).
- Record factual claims and their real sources. Do not invent links or citations.

## Finish every deck

1. Prepare a talk track for **every resulting slide** and write it to that slide's PowerPoint speaker notes. It should explain the actual slide, including visuals and transitions.
2. Create `信息源报告.md` beside the PPTX. List sources by slide, with title, link and access date; list image origins and licenses separately. Label user material, analysis and AI-generated visuals accurately, and flag facts still needing verification.
3. On Windows with PowerPoint, provide `sequence[].notes`, `speaker_notes` keyed by zero-based source page, or `edits[].notes` in the plan, then run `powershell -File scripts/finalize_powerpoint.ps1 -Pptx output.pptx -Plan plan.json -Audit output.pptx.audit.json`. On other hosts, use an available presentation editor or python-pptx notes API and `python scripts/write_source_report.py output.pptx --plan plan.json --audit output.pptx.audit.json`; verify notes were actually saved.
4. Run `python scripts/validate_output.py output.pptx --report validation.json`, render and inspect every slide, then check font size, body contrast, relevant images, slide geometry, notes, and source report. Deliver the PPTX, report and validation summary together.

The repository is agent-neutral. Codex can use this globally installed skill automatically when a PPT request matches its description; the user does not need to open this file or install it again.
