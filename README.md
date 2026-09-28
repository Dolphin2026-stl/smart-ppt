# smart-ppt

A complete-template-first PPT skill for Agent Skills hosts.

[中文说明](README.zh-CN.md) · [SKILL.md](SKILL.md)

## Bring your own templates

No school or company templates are bundled. Clone the project and put your own complete PPTX files or folders under `templates/`. Nested folders are supported. Every complete file is a separate candidate.

```sh
git clone https://github.com/Dolphin2026-stl/smart-ppt.git
cd smart-ppt
python -m pip install -r requirements.txt
python scripts/template_library.py list
python scripts/template_library.py select TEMPLATE_ID work.pptx
```

```text
templates/
  my-school/
    defense-blue.pptx
    defense-red.pptx
  my-company/
    quarterly-report.pptx
```

Default editing preserves all original slides. Add, delete, reorder or change content only according to the user's needs. An empty plan `{}` produces a byte-identical copy. User templates and generated catalogs are Git-ignored and excluded from release ZIPs.

## Three modes

| Mode | Input | Output |
|---|---|---|
| Complete template | Your PPTX and an edit plan | Full deck with only requested changes |
| Creator | Logo, background/palette, font preference | Reusable seven-layout template |
| Style | Content and style preference | Six built-in styles |

```sh
python scripts/layout_inspector.py work.pptx --output layouts.json --table
python scripts/template_engine.py work.pptx plan.json output.pptx
python scripts/validate_output.py output.pptx --report validation.json
```

[Template workflow](references/template-workflow.md) · [Creator workflow](references/creator-workflow.md) · [Style workflow](references/style-workflow.md)

## Self-authored examples

These visuals use the project's own generic demo assets, not institutional templates.

![Example output](docs/assets/hero-screenshot.png)
![Reconstructed workflow](docs/assets/demo-workflow.gif)

The GIF combines actual command output and actual PowerPoint renders; it is not a live screen recording. [Example PPTX](examples/demo-deck/template-demo.pptx)

## Compatibility and tests

Requires Python 3.10+ and filesystem access. Executed in Codex on Windows with Python 3.13.13 / python-pptx 1.0.2. Claude Code, OpenClaw and Hermes are supported by host-independent SKILL.md instructions but have not each been runtime-tested.

Optional multi-agent roles communicate through JSON plans and reports, with a single PPTX writer. A single agent can run the same workflow sequentially.

```sh
python -m unittest discover -s tests -v
```

Text overflow is estimated; render slides to confirm appearance. Chart/image data replacement is separate from the text-edit API. [Verification scope](docs/verification.md)

## Project structure

`SKILL.md` is the entrypoint; `scripts/` implements engines; `references/` contains detailed workflows; `templates/` holds your private files; `assets/` defines styles; `examples/` and `docs/` contain self-authored examples; `tests/` validates behavior.

## Release

Run `python scripts/package_release.py` to create a ZIP and SHA-256 excluding user templates. The public repository includes no private source history or institution-specific decks, logos or screenshots.
