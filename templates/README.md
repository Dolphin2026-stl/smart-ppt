# Put your own complete PPTX templates here

Clone this repository, then place PPTX files or entire template folders under this directory. Nested folders are supported. No organization-specific templates are bundled.

```text
templates/
  my-school/
    blue.pptx
    red.pptx
  my-company/
    quarterly-report.pptx
```

Run `python scripts/template_library.py list` from the project root. Every full PPTX becomes a candidate; all original pages are preserved until your request calls for edits, additions or deletions.

User template files and the generated library.json are ignored by Git and excluded from the release ZIP. Keep working outputs outside templates/ to avoid indexing them as new templates.
