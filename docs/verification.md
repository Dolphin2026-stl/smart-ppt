# Verification scope

The public export contains only general-purpose code and self-authored demo content. No institution templates, logos, screenshots or original private Git history are included.

Core behavior tests cover full-file copying, complete-page retention, targeted text editing, explicit deletion, independent chart/notes duplication, placeholders, picture proportions and pagination in explicit layout mode.

Executed environment: Codex desktop, Windows, Python 3.13.13 and python-pptx 1.0.2. Claude Code / OpenClaw / Hermes use a host-independent SKILL.md design but have not been individually runtime-tested.

Text overflow detection is heuristic; actual rendering remains required. Only self-authored examples are shown in README images.

23 regression tests passed, including recursive custom-template indexing and byte-identical selection. Public PPTX XML and text assets were checked for private institutional references before publication.
