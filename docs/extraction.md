# Extraction provenance

Source: `lucamattiazzi/ai-cdl`, branch `main`, commit
`633ee519c311da9a1af9f7d6081b43194ab24ff1` (including the merged Pair testing and safety changes).

The source checkout was older than its remote branch. This repository was extracted from Git
objects at that remote revision without changing the source working tree. Only selected tracked
files were copied; no original Git history, local environment files, generated session scripts,
conversation exports, certificates, databases, or build artifacts were imported.

The extraction starts from `apps/pair-addin`, `apps/pair-server`, and `packages/pair-cli`, and
includes their complete workspace dependency graph. It also includes the Pair Agent Skill and
deployment files. The original Apache-2.0 license is preserved and included in each library package.

Repository-specific changes:

- Root workspace, TypeScript aliases, Vitest config, Changesets, and checks target retained packages.
- Add-in icons now live in `apps/pair-addin/public`.
- The pairing prompt no longer points skill installation at the original repository.
- Documentation describes this standalone checkout, its retained dependencies, and publication.
- CI verifies Pair and packages without an automatic publishing workflow.

Package names and versions were preserved at extraction time. The independent project was
subsequently renamed Sommelier in `0.2.0-beta.1`, with new `@lucamattiazzi/sommelier*` package
names and reorganized directories. The original `@ai-cdl/*` publications remain available.
See [migration to Sommelier](migration-to-sommelier.md) for the current names.
