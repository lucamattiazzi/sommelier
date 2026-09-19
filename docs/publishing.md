# Publishing Sommelier

The Excel add-in and npm libraries are separate artifacts. The source repository is
[sommelier](https://github.com/lucamattiazzi/sommelier). The first release under the Sommelier name is
`0.2.0-beta.1`. The previous `0.2.0-beta.0` release used the `@ai-cdl/*` namespace.
The new packages use the personal `@lucamattiazzi` scope; install with `@beta` explicitly.
Verify registry availability after each release before announcing installation.
Microsoft Marketplace submission and production hosting are separate steps.

## Libraries

`packages/*` contains ten public packages with ESM/CommonJS entry points, TypeScript declarations,
and Apache-2.0 licenses. `@lucamattiazzi/sommelier-client` is the client/session library; `@lucamattiazzi/sommelier` provides
the local relay plus `sommelier`, its bundled encrypted bridge, MCP tools and native
harness adapters. See [v1 readiness](v1-readiness.md) before release. The root workspace and the two applications remain private npm packages.

Publish using the `lucamattiazzi` npm account. Repository metadata points to
`lucamattiazzi/sommelier`. These are new package names: the old `@ai-cdl/*` versions and
stable tags remain available and are not overwritten. See [migration](migration-to-sommelier.md).

```sh
pnpm build
pnpm smoke:packages
pnpm smoke:consumer
# Already in beta prerelease mode; add a Changeset for the next change.
pnpm changeset
pnpm changeset version
pnpm build
pnpm smoke:consumer
```

Review the generated version changes and release notes. Changesets groups the ten libraries into
one fixed release set. Run `pnpm release:beta` when ready to publish a beta release. This sets the `beta` tag explicitly
for every package and preserves existing stable `latest` tags. For first-ever package
publications, npm also initializes `latest` to that first version; installation instructions
use `@beta` explicitly. Changesets prerelease mode does not accept an additional `--tag` option. Complete npm two-factor
authentication in your terminal when prompted. To inspect the upload without publishing, run
`pnpm release:beta --dry-run`.
There is deliberately no workflow that publishes on push. npm may accept an upload before its
publish-time scan makes the version installable. Wait for every package and dependency to appear
in the registry, then verify an installation in a clean directory. Do not republish or bump versions
merely because an accepted upload is not immediately visible.

To inspect a single package without publishing:

```sh
pnpm --filter @lucamattiazzi/sommelier-client pack --pack-destination ../../artifacts/npm
```

## Add-in and relay

The app builds to `apps/addin/dist`; the bundled relay builds to `apps/server/dist`.
Choose the real public HTTPS origin, then generate and build:

```sh
SOMMELIER_PUBLIC_ORIGIN=https://sommelier.example.com pnpm manifest:production
pnpm sommelier:build
```

Replace `sommelier.example.com` with the deployment domain. The generated manifest is copied into the
add-in build. See [Docker/Caddy deployment](../deploy/sommelier/README.md).

The development manifest remains useful for local sideloading; restore it with
`pnpm manifest:generate` after preparing production artifacts.

Before an AppSource submission, validate the manifest and complete the real Excel/harness checks
in [the testing guide](pair-testing.md). Support/setup/data-handling pages and sized icons are included; see the [asset inventory](marketplace-assets.md).
Production hosting, publisher contact/legal details, real Excel screenshots and Microsoft review
remain separate publication work.
