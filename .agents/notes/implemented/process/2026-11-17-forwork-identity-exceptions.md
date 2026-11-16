# Agent Note: ForWork identity exceptions for protocol tokens and the external Office kit

Status: implemented

English | [中文](2026-11-17-forwork-identity-exceptions.zh.md)

## Problem

[verify-forwork-identity](../../../../scripts/verify-forwork-identity.ts) rejects DeepSeek identity tokens from shipped surfaces: the `@deepseek-ai` package scope, the DeepSeek Harness brand, the legacy `DEEPSEEK_*` environment names, the legacy `dsh` command names, and the legacy artifact prefixes. Two groups of names cannot become ForWork names by a string replacement, and an unexplained pass over them is indistinguishable from a missed violation.

The first group is protocol vocabulary that shipped profiles, published plugin manifests, and existing user homes already carry: the `dsh.bundle`, `dsh.client`, and `dsh.profile` package-manifest keys, the `DSH_*` environment family that includes `$DSH_HOME`, and the `dsh-app://` Electron protocol.

The second group is `@deepseek-ai/libreoffice-kit` together with its `-darwin-*`, `-win32-*`, and `-wasm` engine packages, the external engine behind legacy `.doc`, `.xls`, and `.ppt` conversion.

## Decision

Both groups keep their current names, and each exception names what must change before a name moves. The protocol exception is a decision not to add a rule that names the tokens; the Office exception is one entry in an existing allowlist. Neither relaxes the remaining rules.

### Protocol tokens

No identity rule targets these tokens, so the scan admits them by construction. Each one is a published protocol fact:

- `dsh.bundle` and `dsh.profile` are package-manifest keys. The profile composer resolves patch layers in `dsh.profile.bundles` order ([profile-boot.ts](../../../../apps/cli/src/profile-boot.ts)), and the plugin manager reads published registry metadata at `dsh.bundle` to decide whether an installed version is a bundle and to locate its patch files through `dsh.bundle.patch` ([operations.ts](../../../../packages/boot/plugin-manager/src/operations.ts)).
- `dsh.client` declares a package's browser entry, external runtime specifiers, and platform for the client bundles ([verify-client-packages.ts](../../../../scripts/verify-client-packages.ts)).
- `$DSH_HOME` and the wider `DSH_*` environment family name the harness home for the desktop app, the Python SDK, the e2e harness, and user configuration. `$DSH_HOME` defaults to `~/.dsh`, and existing sessions, settings, credentials, profiles, and attachments live under that path.
- `dsh-app://` is the Electron custom protocol that packaged desktop builds register for the application document and its local assets ([web-document.ts](../../../../apps/desktop/src/web-document.ts)).

Renaming the manifest keys changes the profile composer, the plugin manager, the client build, and every published third-party plugin that declares them; a Harness-side rename alone makes those plugins install as plain dependencies instead of profile layers. Renaming the environment family without an in-place migration leaves existing homes unread, so an upgraded install starts empty while the old data stays on disk.

The tokens remain the shipped protocol vocabulary, not identity violations. Moving them is justified only together with a `SESSION_FORMAT_VERSION` or persistence amendment that defines the successor, plus an in-place migration for existing homes; the release carries neither.

### External Office kit

`@deepseek-ai/libreoffice-kit` and its engine packages are third-party identities that stay external:

- ForWork neither publishes nor vendors the engine, and `THIRD_PARTY_NOTICES.md` names exactly these package identities at MPL-2.0 with corresponding-source pointers to the kit repository.
- The [independent kit ownership](../architecture/2026-09-14-independent-libreoffice-kit.md) decision keeps engine qualification and releases in that repository.
- A ForWork build of the engine would mean publishing a multi-gigabyte native artifact, qualifying it for every platform, and taking over the MPL-2.0 obligations and the `THIRD_PARTY_NOTICES.md` entries that `corepack pnpm run verify-third-party-notices` requires.

The identity scan's `ALLOWED_DEEPSEEK_PACKAGES` holds exactly `@deepseek-ai/libreoffice-kit`. The allowlist matches that literal and its `-<suffix>` engine packages, and nothing else in the `@deepseek-ai` scope. It stays an exact package allowlist; admitting the whole scope would let any upstream package keep its name.

## Alternatives considered

**Migrate the protocol tokens behind a session-format amendment in this release.** The release has no such amendment and no home migration, so the rename would split storage and break installed plugins for a naming gain that the protocol does not require.

**Replace the Office engine with a ForWork-published or vendored kit.** Owning the build means publishing and qualifying the native artifacts for every platform and carrying their source obligations in this repository, while giving up the upstream `THIRD_PARTY_NOTICES.md` entries and corresponding-source links that `corepack pnpm run verify-third-party-notices` verifies.

**Admit the whole `@deepseek-ai` scope in the identity scan.** A scope-wide allowance stops reporting any new upstream dependency, so the scan would no longer distinguish a reviewed third-party engine from a leaked upstream package.

## Consequences

Shipped homes, published plugins, and packaged desktop builds keep working, and the pinned kit version keeps the conversion path and its notices intact.

The cost is that the repository keeps `dsh`-prefixed protocol names and one `@deepseek-ai` package while presenting ForWork identity. The scan cannot show from its output alone that these remaining names are intentional, so the allowlist comment points here as the record a reviewer follows.

`verify-forwork-identity` still fails on any other DeepSeek identity token, and each new external DeepSeek package requires an allowlist edit and an update here. The protocol tokens stay under no rule, so a future rule that adds one must cite this note or supersede it.

Revisit the protocol exception only with a `SESSION_FORMAT_VERSION` or persistence amendment and a home migration, and the Office exception only if ForWork publishes its own kit or replaces the conversion engine.
