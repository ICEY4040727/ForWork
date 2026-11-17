# Contract: ForWork Brand Artwork and Visual Evidence

**Owner**: Brand masters `assets/brand/`; gate `scripts/verify-brand-assets.ts`; evidence suites `apps/web/tests`, `apps/desktop/tests`, `website/tests`
**Purpose**: Keep every shipped ForWork graphic derivable, owned, and visually proven, and keep the retired DeepSeek artwork out of the product (research R21-R24, FR-009, SC-006).

## Shipped assets

The following paths are the contract. A change to a path, viewBox, or pixel size is a contract change and updates the manifest in the same commit.

| Asset | Path | Size | Owner |
|---|---|---|---|
| Monochrome mark | `assets/brand/forwork-mark.svg` | 24×24 viewBox | Brand master |
| App tile master | `assets/brand/forwork-tile.svg` | 1024×1024 | Brand master |
| Lockup | `assets/brand/forwork-lockup.svg` | 182×24 viewBox | Brand master |
| Desktop square crop source | `apps/desktop/resources/icon.svg` | 1104×1104 | Desktop |
| Desktop Windows | `apps/desktop/resources/icon-windows.svg` | 1024×1024 | Desktop |
| Desktop macOS | `apps/desktop/resources/icon-macos.svg` | 1024×1024 | Desktop |
| Desktop PNGs | `apps/desktop/resources/icon*.png` | 1104 or 1024 | Desktop |
| Windows tray | `apps/desktop/resources/tray-windows.ico` | 16-64 multi-size | Desktop |
| Web install icons | `apps/web/public/icon-{192,512}.png`, `icon-maskable-512.png`, `apple-touch-icon.png` | Declared sizes | Web |
| Web favicons | `apps/web/public/favicon{,-dark}.svg` | 50×50 viewBox | Web |
| Website wordmark | `website/public/wordmark.svg` | 182×24 viewBox | Website |
| Website favicon | `website/public/favicon.svg` | 50×50 viewBox | Website |
| Welcome art | `apps/desktop/renderer/assets/welcome-brand.svg` | 472×40 | Desktop |
| Installer art | `apps/desktop/installer/assets/brand{,-2x,-dark,-dark-2x}.png` | 600×196 or 1200×392 | Installer |
| Uninstaller art | `apps/desktop/installer/assets/uninstaller-sidebar.png` | 164×314 | Installer |

The canonical door geometry is the closed path in `assets/brand/README.md`. `icon.svg` keeps a square opaque tile because the Web renderer crops it and the PWA contract requires opaque corners. `icon-windows.svg` and `icon-macos.svg` add the rounded platform masks.

## Manifest

`scripts/brand-assets.manifest.json` is the single inventory.

```json
{
  "version": 1,
  "assets": [
    {
      "path": "assets/brand/forwork-mark.svg",
      "kind": "svg-mark",
      "surface": "brand-master",
      "width": 24,
      "height": 24,
      "sha256": "<64 hex>",
      "generator": "authored",
      "required": true
    }
  ],
  "evidence": [
    {
      "id": "web-home-desktop-light",
      "path": "apps/web/tests/expected/brand/home-desktop-light.png",
      "surface": "web",
      "viewport": { "width": 1440, "height": 900 },
      "colorScheme": "light",
      "locale": "en-US",
      "depicts": ["apps/desktop/resources/icon.svg"],
      "sha256": "<64 hex>",
      "capturedBy": "apps/web/tests/brand-artwork.e2e.ts"
    }
  ],
  "fingerprints": [
    {
      "artifact": "deepseek-whale-mark",
      "kind": "svg-path",
      "value": "<retired path-data fragment>",
      "replacement": "assets/brand/forwork-mark.svg"
    }
  ]
}
```

## Verifier

`verify-brand-assets` scans these brand paths, not every public file:

- `assets/brand/**`
- `apps/desktop/resources/icon*` and `apps/desktop/resources/tray-windows.ico`
- `apps/desktop/installer/assets/brand*` and `apps/desktop/installer/assets/uninstaller-sidebar.png`
- `apps/desktop/renderer/assets/welcome-brand.svg`
- `apps/web/public/icon-*.png`, `apple-touch-icon.png`, and `favicon*.svg`
- `website/public/wordmark.svg` and `website/public/favicon.svg`
- `packages/client/ui-brand-official/src/client/assets/**`

Only `.svg`, `.png`, `.ico`, and `.webp` files count as assets.

It fails, with the path and the corrective command, when:

1. A required asset file has no manifest row.
2. A manifest row points to a missing file or has a stale SHA-256.
3. A file matching those brand paths is present but not referenced by any manifest row.
4. A raster dimension or SVG viewBox differs from the manifest.
5. A required evidence file is missing, has the wrong dimensions, or depicts an asset absent from the manifest.
6. An SVG contains a retired `svg-path` fingerprint, or a raster matches a retired `sha256` fingerprint.
7. The evidence inventory lacks a matrix row promised by `plan.md`, or contains an unreferenced image.

A legitimate asset change updates the manifest and the committed screenshots in the same review. `verify-brand-assets --write` refreshes hashes and dimensions only for rows that already exist; it never creates a row for an unreferenced file. The check mode never rewrites files.

## Home frame layer

The official `@forwork/client-ui-brand-official` package no longer registers an occupant for `conversation.hero.frame`. The slot remains declared; other bundles may occupy it. Acceptance checks:

- `packages/client/ui-brand-official/src/client/index.ts` contains no `PixelWhale` import and no `conversation.hero.frame` registration.
- `packages/client/ui-brand-official/src/client/assets/` contains no `breach-atlas.webp`.
- The home e2e asserts no `[data-home-whale]` node and that a blank-area click does not start an animation.
- The sidebar brand registration and the door hero fallback remain unchanged.

## Screenshot evidence

The suites run under `vitest.brand.config.ts` and write fresh images to `.artifacts/brand/` on every run. The committed review set lives under each owner's `tests/expected/brand/` and is refreshed only with `UPDATE_BRAND_SCREENSHOTS=1`; record the new hashes with `verify-brand-assets --write` in the same review.

- Locale `en-US`, `deviceScaleFactor: 1`, animations disabled except the hover capture.
- Web home: 1440×900 light and dark, 390×844 mobile, collapsed rail, hover state.
- Desktop welcome: 600×700 CSS at 2× device scale, light and dark, through the built renderer with the preload bridge mocked.
- Website navigation: 1440×240 desktop and 390×720 mobile from the built site.
- Installer evidence: 600×196 light and dark assets plus the 164×314 uninstaller sidebar.
- A 960×600 hover GIF is recorded from the real server and attached to the pull request.

Pixel equality against the live browser is not asserted across platforms. The normal lane asserts file existence, dimensions, manifest hashes, and the DOM or geometry checks that make the screenshot meaningful.

## Gate wiring

`verify-brand-assets` is a top-level script and joins the shared hygiene gates beside `verify-forwork-identity`. The Web, desktop welcome, and website screenshot suites run under `vitest.brand.config.ts`; the website suite consumes `website:build` output.
