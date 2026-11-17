# ForWork brand artwork

English | [中文](README.zh.md)

The ForWork mark is a half-open door: a square frame with a leaf swung inward and a round handle. It stands for entering a workspace and getting to work, without borrowing a letter, animal, or AI sparkle.

## Construction

The canonical mark uses a 24-unit square. The frame occupies `x=4..20, y=3..21` with a 2.5-unit left post, a 2.5-unit right post, and a 2.5-unit top bar. The leaf hinges at `(7,7)` and runs to `(16,9)`, `(16,19)`, and `(7,21)`. The handle is a `0.85`-unit radius circle at `(14,14)`; it is a hole in the leaf, not a second ink.

- Mark: `forwork-mark.svg` (monochrome `currentColor`, no handle; use the tile when the background is known)
- Tile: `forwork-tile.svg` (1024 units, `#F4F2EC` field, ink door, handle cutout)
- Lockup: `forwork-lockup.svg` (182 x 24; 156 x 24 without the mark)
- Installer: `installer-brand-light.svg`, `installer-brand-dark.svg` (600 x 196), `uninstaller-sidebar.svg` (164 x 314)

## Palette

| Role | Value |
| --- | --- |
| Tile field | `#F4F2EC` |
| Tile hairline | `#D8D3C8` |
| Mark ink | `#141414` |
| Wordmark ink (dark surface) | `#F9FAFB` |

The mark is flat: no gradient, glow, or secondary hue. The only depth is the tile and mark drop shadow on desktop icon sources.

## Sizes and safety

- Mark: minimum 16 px; sidebar and tray use 24 px and 16 px respectively.
- Tile: 1024-unit master. Maskable Web icons scale the artwork to `0.95` so the door stays inside the `0.4` safe radius.
- Desktop sources keep the 1104-unit canvas and the 40/32 tile inset: `apps/desktop/resources/icon.svg` is the square crop source, `icon-windows.svg` uses a 192 radius, and `icon-macos.svg` uses 230.
- The tray source exposes the door as `id="tray-glyph"` and reuses the tile background.

## Type

The wordmark sets `ForWork` in the platform UI stack (`Segoe UI`, `Helvetica Neue`, Arial) at weight 600. It is not outlined; `textLength="148"` keeps the 182/156 artboard width stable across installed fonts.

## Motion

On the blank home, hovering the mark rotates the leaf `-7deg` around its hinge over 700 ms. The hover state is only installed when hover is available and reduced motion is not requested.

## Derived assets

- Desktop icons: `apps/desktop/resources/icon{,-windows,-macos}.{svg,png}`, `tray-windows.ico`
- Web install icons: `apps/web/public/icon-{192,512}.png`, `icon-maskable-512.png`, `apple-touch-icon.png`
- Favicons: `apps/web/public/favicon{,-dark}.svg`, `website/public/favicon.svg`
- Website wordmark: `website/public/wordmark.svg`
- Welcome and installer: `apps/desktop/renderer/assets/welcome-brand.svg`, `apps/desktop/installer/assets/brand{,-2x,-dark,-dark-2x}.png`

Regenerate the Web icons with `corepack pnpm --filter @forwork/web-frontend run render:app-icons`, the tray with `corepack pnpm run render:tray-icon` in `apps/desktop`, and the installer PNGs from `installer-brand-{light,dark}.svg` with sharp.

## Home frame layer

The official `@forwork/client-ui-brand-official` package no longer occupies `conversation.hero.frame`; the DeepSeek whale atlas and its renderer were removed with T136. The slot stays declared for other bundles, and the blank home keeps the static door mark with its hover leaf.
