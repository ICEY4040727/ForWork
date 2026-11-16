import { FishLogo } from '@forwork/client-ui-primitives'
import type { SidebarBrandMarkOwnerProps, SidebarBrandNameOwnerProps } from '@forwork/client-ui-sidebar/client'

/**
 * Render the official mark with the presentation requested by its host surface.
 * @param props - Host-supplied mark presentation.
 * @returns the current mark glyph.
 * TODO(forwork-brand): replace the fish glyph with approved ForWork mark artwork.
 */
export function OfficialBrandMark({ size }: SidebarBrandMarkOwnerProps) {
  return <FishLogo size={size} />
}

/**
 * Render the official name without its independently slotted mark.
 * @returns the neutral ForWork name.
 */
export function OfficialBrandName({ t }: SidebarBrandNameOwnerProps) {
  return <span>{t('brand.name')}</span>
}
