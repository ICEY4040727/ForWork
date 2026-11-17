import type { IconProps } from './icons/props.ts'

/** Native viewBox of the ForWork mark (square user units). */
export const BRAND_MARK_VIEWBOX = { width: 24, height: 24 }

/** Door frame paths: left post, right post, and top bar. */
export const BRAND_MARK_FRAME_PATH =
  'M4 3h2.5v18H4V3Z M17.5 5.5H20V21h-2.5V5.5Z M4 3h16v2.5H4V3Z'

/** Door leaf path, exported so the hero can animate it around the hinge. */
export const BRAND_MARK_LEAF_PATH = 'M7 7l9 2v10l-9 2V7Z'

/** Complete mark silhouette for consumers that compose their own svg around the same geometry. */
export const BRAND_MARK_PATH = `${BRAND_MARK_FRAME_PATH} ${BRAND_MARK_LEAF_PATH}`

/**
 * Render the ForWork half-open door mark.
 * @param props.size - square edge in px (default 24).
 * @param props.className - extra class for layout placement.
 * @returns the mark svg (aria-hidden; pair with the wordmark for accessibility).
 */
export function BrandMark({ size = 24, className }: IconProps) {
  return (
    <svg
      width={size}
      height={size}
      className={className}
      viewBox={`0 0 ${String(BRAND_MARK_VIEWBOX.width)} ${String(BRAND_MARK_VIEWBOX.height)}`}
      fill="none"
      aria-hidden="true"
    >
      <path d={BRAND_MARK_FRAME_PATH} fill="currentColor" />
      <path d={BRAND_MARK_LEAF_PATH} fill="currentColor" />
    </svg>
  )
}
