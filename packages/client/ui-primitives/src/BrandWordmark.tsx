import type { IconProps } from './icons/props.ts'
import { BRAND_MARK_FRAME_PATH, BRAND_MARK_LEAF_PATH } from './BrandMark.tsx'

/** Display options for the official brand wordmark. */
export interface BrandWordmarkProps extends IconProps {
  /** Whether to include the leading door mark; defaults to true. */
  includeMark?: boolean | undefined
}

/** Font stack the standalone lockup also names, so the component and SVG agree. */
const WORDMARK_FONT_FAMILY = "'Segoe UI', 'Helvetica Neue', Arial, sans-serif"

/**
 * Render the full brand wordmark.
 * @param props.size - height in px (default 24; width follows the selected artwork).
 * @param props.className - extra class for layout placement.
 * @param props.includeMark - whether to include the leading door mark.
 * @returns the wordmark svg (aria-hidden decorative brand art).
 */
export function BrandWordmark({ size = 24, className, includeMark = true }: BrandWordmarkProps) {
  const width = includeMark ? 182 : 156
  return (
    <svg
      width={(size * width) / 24}
      height={size}
      className={className}
      viewBox={includeMark ? '0 0 182 24' : '26 0 156 24'}
      fill="none"
      aria-hidden="true"
    >
      {includeMark && (
        <g transform="translate(2 2) scale(0.833333)" fill="currentColor">
          <path d={BRAND_MARK_FRAME_PATH} />
          <path d={BRAND_MARK_LEAF_PATH} />
        </g>
      )}
      <text
        x="30"
        y="18.4"
        fill="currentColor"
        fontFamily={WORDMARK_FONT_FAMILY}
        fontSize="21"
        fontWeight="600"
        letterSpacing="-0.4"
        textLength="148"
        lengthAdjust="spacingAndGlyphs"
      >
        ForWork
      </text>
    </svg>
  )
}
