import { createContext, useContext, useMemo, useState, type CSSProperties, type ReactNode } from 'react'

export type LookFeel = {
  accent: string
  radius: 'sharp' | 'hub' | 'soft'
  density: 'comfortable' | 'compact'
  buttons: 'filled' | 'outline'
  showFrameTitle: boolean
  headers: 'hub' | 'neutral'
}

export const defaultLook: LookFeel = {
  accent: '#0f69af',
  radius: 'hub',
  density: 'comfortable',
  buttons: 'filled',
  showFrameTitle: false,
  headers: 'hub',
}

const accents = [
  { id: 'hub', label: 'Hub blue', value: '#0f69af' },
  { id: 'm365', label: 'Sky blue', value: '#0078d4' },
  { id: 'teal', label: 'Teal', value: '#0f7b7e' },
  { id: 'green', label: 'Leaf', value: '#218a3a' },
  { id: 'purple', label: 'Violet', value: '#5b5fc7' },
] as const

const LookCtx = createContext<{
  look: LookFeel
  setLook: (next: LookFeel) => void
  reset: () => void
} | null>(null)

export function useLookFeel() {
  const ctx = useContext(LookCtx)
  if (!ctx) {
    throw new Error('useLookFeel needs LookFeelProvider')
  }
  return ctx
}

export function LookFeelProvider({ children }: { children: ReactNode }) {
  const [look, setLook] = useState<LookFeel>(defaultLook)
  const value = useMemo(
    () => ({
      look,
      setLook,
      reset: () => setLook(defaultLook),
    }),
    [look],
  )
  return <LookCtx.Provider value={value}>{children}</LookCtx.Provider>
}

export function lookStyle(look: LookFeel): CSSProperties {
  const radius = look.radius === 'sharp' ? '0px' : look.radius === 'soft' ? '14px' : '5px'
  return {
    ['--wp-accent' as string]: look.accent,
    ['--wp-title' as string]: look.headers === 'hub' ? look.accent : '#1f1f1f',
    ['--wp-radius' as string]: radius,
  }
}

export function LookFeelPanel() {
  const { look, setLook, reset } = useLookFeel()
  return (
    <aside className="look-panel">
      <div className="look-kicker">Look & feel</div>
      <h2>Customize</h2>
      <p className="look-hint">
        Play with this the way a customer would in the web part pane. Color,
        corners, density, and buttons should all move together.
      </p>
      <div
        className="look-preview"
        style={{
          background: `linear-gradient(135deg, ${look.accent}, color-mix(in oklab, ${look.accent} 72%, #7ec8e8))`,
          borderRadius: look.radius === 'sharp' ? 0 : look.radius === 'soft' ? 14 : 5,
        }}
      >
        Live preview
      </div>

      <div className="look-section">Accent color</div>
      <div className="swatches">
        {accents.map((a) => (
          <button
            key={a.id}
            type="button"
            className={`swatch ${look.accent === a.value ? 'on' : ''}`}
            style={{ background: a.value }}
            title={a.label}
            onClick={() => setLook({ ...look, accent: a.value })}
          />
        ))}
      </div>
      <label className="look-color">
        Custom
        <input
          type="color"
          value={look.accent}
          onChange={(e) => setLook({ ...look, accent: e.target.value })}
        />
      </label>

      <div className="look-section">Corners</div>
      <div className="look-choices">
        {(
          [
            ['sharp', 'Sharp'],
            ['hub', 'Hub (5px)'],
            ['soft', 'Soft'],
          ] as const
        ).map(([id, label]) => (
          <button
            key={id}
            type="button"
            className={look.radius === id ? 'on' : ''}
            onClick={() => setLook({ ...look, radius: id })}
          >
            {label}
          </button>
        ))}
      </div>

      <div className="look-section">Density</div>
      <div className="look-choices">
        {(
          [
            ['comfortable', 'Comfortable'],
            ['compact', 'Compact'],
          ] as const
        ).map(([id, label]) => (
          <button
            key={id}
            type="button"
            className={look.density === id ? 'on' : ''}
            onClick={() => setLook({ ...look, density: id })}
          >
            {label}
          </button>
        ))}
      </div>

      <div className="look-section">Buttons</div>
      <div className="look-choices">
        {(
          [
            ['filled', 'Filled'],
            ['outline', 'Outline'],
          ] as const
        ).map(([id, label]) => (
          <button
            key={id}
            type="button"
            className={look.buttons === id ? 'on' : ''}
            onClick={() => setLook({ ...look, buttons: id })}
          >
            {label}
          </button>
        ))}
      </div>

      <div className="look-section">Headers</div>
      <div className="look-choices">
        {(
          [
            ['hub', 'Colored titles'],
            ['neutral', 'Neutral titles'],
          ] as const
        ).map(([id, label]) => (
          <button
            key={id}
            type="button"
            className={look.headers === id ? 'on' : ''}
            onClick={() => setLook({ ...look, headers: id })}
          >
            {label}
          </button>
        ))}
      </div>

      <label className="look-check">
        <input
          type="checkbox"
          checked={look.showFrameTitle}
          onChange={(e) => setLook({ ...look, showFrameTitle: e.target.checked })}
        />
        Show web part title bar
      </label>

      <button type="button" className="look-reset" onClick={reset}>
        Reset to hub default
      </button>
    </aside>
  )
}
