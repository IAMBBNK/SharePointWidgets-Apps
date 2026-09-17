import { useMemo, useState } from 'react'
import { Link } from 'react-router-dom'

const seed = [
  { id: 1, title: 'Harbour QC went live', base: 41 },
  { id: 2, title: 'Jump host, in plain language', base: 28 },
  { id: 3, title: 'Who owns North Plant data', base: 19 },
  { id: 4, title: 'September digest', base: 67 },
]

export function PulsePlayground() {
  const [period, setPeriod] = useState<'week' | 'month'>('week')
  const [reads, setReads] = useState(() => Object.fromEntries(seed.map((p) => [p.id, p.base])))
  const [beat, setBeat] = useState(0)

  const total = useMemo(
    () => Object.values(reads).reduce((s, n) => s + n, 0),
    [reads],
  )
  const openRate = period === 'week' ? 38 + beat : 44 + beat
  const stories = period === 'week' ? 6 : 19

  function bump(id: number) {
    setReads((prev) => ({ ...prev, [id]: prev[id] + 1 }))
    setBeat((b) => Math.min(b + 2, 12))
  }

  return (
    <section id="systems" className="st-chapter st-night" aria-labelledby="systems-title">
      <div className="st-chapter-n" aria-hidden="true">
        04
      </div>
      <p className="st-kicker">04 — Systems</p>
      <h2 id="systems-title" className="st-display" style={{ fontSize: 'clamp(2.4rem, 5vw, 4.2rem)', maxWidth: '14ch' }}>
        Sometimes you need the thing that isn’t in the box.
      </h2>
      <p className="st-lede" style={{ marginTop: 16 }}>
        Custom web parts, dashboards, maps, surveys — built to feel native on a
        modern page. SharePoint is the platform. The experience is the craft.
      </p>
      <p className="st-body" style={{ marginTop: 12 }}>
        Don’t take my word for it. Poke this pulse. Then open the full hub demo
        and change the look of a live web part.
      </p>

      <div className="st-pulse">
        <div className="st-pulse-head">
          <h3>Hub pulse</h3>
          <div className="st-seg" role="group" aria-label="Reporting period">
            <button type="button" className={period === 'week' ? 'on' : ''} onClick={() => setPeriod('week')}>
              Week
            </button>
            <button type="button" className={period === 'month' ? 'on' : ''} onClick={() => setPeriod('month')}>
              Month
            </button>
          </div>
        </div>
        <div className="st-metrics">
          <div className="st-metric">
            <b>{openRate}%</b>
            <span>Pages worth opening</span>
          </div>
          <div className="st-metric">
            <b>{stories}</b>
            <span>Stories published</span>
          </div>
          <div className="st-metric">
            <b>{total}</b>
            <span>Reads on the four below</span>
          </div>
        </div>
        <svg className="st-heart" viewBox="0 0 320 48" aria-hidden="true">
          <path d="M0 28h40l8-16 10 32 8-22 12 8h40l10-18 8 28 12-14h172" />
        </svg>
        <ul className="st-page-list">
          {seed.map((p) => (
            <li key={p.id}>
              <button type="button" onClick={() => bump(p.id)}>
                <span>{p.title}</span>
                <em>{reads[p.id]} reads</em>
                <em>Give it a pulse</em>
              </button>
            </li>
          ))}
        </ul>
      </div>
      <div className="st-split-links" style={{ marginTop: 28 }}>
        <Link className="st-btn" to="/demo/catalog">
          Play with the catalog
        </Link>
        <Link className="st-btn ghost" to="/demo/kpi">
          Open a KPI web part
        </Link>
      </div>
    </section>
  )
}
