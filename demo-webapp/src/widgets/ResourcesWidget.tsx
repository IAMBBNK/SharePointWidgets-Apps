import { useMemo, useState } from 'react'
import { DestinationModal } from '../components/DestinationModal'
import { resources } from '../data/mock'

export function ResourcesWidget() {
  const [q, setQ] = useState('')
  const [cat, setCat] = useState('all')
  const [open, setOpen] = useState<string | null>(resources[0]?.id ?? null)
  const [go, setGo] = useState<(typeof resources)[0] | null>(null)
  const filtered = useMemo(
    () =>
      resources.filter((item) => {
        const hay = `${item.name} ${item.sector} ${item.id} ${item.info}`.toLowerCase()
        const okQ = hay.includes(q.toLowerCase())
        const okC = cat === 'all' || item.category === cat
        return okQ && okC
      }),
    [q, cat],
  )

  return (
    <div>
      <div className="dash-header">
        <div>
          <h2 className="dash-title">Key resources</h2>
          <p className="dash-sub">Use the filters or enter a search term:</p>
        </div>
      </div>
      <div className="sp-filters">
        <button
          type="button"
          className={`round-btn ${cat === 'Standards & Procedures' ? 'active' : ''}`}
          onClick={() =>
            setCat(cat === 'Standards & Procedures' ? 'all' : 'Standards & Procedures')
          }
        >
          <span className="round-btn-label">Standards & Procedures</span>
        </button>
        <button
          type="button"
          className={`round-btn tools ${cat === 'Tools & Applications' ? 'active' : ''}`}
          onClick={() =>
            setCat(cat === 'Tools & Applications' ? 'all' : 'Tools & Applications')
          }
        >
          <span className="round-btn-label">Tools & Applications</span>
        </button>
        <label className="sp-search">
          <span aria-hidden="true">⌕</span>
          <input
            value={q}
            onChange={(e) => setQ(e.target.value)}
            placeholder="Enter your search term here..."
          />
        </label>
      </div>
      {filtered.length === 0 ? (
        <div className="empty-state">No resources match “{q}”. Try another term or clear the category.</div>
      ) : (
        filtered.map((item) => (
          <div key={item.id} className="acc">
            <div
              className={`acc-head tools ${open === item.id ? 'open' : ''}`}
              onClick={() => setOpen(open === item.id ? null : item.id)}
            >
              <div className="acc-main">
                <div className="acc-title">{item.name}</div>
                <div className="acc-cat">{item.sector}</div>
              </div>
              <div className="badge">{item.id}</div>
              <div className="acc-actions">
                <span className="more-info">
                  More info <span className="arrow">{open === item.id ? '▴' : '▾'}</span>
                </span>
                <button
                  type="button"
                  className="go-btn"
                  onClick={(e) => {
                    e.stopPropagation()
                    setGo(item)
                  }}
                >
                  Go to
                </button>
              </div>
            </div>
            <div className={`acc-body ${open === item.id ? 'open' : ''}`}>{item.info}</div>
          </div>
        ))
      )}
      {go ? (
        <DestinationModal title={go.name} body={go.info} onClose={() => setGo(null)} />
      ) : null}
    </div>
  )
}
