import { useMemo, useState } from 'react'
import { DestinationModal } from '../components/DestinationModal'
import { trainings } from '../data/mock'

const domains = [
  { id: 'operating', label: 'IT/OT Operating Model', cls: 'operating' },
  { id: 'technology', label: 'IT/OT Technology & Architecture', cls: 'technology' },
  { id: 'cyber', label: 'IT/OT Cybersecurity', cls: 'cyber' },
  { id: 'service', label: 'IT/OT Service Management', cls: 'service' },
]

export function TrainingsWidget() {
  const [domain, setDomain] = useState('all')
  const [curriculaOnly, setCurriculaOnly] = useState(false)
  const [open, setOpen] = useState<string | null>(null)
  const [go, setGo] = useState<(typeof trainings)[0] | null>(null)
  const items = useMemo(
    () =>
      trainings.filter(
        (t) =>
          (domain === 'all' || t.domain === domain) && (!curriculaOnly || t.curricula),
      ),
    [domain, curriculaOnly],
  )

  return (
    <div>
      <div className="dash-header">
        <div>
          <h2 className="dash-title">Training portal</h2>
          <p className="dash-sub">
            Filter by domain. Open a row for a short insight, or Go to a mock course page.
          </p>
        </div>
      </div>
      <div className="sp-filters filter-quad">
        {domains.map((d) => (
          <button
            key={d.id}
            type="button"
            className={`round-btn ${d.cls} ${domain === d.id ? 'active' : ''}`}
            onClick={() => setDomain(domain === d.id ? 'all' : d.id)}
          >
            <span className="round-btn-label">{d.label}</span>
          </button>
        ))}
      </div>
      <label className="look-check" style={{ margin: '0 0 12px' }}>
        <input
          type="checkbox"
          checked={curriculaOnly}
          onChange={(e) => setCurriculaOnly(e.target.checked)}
        />
        Curricula items only
      </label>
      {items.length === 0 ? (
        <div className="empty-state">No trainings match this filter. Clear a domain or curricula.</div>
      ) : (
        items.map((item) => (
          <div key={item.name} className="acc">
            <div
              className={`acc-head ${item.domain} ${item.curricula ? 'curricula' : ''} ${open === item.name ? 'open' : ''}`}
              onClick={() => setOpen(open === item.name ? null : item.name)}
            >
              <div className="acc-main">
                <div className="acc-title">{item.name}</div>
                <div className="acc-cat">{item.category}</div>
              </div>
              <span className={`badge ${item.internal ? 'internal' : 'external'}`}>
                {item.internal ? 'Internal' : 'External'}
              </span>
              <div className="acc-actions">
                <span className="more-info">
                  More info <span className="arrow">{open === item.name ? '▴' : '▾'}</span>
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
            <div className={`acc-body ${open === item.name ? 'open' : ''}`}>{item.insight}</div>
          </div>
        ))
      )}
      {go ? (
        <DestinationModal title={go.name} body={go.insight} onClose={() => setGo(null)} />
      ) : null}
    </div>
  )
}
