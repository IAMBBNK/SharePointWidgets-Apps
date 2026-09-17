import { useState } from 'react'
import { projects } from '../data/mock'

const statusClass: Record<string, string> = {
  'On track': 'ok',
  'At risk': 'risk',
  Done: 'done',
}

export function ProjectsWidget() {
  const [status, setStatus] = useState('all')
  const [open, setOpen] = useState<string | null>(null)
  const visible = projects.filter((p) => status === 'all' || p.status === status)

  return (
    <div>
      <div className="dash-header">
        <div>
          <h2 className="dash-title">Projects board</h2>
          <p className="dash-sub">Portfolio items for the hub program. Open a card for the next step.</p>
        </div>
      </div>
      <div className="map-filters">
        {['all', 'On track', 'At risk', 'Done'].map((s) => (
          <button
            key={s}
            type="button"
            className={`sp-btn ${status === s ? 'active' : 'ghost'}`}
            onClick={() => setStatus(s)}
          >
            {s === 'all' ? 'All' : s}
          </button>
        ))}
      </div>
      {visible.length === 0 ? (
        <div className="empty-state">No projects in this status.</div>
      ) : (
        <div className="proj-grid">
          {visible.map((p) => (
            <button
              key={p.title}
              type="button"
              className={`proj-cell as-btn ${open === p.title ? 'open' : ''}`}
              onClick={() => setOpen(open === p.title ? null : p.title)}
            >
              <h3>{p.title}</h3>
              <span className={`status ${statusClass[p.status]}`}>{p.status}</span>
              <div className="proj-owner">{p.owner}</div>
              <p>{p.note}</p>
              {open === p.title ? <p className="tl-outcome">Next: {p.next}</p> : null}
            </button>
          ))}
        </div>
      )}
    </div>
  )
}
