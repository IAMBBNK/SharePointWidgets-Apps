import { useMemo, useState } from 'react'
import { filterSites, sites, type Site } from '../data/mock'

/** Section meters on each site row. */
const sectionLabels = [
  ['general', 'General'],
  ['people', 'People'],
  ['maturity', 'Maturity'],
  ['counts', 'Counts'],
] as const

export function CompletionWidget() {
  const [sector, setSector] = useState('')
  const [region, setRegion] = useState('')
  const [q, setQ] = useState('')
  const [open, setOpen] = useState<number | null>(null)
  const [sort, setSort] = useState<'asc' | 'desc'>('asc')

  const visible = useMemo(() => {
    const base = filterSites(sites, sector, region).filter((s) => {
      const hay = `${s.name} ${s.location} ${s.unit}`.toLowerCase()
      return hay.includes(q.toLowerCase())
    })
    return [...base].sort((a, b) =>
      sort === 'asc' ? a.completeness - b.completeness : b.completeness - a.completeness,
    )
  }, [sector, region, q, sort])

  const avg = visible.length
    ? Math.round(visible.reduce((sum, s) => sum + s.completeness, 0) / visible.length)
    : 0

  return (
    <div>
      <div className="dash-header">
        <div>
          <h2 className="dash-title">Site completion dashboard</h2>
          <p className="dash-sub">
            Field-level completeness across sites so content owners know what is still missing.
          </p>
        </div>
      </div>
      <section className="sp-filters" aria-label="Completion filters">
        <div className="filter-group">
          <label htmlFor="cRegion">Region</label>
          <select id="cRegion" value={region} onChange={(e) => setRegion(e.target.value)}>
            <option value="">All Regions</option>
            <option value="EMEA">EMEA</option>
            <option value="AMER">AMER</option>
            <option value="APAC">APAC</option>
          </select>
        </div>
        <div className="filter-group">
          <label htmlFor="cSector">Sector</label>
          <select id="cSector" value={sector} onChange={(e) => setSector(e.target.value)}>
            <option value="">All Sectors</option>
            <option>Manufacturing</option>
            <option>R&D</option>
            <option>Corporate</option>
          </select>
        </div>
        <label className="sp-search">
          <span aria-hidden="true">⌕</span>
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search site or location" />
        </label>
        <button type="button" className="reset-button" onClick={() => setSort(sort === 'asc' ? 'desc' : 'asc')}>
          Sort {sort === 'asc' ? 'low → high' : 'high → low'}
        </button>
      </section>
      <div className="stat-row">
        <div className="stat">
          <b>{visible.length}</b>
          <span>Sites</span>
        </div>
        <div className="stat">
          <b>{avg}%</b>
          <span>Average completeness</span>
        </div>
        <div className="stat">
          <b>{visible.filter((s) => s.completeness === 100).length}</b>
          <span>Fully complete</span>
        </div>
        <div className="stat">
          <b>{visible.filter((s) => s.completeness < 70).length}</b>
          <span>Need attention</span>
        </div>
      </div>
      {visible.length === 0 ? (
        <div className="empty-state">No sites match these filters.</div>
      ) : (
        visible.map((site) => <CompleteRow key={site.id} site={site} open={open} setOpen={setOpen} />)
      )}
    </div>
  )
}

function CompleteRow({
  site,
  open,
  setOpen,
}: {
  site: Site
  open: number | null
  setOpen: (id: number | null) => void
}) {
  const expanded = open === site.id
  return (
    <div className="complete-row">
      <header>
        <button type="button" className="linkish" onClick={() => setOpen(expanded ? null : site.id)}>
          {site.name} · {site.location}
        </button>
        <strong>{site.completeness}%</strong>
      </header>
      <div className="meter">
        <i style={{ width: `${site.completeness}%` }} />
      </div>
      <div className="section-meters">
        {sectionLabels.map(([key, label]) => (
          <div key={key} className="section-meter">
            <span>
              {label} {site.sections[key]}%
            </span>
            <div className="meter">
              <i style={{ width: `${site.sections[key]}%` }} />
            </div>
          </div>
        ))}
      </div>
      {expanded ? (
        <div className="missing-box">
          {site.missing.length === 0 ? (
            'No missing fields.'
          ) : (
            <>
              Still missing: {site.missing.join(' · ')}
            </>
          )}
        </div>
      ) : (
        <div className="hint-line">Click the site name to see missing fields.</div>
      )}
    </div>
  )
}
