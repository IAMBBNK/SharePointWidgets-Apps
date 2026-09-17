import { useMemo, useState } from 'react'
import { filterSites, otKpiPeriods, pageKpis, sites, tally } from '../data/mock'

export function KpiWidget() {
  const [sector, setSector] = useState('')
  const [region, setRegion] = useState('')
  const visible = useMemo(() => filterSites(sites, sector, region), [sector, region])
  const reviewed = visible.filter((s) => s.completeness >= 90).length
  const charts = [
    { title: 'System integration', categories: tally(visible, 'integration') },
    { title: 'IT assessment', categories: tally(visible, 'itAssessment') },
    { title: 'OT assessment', categories: tally(visible, 'otAssessment') },
    { title: 'Target platform', categories: tally(visible, 'platform') },
  ]
  const otApps = visible.reduce((s, x) => s + x.otApps, 0)
  const otToMigrate = visible.reduce((s, x) => s + x.otToMigrate, 0)
  const integrated = otApps ? Math.round(((otApps - otToMigrate) / otApps) * 100) : 0

  return (
    <div>
      <div className="dash-header">
        <div>
          <h2 className="dash-title">Site Master Data — KPI Dashboard</h2>
          <p className="dash-sub">
            Live overview of site assessments, implementation status and roadmap maturity.
          </p>
        </div>
        <div className="dash-sub">Last updated 12 Sep 2026</div>
      </div>
      <section className="sp-filters" aria-label="Dashboard filters">
        <div className="filter-group">
          <label htmlFor="filterRegion">Region</label>
          <select id="filterRegion" value={region} onChange={(e) => setRegion(e.target.value)}>
            <option value="">All Regions</option>
            <option value="EMEA">EMEA</option>
            <option value="AMER">AMER</option>
            <option value="APAC">APAC</option>
          </select>
        </div>
        <div className="filter-group">
          <label htmlFor="filterSector">Sector</label>
          <select id="filterSector" value={sector} onChange={(e) => setSector(e.target.value)}>
            <option value="">All Sectors</option>
            <option>Manufacturing</option>
            <option>R&D</option>
            <option>Corporate</option>
          </select>
        </div>
        <button
          type="button"
          className="reset-button"
          onClick={() => {
            setSector('')
            setRegion('')
          }}
        >
          Reset filters
        </button>
      </section>
      <section className="kpi-tiles">
        <div className="kpi-tile">
          <div className="kpi-tile-label">Total Sites</div>
          <div className="kpi-tile-value">{visible.length}</div>
          <div className="kpi-tile-sub">All sites currently in scope</div>
        </div>
        <div className="kpi-tile secondary">
          <div className="kpi-tile-label">Reviewed / Updated</div>
          <div className="kpi-tile-value">{reviewed}</div>
          <div className="kpi-tile-sub">Completeness ≥ 90%</div>
        </div>
      </section>
      <div className="ot-strip">
        <div>
          <b>{otApps}</b>
          <span>OT apps in scope</span>
        </div>
        <div>
          <b>{otToMigrate}</b>
          <span>Still to migrate</span>
        </div>
        <div>
          <b>{integrated}%</b>
          <span>Already on target platform</span>
        </div>
      </div>
      <div className="charts">
        {charts.map((kpi) => {
          const total = kpi.categories.reduce((s, c) => s + c.value, 0) || 1
          return (
            <div key={kpi.title} className="chart-card">
              <h3 className="chart-title">{kpi.title}</h3>
              <div className="stack">
                {kpi.categories.map((c, idx) => (
                  <span
                    key={c.label}
                    className={`seg-${idx % 4}`}
                    style={{ width: `${(c.value / total) * 100}%` }}
                    title={`${c.label}: ${c.value}`}
                  />
                ))}
              </div>
              <ul className="legend">
                {kpi.categories.map((c, idx) => (
                  <li key={c.label}>
                    <i className={`swatch-box seg-${idx % 4}`} />
                    <span className="legend-label">{c.label}</span>
                    <strong>{c.value}</strong>
                  </li>
                ))}
              </ul>
            </div>
          )
        })}
      </div>
    </div>
  )
}

const periodKeys = Object.keys(otKpiPeriods) as Array<keyof typeof otKpiPeriods>

export function OtKpiWidget() {
  const [period, setPeriod] = useState<(typeof periodKeys)[number]>('Q3 2026')
  const [openInfo, setOpenInfo] = useState<string | null>(null)
  const [collapsed, setCollapsed] = useState<Record<string, boolean>>({})
  const groups = otKpiPeriods[period]

  return (
    <div>
      <div className="dash-header">
        <div>
          <h2 className="dash-title">Operations KPI dashboard</h2>
          <p className="dash-sub">Published operational KPIs grouped by category, with targets and formulas.</p>
        </div>
        <div className="filter-group" style={{ minWidth: 160 }}>
          <label htmlFor="otPeriod">Reporting period</label>
          <select
            id="otPeriod"
            value={period}
            onChange={(e) => setPeriod(e.target.value as (typeof periodKeys)[number])}
          >
            {periodKeys.map((p) => (
              <option key={p}>{p}</option>
            ))}
          </select>
        </div>
      </div>
      {groups.map((group) => {
        const avg = Math.round(
          group.items.reduce((s, i) => s + i.value, 0) / group.items.length,
        )
        const shut = collapsed[group.category]
        return (
          <div key={group.category} className="ot-cat">
            <button
              type="button"
              className="ot-cat-head"
              onClick={() => setCollapsed({ ...collapsed, [group.category]: !shut })}
            >
              <h3>{group.category}</h3>
              <span>Avg {avg}%</span>
            </button>
            {shut
              ? null
              : group.items.map((item) => {
                  const tone = item.value >= item.target ? 'ok' : item.value >= item.target - 15 ? 'warn' : 'bad'
                  return (
                    <div key={item.title} className="ot-row">
                      <header>
                        <span>
                          {item.title}{' '}
                          <button
                            type="button"
                            className="info-dot"
                            aria-label={`About ${item.title}`}
                            onClick={() =>
                              setOpenInfo(openInfo === item.title ? null : item.title)
                            }
                          >
                            ⓘ
                          </button>
                        </span>
                        <span>
                          {item.value}% / target {item.target}%
                        </span>
                      </header>
                      <div className={`meter tone-${tone}`}>
                        <i style={{ width: `${item.value}%` }} />
                      </div>
                      {openInfo === item.title ? (
                        <p className="ot-info">
                          {item.description} Formula: {item.formula}.
                        </p>
                      ) : null}
                    </div>
                  )
                })}
          </div>
        )
      })}
    </div>
  )
}

export function PageKpiWidget() {
  const [q, setQ] = useState('')
  const rows = pageKpis.topPages.filter((p) => {
    const hay = `${p.title} ${p.section} ${p.type}`.toLowerCase()
    return hay.includes(q.toLowerCase())
  })
  const menus = rows.filter((p) => p.type === 'menu')
  const articles = rows.filter((p) => p.type === 'article')

  return (
    <div>
      <div className="dash-header">
        <div>
          <h2 className="dash-title">Page & engagement KPIs</h2>
          <p className="dash-sub">Top pages, hub sections, viewers, and Viva referrals.</p>
        </div>
        <div className="dash-sub">Last updated {pageKpis.lastUpdated}</div>
      </div>
      <label className="sp-search" style={{ maxWidth: 360, marginBottom: 16 }}>
        <span aria-hidden="true">⌕</span>
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search title, section, or type" />
      </label>
      <div className="stat-row">
        {pageKpis.totals.map((t) => (
          <div key={t.label} className="stat">
            <b>{t.value}</b>
            <span>{t.label}</span>
          </div>
        ))}
      </div>
      <PageTable title="Menu pages" rows={menus} />
      <PageTable title="Articles" rows={articles} />
    </div>
  )
}

function PageTable({
  title,
  rows,
}: {
  title: string
  rows: (typeof pageKpis.topPages)[number][]
}) {
  return (
    <div className="page-block">
      <h3 className="chart-title">{title}</h3>
      {rows.length === 0 ? (
        <div className="empty-state">No pages in this group.</div>
      ) : (
        <table className="simple-table">
          <thead>
            <tr>
              <th>Page</th>
              <th>Hub section</th>
              <th>Type</th>
              <th>Views</th>
              <th>Viewers</th>
              <th>Viva</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((p) => (
              <tr key={p.title}>
                <td>{p.title}</td>
                <td>{p.section}</td>
                <td>
                  <span className="badge">{p.type}</span>
                </td>
                <td>{p.views.toLocaleString()}</td>
                <td>{p.viewers.toLocaleString()}</td>
                <td>{p.viva}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  )
}
