import { useMemo, useState, type FormEvent } from 'react'
import { maintenanceEntries, sites } from '../data/mock'

const assessmentOpts = ['Completed', 'Ongoing', 'Planning', 'N/A']
const north = sites.find((s) => s.id === 1)!

const siteDefaults = {
  name: north.name,
  sector: north.sector,
  location: north.location,
  employees: String(north.employees),
  itAssessment: north.itAssessment as string,
  otAssessment: north.otAssessment as string,
  integration: north.integration as string,
  platform: north.platform as string,
  notes: 'Integration wave 2 scheduled for Q4.',
}

export function SiteFormWidget() {
  const [done, setDone] = useState(false)
  const [error, setError] = useState('')
  const [form, setForm] = useState(siteDefaults)

  function revert() {
    setForm(siteDefaults)
    setError('')
  }

  function onSubmit(e: FormEvent) {
    e.preventDefault()
    if (!form.name.trim() || !form.location.trim()) {
      setError('Site name and location are required.')
      return
    }
    if (!form.employees.trim() || Number.isNaN(Number(form.employees))) {
      setError('Employees must be a number.')
      return
    }
    setError('')
    setDone(true)
  }

  if (done) {
    return (
      <div className="banner-ok">
        Saved in this demo only. In production this writes to a SharePoint list.
      </div>
    )
  }

  return (
    <form onSubmit={onSubmit}>
      <div className="dash-header">
        <div>
          <h2 className="dash-title">Site Master Data update</h2>
          <p className="dash-sub">Signed in as Jordan Lee (site owner). Update fields for your plant.</p>
        </div>
      </div>
      {error ? <div className="banner-err">{error}</div> : null}

      <h3 className="form-section">Site</h3>
      <div className="form-grid">
        <div className="field">
          <label>Site name</label>
          <input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
        </div>
        <div className="field">
          <label>Sector</label>
          <select value={form.sector} onChange={(e) => setForm({ ...form, sector: e.target.value })}>
            <option>Manufacturing</option>
            <option>R&D</option>
            <option>Corporate</option>
          </select>
        </div>
        <div className="field">
          <label>Location</label>
          <input value={form.location} onChange={(e) => setForm({ ...form, location: e.target.value })} />
        </div>
        <div className="field">
          <label>Employees</label>
          <input value={form.employees} onChange={(e) => setForm({ ...form, employees: e.target.value })} />
        </div>
      </div>

      <h3 className="form-section">People</h3>
      <div className="chip-row">
        <span className="person-chip">Jordan Lee · Site owner</span>
        <span className="person-chip">A. Keller · Architecture</span>
        <span className="person-chip ghost">Add backup (demo)</span>
      </div>

      <h3 className="form-section">Assessments</h3>
      <div className="form-grid">
        <div className="field">
          <label>IT assessment</label>
          <select value={form.itAssessment} onChange={(e) => setForm({ ...form, itAssessment: e.target.value })}>
            {assessmentOpts.map((o) => (
              <option key={o}>{o}</option>
            ))}
          </select>
        </div>
        <div className="field">
          <label>OT assessment</label>
          <select value={form.otAssessment} onChange={(e) => setForm({ ...form, otAssessment: e.target.value })}>
            {assessmentOpts.map((o) => (
              <option key={o}>{o}</option>
            ))}
          </select>
        </div>
        <div className="field">
          <label>Integration</label>
          <select value={form.integration} onChange={(e) => setForm({ ...form, integration: e.target.value })}>
            {assessmentOpts.map((o) => (
              <option key={o}>{o}</option>
            ))}
          </select>
        </div>
        <div className="field">
          <label>Target platform</label>
          <select value={form.platform} onChange={(e) => setForm({ ...form, platform: e.target.value })}>
            <option>Platinum</option>
            <option>Premium</option>
            <option>Essential</option>
            <option>N/A</option>
          </select>
        </div>
        <div className="field full">
          <label>Notes</label>
          <textarea rows={3} value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} />
        </div>
      </div>
      <div className="form-actions">
        <button type="button" className="reset-button" onClick={revert}>
          Revert
        </button>
        <button type="submit" className="sp-btn">
          Submit
        </button>
      </div>
    </form>
  )
}

export function MaintenanceFormWidget() {
  const [done, setDone] = useState(false)
  const [error, setError] = useState('')
  const [entry, setEntry] = useState('new')
  const [siteId, setSiteId] = useState(String(sites[0].id))
  const [asset, setAsset] = useState('')
  const [severity, setSeverity] = useState('Medium')
  const [observation, setObservation] = useState('')
  const site = useMemo(
    () => sites.find((s) => String(s.id) === siteId) ?? sites[0],
    [siteId],
  )

  function onSubmit(e: FormEvent) {
    e.preventDefault()
    if (!asset.trim()) {
      setError('Asset / system is required.')
      return
    }
    if (!observation.trim()) {
      setError('Describe the observation so the CoE can route it.')
      return
    }
    setError('')
    setDone(true)
  }

  if (done) {
    return (
      <div className="banner-ok">
        Saved in this demo only. In production this writes to a SharePoint list
        {entry !== 'new' ? ` as an update to ${entry}` : ''}.
      </div>
    )
  }

  return (
    <form onSubmit={onSubmit}>
      <div className="dash-header">
        <div>
          <h2 className="dash-title">Predictive maintenance intake</h2>
          <p className="dash-sub">
            Capture a predictive signal so the CoE can route it to the right plant owner.
          </p>
        </div>
      </div>
      {error ? <div className="banner-err">{error}</div> : null}
      <div className="form-grid">
        <div className="field">
          <label>Entry</label>
          <select
            value={entry}
            onChange={(e) => {
              setEntry(e.target.value)
              if (e.target.value === 'PM-1044') {
                setSiteId('1')
                setAsset('Mixer-04 historian')
              }
              if (e.target.value === 'PM-0981') {
                setSiteId('10')
                setAsset('Pack-line PLC')
              }
            }}
          >
            {maintenanceEntries.map((m) => (
              <option key={m.id} value={m.id}>
                {m.label}
              </option>
            ))}
          </select>
        </div>
        <div className="field">
          <label>Site</label>
          <select value={siteId} onChange={(e) => setSiteId(e.target.value)}>
            {sites.map((s) => (
              <option key={s.id} value={s.id}>
                {s.name}
              </option>
            ))}
          </select>
        </div>
        <div className="field">
          <label>Prefill</label>
          <input readOnly value={`${site.sector} · ${site.location}`} />
        </div>
        <div className="field">
          <label>Asset / system</label>
          <input value={asset} onChange={(e) => setAsset(e.target.value)} placeholder="e.g. Mixer-04 historian" />
        </div>
        <div className="field">
          <label>Severity</label>
          <select value={severity} onChange={(e) => setSeverity(e.target.value)}>
            <option>Low</option>
            <option>Medium</option>
            <option>High</option>
          </select>
        </div>
        <div className="field">
          <label>Owner</label>
          <div className="chip-row" style={{ marginTop: 6 }}>
            <span className="person-chip">Jordan Lee · Site owner</span>
          </div>
        </div>
        <div className="field full">
          <label>Observation</label>
          <textarea
            rows={3}
            value={observation}
            onChange={(e) => setObservation(e.target.value)}
            placeholder="What did you see, and since when?"
          />
        </div>
      </div>
      <div className="form-actions">
        <button type="submit" className="sp-btn">
          Submit
        </button>
      </div>
    </form>
  )
}
