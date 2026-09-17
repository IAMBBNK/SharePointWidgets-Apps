import { useState } from 'react'
import { surveyMilestones } from '../data/mock'

const labels = ['#1 Best', '#2', '#3', '#4', '#5', '#6 Worst']

export function SurveyWidget() {
  const [role, setRole] = useState('Site lead')
  const [slots, setSlots] = useState<Array<(typeof surveyMilestones)[0] | null>>(
    Array(6).fill(null),
  )
  const [picked, setPicked] = useState<string | null>(null)
  const [worked, setWorked] = useState('')
  const [didnt, setDidnt] = useState('')
  const [consent, setConsent] = useState(false)
  const [error, setError] = useState('')
  const [done, setDone] = useState(false)

  const pool = surveyMilestones.filter((m) => !slots.some((s) => s?.id === m.id))

  function placeIn(index: number) {
    const next = [...slots]
    if (picked) {
      const from = surveyMilestones.find((m) => m.id === picked)
      if (!from) return
      const existing = next[index]
      const oldIndex = next.findIndex((s) => s?.id === picked)
      if (oldIndex >= 0) next[oldIndex] = existing
      next[index] = from
      setSlots(next)
      setPicked(null)
      return
    }
    if (next[index]) {
      next[index] = null
      setSlots(next)
    }
  }

  function submit() {
    if (slots.some((s) => !s)) {
      setError('Rank all six milestones from best to worst.')
      return
    }
    if (!worked.trim() || !didnt.trim()) {
      setError('Tell us what worked and what didn’t.')
      return
    }
    if (!consent) {
      setError('Confirm follow-up consent to submit in this demo.')
      return
    }
    setError('')
    setDone(true)
  }

  if (done) {
    return (
      <div className="banner-ok">
        Thank you, {role}. In production this stores a response in a SharePoint list.
      </div>
    )
  }

  return (
    <div>
      <div className="dash-header">
        <div>
          <h2 className="dash-title">OT integration survey</h2>
          <p className="dash-sub">
            Rank all six milestones. Click a pill, then an empty slot (or swap with a filled one).
          </p>
        </div>
      </div>
      {error ? <div className="banner-err">{error}</div> : null}
      <div className="field" style={{ maxWidth: 280, marginBottom: 12 }}>
        <label>Your role</label>
        <select value={role} onChange={(e) => setRole(e.target.value)}>
          <option>Site lead</option>
          <option>Architect</option>
          <option>Service manager</option>
        </select>
      </div>
      <strong>Unranked milestones</strong>
      <div className="survey-pool">
        {pool.length === 0 ? <span className="hint-line">All placed.</span> : null}
        {pool.map((m) => (
          <button
            key={m.id}
            type="button"
            className={`pill ${picked === m.id ? 'on' : ''}`}
            onClick={() => setPicked(picked === m.id ? null : m.id)}
          >
            {m.id}: {m.name}
          </button>
        ))}
      </div>
      <div className="rank-board">
        {labels.map((label, i) => (
          <button
            key={label}
            type="button"
            className={`rank-slot ${slots[i] ? 'filled' : ''}`}
            onClick={() => placeIn(i)}
          >
            <span>{label}</span>
            {slots[i] ? `${slots[i]!.id}: ${slots[i]!.name}` : 'Click to place'}
          </button>
        ))}
      </div>
      <div className="form-grid">
        <div className="field">
          <label>What worked (best milestone)</label>
          <textarea rows={2} value={worked} onChange={(e) => setWorked(e.target.value)} />
        </div>
        <div className="field">
          <label>What didn’t work (worst milestone)</label>
          <textarea rows={2} value={didnt} onChange={(e) => setDidnt(e.target.value)} />
        </div>
      </div>
      <label className="look-check">
        <input type="checkbox" checked={consent} onChange={(e) => setConsent(e.target.checked)} />
        Follow-up is OK (demo consent — not stored)
      </label>
      <button type="button" className="sp-btn" style={{ marginTop: 12 }} onClick={submit}>
        Submit survey
      </button>
    </div>
  )
}
