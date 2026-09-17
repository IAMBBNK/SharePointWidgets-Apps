import { useState } from 'react'
import { newsletterEditions } from '../data/mock'

export function NewsletterWidget() {
  const [i, setI] = useState(0)
  const ed = newsletterEditions[i]
  const openRate = Math.round((ed.opens7[ed.opens7.length - 1] / ed.delivered) * 100)
  const clickRate = Math.round((ed.clicks7[ed.clicks7.length - 1] / ed.delivered) * 100)
  const maxOpen = Math.max(...ed.opens24)
  const maxOpen7 = Math.max(...ed.opens7)
  const maxClick7 = Math.max(...ed.clicks7)

  return (
    <div>
      <div className="sp-filters">
        {newsletterEditions.map((e, idx) => (
          <button
            key={e.title}
            type="button"
            className={`sp-btn ${i === idx ? 'active' : 'ghost'}`}
            onClick={() => setI(idx)}
          >
            {e.title}
          </button>
        ))}
      </div>
      <p className="dash-sub" style={{ marginBottom: 12 }}>
        Sent {ed.sent} · Delivered to {ed.delivered.toLocaleString()} · {ed.notes}
      </p>
      <div className="stat-row">
        <div className="stat">
          <b>{openRate}%</b>
          <span>7-day open rate</span>
        </div>
        <div className="stat">
          <b>{clickRate}%</b>
          <span>7-day click rate</span>
        </div>
        <div className="stat">
          <b>{ed.opens24[ed.opens24.length - 1]}</b>
          <span>Opens in 24h</span>
        </div>
        <div className="stat">
          <b>{ed.clicks24[ed.clicks24.length - 1]}</b>
          <span>Clicks in 24h</span>
        </div>
      </div>
      <div className="chart-title">Opens over 24 hours</div>
      <div className="spark" aria-hidden="true">
        {ed.opens24.map((v, idx) => (
          <i key={idx} style={{ height: `${(v / maxOpen) * 100}%` }} />
        ))}
      </div>
      <div className="chart-title">Clicks over 24 hours</div>
      <div className="spark spark-alt" aria-hidden="true">
        {ed.clicks24.map((v, idx) => (
          <i key={idx} style={{ height: `${(v / Math.max(...ed.clicks24)) * 100}%` }} />
        ))}
      </div>
      <div className="chart-title">7-day opens</div>
      <div className="spark" aria-hidden="true">
        {ed.opens7.map((v, idx) => (
          <i key={`o7-${idx}`} style={{ height: `${(v / maxOpen7) * 100}%` }} />
        ))}
      </div>
      <div className="chart-title">7-day clicks</div>
      <div className="spark spark-alt" aria-hidden="true">
        {ed.clicks7.map((v, idx) => (
          <i key={`c7-${idx}`} style={{ height: `${(v / maxClick7) * 100}%` }} />
        ))}
      </div>
      <h3 className="chart-title">Stories</h3>
      <table className="simple-table">
        <thead>
          <tr>
            <th>Story</th>
            <th>Clicks</th>
            <th>% of page views</th>
          </tr>
        </thead>
        <tbody>
          {ed.stories.map((s) => (
            <tr key={s.title}>
              <td>{s.title}</td>
              <td>{s.clicks}</td>
              <td>{Math.round((s.clicks / s.pageViews) * 100)}%</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
