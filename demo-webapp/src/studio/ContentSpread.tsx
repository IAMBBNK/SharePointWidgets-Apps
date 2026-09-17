import { useState } from 'react'

const files = [
  { name: 'Copy of site pages (2)', rot: -8, x: '8%', y: '10%', after: 'Hub home', kind: 'Page' },
  { name: 'governance_FINAL_really.xlsx', rot: 6, x: '18%', y: '28%', after: 'Role matrix', kind: 'List' },
  { name: 'IMG_4401.png', rot: -3, x: '28%', y: '48%', after: 'Training catalog', kind: 'Library' },
  { name: 'who-owns-this.txt', rot: 9, x: '12%', y: '64%', after: 'Owners & backups', kind: 'Governance' },
  { name: 'news dump', rot: -11, x: '36%', y: '18%', after: 'Stories', kind: 'News' },
]

export function ContentSpread() {
  const [tidy, setTidy] = useState(false)

  return (
    <section id="content" className="st-chapter" aria-labelledby="content-title">
      <div className="st-chapter-n" aria-hidden="true">
        02
      </div>
      <p className="st-kicker">02 — Content</p>
      <h2 id="content-title" className="st-display" style={{ fontSize: 'clamp(2.4rem, 5vw, 4.2rem)', maxWidth: '12ch' }}>
        Filing cabinets don’t have a pulse.
      </h2>
      <div className="st-content-grid" style={{ marginTop: 36 }}>
        <div>
          <p className="st-body">
            I take ownership of the SharePoint people already have. Information
            architecture, audits, governance, page hygiene, findability. The
            unglamorous work that makes everything else possible.
          </p>
          <p className="st-body" style={{ marginTop: 16 }}>
            Sometimes that means deleting forty “final” documents. Sometimes it
            means teaching a site owner how to publish without a ticket.
          </p>
          <button type="button" className="st-btn" style={{ marginTop: 24 }} onClick={() => setTidy(!tidy)}>
            {tidy ? 'Show the mess' : 'Tidy the mess'}
          </button>
        </div>
        <div className={`st-messy ${tidy ? 'tidy' : ''}`} aria-live="polite">
          {files.map((f) => (
            <article
              key={f.name}
              className="st-file"
              style={
                tidy
                  ? undefined
                  : { transform: `rotate(${f.rot}deg)`, left: f.x, top: f.y }
              }
            >
              {tidy ? f.after : f.name}
              <small>{tidy ? f.kind : 'orphaned'}</small>
            </article>
          ))}
        </div>
      </div>
    </section>
  )
}
