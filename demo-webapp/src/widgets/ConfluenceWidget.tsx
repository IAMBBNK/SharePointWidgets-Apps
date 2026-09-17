import { useMemo, useState } from 'react'
import { confluencePages } from '../data/mock'

export function ConfluenceWidget() {
  const [id, setId] = useState('outage')
  const [q, setQ] = useState('')
  const page = confluencePages.find((p) => p.id === id) ?? confluencePages[0]
  const crumbs = useMemo(() => {
    const path = []
    let cur: (typeof confluencePages)[0] | undefined = page
    while (cur) {
      path.unshift(cur)
      cur = confluencePages.find((p) => p.id === cur?.parent)
    }
    return path
  }, [page])
  const visible = confluencePages.filter((p) =>
    p.title.toLowerCase().includes(q.toLowerCase()),
  )

  function indent(p: (typeof confluencePages)[0]) {
    let n = 0
    let cur: (typeof confluencePages)[0] | undefined = p
    while (cur?.parent) {
      n += 1
      cur = confluencePages.find((x) => x.id === cur?.parent)
    }
    return n
  }

  return (
    <div className="conf">
      <nav className="conf-tree" aria-label="Space tree">
        <label className="sp-search" style={{ marginBottom: 10 }}>
          <span aria-hidden="true">⌕</span>
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search pages" />
        </label>
        {visible.length === 0 ? <div className="empty-state">No pages.</div> : null}
        {visible.map((p) => (
          <button
            key={p.id}
            type="button"
            className={id === p.id ? 'on' : ''}
            style={{ paddingLeft: 8 + indent(p) * 14 }}
            onClick={() => setId(p.id)}
          >
            {p.title}
          </button>
        ))}
      </nav>
      <article className="conf-page">
        <div className="conf-crumb">
          {crumbs.map((c, i) => (
            <span key={c.id}>
              {i > 0 ? ' / ' : ''}
              <button type="button" className="linkish" onClick={() => setId(c.id)}>
                {c.title}
              </button>
            </span>
          ))}
        </div>
        <h2>{page.title}</h2>
        <p className="conf-body">{page.body}</p>
        <div className="macro macro-info">
          <strong>Info.</strong> Keep this page in sync with the SharePoint site master data for the plant.
        </div>
        {page.attach ? (
          <div className="macro macro-attach">
            <strong>Attachments.</strong> {page.attach}
          </div>
        ) : null}
        <div className="macro macro-people">
          <strong>People.</strong> {page.people}
        </div>
      </article>
    </div>
  )
}
