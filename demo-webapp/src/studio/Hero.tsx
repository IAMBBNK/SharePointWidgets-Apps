import { useEffect, useRef } from 'react'
import { Link } from 'react-router-dom'

const ignored = [
  'Q3_update_FINAL2.pptx',
  'Please read',
  'New site request',
  'Policy_v12.docx',
  'All-hands deck',
  'untitled page',
  'Copy of Copy of News',
  'Important!!!',
]

const worth = [
  'The packing line is live',
  "Maya's first week on nights",
  'Why the jump host changed',
  'Harbour QC, in their words',
  'A map people actually use',
  'The forum nobody skipped',
]

const scraps = [
  { t: 'final_v7.docx', x: '42%', y: '10%' },
  { t: 'news — draft', x: '52%', y: '22%' },
  { t: 'KPI — broken', x: '38%', y: '68%' },
  { t: 'who owns this?', x: '8%', y: '78%' },
]

export function Hero() {
  const field = useRef<HTMLElement>(null)

  useEffect(() => {
    const root = field.current
    if (!root) return
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return
    if (window.matchMedia('(pointer: coarse)').matches) return

    const onMove = (e: PointerEvent) => {
      const r = root.getBoundingClientRect()
      const px = (e.clientX - r.left) / r.width - 0.5
      const py = (e.clientY - r.top) / r.height - 0.5
      root.querySelectorAll<HTMLElement>('.st-floater').forEach((el, i) => {
        const d = (i + 1) * 8
        el.style.transform = `translate(${px * d}px, ${py * d}px)`
      })
    }
    root.addEventListener('pointermove', onMove)
    return () => root.removeEventListener('pointermove', onMove)
  }, [])

  return (
    <section id="problem" className="st-hero" ref={field} aria-labelledby="hero-title">
      <div className="st-floaters" aria-hidden="true">
        {scraps.map((s) => (
          <span key={s.t} className="st-floater" style={{ left: s.x, top: s.y }}>
            {s.t}
          </span>
        ))}
      </div>
      <div className="st-hero-inner">
        <p className="st-kicker">01 — The problem</p>
        <h1 id="hero-title" className="st-display st-hero-title">
          Make your
          <br />
          intranet
          <br />
          <em>worth</em>
          <br />
          opening.
        </h1>
      </div>
      <div className="st-hero-copy">
        <p className="st-lede">
          I keep SharePoint useful, alive, and worth visiting. I talk to people,
          structure the pages, and build the pieces that don’t exist yet.
        </p>
        <div className="st-disciplines" aria-label="Three disciplines">
          <a className="st-disc" href="#content">
            <b>Content</b>
            <span>Keep the place structured.</span>
          </a>
          <a className="st-disc" href="#stories">
            <b>Stories</b>
            <span>Find what’s interesting.</span>
          </a>
          <a className="st-disc" href="#systems">
            <b>Systems</b>
            <span>Build what’s missing.</span>
          </a>
        </div>
        <div className="st-split-links">
          <Link className="st-btn" to="/demo">
            Step into a living hub
          </Link>
          <a className="st-btn ghost" href="#talk">
            Start a conversation
          </a>
        </div>
      </div>
      <div className="st-marquee" aria-hidden="true">
        <div className="st-marquee-lab">Nobody opens</div>
        <div className="st-marquee-clip">
          <div className="st-marquee-track">
            {[...ignored, ...ignored].map((t, i) => (
              <span key={`${t}-${i}`}>{t}</span>
            ))}
          </div>
        </div>
      </div>
      <div className="st-marquee reverse" aria-hidden="true">
        <div className="st-marquee-lab good">People open</div>
        <div className="st-marquee-clip">
          <div className="st-marquee-track">
            {[...worth, ...worth].map((t, i) => (
              <span key={`${t}-${i}`}>
                <b>{t}</b>
              </span>
            ))}
          </div>
        </div>
      </div>
    </section>
  )
}
