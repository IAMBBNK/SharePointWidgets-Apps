import { useEffect, useMemo, useState } from 'react'
import { carouselSlides } from '../data/mock'

export function CarouselWidget() {
  const [open, setOpen] = useState(false)
  const [index, setIndex] = useState(0)
  const [visible, setVisible] = useState<1 | 3>(3)
  const count = visible
  const len = carouselSlides.length

  const cards = useMemo(() => {
    const out = []
    for (let i = 0; i < count; i++) {
      out.push(carouselSlides[(index + i) % len])
    }
    return out
  }, [index, count, len])

  useEffect(() => {
    if (!open) return
    const prev = document.body.style.overflow
    document.body.style.overflow = 'hidden'
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(false)
      if (e.key === 'ArrowLeft') setIndex((i) => (i - 1 + len) % len)
      if (e.key === 'ArrowRight') setIndex((i) => (i + 1) % len)
    }
    window.addEventListener('keydown', onKey)
    return () => {
      document.body.style.overflow = prev
      window.removeEventListener('keydown', onKey)
    }
  }, [open, len])

  return (
    <div>
      <div className="dash-header">
        <div>
          <h2 className="dash-title">Image gallery</h2>
          <p className="dash-sub">A button opens a popup carousel — 1 or 3 cards, as on the hub. Esc and arrows work while it is open.</p>
        </div>
      </div>
      <div className="sp-filters">
        <button type="button" className="imgc-trigger" onClick={() => setOpen(true)}>
          View gallery
        </button>
        <button
          type="button"
          className={`sp-btn ghost ${visible === 1 ? 'active' : ''}`}
          onClick={() => setVisible(1)}
        >
          1 picture
        </button>
        <button
          type="button"
          className={`sp-btn ghost ${visible === 3 ? 'active' : ''}`}
          onClick={() => setVisible(3)}
        >
          3 pictures
        </button>
      </div>
      {open ? (
        <div className="imgc-overlay" onClick={(e) => e.target === e.currentTarget && setOpen(false)}>
          <button type="button" className="imgc-close" onClick={() => setOpen(false)}>
            ×
          </button>
          <div>
            <div className="imgc-stage" style={{ width: visible === 1 ? 440 : undefined }}>
              <button
                type="button"
                className="imgc-nav"
                onClick={() => setIndex((i) => (i - 1 + len) % len)}
              >
                ‹
              </button>
              <div className="imgc-cards">
                {cards.map((card) => (
                  <article key={`${card.title}-${card.tone}`} className="imgc-card">
                    <div className={`imgc-photo tone-${card.tone}`} role="img" aria-label={card.title} />
                    <h3>{card.title}</h3>
                    <p>{card.description}</p>
                  </article>
                ))}
              </div>
              <button
                type="button"
                className="imgc-nav"
                onClick={() => setIndex((i) => (i + 1) % len)}
              >
                ›
              </button>
            </div>
            <div className="imgc-dots">
              {carouselSlides.map((s, i) => (
                <button
                  key={`${s.title}-${i}`}
                  type="button"
                  className={`imgc-dot ${i === index ? 'on' : ''}`}
                  onClick={() => setIndex(i)}
                  aria-label={`Go to ${s.title}`}
                />
              ))}
            </div>
          </div>
        </div>
      ) : null}
    </div>
  )
}
