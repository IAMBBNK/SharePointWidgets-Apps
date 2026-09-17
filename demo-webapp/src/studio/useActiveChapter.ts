import { useEffect, useState } from 'react'
import { chapters, type ChapterId } from './chapters'

export function useActiveChapter() {
  const [active, setActive] = useState<ChapterId>(chapters[0].id)

  useEffect(() => {
    const nodes = chapters
      .map((c) => document.getElementById(c.id))
      .filter((el): el is HTMLElement => Boolean(el))
    if (!nodes.length) return

    const seen = new Map<string, number>()
    const io = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          seen.set(entry.target.id, entry.intersectionRatio)
        }
        let best: ChapterId = chapters[0].id
        let ratio = -1
        for (const c of chapters) {
          const r = seen.get(c.id) ?? 0
          if (r > ratio) {
            ratio = r
            best = c.id
          }
        }
        if (ratio > 0) setActive(best)
      },
      { threshold: [0, 0.2, 0.45, 0.7], rootMargin: '-20% 0px -40% 0px' },
    )
    nodes.forEach((n) => io.observe(n))
    return () => io.disconnect()
  }, [])

  return active
}
