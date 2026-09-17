import { chapters, type ChapterId } from './chapters'

export function ChapterNav({ active }: { active: ChapterId }) {
  return (
    <>
      <nav className="st-rail" aria-label="Chapters">
        {chapters.map((c) => (
          <a key={c.id} href={`#${c.id}`} className={active === c.id ? 'on' : undefined}>
            <span className="st-rail-n">{c.n}</span>
            <span className="st-rail-lab">{c.label}</span>
          </a>
        ))}
      </nav>
      <nav className="st-chipbar" aria-label="Chapters">
        {chapters.map((c) => (
          <a key={c.id} href={`#${c.id}`} className={active === c.id ? 'on' : undefined}>
            {c.n} {c.label}
          </a>
        ))}
      </nav>
    </>
  )
}
