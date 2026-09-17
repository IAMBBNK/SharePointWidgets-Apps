import type { ReactNode } from 'react'
import type { CatalogItem } from '../data/catalog'

export function WebPartFrame({
  title,
  children,
}: {
  title: string
  children: ReactNode
}) {
  return (
    <section className="wp">
      {title ? <div className="wp-title">{title}</div> : null}
      <div className="wp-body">{children}</div>
    </section>
  )
}

export function DemoBlurb({ item }: { item: CatalogItem }) {
  return (
    <div className="demo-blurb">
      <span className={`chip chip-${item.usedFor === 'Add-on' ? 'addon' : item.usedFor.toLowerCase()}`}>
        {item.usedFor}
      </span>
      <p>{item.blurb}</p>
    </div>
  )
}
