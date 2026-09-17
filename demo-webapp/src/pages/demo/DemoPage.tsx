import type { ReactNode } from 'react'
import { itemBySlug, type CatalogItem } from '../../data/catalog'
import { DemoBlurb, WebPartFrame } from '../../components/WebPartFrame'
import { LookFeelPanel, LookFeelProvider, lookStyle, useLookFeel } from '../../theme/LookFeel'

function Stage({
  item,
  children,
}: {
  item: CatalogItem
  children: ReactNode
}) {
  const { look } = useLookFeel()
  return (
    <div className="demo-stage">
      <div>
        <DemoBlurb item={item} />
        <div
          className={`wp-stage ${look.density === 'compact' ? 'is-compact' : ''} ${look.buttons === 'outline' ? 'is-outline' : ''}`}
          style={lookStyle(look)}
        >
          <WebPartFrame title={look.showFrameTitle ? item.title : ''}>{children}</WebPartFrame>
        </div>
      </div>
      <LookFeelPanel />
    </div>
  )
}

export function DemoPage({
  slug,
  children,
}: {
  slug: string
  children: ReactNode
}) {
  const item = itemBySlug(slug) as CatalogItem
  return (
    <LookFeelProvider>
      <Stage item={item}>{children}</Stage>
    </LookFeelProvider>
  )
}
