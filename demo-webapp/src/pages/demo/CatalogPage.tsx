import { Link } from 'react-router-dom'
import { catalog, groupLabels, type OfferGroup } from '../../data/catalog'
import { WebPartFrame } from '../../components/WebPartFrame'

const order: OfferGroup[] = ['content', 'extension', 'confluence', 'addon']

export function CatalogPage() {
  return (
    <WebPartFrame title="Web part catalog">
      <p className="catalog-intro">
        Every demo uses mock data and SharePoint-like chrome. Open a card,
        then use the look-and-feel pane on the right — color, corners, and
        density should follow the widget, not just the buttons.
      </p>
      {order.map((group) => (
        <div key={group}>
          <div className="group-label">{groupLabels[group]}</div>
          <div className="catalog-grid">
            {catalog
              .filter((item) => item.group === group)
              .map((item) => (
                <Link key={item.slug} className="catalog-card" to={item.path}>
                  <span className="chip">{item.usedFor}</span>
                  <h3>{item.title}</h3>
                  <p>{item.blurb}</p>
                </Link>
              ))}
          </div>
        </div>
      ))}
    </WebPartFrame>
  )
}
