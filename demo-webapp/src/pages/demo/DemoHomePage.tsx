import { Link } from 'react-router-dom'
import { catalog } from '../../data/catalog'
import { newsItems } from '../../data/mock'
import { WebPartFrame } from '../../components/WebPartFrame'

export function DemoHomePage() {
  const featured = catalog.filter((item) =>
    ['kpi', 'map', 'trainings', 'carousel', 'engage', 'newsletter'].includes(item.slug),
  )

  return (
    <>
      <div className="welcome">
        <p className="eyebrow">You’re in the hub</p>
        <h2>Pick something to try. Then play with the look.</h2>
        <p>
          This is a sample intranet — mock data only, no live tenant. Open a
          web part from the cards below, then use the side panel to change
          color, corners, density, and buttons. It should feel like sitting
          next to someone who can show you the page, not a spec.
        </p>
      </div>
      <WebPartFrame title="Start here">
        <div className="highlights">
          {featured.map((item) => (
            <Link key={item.slug} className="catalog-card" to={item.path}>
              <span className="chip">{item.usedFor}</span>
              <h3>{item.title}</h3>
              <p>{item.blurb}</p>
            </Link>
          ))}
        </div>
      </WebPartFrame>
      <div className="news-grid">
        <WebPartFrame title="What’s going on">
          {newsItems.map((item) => (
            <div key={item.title} className="news-item">
              <div className="kicker">{item.kicker}</div>
              <h3>{item.title}</h3>
              <p>{item.text}</p>
            </div>
          ))}
        </WebPartFrame>
        <WebPartFrame title="Handy links">
          <p>
            <Link to="/demo/catalog">Full web part catalog</Link>
          </p>
          <p>
            <Link to="/demo/confluence">Operations space (Confluence)</Link>
          </p>
          <p>
            <Link to="/demo/survey">Integration feedback survey</Link>
          </p>
          <p>
            <Link to="/">Back to services</Link>
          </p>
        </WebPartFrame>
      </div>
    </>
  )
}
