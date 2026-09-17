import { Link } from 'react-router-dom'

export function SharePointServicePage() {
  return (
    <article className="st-article">
      <p className="st-kicker">SharePoint</p>
      <h1>Pages people actually open.</h1>
      <p>
        Two tracks, usually on the same hub: keep the content trustworthy, and
        extend the canvas with things that feel like they belong there. I write,
        I structure, and I build.
      </p>
      <h2>Content</h2>
      <ul>
        <li>Portals that filter instead of dumping a library on someone</li>
        <li>Governance that names a human, not a generic mailbox</li>
        <li>Forms local owners can finish without a ticket</li>
        <li>Completion views so missing fields stop being a rumour</li>
      </ul>
      <h2>Stories</h2>
      <p>
        Interviews with the people who actually do the work. Then a page, a
        digest line, an announcement — written so it gets finished.
      </p>
      <h2>Systems</h2>
      <ul>
        <li>KPI dashboards that recompute from the list, not a screenshot</li>
        <li>Maps, galleries, timelines, surveys that feel native</li>
        <li>The odd widget that isn’t in the box, on purpose</li>
      </ul>
      <div className="st-split-links">
        <Link className="st-btn" to="/demo/catalog">
          Open the hub catalog
        </Link>
        <Link className="st-btn ghost" to="/">
          Back to field.
        </Link>
      </div>
    </article>
  )
}

export function ConfluenceServicePage() {
  return (
    <article className="st-article">
      <p className="st-kicker">Also: Confluence</p>
      <h1>When the runbook shouldn’t live on the hub.</h1>
      <p>
        Some knowledge belongs next to Jira. I set up spaces, trees, and macros
        so the page is navigable — not a warehouse of unloved procedures.
      </p>
      <ul>
        <li>Space trees and templates for operations knowledge</li>
        <li>Macros that read like Confluence, not a pasted Word file</li>
        <li>A cross-link from the SharePoint hub when both exist</li>
      </ul>
      <p>
        Confluence is a neighbour, not a replacement. Hub for community and
        data. Space for the controlled how-to.
      </p>
      <div className="st-split-links">
        <Link className="st-btn" to="/demo/confluence">
          Open the space demo
        </Link>
        <Link className="st-btn ghost" to="/">
          Back to field.
        </Link>
      </div>
    </article>
  )
}
