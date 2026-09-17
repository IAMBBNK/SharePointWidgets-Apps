import { NavLink, Outlet } from 'react-router-dom'
import { useEffect } from 'react'

export function DemoLayout() {
  useEffect(() => {
    const prev = document.title
    document.title = 'Contoso Hub — Demo'
    return () => {
      document.title = prev
    }
  }, [])
  return (
    <div className="demo-shell">
      <div className="suite-bar">
        <button type="button" className="waffle" aria-label="App launcher">
          <span />
          <span />
          <span />
          <span />
          <span />
          <span />
          <span />
          <span />
          <span />
        </button>
        <div className="suite-word">SharePoint</div>
        <div className="suite-search">
          <input type="search" placeholder="Search this site" aria-label="Search" />
        </div>
        <NavLink to="/" className="suite-exit">
          Back to field.
        </NavLink>
        <div className="suite-user" aria-hidden="true">
          JD
        </div>
      </div>
      <header className="site-header">
        <div className="site-logo" aria-hidden="true" />
        <div>
          <div className="site-name">Contoso Hub</div>
          <div className="site-sub">Come in — this is a demo intranet</div>
        </div>
        <nav className="site-nav">
          <NavLink to="/demo" end>
            Home
          </NavLink>
          <NavLink to="/demo/catalog">Catalog</NavLink>
          <NavLink to="/demo/confluence">Confluence</NavLink>
          <NavLink to="/demo/engage">Engage</NavLink>
          <NavLink to="/demo/newsletter">Newsletter</NavLink>
        </nav>
      </header>
      <main className="page-canvas">
        <Outlet />
      </main>
    </div>
  )
}
