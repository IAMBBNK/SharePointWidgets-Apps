import { Link, NavLink, Outlet, useLocation } from 'react-router-dom'
import { useEffect } from 'react'

export function MarketingLayout() {
  const location = useLocation()

  useEffect(() => {
    document.body.classList.add('studio-body')
    const prev = document.title
    document.title = 'field. — SharePoint, but make it human'
    return () => {
      document.body.classList.remove('studio-body')
      document.title = prev
    }
  }, [])

  useEffect(() => {
    if (!location.hash) return
    const el = document.querySelector(location.hash)
    el?.scrollIntoView({ behavior: 'smooth' })
  }, [location.pathname, location.hash])

  return (
    <div className="studio">
      <a className="studio-skip" href="#main">
        Skip to content
      </a>
      <header className="studio-bar">
        <NavLink to="/" className="studio-wordmark">
          field<span>.</span>
        </NavLink>
        <nav className="studio-nav" aria-label="Primary">
          <Link to={{ pathname: '/', hash: 'content' }}>Content</Link>
          <Link to={{ pathname: '/', hash: 'stories' }}>Stories</Link>
          <Link to={{ pathname: '/', hash: 'systems' }}>Systems</Link>
          <NavLink to="/demo">Hub</NavLink>
          <Link to={{ pathname: '/', hash: 'talk' }}>Talk</Link>
        </nav>
      </header>
      <main id="main" className="studio-main">
        <Outlet />
      </main>
      <footer className="studio-foot">
        <span>SharePoint is the platform. The experience is the craft.</span>
        <span>
          <Link to="/services/sharepoint">SharePoint</Link>
          {' · '}
          <Link to="/services/confluence">Confluence</Link>
          {' · '}
          Sample work, mock data, no live tenant.
        </span>
      </footer>
    </div>
  )
}
