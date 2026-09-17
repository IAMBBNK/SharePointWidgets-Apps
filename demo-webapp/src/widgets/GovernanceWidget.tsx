import { useState } from 'react'
import { governanceBoards, governanceTabs } from '../data/mock'

function initials(name: string) {
  if (name === '—') return '?'
  return name
    .split(' ')
    .map((p) => p[0])
    .join('')
    .slice(0, 2)
}

export function GovernanceWidget() {
  const [tab, setTab] = useState<(typeof governanceTabs)[number]['id']>('bodies')
  const [backup, setBackup] = useState(false)
  const board = governanceBoards[tab]

  return (
    <div>
      <div className="dash-header">
        <div>
          <h2 className="dash-title">Governance model</h2>
          <p className="dash-sub">
            {tab === 'bodies'
              ? 'Core team and CoE membership by sector. Flip to backups.'
              : `${board.headers[0]} and related roles for this community of practice.`}
          </p>
        </div>
      </div>
      <div className="gov-tabs">
        {governanceTabs.map((t) => (
          <button
            key={t.id}
            type="button"
            className={`sp-btn ${tab === t.id ? 'active' : 'ghost'}`}
            onClick={() => setTab(t.id)}
          >
            {t.label}
          </button>
        ))}
        <button type="button" className="sp-btn" onClick={() => setBackup(!backup)}>
          {backup ? 'Show primary' : 'Show backups'}
        </button>
      </div>
      <table className="gov-table">
        <thead>
          <tr>
            <th>Sector</th>
            {board.headers.map((h) => (
              <th key={h}>
                {h}
                <div className="gov-subhead">{backup ? 'Backup' : 'Members'}</div>
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {board.rows.map((row) => (
            <tr key={row.sector}>
              <td className="sector" style={{ background: row.color }}>
                {row.sector}
              </td>
              {row.cells.map((cell) => {
                const name = backup ? cell.backup : cell.primary
                return (
                  <td key={`${row.sector}-${cell.primary}`}>
                    <div className="gov-person">
                      <span className="gov-avatar" aria-hidden="true">
                        {initials(name)}
                      </span>
                      {name === '—' ? (
                        name
                      ) : (
                        <a href={`mailto:${cell.email}`}>{name}</a>
                      )}
                    </div>
                  </td>
                )
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
