import { useEffect, useRef, useState } from 'react'
import L from 'leaflet'
import 'leaflet/dist/leaflet.css'
import { sites, type Site } from '../data/mock'

const colors: Record<string, string> = {
  Manufacturing: '#0078d4',
  'R&D': '#ca5010',
  Corporate: '#5c2e91',
}

const units = ['all', ...Array.from(new Set(sites.map((s) => s.businessUnit)))]

export function MapWidget() {
  const el = useRef<HTMLDivElement>(null)
  const mapRef = useRef<L.Map | null>(null)
  const [sector, setSector] = useState('all')
  const [unit, setUnit] = useState('all')
  const [selected, setSelected] = useState<Site | null>(sites[0])

  useEffect(() => {
    if (!el.current || mapRef.current) return
    const map = L.map(el.current, { scrollWheelZoom: false }).setView([20, 10], 2)
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '&copy; OpenStreetMap',
    }).addTo(map)
    mapRef.current = map
    return () => {
      map.remove()
      mapRef.current = null
    }
  }, [])

  useEffect(() => {
    const map = mapRef.current
    if (!map) return
    const layer = L.layerGroup().addTo(map)
    const visible = sites.filter(
      (s) =>
        (sector === 'all' || s.sector === sector) &&
        (unit === 'all' || s.businessUnit === unit),
    )
    visible.forEach((site) => {
      const marker = L.circleMarker([site.lat, site.lng], {
        radius: 8,
        color: '#fff',
        weight: 2,
        fillColor: colors[site.sector],
        fillOpacity: 1,
      }).addTo(layer)
      marker.on('click', () => setSelected(site))
    })
    return () => {
      layer.remove()
    }
  }, [sector, unit])

  return (
    <div>
      <div className="dash-header">
        <div>
          <h2 className="dash-title">Global site map</h2>
          <p className="dash-sub">Filter by sector or business unit, then click a pin for site details.</p>
        </div>
      </div>
      <div className="map-filters">
        {['all', 'Manufacturing', 'R&D', 'Corporate'].map((s) => (
          <button
            key={s}
            type="button"
            className={`sp-btn ${sector === s ? 'active' : ''}`}
            onClick={() => setSector(s)}
          >
            {s === 'all' ? 'All sectors' : s}
          </button>
        ))}
      </div>
      <div className="map-filters">
        {units.map((u) => (
          <button
            key={u}
            type="button"
            className={`sp-btn ghost ${unit === u ? 'active' : ''}`}
            onClick={() => setUnit(u)}
          >
            {u === 'all' ? 'All units' : u}
          </button>
        ))}
      </div>
      <ul className="map-legend">
        {Object.entries(colors).map(([label, color]) => (
          <li key={label}>
            <i style={{ background: color }} />
            {label}
          </li>
        ))}
      </ul>
      <div ref={el} className="map-wrap" />
      {selected ? (
        <div className="map-card">
          <strong>
            {selected.name} / {selected.unit}
          </strong>
          <p>
            {selected.location} · {selected.businessUnit} · {selected.employees} employees
          </p>
          <div className="chip-row">
            <span className="status ok">IT {selected.itAssessment}</span>
            <span className="status risk">OT {selected.otAssessment}</span>
            <span className="status done">Integration {selected.integration}</span>
          </div>
          <p>{selected.services}</p>
          <p>{selected.maturity}</p>
          <p>OT systems: {selected.otSystems}</p>
        </div>
      ) : null}
    </div>
  )
}
