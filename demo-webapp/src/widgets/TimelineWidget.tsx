import { timeline } from '../data/mock'

export function TimelineWidget() {
  return (
    <div>
      <div className="dash-header">
        <div>
          <h2 className="dash-title">Program timeline</h2>
          <p className="dash-sub">Milestones for the hub journey — each year has a one-line outcome.</p>
        </div>
      </div>
      <div className="tl">
        {timeline.map((item) => (
          <div key={item.year} className="tl-item">
            <div className="tl-hex">{item.year}</div>
            <div>
              <h3>{item.title}</h3>
              <p>{item.text}</p>
              <p className="tl-outcome">{item.outcome}</p>
            </div>
          </div>
        ))}
      </div>
    </div>
  )
}
