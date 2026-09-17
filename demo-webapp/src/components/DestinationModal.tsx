export function DestinationModal({
  title,
  body,
  onClose,
}: {
  title: string
  body: string
  onClose: () => void
}) {
  return (
    <div
      className="imgc-overlay"
      onClick={(e) => e.target === e.currentTarget && onClose()}
      role="dialog"
      aria-modal="true"
      aria-labelledby="dest-title"
    >
      <div className="dest-card">
        <h3 id="dest-title">{title}</h3>
        <p>{body}</p>
        <p className="dest-note">In production this opens the SharePoint page or course.</p>
        <button type="button" className="sp-btn" onClick={onClose}>
          Close
        </button>
      </div>
    </div>
  )
}
