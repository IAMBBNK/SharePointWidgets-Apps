import { useState } from 'react'
import { Link } from 'react-router-dom'

export function ResultAndTalk() {
  const [note, setNote] = useState('')
  const [sent, setSent] = useState(false)

  return (
    <>
      <section id="result" className="st-chapter" aria-labelledby="result-title">
        <div className="st-chapter-n" aria-hidden="true">
          05
        </div>
        <p className="st-kicker">05 — The result</p>
        <div className="st-result">
          <div>
            <h2 id="result-title" className="st-pull">
              “We stopped hunting for the current version.
              <span> People started sending links instead of attachments.</span>”
            </h2>
            <p className="st-byline">Plant lead · after a content + stories pass, not a replatform</p>
          </div>
          <div className="st-stats">
            <div>
              <b>1 person</b>
              <span>who can write, structure, and ship a web part</span>
            </div>
            <div>
              <b>3 crafts</b>
              <span>content, interviews, custom experiences</span>
            </div>
            <div>
              <b>0 decks</b>
              <span>that pretend a screenshot is a working page</span>
            </div>
          </div>
        </div>
      </section>

      <section id="talk" className="st-chapter st-talk" aria-labelledby="talk-title">
        <div className="st-chapter-n" aria-hidden="true">
          06
        </div>
        <p className="st-kicker">06 — Let’s talk</p>
        <h2 id="talk-title" className="st-display st-talk-title">
          Write me.
        </h2>
        <p className="st-lede">
          Tell me what your hub is supposed to do that it currently doesn’t.
          I’ll tell you whether it needs an editor, a builder, or both.
        </p>
        {sent ? (
          <p className="st-note-ok" role="status">
            Noted — in this demo it stays in your browser. In production it
            would land in my inbox.
          </p>
        ) : (
          <form
            className="st-note"
            onSubmit={(e) => {
              e.preventDefault()
              if (!note.trim()) return
              setSent(true)
            }}
          >
            <label htmlFor="field-note" className="st-kicker">
              A short note
            </label>
            <textarea
              id="field-note"
              value={note}
              onChange={(e) => setNote(e.target.value)}
              placeholder="Our hub is a graveyard of untitled pages…"
            />
            <button type="submit" className="st-btn">
              Send the note
            </button>
          </form>
        )}
        <div className="st-split-links">
          <Link className="st-btn ghost" to="/demo">
            Or just open the hub
          </Link>
        </div>
      </section>
    </>
  )
}
