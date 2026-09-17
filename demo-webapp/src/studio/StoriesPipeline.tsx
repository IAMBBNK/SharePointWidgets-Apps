import { useState } from 'react'

const steps = [
  {
    id: 'person',
    n: '01',
    title: 'Person',
    hint: 'A shift lead who never writes “news.”',
  },
  {
    id: 'conversation',
    n: '02',
    title: 'Conversation',
    hint: 'Notes, not a questionnaire.',
  },
  {
    id: 'ideas',
    n: '03',
    title: 'Ideas',
    hint: 'What is actually new.',
  },
  {
    id: 'story',
    n: '04',
    title: 'Story',
    hint: 'A page people finish.',
  },
  {
    id: 'sharepoint',
    n: '05',
    title: 'SharePoint',
    hint: 'Published where work already lives.',
  },
] as const

type StepId = (typeof steps)[number]['id']

export function StoriesPipeline() {
  const [step, setStep] = useState<StepId>('person')

  return (
    <section id="stories" className="st-chapter" aria-labelledby="stories-title">
      <div className="st-chapter-n" aria-hidden="true">
        03
      </div>
      <p className="st-kicker">03 — Stories</p>
      <h2 id="stories-title" className="st-display" style={{ fontSize: 'clamp(2.4rem, 5vw, 4.2rem)', maxWidth: '16ch' }}>
        I don’t just publish content. I discover it.
      </h2>
      <p className="st-lede" style={{ marginTop: 16 }}>
        Your people already have stories. I sit with them, find the interesting
        stuff, and turn it into communication they actually want to read.
      </p>
      <div className="st-pipeline">
        <div className="st-steps" role="tablist" aria-label="How a story is made">
          {steps.map((s) => (
            <button
              key={s.id}
              type="button"
              role="tab"
              aria-selected={step === s.id}
              className={`st-step ${step === s.id ? 'on' : ''}`}
              onClick={() => setStep(s.id)}
            >
              <i>{s.n}</i>
              <span>
                <strong>{s.title}</strong>
                <p>{s.hint}</p>
              </span>
            </button>
          ))}
        </div>
        <div className="st-stage" role="tabpanel" key={step}>
          {step === 'person' ? (
            <>
              <div className="st-stage-lab">Field note · Harbour QC</div>
              <h3 className="st-person-name">Maya Chen</h3>
              <p className="st-body">
                Shift lead. Twelve years on the line. Has never opened the hub
                except to find a holiday calendar.
              </p>
              <p className="st-body" style={{ marginTop: 12 }}>
                She mentioned, unprompted, that the new packing line went live
                on Tuesday. The intranet still showed last year’s photo.
              </p>
            </>
          ) : null}
          {step === 'conversation' ? (
            <>
              <div className="st-stage-lab">Raw notes</div>
              <div className="st-notes">
                <p>— night shift found out via WhatsApp</p>
                <p>— “the photo on the hub is the old line??”</p>
                <p>— vendors on site Thu / time-boxed tunnel</p>
                <p>— nobody briefed quality</p>
              </div>
              <blockquote className="st-quote">
                “Nobody knows the new packing line is live{' '}
                <mark>except the six of us.</mark>”
              </blockquote>
            </>
          ) : null}
          {step === 'ideas' ? (
            <>
              <div className="st-stage-lab">What to say</div>
              <p className="st-body">Not a go-live checklist. A human fact with a reason to care.</p>
              <div className="st-ideas">
                <span>packing line live</span>
                <span>night shift</span>
                <span>quality not briefed</span>
                <span>photo is wrong</span>
                <span>six people know</span>
              </div>
            </>
          ) : null}
          {step === 'story' ? (
            <>
              <div className="st-article-kicker">Internal story · 3 min</div>
              <h3 className="st-article-hed">The packing line went live. The intranet didn’t notice.</h3>
              <p className="st-body">
                On Tuesday night, Harbour QC started running the new line. By
                Wednesday morning, six people knew. The hub still showed a
                photograph from 2019. Here’s what changed, who to ask, and why
                the picture matters.
              </p>
            </>
          ) : null}
          {step === 'sharepoint' ? (
            <>
              <div className="st-stage-lab">On the hub</div>
              <div className="st-sp-chrome">
                <div className="st-sp-bar">
                  <span className="st-sp-dot" />
                  <span className="st-sp-dot" />
                  <span className="st-sp-dot" />
                  Contoso Hub · News
                </div>
                <div className="st-sp-body">
                  <div className="st-article-kicker" style={{ color: '#0f6cbd' }}>
                    Site news
                  </div>
                  <h3>The packing line went live. The intranet didn’t notice.</h3>
                  <p>
                    Maya Chen, shift lead · Harbour QC. Published to the hub,
                    the plant channel, and the Thursday digest.
                  </p>
                </div>
              </div>
            </>
          ) : null}
        </div>
      </div>
    </section>
  )
}
