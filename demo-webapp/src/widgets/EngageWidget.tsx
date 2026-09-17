import { useState } from 'react'
import { engagePosts } from '../data/mock'

type Post = (typeof engagePosts)[number]

export function EngageWidget() {
  const [posts, setPosts] = useState<Post[]>(engagePosts)
  const [likes, setLikes] = useState<Record<number, number>>(() =>
    Object.fromEntries(engagePosts.map((p) => [p.id, p.likes])),
  )
  const [liked, setLiked] = useState<Record<number, boolean>>({})
  const [openReplies, setOpenReplies] = useState<Record<number, boolean>>({ 1: true })
  const [draft, setDraft] = useState('')

  function compose() {
    const text = draft.trim()
    if (!text) return
    const next: Post = {
      id: Date.now(),
      author: 'Jordan Lee',
      role: 'You',
      time: 'just now',
      announcement: false,
      text,
      likes: 0,
      replies: [],
    }
    setPosts([next, ...posts])
    setLikes({ ...likes, [next.id]: 0 })
    setDraft('')
  }

  return (
    <div>
      <div className="dash-header">
        <div>
          <h2 className="dash-title">Community feed</h2>
          <p className="dash-sub">Announcements, likes, and replies. Compose stays in this browser session.</p>
        </div>
      </div>
      <div className="compose">
        <textarea
          rows={2}
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          placeholder="Share an update with the community…"
        />
        <button type="button" className="sp-btn" onClick={compose} disabled={!draft.trim()}>
          Post
        </button>
      </div>
      {posts.map((post) => (
        <article key={post.id} className="feed-post">
          {post.announcement ? <div className="ann">Announcement</div> : null}
          <div className="feed-meta">
            {post.author} · {post.role} · {post.time}
          </div>
          <p>{post.text}</p>
          <div className="feed-actions">
            <button
              type="button"
              className={`sp-btn ${liked[post.id] ? 'active' : 'ghost'}`}
              onClick={() => {
                const on = !liked[post.id]
                setLiked({ ...liked, [post.id]: on })
                setLikes({
                  ...likes,
                  [post.id]: (likes[post.id] ?? post.likes) + (on ? 1 : -1),
                })
              }}
            >
              Like · {likes[post.id] ?? post.likes}
            </button>
            {post.replies.length ? (
              <button
                type="button"
                className="sp-btn ghost"
                onClick={() => setOpenReplies({ ...openReplies, [post.id]: !openReplies[post.id] })}
              >
                {openReplies[post.id] ? 'Hide' : 'Show'} replies ({post.replies.length})
              </button>
            ) : null}
          </div>
          {openReplies[post.id]
            ? post.replies.map((r) => (
                <div key={r.text} className="reply">
                  <strong>{r.author}</strong> — {r.text}
                </div>
              ))
            : null}
        </article>
      ))}
    </div>
  )
}
