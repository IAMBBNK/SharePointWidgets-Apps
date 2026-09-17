export const chapters = [
  { id: 'problem', n: '01', label: 'Problem' },
  { id: 'content', n: '02', label: 'Content' },
  { id: 'stories', n: '03', label: 'Stories' },
  { id: 'systems', n: '04', label: 'Systems' },
  { id: 'result', n: '05', label: 'Result' },
  { id: 'talk', n: '06', label: 'Talk' },
] as const

export type ChapterId = (typeof chapters)[number]['id']
