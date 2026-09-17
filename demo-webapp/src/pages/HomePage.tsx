import { ChapterNav } from '../studio/ChapterNav'
import { ContentSpread } from '../studio/ContentSpread'
import { Hero } from '../studio/Hero'
import { PulsePlayground } from '../studio/PulsePlayground'
import { ResultAndTalk } from '../studio/ResultAndTalk'
import { StoriesPipeline } from '../studio/StoriesPipeline'
import { useActiveChapter } from '../studio/useActiveChapter'

export function HomePage() {
  const active = useActiveChapter()

  return (
    <>
      <ChapterNav active={active} />
      <Hero />
      <ContentSpread />
      <StoriesPipeline />
      <PulsePlayground />
      <ResultAndTalk />
    </>
  )
}
