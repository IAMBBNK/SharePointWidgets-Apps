export type OfferGroup = 'content' | 'extension' | 'confluence' | 'addon'
export type UsedFor = 'SharePoint' | 'Confluence' | 'Add-on'

export type CatalogItem = {
  slug: string
  title: string
  blurb: string
  group: OfferGroup
  usedFor: UsedFor
  path: string
}

export const groupLabels: Record<OfferGroup, string> = {
  content: 'SharePoint content management',
  extension: 'SharePoint feature extension',
  confluence: 'Confluence',
  addon: 'Optional add-ons',
}

export const catalog: CatalogItem[] = [
  {
    slug: 'resources',
    title: 'Resources portal',
    blurb: 'Searchable library of standards and tools, with a mock destination on Go to.',
    group: 'content',
    usedFor: 'SharePoint',
    path: '/demo/resources',
  },
  {
    slug: 'trainings',
    title: 'Trainings catalog',
    blurb: 'Accordion catalog with domain and curricula filters, plus a mock course destination.',
    group: 'content',
    usedFor: 'SharePoint',
    path: '/demo/trainings',
  },
  {
    slug: 'governance',
    title: 'Governance tables',
    blurb: 'Tab-specific role matrices with primary/backup flip, initials, and mailto.',
    group: 'content',
    usedFor: 'SharePoint',
    path: '/demo/governance',
  },
  {
    slug: 'site-form',
    title: 'Site master data form',
    blurb: 'Sectioned site form with people chips, validation, and revert.',
    group: 'content',
    usedFor: 'SharePoint',
    path: '/demo/site-form',
  },
  {
    slug: 'maintenance-form',
    title: 'Predictive maintenance form',
    blurb: 'Intake with entry picker, site prefill, and required-field checks.',
    group: 'content',
    usedFor: 'SharePoint',
    path: '/demo/maintenance-form',
  },
  {
    slug: 'completion',
    title: 'Site completion dashboard',
    blurb: 'Region/sector filters, section meters, and missing-field lists per site.',
    group: 'content',
    usedFor: 'SharePoint',
    path: '/demo/completion',
  },
  {
    slug: 'kpi',
    title: 'Site KPI dashboard',
    blurb: 'Stacked-bar KPIs that recompute from region and sector filters.',
    group: 'extension',
    usedFor: 'SharePoint',
    path: '/demo/kpi',
  },
  {
    slug: 'ot-kpi',
    title: 'Operations KPI dashboard',
    blurb: 'Period switch, target tones, and ⓘ formulas on published OT KPIs.',
    group: 'extension',
    usedFor: 'SharePoint',
    path: '/demo/ot-kpi',
  },
  {
    slug: 'page-kpi',
    title: 'Page & engagement KPIs',
    blurb: 'Search, menu vs article split, viewers and Viva columns.',
    group: 'extension',
    usedFor: 'SharePoint',
    path: '/demo/page-kpi',
  },
  {
    slug: 'map',
    title: 'Interactive site map',
    blurb: 'Global pins with sector and business-unit filters and a detail card.',
    group: 'extension',
    usedFor: 'SharePoint',
    path: '/demo/map',
  },
  {
    slug: 'carousel',
    title: 'Image carousel',
    blurb: 'Gallery overlay with 1- or 3-card view, keyboard, and scroll lock.',
    group: 'extension',
    usedFor: 'SharePoint',
    path: '/demo/carousel',
  },
  {
    slug: 'timeline',
    title: 'Program timeline',
    blurb: 'Hex-year journey with a one-line outcome on each milestone.',
    group: 'extension',
    usedFor: 'SharePoint',
    path: '/demo/timeline',
  },
  {
    slug: 'survey',
    title: 'Feedback survey',
    blurb: 'Six-slot rank, best/worst comments, and follow-up consent.',
    group: 'extension',
    usedFor: 'SharePoint',
    path: '/demo/survey',
  },
  {
    slug: 'projects',
    title: 'Projects board',
    blurb: 'Status-filtered portfolio cards with a next-step expand.',
    group: 'extension',
    usedFor: 'SharePoint',
    path: '/demo/projects',
  },
  {
    slug: 'confluence',
    title: 'Confluence space',
    blurb: 'Nested space tree, breadcrumbs, search, and macros.',
    group: 'confluence',
    usedFor: 'Confluence',
    path: '/demo/confluence',
  },
  {
    slug: 'engage',
    title: 'Viva Engage community',
    blurb: 'Feed with compose, like counts, and expandable replies.',
    group: 'addon',
    usedFor: 'Add-on',
    path: '/demo/engage',
  },
  {
    slug: 'newsletter',
    title: 'Newsletter',
    blurb: 'Edition switch, 24h and 7-day curves, and a story click table.',
    group: 'addon',
    usedFor: 'Add-on',
    path: '/demo/newsletter',
  },
]

export function itemBySlug(slug: string): CatalogItem | undefined {
  return catalog.find((item) => item.slug === slug)
}
