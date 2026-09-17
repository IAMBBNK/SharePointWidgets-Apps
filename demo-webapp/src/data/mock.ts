export type Assessment = 'Completed' | 'Ongoing' | 'Planning' | 'N/A'
export type Platform = 'Platinum' | 'Premium' | 'Essential' | 'N/A'
export type Region = 'EMEA' | 'AMER' | 'APAC'

export type Site = {
  id: number
  name: string
  unit: string
  sector: string
  location: string
  region: Region
  lat: number
  lng: number
  employees: number
  businessUnit: string
  itAssessment: Assessment
  otAssessment: Assessment
  integration: Assessment
  platform: Platform
  completeness: number
  sections: { general: number; people: number; maturity: number; counts: number }
  missing: string[]
  services: string
  maturity: string
  otSystems: string
  otApps: number
  otToMigrate: number
}

export const sites: Site[] = [
  {
    id: 1,
    name: 'North Plant',
    unit: 'Manufacturing',
    sector: 'Manufacturing',
    location: 'Hamburg, DE',
    region: 'EMEA',
    lat: 53.55,
    lng: 9.99,
    employees: 420,
    businessUnit: 'Operations',
    itAssessment: 'Completed',
    otAssessment: 'Ongoing',
    integration: 'Planning',
    platform: 'Premium',
    completeness: 92,
    sections: { general: 100, people: 100, maturity: 80, counts: 88 },
    missing: ['OT target platform date'],
    services: 'MES, historian, quality lab',
    maturity: 'Current 2.5 · target 3.5',
    otSystems: 'Historian, batch, SCADA',
    otApps: 14,
    otToMigrate: 4,
  },
  {
    id: 2,
    name: 'River Labs',
    unit: 'R&D',
    sector: 'R&D',
    location: 'Basel, CH',
    region: 'EMEA',
    lat: 47.56,
    lng: 7.59,
    employees: 180,
    businessUnit: 'Science',
    itAssessment: 'Completed',
    otAssessment: 'Completed',
    integration: 'Completed',
    platform: 'Platinum',
    completeness: 100,
    sections: { general: 100, people: 100, maturity: 100, counts: 100 },
    missing: [],
    services: 'LIMS, instrument PCs',
    maturity: 'Current 4 · target 4',
    otSystems: 'LIMS, chromatography',
    otApps: 9,
    otToMigrate: 0,
  },
  {
    id: 3,
    name: 'East Campus',
    unit: 'Corporate',
    sector: 'Corporate',
    location: 'Boston, US',
    region: 'AMER',
    lat: 42.36,
    lng: -71.06,
    employees: 95,
    businessUnit: 'Corporate',
    itAssessment: 'Ongoing',
    otAssessment: 'N/A',
    integration: 'N/A',
    platform: 'Essential',
    completeness: 64,
    sections: { general: 90, people: 50, maturity: 40, counts: 70 },
    missing: ['Backup SPOC', 'OT assessment N/A confirmed', 'Asset count'],
    services: 'Identity, service desk',
    maturity: 'Current 1.5 · target 2.5',
    otSystems: 'None on campus',
    otApps: 2,
    otToMigrate: 1,
  },
  {
    id: 4,
    name: 'Valley Works',
    unit: 'Manufacturing',
    sector: 'Manufacturing',
    location: 'Lyon, FR',
    region: 'EMEA',
    lat: 45.76,
    lng: 4.84,
    employees: 310,
    businessUnit: 'Operations',
    itAssessment: 'Planning',
    otAssessment: 'Planning',
    integration: 'Planning',
    platform: 'Essential',
    completeness: 48,
    sections: { general: 70, people: 40, maturity: 30, counts: 50 },
    missing: ['Site head', 'IT assessment owner', 'OT app list', 'DPI flag'],
    services: 'Packaging line, warehouse WMS',
    maturity: 'Current 1 · target 3',
    otSystems: 'PLC cells, WMS',
    otApps: 11,
    otToMigrate: 8,
  },
  {
    id: 5,
    name: 'Pacific Hub',
    unit: 'R&D',
    sector: 'R&D',
    location: 'Singapore, SG',
    region: 'APAC',
    lat: 1.35,
    lng: 103.82,
    employees: 140,
    businessUnit: 'Science',
    itAssessment: 'Completed',
    otAssessment: 'Ongoing',
    integration: 'Ongoing',
    platform: 'Premium',
    completeness: 81,
    sections: { general: 100, people: 80, maturity: 70, counts: 75 },
    missing: ['Backup architect', 'Integration wave date'],
    services: 'Pilot plant, LIMS',
    maturity: 'Current 3 · target 3.5',
    otSystems: 'Pilot DCS, LIMS',
    otApps: 7,
    otToMigrate: 2,
  },
  {
    id: 6,
    name: 'Prairie Fill',
    unit: 'Manufacturing',
    sector: 'Manufacturing',
    location: 'Chicago, US',
    region: 'AMER',
    lat: 41.88,
    lng: -87.63,
    employees: 260,
    businessUnit: 'Operations',
    itAssessment: 'Completed',
    otAssessment: 'Completed',
    integration: 'Ongoing',
    platform: 'Premium',
    completeness: 88,
    sections: { general: 100, people: 90, maturity: 80, counts: 80 },
    missing: ['Target BPOG date'],
    services: 'Fill-finish, serialization',
    maturity: 'Current 3 · target 4',
    otSystems: 'Serialization, MES',
    otApps: 12,
    otToMigrate: 3,
  },
  {
    id: 7,
    name: 'Andes Site',
    unit: 'Manufacturing',
    sector: 'Manufacturing',
    location: 'São Paulo, BR',
    region: 'AMER',
    lat: -23.55,
    lng: -46.63,
    employees: 190,
    businessUnit: 'Operations',
    itAssessment: 'Ongoing',
    otAssessment: 'Planning',
    integration: 'Planning',
    platform: 'N/A',
    completeness: 55,
    sections: { general: 80, people: 45, maturity: 35, counts: 55 },
    missing: ['Legal entity', 'Service SPOC', 'OT inventory'],
    services: 'Solids manufacturing',
    maturity: 'Current 1.5 · target 3',
    otSystems: 'MES (legacy)',
    otApps: 8,
    otToMigrate: 6,
  },
  {
    id: 8,
    name: 'Harbour QC',
    unit: 'R&D',
    sector: 'R&D',
    location: 'Osaka, JP',
    region: 'APAC',
    lat: 34.69,
    lng: 135.5,
    employees: 75,
    businessUnit: 'Science',
    itAssessment: 'Completed',
    otAssessment: 'N/A',
    integration: 'N/A',
    platform: 'Essential',
    completeness: 72,
    sections: { general: 90, people: 80, maturity: 50, counts: 70 },
    missing: ['OT N/A sign-off', 'Instrument inventory'],
    services: 'QC lab only',
    maturity: 'Current 2 · target 2.5',
    otSystems: 'QC instruments',
    otApps: 4,
    otToMigrate: 0,
  },
  {
    id: 9,
    name: 'Capitol HQ',
    unit: 'Corporate',
    sector: 'Corporate',
    location: 'London, UK',
    region: 'EMEA',
    lat: 51.51,
    lng: -0.13,
    employees: 220,
    businessUnit: 'Corporate',
    itAssessment: 'Completed',
    otAssessment: 'N/A',
    integration: 'Completed',
    platform: 'Platinum',
    completeness: 96,
    sections: { general: 100, people: 100, maturity: 90, counts: 95 },
    missing: ['Newsletter owner'],
    services: 'Identity, intranet, service catalog',
    maturity: 'Current 4 · target 4',
    otSystems: 'None',
    otApps: 3,
    otToMigrate: 0,
  },
  {
    id: 10,
    name: 'Delta Pack',
    unit: 'Manufacturing',
    sector: 'Manufacturing',
    location: 'Sydney, AU',
    region: 'APAC',
    lat: -33.87,
    lng: 151.21,
    employees: 155,
    businessUnit: 'Operations',
    itAssessment: 'Planning',
    otAssessment: 'Ongoing',
    integration: 'Ongoing',
    platform: 'Premium',
    completeness: 69,
    sections: { general: 85, people: 60, maturity: 55, counts: 75 },
    missing: ['Architecture lead', 'Backup service owner'],
    services: 'Packaging, warehouse',
    maturity: 'Current 2.5 · target 3.5',
    otSystems: 'WMS, pack-line PLCs',
    otApps: 10,
    otToMigrate: 5,
  },
  {
    id: 11,
    name: 'Nordic Trials',
    unit: 'R&D',
    sector: 'R&D',
    location: 'Stockholm, SE',
    region: 'EMEA',
    lat: 59.33,
    lng: 18.07,
    employees: 60,
    businessUnit: 'Science',
    itAssessment: 'Ongoing',
    otAssessment: 'Ongoing',
    integration: 'Planning',
    platform: 'Essential',
    completeness: 58,
    sections: { general: 75, people: 55, maturity: 40, counts: 60 },
    missing: ['Site master owner', 'IT assessment evidence', 'OT owners'],
    services: 'Clinical supplies',
    maturity: 'Current 2 · target 3',
    otSystems: 'Trial labeling',
    otApps: 5,
    otToMigrate: 3,
  },
  {
    id: 12,
    name: 'West Admin',
    unit: 'Corporate',
    sector: 'Corporate',
    location: 'Austin, US',
    region: 'AMER',
    lat: 30.27,
    lng: -97.74,
    employees: 48,
    businessUnit: 'Corporate',
    itAssessment: 'Completed',
    otAssessment: 'N/A',
    integration: 'N/A',
    platform: 'N/A',
    completeness: 41,
    sections: { general: 60, people: 30, maturity: 25, counts: 50 },
    missing: ['People section', 'Maturity scores', 'Review status'],
    services: 'Shared services desk',
    maturity: 'Current 1 · target 2',
    otSystems: 'None',
    otApps: 1,
    otToMigrate: 1,
  },
]

export function filterSites(list: Site[], sector: string, region: string) {
  return list.filter((s) => {
    const okS = !sector || s.sector === sector
    const okR = !region || s.region === region
    return okS && okR
  })
}

export function tally(
  list: Site[],
  key: 'itAssessment' | 'otAssessment' | 'integration' | 'platform',
) {
  const labels =
    key === 'platform'
      ? (['Platinum', 'Premium', 'Essential', 'N/A'] as const)
      : (['Completed', 'Ongoing', 'Planning', 'N/A'] as const)
  return labels.map((label) => ({
    label,
    value: list.filter((s) => s[key] === label).length,
  }))
}

export const resources = [
  {
    name: 'Change control SOP',
    category: 'Standards & Procedures',
    sector: 'Manufacturing',
    id: 'SOP-104',
    info: 'How to raise and close change requests for plant systems, including emergency changes and CAB cadence.',
    hrefLabel: 'Open SOP',
  },
  {
    name: 'Architecture decision record template',
    category: 'Standards & Procedures',
    sector: 'Corporate',
    id: 'TPL-012',
    info: 'Lightweight ADR for intranet and integration choices. One page, named owner, review date.',
    hrefLabel: 'Open template',
  },
  {
    name: 'Asset register workbook',
    category: 'Tools & Applications',
    sector: 'Manufacturing',
    id: 'APP-221',
    info: 'Inventory of OT applications, owners, and target platform. Export monthly to the KPI list.',
    hrefLabel: 'Open workbook',
  },
  {
    name: 'Service catalog portal',
    category: 'Tools & Applications',
    sector: 'Corporate',
    id: 'APP-018',
    info: 'Request path for new integrations, access, and plant onboarding.',
    hrefLabel: 'Open catalog',
  },
  {
    name: 'Incident severity guide',
    category: 'Standards & Procedures',
    sector: 'R&D',
    id: 'SOP-088',
    info: 'Severity, escalation, and communication for lab systems and instrument PCs.',
    hrefLabel: 'Open guide',
  },
  {
    name: 'Jump-host naming standard',
    category: 'Standards & Procedures',
    sector: 'Manufacturing',
    id: 'STD-044',
    info: 'Hostnames, OU placement, and break-glass accounts for plant remote access.',
    hrefLabel: 'Open standard',
  },
  {
    name: 'Historian failover runbook',
    category: 'Standards & Procedures',
    sector: 'Manufacturing',
    id: 'RB-019',
    info: 'Buffer node, notify service desk, capture a timeline. Keep in sync with Confluence.',
    hrefLabel: 'Open runbook',
  },
  {
    name: 'CMDB connector',
    category: 'Tools & Applications',
    sector: 'Corporate',
    id: 'APP-090',
    info: 'Pushes named OT assets into the configuration database once a week.',
    hrefLabel: 'Open tool',
  },
  {
    name: 'Lab instrument checklist',
    category: 'Standards & Procedures',
    sector: 'R&D',
    id: 'CHK-031',
    info: 'Onboarding a new instrument PC: image, patch ring, and owner in master data.',
    hrefLabel: 'Open checklist',
  },
  {
    name: 'Access request bot',
    category: 'Tools & Applications',
    sector: 'Corporate',
    id: 'APP-055',
    info: 'Teams bot that opens a catalog ticket for plant network access.',
    hrefLabel: 'Open bot',
  },
]

export const trainings = [
  {
    name: 'Operating model fundamentals',
    domain: 'operating',
    category: 'IT/OT Operating Model',
    insight: 'Roles, forums, and how work is funded. Required for new site leads in the first 30 days.',
    internal: true,
    curricula: true,
  },
  {
    name: 'Funding and demand intake',
    domain: 'operating',
    category: 'IT/OT Operating Model',
    insight: 'How a plant request becomes a funded integration. Includes the demand form walkthrough.',
    internal: true,
    curricula: false,
  },
  {
    name: 'Plant network zones',
    domain: 'technology',
    category: 'Technology & Architecture',
    insight: 'Purdue model, jump hosts, and remote access patterns used on the hub.',
    internal: true,
    curricula: false,
  },
  {
    name: 'Integration design clinic',
    domain: 'technology',
    category: 'Technology & Architecture',
    insight: 'Patterns for MES, historian, and ERP links. Bring a live site example.',
    internal: true,
    curricula: true,
  },
  {
    name: 'Reference architecture v3',
    domain: 'technology',
    category: 'Technology & Architecture',
    insight: 'Zone model published in Engage. Use it for any plant network change from Monday.',
    internal: true,
    curricula: true,
  },
  {
    name: 'OT cybersecurity awareness',
    domain: 'cyber',
    category: 'Cybersecurity',
    insight: 'Threats specific to industrial environments. 20 minutes, required for site OT leads.',
    internal: false,
    curricula: true,
  },
  {
    name: 'Identity on the plant floor',
    domain: 'cyber',
    category: 'Cybersecurity',
    insight: 'MFA for jump hosts, break-glass, and what not to share on a shared HMI.',
    internal: true,
    curricula: false,
  },
  {
    name: 'Incident to problem handover',
    domain: 'service',
    category: 'Service Management',
    insight: 'When a ticket becomes a problem record, and who owns the known-error article.',
    internal: true,
    curricula: false,
  },
  {
    name: 'Service catalog for plants',
    domain: 'service',
    category: 'Service Management',
    insight: 'Which requests go through the catalog versus a local change. Includes SLA mapping.',
    internal: true,
    curricula: true,
  },
  {
    name: 'Major incident comms',
    domain: 'service',
    category: 'Service Management',
    insight: 'Who speaks, which channel, and when to post on the hub versus Engage.',
    internal: false,
    curricula: false,
  },
  {
    name: 'Forum cadence for CoEs',
    domain: 'operating',
    category: 'IT/OT Operating Model',
    insight: 'Monthly architecture forum, quarterly portfolio, and when backups must attend.',
    internal: true,
    curricula: false,
  },
  {
    name: 'Secure remote vendor access',
    domain: 'cyber',
    category: 'Cybersecurity',
    insight: 'Time-boxed vendor sessions, recording, and closing the tunnel.',
    internal: true,
    curricula: true,
  },
]

export const otKpiPeriods = {
  'Q2 2026': [
    {
      category: 'Organization & Governance',
      items: [
        { title: 'Site organization coverage', value: 71, target: 90, description: 'Share of in-scope sites with a named operating-model owner.', formula: 'sites with owner / sites in scope' },
        { title: 'Staffing adequacy', value: 54, target: 80, description: 'Self-assessed capacity versus the role matrix.', formula: 'roles filled / roles required' },
        { title: 'Sector roles assigned', value: 80, target: 100, description: 'Primary named for architecture, portfolio, service, security.', formula: 'named cells / total cells' },
      ],
    },
    {
      category: 'Training & Compliance',
      items: [
        { title: 'Foundational training', value: 86, target: 95, description: 'Operating model fundamentals completed.', formula: 'completions / required audience' },
        { title: 'Role-based training', value: 48, target: 80, description: 'Path assigned to the person’s CoE role.', formula: 'role paths done / people in role' },
      ],
    },
    {
      category: 'Asset & Inventory',
      items: [
        { title: 'Assets in CMDB', value: 62, target: 85, description: 'OT applications with a CI.', formula: 'CIs / declared OT apps' },
        { title: 'Owners named', value: 77, target: 90, description: 'Each CI has a named technical owner.', formula: 'owned CIs / CIs' },
      ],
    },
    {
      category: 'Lifecycle Management',
      items: [
        { title: 'In-support versions', value: 58, target: 80, description: 'OT apps on a vendor-supported version.', formula: 'in-support / OT apps' },
        { title: 'Roadmap approved', value: 41, target: 70, description: 'Site has an approved 18-month OT roadmap.', formula: 'approved sites / sites' },
      ],
    },
  ],
  'Q3 2026': [
    {
      category: 'Organization & Governance',
      items: [
        { title: 'Site organization coverage', value: 78, target: 90, description: 'Share of in-scope sites with a named operating-model owner.', formula: 'sites with owner / sites in scope' },
        { title: 'Staffing adequacy', value: 61, target: 80, description: 'Self-assessed capacity versus the role matrix.', formula: 'roles filled / roles required' },
        { title: 'Sector roles assigned', value: 88, target: 100, description: 'Primary named for architecture, portfolio, service, security.', formula: 'named cells / total cells' },
      ],
    },
    {
      category: 'Training & Compliance',
      items: [
        { title: 'Foundational training', value: 92, target: 95, description: 'Operating model fundamentals completed.', formula: 'completions / required audience' },
        { title: 'Role-based training', value: 54, target: 80, description: 'Path assigned to the person’s CoE role.', formula: 'role paths done / people in role' },
      ],
    },
    {
      category: 'Asset & Inventory',
      items: [
        { title: 'Assets in CMDB', value: 71, target: 85, description: 'OT applications with a CI.', formula: 'CIs / declared OT apps' },
        { title: 'Owners named', value: 84, target: 90, description: 'Each CI has a named technical owner.', formula: 'owned CIs / CIs' },
      ],
    },
    {
      category: 'Lifecycle Management',
      items: [
        { title: 'In-support versions', value: 66, target: 80, description: 'OT apps on a vendor-supported version.', formula: 'in-support / OT apps' },
        { title: 'Roadmap approved', value: 49, target: 70, description: 'Site has an approved 18-month OT roadmap.', formula: 'approved sites / sites' },
      ],
    },
  ],
} as const

export const otKpis = otKpiPeriods['Q3 2026']

export const pageKpis = {
  lastUpdated: '12 Sep 2026',
  totals: [
    { label: 'Page views (30d)', value: '18,420' },
    { label: 'Unique viewers', value: '1,204' },
    { label: 'News posts', value: '14' },
    { label: 'Viva referrals', value: '312' },
  ],
  topPages: [
    { title: 'Home', views: 4200, viewers: 890, viva: 40, section: 'General', type: 'menu' as const },
    { title: 'Architecture playbook', views: 1880, viewers: 410, viva: 86, section: 'Architecture', type: 'article' as const },
    { title: 'Training catalog', views: 1640, viewers: 520, viva: 22, section: 'Upskilling', type: 'menu' as const },
    { title: 'Service catalog', views: 1210, viewers: 330, viva: 18, section: 'Service Management', type: 'menu' as const },
    { title: 'Q3 program update', views: 980, viewers: 610, viva: 74, section: 'News', type: 'article' as const },
    { title: 'Site master data', views: 870, viewers: 190, viva: 9, section: 'Governance', type: 'menu' as const },
    { title: 'Jump-host standard', views: 640, viewers: 210, viva: 31, section: 'Architecture', type: 'article' as const },
    { title: 'Onboarding checklist', views: 510, viewers: 160, viva: 12, section: 'Upskilling', type: 'article' as const },
  ],
}

export const newsletterEditions = [
  {
    title: 'September hub digest',
    sent: '4 Sep 2026',
    delivered: 1180,
    notes: 'Program update + two new trainings',
    opens24: [12, 40, 88, 140, 190, 210, 226, 240, 248, 255, 260, 262, 264, 265, 266, 268, 270, 271, 272, 273, 274, 274, 275, 276],
    clicks24: [2, 8, 18, 28, 36, 41, 44, 47, 49, 51, 52, 53, 54, 54, 55, 55, 56, 56, 57, 57, 57, 58, 58, 58],
    opens7: [276, 310, 338, 352, 361, 368, 372],
    clicks7: [58, 71, 79, 84, 88, 90, 92],
    stories: [
      { title: 'Q3 program update', clicks: 41, pageViews: 980 },
      { title: 'New training path', clicks: 28, pageViews: 1640 },
      { title: 'Map refresh', clicks: 15, pageViews: 420 },
    ],
  },
  {
    title: 'August architecture spotlight',
    sent: '7 Aug 2026',
    delivered: 1102,
    notes: 'Reference architecture for plant networks',
    opens24: [8, 22, 50, 90, 120, 148, 160, 170, 176, 180, 184, 186, 188, 190, 191, 192, 193, 194, 194, 195, 195, 196, 196, 197],
    clicks24: [1, 4, 10, 16, 22, 26, 29, 31, 33, 34, 35, 35, 36, 36, 37, 37, 37, 38, 38, 38, 38, 39, 39, 39],
    opens7: [197, 220, 238, 249, 255, 260, 264],
    clicks7: [39, 48, 54, 58, 61, 63, 64],
    stories: [
      { title: 'Reference architecture v3', clicks: 33, pageViews: 1880 },
      { title: 'Jump-host standard', clicks: 19, pageViews: 640 },
      { title: 'Zone model FAQ', clicks: 12, pageViews: 310 },
    ],
  },
  {
    title: 'July enablement digest',
    sent: '3 Jul 2026',
    delivered: 1044,
    notes: 'Curricula launch and completion dashboard',
    opens24: [10, 28, 60, 100, 132, 150, 162, 170, 176, 180, 183, 185, 186, 187, 188, 189, 190, 190, 191, 191, 192, 192, 193, 193],
    clicks24: [2, 6, 14, 22, 28, 32, 35, 37, 38, 39, 40, 40, 41, 41, 41, 42, 42, 42, 42, 43, 43, 43, 43, 43],
    opens7: [193, 214, 228, 239, 246, 250, 254],
    clicks7: [43, 51, 57, 61, 64, 66, 67],
    stories: [
      { title: 'Training curricula', clicks: 30, pageViews: 1640 },
      { title: 'Completion dashboard how-to', clicks: 21, pageViews: 870 },
      { title: 'Site form reminder', clicks: 16, pageViews: 510 },
    ],
  },
]

export const carouselSlides = [
  { tone: 0, title: 'Autumn', description: 'Season kickoff at the North Plant campus.' },
  { tone: 1, title: 'Lab companion', description: 'R&D mascot day during the summer offsite.' },
  { tone: 2, title: 'Night in the city', description: 'East Campus skyline after the architecture forum.' },
  { tone: 3, title: 'Mountain lake', description: 'Leadership offsite — strategy and site roadmaps.' },
  { tone: 4, title: 'Forest path', description: 'Safety walkaround at Valley Works.' },
  { tone: 0, title: 'Harbour morning', description: 'QC shift handover at Harbour QC, Osaka.' },
  { tone: 2, title: 'Prairie line', description: 'Serialization go-live at Prairie Fill.' },
  { tone: 3, title: 'Capitol atrium', description: 'Community of practice meetup at HQ.' },
]

export const timeline = [
  { year: '2022', title: 'Pilot', text: 'Three sites, a shared list, and a weekly forum.', outcome: 'Proved owners would update data if the page was short.' },
  { year: '2023', title: 'Foundation', text: 'Hub launched, master data list, and first community of practice.', outcome: 'Bodies table went live for three sectors.' },
  { year: '2024', title: 'Scale', text: 'Trainings portal, KPI dashboards, and site forms for local owners.', outcome: 'Completion moved from email chasing to a dashboard.' },
  { year: '2025', title: 'Connect', text: 'Page analytics, newsletter, and Viva Engage referrals.', outcome: 'Program news started showing up where people already talk.' },
  { year: '2026', title: 'Extend', text: 'Map, carousel, surveys, and Confluence knowledge spaces.', outcome: 'Same story on the hub and in the runbook space.' },
  { year: '2027', title: 'Next', text: 'Planned: auto-reminders from completeness gaps and a second language pack.', outcome: 'Still a plan — not in this demo.' },
]

export const projects = [
  { title: 'Identity hardening', status: 'On track', owner: 'Security CoE', note: 'MFA rollout for plant jump hosts.', next: 'Close Prairie Fill exceptions by 30 Sep.' },
  { title: 'Historian standard', status: 'At risk', owner: 'Architecture', note: 'Vendor shortlist delayed one sprint.', next: 'Forum decision on 24 Sep or slip to Q4.' },
  { title: 'Service catalog v2', status: 'On track', owner: 'Service Mgmt', note: 'Request types mapped to SLAs.', next: 'Pilot three request types at North Plant.' },
  { title: 'Training curricula', status: 'Done', owner: 'Enablement', note: 'Four role paths published.', next: 'Measure completion in the Q4 OT KPI.' },
  { title: 'CMDB connector', status: 'On track', owner: 'Architecture', note: 'Weekly push of named OT assets.', next: 'Turn on for APAC sites.' },
  { title: 'Andes onboarding', status: 'At risk', owner: 'Operations', note: 'Master data still below 60%.', next: 'Site form workshop 18 Sep.' },
  { title: 'Newsletter cadence', status: 'Done', owner: 'Communications', note: 'Monthly digest with story clicks.', next: 'Keep September story table in the add-on demo.' },
  { title: 'Confluence runbooks', status: 'On track', owner: 'Service Mgmt', note: 'Historian failover page is the template.', next: 'Clone for WMS and serialization.' },
]

export const surveyMilestones = [
  { id: 'M1', name: 'Service design' },
  { id: 'M2', name: 'Demand submission' },
  { id: 'M3', name: 'Solution design' },
  { id: 'M4', name: 'Deployment' },
  { id: 'M5', name: 'Testing' },
  { id: 'M6', name: 'Handover' },
]

export const governanceTabs = [
  { id: 'bodies', label: 'Bodies' },
  { id: 'architecture', label: 'Architecture' },
  { id: 'portfolio', label: 'Portfolio' },
  { id: 'service', label: 'Service' },
  { id: 'security', label: 'Security' },
] as const

type GovCell = { primary: string; backup: string; email: string }
type GovRow = { sector: string; color: string; cells: GovCell[] }

const email = (name: string) =>
  `${name.replace(/\s|\./g, '').toLowerCase()}@contoso.example`

function cell(primary: string, backup: string): GovCell {
  return { primary, backup, email: email(primary) }
}

export const governanceBoards: Record<(typeof governanceTabs)[number]['id'], { headers: string[]; rows: GovRow[] }> = {
  bodies: {
    headers: ['Core team', 'Architecture CoE', 'Service CoE', 'Security CoE'],
    rows: [
      { sector: 'Manufacturing', color: '#0078d4', cells: [cell('A. Keller', 'M. Shaw'), cell('J. Ortiz', 'P. Ng'), cell('L. Berg', 'S. Cole'), cell('R. Das', 'K. Moon')] },
      { sector: 'R&D', color: '#ca5010', cells: [cell('N. Patel', 'T. Ruiz'), cell('C. Young', 'D. Fox'), cell('E. Haas', 'B. Lin'), cell('F. Adeyemi', 'H. Cho')] },
      { sector: 'Corporate', color: '#5c2e91', cells: [cell('G. Rossi', 'I. Novak'), cell('Q. Blake', 'U. Stein'), cell('V. Holm', 'W. Park'), cell('Y. Diaz', 'Z. Klein')] },
    ],
  },
  architecture: {
    headers: ['Person', 'Backup', 'Network', 'Labs'],
    rows: [
      { sector: 'Manufacturing', color: '#0078d4', cells: [cell('J. Ortiz', 'P. Ng'), cell('P. Ng', 'A. Keller'), cell('S. Reed', 'D. Cho'), cell('H. Lang', 'M. Shaw')] },
      { sector: 'R&D', color: '#ca5010', cells: [cell('C. Young', 'D. Fox'), cell('D. Fox', 'N. Patel'), cell('I. Berg', 'T. Ruiz'), cell('A. Shah', 'B. Lin')] },
      { sector: 'Corporate', color: '#5c2e91', cells: [cell('Q. Blake', 'U. Stein'), cell('U. Stein', 'G. Rossi'), cell('K. Ade', 'W. Park'), cell('—', '—')] },
    ],
  },
  portfolio: {
    headers: ['Demand', 'Funding', 'Roadmap', 'Benefits'],
    rows: [
      { sector: 'Manufacturing', color: '#0078d4', cells: [cell('L. Berg', 'S. Cole'), cell('M. Shaw', 'A. Keller'), cell('P. Ng', 'J. Ortiz'), cell('S. Cole', 'L. Berg')] },
      { sector: 'R&D', color: '#ca5010', cells: [cell('E. Haas', 'B. Lin'), cell('T. Ruiz', 'N. Patel'), cell('D. Fox', 'C. Young'), cell('B. Lin', 'E. Haas')] },
      { sector: 'Corporate', color: '#5c2e91', cells: [cell('V. Holm', 'W. Park'), cell('I. Novak', 'G. Rossi'), cell('U. Stein', 'Q. Blake'), cell('W. Park', 'V. Holm')] },
    ],
  },
  service: {
    headers: ['Service owner', 'Process', 'Tooling', 'Major incident'],
    rows: [
      { sector: 'Manufacturing', color: '#0078d4', cells: [cell('L. Berg', 'S. Cole'), cell('S. Cole', 'R. Das'), cell('K. Moon', 'P. Ng'), cell('A. Keller', 'M. Shaw')] },
      { sector: 'R&D', color: '#ca5010', cells: [cell('E. Haas', 'B. Lin'), cell('B. Lin', 'H. Cho'), cell('F. Adeyemi', 'C. Young'), cell('N. Patel', 'T. Ruiz')] },
      { sector: 'Corporate', color: '#5c2e91', cells: [cell('V. Holm', 'W. Park'), cell('W. Park', 'Y. Diaz'), cell('Z. Klein', 'Q. Blake'), cell('G. Rossi', 'I. Novak')] },
    ],
  },
  security: {
    headers: ['Cyber lead', 'Identity', 'OT security', 'Vendor access'],
    rows: [
      { sector: 'Manufacturing', color: '#0078d4', cells: [cell('R. Das', 'K. Moon'), cell('K. Moon', 'A. Keller'), cell('J. Ortiz', 'P. Ng'), cell('M. Shaw', 'S. Cole')] },
      { sector: 'R&D', color: '#ca5010', cells: [cell('F. Adeyemi', 'H. Cho'), cell('H. Cho', 'N. Patel'), cell('C. Young', 'D. Fox'), cell('T. Ruiz', 'B. Lin')] },
      { sector: 'Corporate', color: '#5c2e91', cells: [cell('Y. Diaz', 'Z. Klein'), cell('Z. Klein', 'G. Rossi'), cell('Q. Blake', 'U. Stein'), cell('I. Novak', 'W. Park')] },
    ],
  },
}

export const confluencePages = [
  {
    id: 'home',
    title: 'Operations space',
    parent: null,
    body: 'Welcome to the operations knowledge space. Use the tree for standards, runbooks, and onboarding. Keep steps short and name an owner on every page.',
    attach: 'space-conventions.pdf',
    people: 'Owner A. Keller · Backup M. Shaw',
  },
  {
    id: 'standards',
    title: 'Standards',
    parent: 'home',
    body: 'Approved ways of working. If a plant needs an exception, record it as an ADR and bring it to the architecture forum.',
    attach: 'exception-template.docx',
    people: 'Owner J. Ortiz · Backup P. Ng',
  },
  {
    id: 'naming',
    title: 'Jump-host naming',
    parent: 'standards',
    body: 'Hostnames follow PLT-<site>-JH-<nn>. Place computer objects in the plant OU. Break-glass accounts are stored in the password vault, not on the desktop.',
    attach: 'naming-std-044.pdf',
    people: 'Owner S. Reed · Backup D. Cho',
  },
  {
    id: 'runbooks',
    title: 'Runbooks',
    parent: 'home',
    body: 'Operational runbooks for plant integrations. Clone the historian page when you add WMS or serialization.',
    attach: 'runbook-template.docx',
    people: 'Owner L. Berg · Backup S. Cole',
  },
  {
    id: 'outage',
    title: 'Historian outage',
    parent: 'runbooks',
    body: 'If the historian is unreachable, fail over to the buffer node, notify the service desk, and capture a timeline. Do not reboot the primary until the buffer has 15 minutes of clean data.',
    attach: 'historian-failover.pdf',
    people: 'Owner A. Keller · Backup M. Shaw',
  },
  {
    id: 'wms',
    title: 'WMS reconnect',
    parent: 'runbooks',
    body: 'Warehouse management reconnects through the site jump host. If the tunnel drops, do not share a personal VPN profile — open a catalog ticket.',
    attach: 'wms-reconnect.pdf',
    people: 'Owner K. Moon · Backup P. Ng',
  },
  {
    id: 'onboarding',
    title: 'New site onboarding',
    parent: 'home',
    body: 'Checklist for bringing a site onto the hub: master data, owners, training path, and first architecture review. Completeness should be above 70% before the site is announced in Engage.',
    attach: 'onboarding-checklist.xlsx',
    people: 'Owner G. Rossi · Backup I. Novak',
  },
  {
    id: 'faq',
    title: 'FAQ',
    parent: 'home',
    body: 'The hub is not a ticket system. Use the service catalog for access. Use this space for how we work. Use Engage for announcements.',
    attach: '',
    people: 'Owner V. Holm · Backup W. Park',
  },
]

export const engagePosts = [
  {
    id: 1,
    author: 'Alex Rivera',
    role: 'Architecture',
    time: '2h',
    announcement: true,
    text: 'Reference architecture v3 is published. Please use the new zone model for any plant network change from Monday.',
    likes: 18,
    replies: [{ author: 'Priya Shah', text: 'Does this change jump-host naming?' }],
  },
  {
    id: 2,
    author: 'Sam Okonkwo',
    role: 'North Plant',
    time: 'Yesterday',
    announcement: false,
    text: 'We closed the last open master-data fields for North Plant. Completion dashboard should now show 92%.',
    likes: 9,
    replies: [],
  },
  {
    id: 3,
    author: 'Jordan Lee',
    role: 'Enablement',
    time: '3d',
    announcement: false,
    text: 'New cybersecurity awareness module is in the trainings catalog. 20 minutes, required for site OT leads.',
    likes: 24,
    replies: [{ author: 'Marta Kovacs', text: 'Added it to the R&D curricula. Thanks!' }],
  },
  {
    id: 4,
    author: 'Mei Chen',
    role: 'Pacific Hub',
    time: '4d',
    announcement: false,
    text: 'Pilot DCS vendor session is booked for Thursday. Using the time-boxed tunnel in the cyber training — not a standing VPN.',
    likes: 7,
    replies: [{ author: 'R. Das', text: 'Good. Close the tunnel in the vault when they hang up.' }],
  },
  {
    id: 5,
    author: 'G. Rossi',
    role: 'Corporate',
    time: '1w',
    announcement: true,
    text: 'September hub digest went out Friday. Story clicks are on the newsletter demo if you want to see which pages moved.',
    likes: 11,
    replies: [],
  },
]

export const newsItems = [
  { title: 'Q3 program update', kicker: 'News', text: 'Four sites completed assessments; map and KPI views refreshed.' },
  { title: 'New training path', kicker: 'Enablement', text: 'Service management curricula now includes incident-to-problem.' },
  { title: 'Newsletter cadence', kicker: 'Add-on', text: 'Monthly digest with open/click analytics is live in the demo.' },
]

export const maintenanceEntries = [
  { id: 'new', label: 'Create a new signal' },
  { id: 'PM-1044', label: 'PM-1044 · Mixer-04 historian (North Plant)' },
  { id: 'PM-0981', label: 'PM-0981 · Pack-line PLC (Delta Pack)' },
]
