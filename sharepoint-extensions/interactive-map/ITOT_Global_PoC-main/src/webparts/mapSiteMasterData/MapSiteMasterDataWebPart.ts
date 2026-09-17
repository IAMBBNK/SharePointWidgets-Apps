import { Version } from '@microsoft/sp-core-library'
import {
	type IPropertyPaneConfiguration,
	PropertyPaneTextField,
} from '@microsoft/sp-property-pane'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'

import * as strings from 'MapSiteMasterDataWebPartStrings'

const BING_MAPS_KEY =
	'AgJMis6__ur7UfvExgEQb3R3EWCwZTTAYvhjRDPtG9ad-JlxHO3WzhzLIc9DHRzw'
const SITE_WEB_URL = 'https://mdigital.sharepoint.com/sites/ITOTCommunityHub'
const LIST_TITLE = 'Site Master Data'
const LIST_PATH = '/sites/ITOTCommunityHub/Lists/Site Master Data'
const MAP_ELEMENT_ID = 'map-smd-map'
const ICON_BASE =
	'https://mdigital.sharepoint.com/sites/ITOTCommunityHub/SiteAssets/Visuals/Buttons%20%26%20Icons/'

export interface IMapSiteMasterDataWebPartProps {
	description: string
}

interface IListField {
	InternalName: string
	Title: string
	TypeAsString?: string
	Hidden?: boolean
}

interface IFieldSpec {
	key: keyof IFormattedSite
	aliases: string[]
}

interface IFormattedSite {
	ID: number
	Title: string
	Sector: string
	SiteName: string
	BusinessUnit: string
	Location: string
	LegalEntityName: string
	LegalEntityCode: string
	Services: string
	Employees: string
	TargetDpi: string
	OtSystemsIntegrated: string
	ItAssessment: string
	OtAssessment: string
	DpiCoreImplementation: string
	OtSystemIntegration: string
	CurrentMaturity: string
	TargetMaturity: string
	Lattitude?: number
	Longitude?: number
}

const FIELD_SPECS: IFieldSpec[] = [
	{ key: 'Title', aliases: ['title'] },
	{ key: 'Sector', aliases: ['sector'] },
	{
		key: 'SiteName',
		aliases: ['site name', 'sitename', 'field_2'],
	},
	{
		key: 'BusinessUnit',
		aliases: ['business unit', 'businessunit', 'field_5'],
	},
	{
		key: 'Location',
		aliases: ['site address', 'location', 'city', 'field_8'],
	},
	{
		key: 'LegalEntityName',
		aliases: ['legal entity name', 'legalentityname', 'field_7'],
	},
	{
		key: 'LegalEntityCode',
		aliases: [
			'legal entity code',
			'legal entity number',
			'legalentitynumber',
			'site id',
			'siteid',
		],
	},
	{
		key: 'Services',
		aliases: [
			'business field / type / service',
			'business field',
			'services',
			'field_9',
		],
	},
	{
		key: 'Employees',
		aliases: ['# of employees', 'number of employees', 'employees', 'field_6'],
	},
	{
		key: 'TargetDpi',
		aliases: ['(target) dpi variant', 'target dpi variant', 'dpivariant'],
	},
	{
		key: 'OtSystemsIntegrated',
		aliases: [
			'total number of ot applications',
			'# of ot systems',
			'number of migrated ot applications',
			'_x0023_ofotsystems',
			'_x0023_ofmigratedotsystems',
			'odata__x0023_ofotsystems',
		],
	},
	{ key: 'ItAssessment', aliases: ['it assessment', 'itassessment'] },
	{ key: 'OtAssessment', aliases: ['ot assessment', 'otassessment'] },
	{
		key: 'DpiCoreImplementation',
		aliases: ['dpi core implementation', 'dpicoreimplementation'],
	},
	{
		key: 'OtSystemIntegration',
		aliases: ['ot system integration', 'otsystemintegration'],
	},
	{
		key: 'CurrentMaturity',
		aliases: [
			'current maturity (bpog)',
			'current maturity',
			'currentmaturity_x0028_bpog_x0029_',
			'currentmaturity_x0028_bpog_x0029',
		],
	},
	{
		key: 'TargetMaturity',
		aliases: [
			'target maturity (bpog)',
			'target maturity',
			'targetmaturity_x0028_bpog_x0029_',
			'targetmaturity_x0028_bpog_x0029',
		],
	},
]

const SECTOR_COLORS: Record<string, { fill: string; css: string }> = {
	healthcare: { fill: '#eb3c96', css: 'sector-pink' },
	'health care': { fill: '#eb3c96', css: 'sector-pink' },
	'life science': { fill: '#ffc832', css: 'sector-yellow' },
	electronics: { fill: '#4ca5e9', css: 'sector-blue' },
}

declare global {
	interface Window {
		Microsoft?: any
		initMapSiteMaster?: () => void
	}
}

export default class MapSiteMasterDataWebPart extends BaseClientSideWebPart<IMapSiteMasterDataWebPartProps> {
	private _mapContainer: HTMLDivElement | null = null
	private _mapId = MAP_ELEMENT_ID
	private _mapReadyGen = -1
	private _initGen = 0

	public render(): void {
		this._initGen++
		this.domElement.innerHTML = ''
		const container = document.createElement('div')
		container.className = 'map-itot-container'
		container.innerHTML = `
      <div class="filter-container">
        <div>Select a sector:</div>
        <button type="button" class="sector-filter active-filter" data-id="all">All</button>
        <button type="button" class="sector-filter" data-id="Life Science">
          <span class="sector-icons life-science-icon"></span>Life Science
        </button>
        <button type="button" class="sector-filter" data-id="Healthcare">
          <span class="sector-icons healthcare-icon"></span>Healthcare
        </button>
        <button type="button" class="sector-filter" data-id="Electronics">
          <span class="sector-icons electronics-icon"></span>Electronics
        </button>
        <div class="accordion-container">
          <button type="button" class="accordion" id="smd-unit-toggle">Select a Business Unit</button>
          <div class="panel" id="smd-dynamic-unit">
            <button type="button" class="apply-filter" id="smd-apply-filter">Apply</button>
          </div>
        </div>
      </div>
      <div class="smd-status" id="smd-status">Loading Site Master Data…</div>`
		this._mapId =
			MAP_ELEMENT_ID + '-' + String(this.instanceId || '0').replace(/[^a-z0-9]/gi, '')
		this._mapContainer = document.createElement('div')
		this._mapContainer.id = this._mapId
		this._mapContainer.style.height = '700px'
		this._mapContainer.style.minHeight = '700px'
		this._mapContainer.style.width = '100%'
		this._mapContainer.style.minWidth = '100%'
		this._mapContainer.style.position = 'relative'
		container.appendChild(this._mapContainer)
		this._injectStyles()
		this.domElement.appendChild(container)
		this._loadBingAndInit()
	}

	private _injectStyles(): void {
		const style = document.createElement('style')
		style.textContent = `
      .map-itot-container { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
      .filter-container {
        display: flex;
        gap: 15px;
        align-items: center;
        margin-bottom: 8px;
        font-weight: bold;
        justify-content: center;
        flex-wrap: wrap;
      }
      .smd-status {
        text-align: center;
        font-size: 13px;
        font-weight: 500;
        color: #555;
        margin-bottom: 10px;
      }
      .smd-status.error { color: #a4262c; }
      .sector-icons {
        width: 20px;
        height: 20px;
        border-radius: 50%;
        display: inline-block;
      }
      .life-science-icon { background-color: #ffc832; }
      .healthcare-icon { background-color: #eb3c96; }
      .electronics-icon { background-color: #4ca5e9; }
      .sector-filter, .accordion, .apply-filter {
        background-color: #1872b9;
        padding: 15px;
        color: white;
        text-transform: uppercase;
        cursor: pointer;
        border-radius: 5px;
        border: none;
        font-family: inherit;
        font-size: 13px;
        font-weight: 700;
      }
      .sector-filter { display: flex; gap: 10px; align-items: center; }
      .active-filter { background-color: #419ce3; }
      .accordion-container { position: relative; width: 330px; }
      .accordion.active { border-radius: 5px 5px 0 0; }
      .accordion:after {
        content: url(https://mdigital.sharepoint.com/sites/ITOTCommunityHub/SiteAssets/Visuals/Buttons%20%26%20Icons/dropdown.png);
        float: left;
        margin-top: 10px;
        margin-right: 25px;
        zoom: 25%;
      }
      .panel {
        padding: 15px;
        display: none;
        background-color: #419ce3;
        color: white;
        overflow: auto;
        max-height: 280px;
        flex-direction: column;
        gap: 5px;
        position: absolute;
        width: 300px;
        left: 0;
        z-index: 1003;
        border-radius: 0 0 5px 5px;
      }
      .panel.open { display: flex; }
      .panel-item { display: flex; gap: 5px; align-items: center; font-weight: 500; }
      .apply-filter { margin-top: 10px; width: fit-content; }
      .map-tooltip {
        background-color: white;
        color: #333;
        padding: 15px;
        border-radius: 8px;
        overflow: hidden;
        width: 500px;
      }
      .map-tooltip.sector-pink { border-top: 15px solid #eb3c96; }
      .map-tooltip.sector-yellow { border-top: 15px solid #ffc832; }
      .map-tooltip.sector-blue { border-top: 15px solid #4ca5e9; }
      .map-tooltip.sector-mixed { border-top: 15px solid #503291; }
      .map-tooltip.map-tooltip-cluster { width: 360px; }
      .map-tooltip-header { position: relative; padding-bottom: 10px; }
      .map-tooltip-headline { font-size: 18px; font-weight: 600; margin: 0; padding-right: 24px; }
      .map-tooltip-close {
        position: absolute; top: -10px; right: 0;
        font-size: 25px; font-weight: 600; cursor: pointer; line-height: 1; color: #666;
      }
      .map-tooltip-items { display: flex; margin-top: 5px; gap: 25px; }
      .map-tooltip-block { flex: 1; min-width: 0; }
      .map-tooltip-item {
        display: flex;
        gap: 5px;
        align-items: center;
        margin-bottom: 5px;
        font-size: 14px;
      }
      .map-tooltip-item img { width: 35px; object-fit: contain; }
      .tooltip-legend {
        justify-content: center;
        background-color: #f4f4f4;
        padding: 10px;
        border-radius: 10px;
        display: flex;
        gap: 10px;
        flex: 1;
        flex-wrap: wrap;
      }
      .tooltip-legend > div {
        display: flex;
        gap: 6px;
        font-size: 11px;
        align-items: center;
      }
      .tooltip-legend img { width: 22px; height: 22px; object-fit: contain; }
      .map-tooltip-footer {
        display: flex;
        gap: 10px;
        align-items: center;
        margin-top: 8px;
      }
      .map-tooltip .see-all {
        cursor: pointer;
        padding: 8px 15px;
        background-color: #0f69af;
        color: white;
        text-decoration: none;
        border-radius: 10px;
        font-size: 12px;
        font-weight: 500;
        white-space: nowrap;
      }
      .map-cluster-list { list-style: none; margin: 0; padding: 0; }
      .map-cluster-list-paged { min-height: 210px; }
      .map-cluster-item {
        display: flex;
        align-items: center;
        gap: 10px;
        width: 100%;
        margin: 0;
        padding: 10px 6px;
        border: none;
        border-bottom: 1px solid #eee;
        background: transparent;
        text-align: left;
        cursor: pointer;
        font-family: inherit;
        font-size: 14px;
        color: #333;
      }
      .map-cluster-item:last-child { border-bottom: none; }
      .map-cluster-item:hover { background: #f5f5f5; }
      .map-cluster-dot {
        width: 12px;
        height: 12px;
        border-radius: 50%;
        border: 2px solid #0f69af;
        flex-shrink: 0;
      }
      .map-cluster-item-text { min-width: 0; flex: 1; }
      .map-cluster-pager {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
        padding-top: 10px;
        margin-top: 4px;
        border-top: 1px solid #eee;
      }
      .map-cluster-page-btn {
        width: 32px;
        height: 32px;
        padding: 0;
        border: 1px solid #d0d0d0;
        border-radius: 6px;
        background: #fff;
        color: #333;
        font-size: 20px;
        line-height: 1;
        cursor: pointer;
        font-family: inherit;
      }
      .map-cluster-page-btn:hover:not(:disabled) { background: #f5f5f5; }
      .map-cluster-page-btn:disabled { opacity: 0.35; cursor: default; }
      .map-cluster-page-label { font-size: 13px; color: #555; font-weight: 600; }
    `
		this.domElement.appendChild(style)
	}

	private _hideBingCredentialsError(): void {
		const credsText = 'The specified credentials are invalid'
		const root = this._mapContainer || this.domElement
		const divs = root.getElementsByTagName('div')
		for (let i = 0; i < divs.length; i++) {
			const el = divs[i]
			const style = el.getAttribute('style') || ''
			const hasCredsText =
				el.textContent && (el.textContent as string).indexOf(credsText) >= 0
			const hasCredsBackground =
				style.indexOf('248') >= 0 &&
				style.indexOf('247') >= 0 &&
				style.indexOf('245') >= 0
			if (hasCredsText && hasCredsBackground) {
				el.style.display = 'none'
				break
			}
		}
	}

	private _sectorKey(value: string): string {
		return (value || '').toLowerCase().replace(/[\s_-]+/g, '')
	}

	private _sectorsEqual(a: string, b: string): boolean {
		const na = this._sectorKey(a)
		const nb = this._sectorKey(b)
		if (!na || !nb) return false
		return na === nb || na.indexOf(nb) >= 0 || nb.indexOf(na) >= 0
	}

	private _sectorColor(sector: string): { fill: string; css: string } {
		const key = (sector || '').toLowerCase().trim()
		if (SECTOR_COLORS[key]) return SECTOR_COLORS[key]
		const names = Object.keys(SECTOR_COLORS)
		for (let i = 0; i < names.length; i++) {
			if (key.indexOf(names[i]) >= 0) return SECTOR_COLORS[names[i]]
		}
		return SECTOR_COLORS.electronics
	}

	private _fieldValue(val: unknown): string {
		if (val == null || val === '') return '-'
		if (typeof val === 'string') {
			const t = val.trim()
			return t === '' ? '-' : t
		}
		if (typeof val === 'number' || typeof val === 'boolean') return String(val)
		if (typeof val === 'object') {
			const obj = val as {
				Title?: string
				Label?: string
				Results?: unknown[]
				results?: unknown[]
			}
			if (obj.Title && String(obj.Title).trim()) return String(obj.Title).trim()
			if (obj.Label && String(obj.Label).trim()) return String(obj.Label).trim()
			const nested = obj.Results || obj.results
			if (Array.isArray(nested) && nested[0] != null) {
				return this._fieldValue(nested[0])
			}
		}
		return '-'
	}

	private _text(item: Record<string, unknown>, keys: string[]): string {
		return this._fieldValue(this._readItemValue(item, keys))
	}

	private _normName(s: string): string {
		return (s || '')
			.toLowerCase()
			.replace(/_x0020_/g, ' ')
			.replace(/_x0023_/g, '#')
			.replace(/_x0028_/g, '(')
			.replace(/_x0029_/g, ')')
			.replace(/[^a-z0-9]/g, '')
	}

	private _toODataSelectName(internalName: string): string {
		const n = (internalName || '').trim()
		if (!n) return n
		if (n.indexOf('OData_') === 0) return n
		if (n.charAt(0) === '_') return 'OData_' + n
		return n
	}

	private _parseCoord(raw: unknown): number | undefined {
		if (raw == null || raw === '') return undefined
		if (typeof raw === 'object') {
			const obj = raw as {
				Latitude?: number | string
				Longitude?: number | string
				results?: unknown[]
				Results?: unknown[]
			}
			if (obj.Latitude != null) return this._parseCoord(obj.Latitude)
			if (obj.Longitude != null) return this._parseCoord(obj.Longitude)
			const nested = obj.Results || obj.results
			if (Array.isArray(nested) && nested[0] != null)
				return this._parseCoord(nested[0])
		}
		const n = parseFloat(String(raw).replace(',', '.'))
		return typeof n === 'number' && isFinite(n) && !isNaN(n) ? n : undefined
	}

	private _normalizeCoords(
		lat?: number,
		lng?: number,
	): { lat: number; lng: number } | undefined {
		if (lat == null || lng == null || isNaN(lat) || isNaN(lng)) return undefined
		const latOk = lat >= -90 && lat <= 90
		const lngOk = lng >= -180 && lng <= 180
		if (latOk && lngOk && lat !== 0 && lng !== 0) return { lat, lng }
		if (!latOk && lng >= -90 && lng <= 90 && lat >= -180 && lat <= 180 && lng !== 0) {
			console.warn('[MapSMD] Swapping lat/lng (latitude was out of range)', lat, lng)
			return { lat: lng, lng: lat }
		}
		return undefined
	}

	private _itemsFromJson(json: any): Record<string, unknown>[] | undefined {
		if (!json || json.error) return undefined
		if (json.d && Array.isArray(json.d.results)) return json.d.results
		if (json.d && json.d.RenderListData && Array.isArray(json.d.RenderListData.Row))
			return json.d.RenderListData.Row
		if (json.RenderListData && Array.isArray(json.RenderListData.Row))
			return json.RenderListData.Row
		if (Array.isArray(json.Row)) return json.Row
		if (Array.isArray(json.value)) return json.value
		if (Array.isArray(json)) return json
		return undefined
	}

	private _nextLinkFromJson(json: any): string | undefined {
		if (!json) return undefined
		const next =
			json['@odata.nextLink'] ||
			(json.d && json.d.__next) ||
			json.NextHref ||
			(json.d && json.d.NextHref) ||
			(json.d && json.d.RenderListData && json.d.RenderListData.NextHref) ||
			(json.RenderListData && json.RenderListData.NextHref)
		return typeof next === 'string' && next ? next : undefined
	}

	private _readItemValue(
		item: Record<string, unknown>,
		names: string[],
	): unknown {
		for (let i = 0; i < names.length; i++) {
			const name = names[i]
			if (!name) continue
			if (Object.prototype.hasOwnProperty.call(item, name)) {
				const val = item[name]
				if (val != null && val !== '') return val
			}
			const odata = this._toODataSelectName(name)
			if (odata !== name && Object.prototype.hasOwnProperty.call(item, odata)) {
				const val = item[odata]
				if (val != null && val !== '') return val
			}
		}
		const wanted = names.map((n) => this._normName(n)).filter(Boolean)
		if (!wanted.length) return undefined
		const keys = Object.keys(item)
		for (let i = 0; i < keys.length; i++) {
			const key = keys[i]
			if (key === '__metadata') continue
			const nk = this._normName(key)
			if (wanted.indexOf(nk) >= 0 && item[key] != null && item[key] !== '')
				return item[key]
		}
		return undefined
	}

	private _matchField(
		fields: IListField[],
		aliases: string[],
	): IListField | undefined {
		const hits: IListField[] = []
		for (let i = 0; i < fields.length; i++) {
			const f = fields[i]
			const title = (f.Title || '').toLowerCase().trim()
			const internal = this._normName(f.InternalName)
			for (let a = 0; a < aliases.length; a++) {
				const alias = aliases[a].toLowerCase().trim()
				if (
					title === alias ||
					this._normName(alias) === internal ||
					(f.InternalName || '').toLowerCase() === alias
				) {
					hits.push(f)
					break
				}
			}
		}
		if (!hits.length) return undefined
		const visible = hits.filter((f) => !f.Hidden)
		return visible[0] || hits[0]
	}

	private _listApiRoots(): string[] {
		return [
			SITE_WEB_URL +
				"/_api/web/lists/getbytitle('" +
				LIST_TITLE.replace(/'/g, "''") +
				"')",
			SITE_WEB_URL +
				'/_api/web/lists/getbytitle(%27' +
				encodeURIComponent(LIST_TITLE) +
				'%27)',
			SITE_WEB_URL +
				"/_api/web/GetList('" +
				LIST_PATH.replace(/ /g, '%20') +
				"')",
		]
	}

	private async _spRequest(
		url: string,
		method: 'GET' | 'POST',
		body?: string,
		extraHeaders?: Record<string, string>,
	): Promise<{ ok: boolean; status: number; json?: any; text: string }> {
		const headers: Record<string, string> = {
			Accept: 'application/json;odata=verbose',
			...(extraHeaders || {}),
		}
		if (method === 'POST' && !headers['Content-Type'] && !headers['content-type']) {
			headers['Content-Type'] = 'application/json;odata=verbose'
		}
		try {
			const res = await fetch(url, {
				method,
				headers,
				credentials: 'include',
				body,
			})
			const text = await res.text()
			let json: any
			try {
				json = text ? JSON.parse(text) : undefined
			} catch (_e) {
				json = undefined
			}
			if (!res.ok) {
				console.warn('[MapSMD] fetch', method, res.status, url, text.slice(0, 250))
			}
			return { ok: res.ok, status: res.status, json, text }
		} catch (err) {
			console.warn('[MapSMD] Request error', method, url, err)
			return { ok: false, status: 0, text: String(err) }
		}
	}

	private async _getDigest(): Promise<string | undefined> {
		const res = await this._spRequest(SITE_WEB_URL + '/_api/contextinfo', 'POST')
		if (!res.ok || !res.json) return undefined
		const digest =
			res.json.d &&
			res.json.d.GetContextWebInformation &&
			res.json.d.GetContextWebInformation.FormDigestValue
		return typeof digest === 'string' ? digest : undefined
	}

	private async _fetchPagedItems(startUrl: string): Promise<Record<string, unknown>[] | undefined> {
		const collected: Record<string, unknown>[] = []
		let nextUrl: string | undefined = startUrl
		let pages = 0
		while (nextUrl && pages < 50) {
			pages++
			const absolute =
				nextUrl.indexOf('http') === 0
					? nextUrl
					: SITE_WEB_URL.replace(/\/$/, '') +
					  (nextUrl.charAt(0) === '/' ? nextUrl : '/' + nextUrl)
			const res = await this._spRequest(absolute, 'GET')
			if (!res.ok) {
				console.warn(
					'[MapSMD] Items page failed',
					res.status,
					absolute,
					(res.text || '').slice(0, 300),
				)
				return collected.length ? collected : undefined
			}
			const page = this._itemsFromJson(res.json)
			if (!page) {
				console.warn('[MapSMD] Unexpected items payload', absolute, res.json && res.json.error)
				return collected.length ? collected : undefined
			}
			collected.push.apply(collected, page)
			const rawNext = this._nextLinkFromJson(res.json)
			nextUrl = rawNext
			console.log(
				'[MapSMD] Page',
				pages,
				'items',
				page.length,
				'total',
				collected.length,
				nextUrl ? '(more)' : '(done)',
			)
		}
		return collected
	}

	private _siteHeadline(meta: IFormattedSite): string {
		const name = meta.SiteName && meta.SiteName !== '-' ? meta.SiteName : ''
		const bu =
			meta.BusinessUnit && meta.BusinessUnit !== '-' ? meta.BusinessUnit : ''
		if (name && bu) return name + ' / ' + bu
		return name || bu || 'Site'
	}

	private _setStatus(message: string, isError?: boolean): void {
		const el = this.domElement.querySelector('#smd-status') as HTMLElement | null
		if (!el) return
		el.textContent = message
		if (isError) el.classList.add('error')
		else el.classList.remove('error')
		console.log('[MapSMD] Status:', message)
	}

	private _coordsFromItem(item: Record<string, unknown>): {
		lat?: number
		lng?: number
	} {
		const latRaw =
			item.Lattitude ??
			item.Latitude ??
			item.lat ??
			item.Lat ??
			this._readItemValue(item, ['Lattitude', 'Latitude', 'lat', 'Lat'])
		const lngRaw =
			item.Longitude ??
			item.Longtitude ??
			item.lng ??
			item.Lon ??
			this._readItemValue(item, ['Longitude', 'Longtitude', 'lng', 'Lon', 'Long'])
		const lat = this._parseCoord(latRaw)
		const lng = this._parseCoord(lngRaw)
		if (latRaw && typeof latRaw === 'object') {
			const geo = latRaw as { Latitude?: unknown; Longitude?: unknown }
			return (
				this._normalizeCoords(
					this._parseCoord(geo.Latitude) ?? lat,
					this._parseCoord(geo.Longitude) ?? lng,
				) || {}
			)
		}
		return this._normalizeCoords(lat, lng) || {}
	}

	private _assessmentIcon(value: string): string {
		const raw = (value || '').trim()
		const file =
			!raw || raw === '-' || raw.toUpperCase() === 'N/A' ? 'NA' : raw
		return ICON_BASE + encodeURIComponent(file) + '.svg'
	}

	private async _getSiteMasterData(): Promise<IFormattedSite[]> {
		this._setStatus('Loading Site Master Data…')
		console.log('[MapSMD] Fetching Site Master Data from', SITE_WEB_URL)
		const fields: IListField[] = []
		const resolved: Partial<Record<keyof IFormattedSite, IListField>> = {}
		if (fields.length) {
			for (let i = 0; i < FIELD_SPECS.length; i++) {
				const spec = FIELD_SPECS[i]
				const field = this._matchField(fields, spec.aliases)
				if (field) resolved[spec.key] = field
				else console.warn('[MapSMD] Column not resolved for', spec.key, spec.aliases)
			}
		}
		const latField = this._matchField(fields, [
			'lattitude',
			'latitude',
			'lat',
		])
		const lngField = this._matchField(fields, [
			'longitude',
			'longtitude',
			'lng',
			'lon',
			'long',
		])
		console.log('[MapSMD] Resolved columns', {
			siteName: resolved.SiteName && resolved.SiteName.InternalName,
			sector: resolved.Sector && resolved.Sector.InternalName,
			businessUnit: resolved.BusinessUnit && resolved.BusinessUnit.InternalName,
			location: resolved.Location && resolved.Location.InternalName,
			lat: latField && latField.InternalName,
			lng: lngField && lngField.InternalName,
			itAssessment: resolved.ItAssessment && resolved.ItAssessment.InternalName,
			dpiVariant: resolved.TargetDpi && resolved.TargetDpi.InternalName,
		})

		const addSelect = (names: string[], field?: IListField, extra?: string[]): void => {
			const candidates = extra ? extra.slice() : []
			if (field) {
				candidates.unshift(field.InternalName)
				candidates.unshift(this._toODataSelectName(field.InternalName))
			}
			for (let i = 0; i < candidates.length; i++) {
				const n = candidates[i]
				if (n && names.indexOf(n) < 0) names.push(n)
			}
		}

		const fullSelect: string[] = ['ID', 'Id', 'Title']
		addSelect(fullSelect, resolved.Sector, ['Sector'])
		addSelect(fullSelect, resolved.SiteName, ['field_2'])
		addSelect(fullSelect, resolved.BusinessUnit, ['field_5'])
		addSelect(fullSelect, resolved.Employees, ['field_6'])
		addSelect(fullSelect, resolved.LegalEntityName, ['field_7'])
		addSelect(fullSelect, resolved.Location, ['field_8'])
		addSelect(fullSelect, resolved.Services, ['field_9'])
		addSelect(fullSelect, resolved.LegalEntityCode, ['LegalEntityNumber'])
		addSelect(fullSelect, resolved.TargetDpi, ['DPIVariant'])
		addSelect(fullSelect, resolved.ItAssessment, ['ITAssessment'])
		addSelect(fullSelect, resolved.OtAssessment, ['OTAssessment'])
		addSelect(fullSelect, resolved.DpiCoreImplementation, ['DPICoreImplementation'])
		addSelect(fullSelect, resolved.OtSystemIntegration, ['OTSystemIntegration'])
		addSelect(fullSelect, resolved.CurrentMaturity, [
			'CurrentMaturity_x0028_BPOG_x0029_',
			'CurrentMaturity_x0028_BPOG_x0029',
		])
		addSelect(fullSelect, resolved.TargetMaturity, [
			'TargetMaturity_x0028_BPOG_x0029_',
			'TargetMaturity_x0028_BPOG_x0029',
		])
		addSelect(fullSelect, resolved.OtSystemsIntegrated, [
			'_x0023_ofOTSystems',
			'OData__x0023_ofOTSystems',
			'_x0023_ofmigratedOTSystems',
			'OData__x0023_ofmigratedOTSystems',
		])
		addSelect(fullSelect, latField, ['Lattitude', 'Latitude'])
		addSelect(fullSelect, lngField, ['Longitude'])

		const coreSelect: string[] = ['ID', 'Id', 'Title', 'Sector', 'field_2', 'field_5', 'field_6', 'field_7', 'field_8', 'field_9', 'LegalEntityNumber', 'Lattitude', 'Latitude', 'Longitude']
		addSelect(coreSelect, resolved.SiteName)
		addSelect(coreSelect, resolved.Sector)
		addSelect(coreSelect, latField)
		addSelect(coreSelect, lngField)

		const provenSelect =
			'ID,Title,Sector,field_2,field_5,field_6,field_7,field_8,field_9,LegalEntityNumber,Lattitude,Longitude'
		const widgetSelect =
			'ID,Title,Sector,field_2,field_5,field_6,field_7,field_8,field_9,DPIVariant,TargetMaturity_x0028_BPOG_x0029_,ITAssessment,OTAssessment,OTSystemIntegration,DPICoreImplementation,CurrentMaturity_x0028_BPOG_x0029,Lattitude,Longitude,OData__x0023_ofOTSystems,OData__x0023_ofmigratedOTSystems,LegalEntityNumber'
		const byTitle =
			SITE_WEB_URL +
			"/_api/web/lists/getbytitle('Site Master Data')/items?$top=1000"
		const urls: string[] = [
			byTitle + '&$select=' + provenSelect,
			SITE_WEB_URL +
				'/_api/web/lists/getbytitle(%27' +
				encodeURIComponent(LIST_TITLE) +
				'%27)/items?$top=1000&$select=' +
				provenSelect,
			byTitle + '&$select=' + widgetSelect,
			byTitle,
		]
		const roots = this._listApiRoots()
		for (let i = 0; i < roots.length; i++) {
			urls.push(
				roots[i] +
					'/items?$top=1000&$select=' +
					provenSelect,
			)
			urls.push(
				roots[i] +
					'/items?$top=5000&$select=' +
					encodeURIComponent(coreSelect.join(',')),
			)
			urls.push(roots[i] + '/items?$top=5000')
		}

		let results: Record<string, unknown>[] | undefined
		for (let i = 0; i < urls.length; i++) {
			console.log('[MapSMD] Trying items URL', i + 1 + '/' + urls.length, urls[i])
			const items = await this._fetchPagedItems(urls[i])
			if (items && items.length) {
				results = items
				console.log('[MapSMD] Loaded', items.length, 'items from', urls[i])
				break
			}
			if (items && !items.length) {
				console.warn('[MapSMD] URL returned 0 items', urls[i])
			}
		}

		if (!results) {
			console.warn('[MapSMD] REST items failed; trying RenderListDataAsStream')
			results = await this._fetchRenderListData(
				fullSelect.filter((n) => n !== 'Id' && n.indexOf('/') < 0),
			)
		}

		if (!results) {
			console.error('[MapSMD] Site Master Data fetch failed')
			this._setStatus(
				'Could not load Site Master Data. Open the console (F12) and look for [MapSMD] errors.',
				true,
			)
			return []
		}

		const namesFor = (key: keyof IFormattedSite, extras: string[]): string[] => {
			const field = resolved[key]
			const names = extras.slice()
			if (field) {
				names.unshift(field.Title)
				names.unshift(this._toODataSelectName(field.InternalName))
				names.unshift(field.InternalName)
			}
			return names
		}

		const mapped = results.map((item) => {
			const coords = this._coordsFromItem(item)
			const lat = coords.lat
			const lng = coords.lng
			const idRaw = this._readItemValue(item, ['ID', 'Id', 'ID0'])
			return {
				ID: Number(idRaw),
				Title: this._text(item, namesFor('Title', ['Title'])),
				Sector: this._text(item, namesFor('Sector', ['Sector'])),
				SiteName: this._text(
					item,
					namesFor('SiteName', ['field_2', 'Site Name', 'Title']),
				),
				BusinessUnit: this._text(
					item,
					namesFor('BusinessUnit', ['field_5', 'Business Unit']),
				),
				Location: this._text(
					item,
					namesFor('Location', ['field_8', 'Site Address', 'City']),
				),
				LegalEntityName: this._text(
					item,
					namesFor('LegalEntityName', ['field_7', 'Legal Entity Name']),
				),
				LegalEntityCode: this._text(
					item,
					namesFor('LegalEntityCode', ['LegalEntityNumber', 'Site ID']),
				),
				Services: this._text(
					item,
					namesFor('Services', ['field_9', 'Business Field / Type / Service']),
				),
				Employees: this._text(
					item,
					namesFor('Employees', ['field_6', '# of Employees']),
				),
				TargetDpi: this._text(
					item,
					namesFor('TargetDpi', ['DPIVariant', '(Target) DPI Variant']),
				),
				OtSystemsIntegrated: this._text(
					item,
					namesFor('OtSystemsIntegrated', [
						'OData__x0023_ofOTSystems',
						'_x0023_ofOTSystems',
						'OData__x0023_ofmigratedOTSystems',
						'_x0023_ofmigratedOTSystems',
					]),
				),
				ItAssessment: this._text(
					item,
					namesFor('ItAssessment', ['ITAssessment', 'IT Assessment']),
				),
				OtAssessment: this._text(
					item,
					namesFor('OtAssessment', ['OTAssessment', 'OT Assessment']),
				),
				DpiCoreImplementation: this._text(
					item,
					namesFor('DpiCoreImplementation', [
						'DPICoreImplementation',
						'DPI Core Implementation',
					]),
				),
				OtSystemIntegration: this._text(
					item,
					namesFor('OtSystemIntegration', [
						'OTSystemIntegration',
						'OT System Integration',
					]),
				),
				CurrentMaturity: this._text(
					item,
					namesFor('CurrentMaturity', [
						'CurrentMaturity_x0028_BPOG_x0029_',
						'CurrentMaturity_x0028_BPOG_x0029',
						'Current Maturity (BPOG)',
					]),
				),
				TargetMaturity: this._text(
					item,
					namesFor('TargetMaturity', [
						'TargetMaturity_x0028_BPOG_x0029_',
						'TargetMaturity_x0028_BPOG_x0029',
						'Target Maturity (BPOG)',
					]),
				),
				Lattitude: lat,
				Longitude: lng,
			}
		})

		const withCoords = mapped.filter(
			(s) => s.Lattitude != null && s.Longitude != null,
		)
		const sectors: string[] = []
		for (let i = 0; i < mapped.length; i++) {
			const sector = mapped[i].Sector
			if (sector && sector !== '-' && sectors.indexOf(sector) < 0)
				sectors.push(sector)
		}
		console.log('[MapSMD] --- Data fetch summary ---')
		console.log('[MapSMD] Items:', mapped.length, '| with coords:', withCoords.length)
		console.log('[MapSMD] Sectors:', sectors)
		console.log(
			'[MapSMD] First raw keys:',
			results[0] ? Object.keys(results[0]).filter((k) => k !== '__metadata') : [],
		)
		if (mapped[0]) {
			console.log('[MapSMD] First mapped site:', {
				ID: mapped[0].ID,
				SiteName: mapped[0].SiteName,
				Sector: mapped[0].Sector,
				BusinessUnit: mapped[0].BusinessUnit,
				Location: mapped[0].Location,
				Lattitude: mapped[0].Lattitude,
				Longitude: mapped[0].Longitude,
				ItAssessment: mapped[0].ItAssessment,
				CurrentMaturity: mapped[0].CurrentMaturity,
			})
		}
		if (!withCoords.length) {
			console.warn(
				'[MapSMD] No sites have coordinates. Check Lattitude/Longitude columns and the first-row keys above.',
			)
			this._setStatus(
				'Loaded ' +
					mapped.length +
					' sites but none have Lattitude/Longitude. Pins cannot be drawn.',
				true,
			)
		} else {
			this._setStatus(
				'Loaded ' +
					mapped.length +
					' sites, ' +
					withCoords.length +
					' with coordinates. Filter: All.',
			)
		}
		console.log('[MapSMD] --- end summary ---')
		return mapped
	}

	private async _fetchRenderListData(
		viewFields: string[],
	): Promise<Record<string, unknown>[] | undefined> {
		const digest = await this._getDigest()
		if (!digest) {
			console.warn('[MapSMD] No form digest; skipping RenderListDataAsStream')
			return undefined
		}
		const unique: string[] = []
		for (let i = 0; i < viewFields.length; i++) {
			const raw = viewFields[i]
			if (!raw) continue
			const name = raw.indexOf('OData_') === 0 ? raw.slice('OData_'.length) : raw
			if (unique.indexOf(name) < 0) unique.push(name)
		}
		const fieldRefs = unique
			.map((n) => "<FieldRef Name='" + String(n).replace(/'/g, '&apos;') + "' />")
			.join('')
		const viewXml =
			"<View Scope='RecursiveAll'><ViewFields>" +
			fieldRefs +
			"</ViewFields><RowLimit Paged='TRUE'>5000</RowLimit></View>"
		const collected: Record<string, unknown>[] = []
		let nextHref: string | undefined
		let pages = 0
		const startUrl = this._listApiRoots()[0] + '/RenderListDataAsStream'
		while (pages < 50) {
			pages++
			const url = nextHref
				? SITE_WEB_URL.replace(/\/$/, '') +
				  (nextHref.charAt(0) === '/' ? nextHref : '/' + nextHref)
				: startUrl
			console.log('[MapSMD] RenderListDataAsStream page', pages, url)
			const res = await this._spRequest(
				url,
				'POST',
				JSON.stringify({
					parameters: { RenderOptions: 2, ViewXml: viewXml },
				}),
				{ 'X-RequestDigest': digest },
			)
			if (!res.ok) {
				console.warn(
					'[MapSMD] RenderListDataAsStream failed',
					res.status,
					(res.text || '').slice(0, 300),
				)
				return collected.length ? collected : undefined
			}
			const page = this._itemsFromJson(res.json)
			if (!page) {
				console.warn('[MapSMD] Unexpected RenderListData payload')
				return collected.length ? collected : undefined
			}
			collected.push.apply(collected, page)
			nextHref = this._nextLinkFromJson(res.json)
			if (!nextHref) break
		}
		console.log('[MapSMD] RenderListDataAsStream loaded', collected.length, 'rows')
		return collected.length ? collected : undefined
	}

	private _loadBingAndInit(): void {
		const self = this
		const start = (): void => {
			void self._initMap()
		}
		window.initMapSiteMaster = start
		if (window.Microsoft && window.Microsoft.Maps) {
			start()
			return
		}
		if (!document.getElementById('bing-maps-script-itot')) {
			const script = document.createElement('script')
			script.id = 'bing-maps-script-itot'
			script.type = 'text/javascript'
			script.src = `https://www.bing.com/api/maps/mapcontrol?key=${BING_MAPS_KEY}&callback=initMapSiteMaster`
			script.async = true
			script.defer = true
			document.head.appendChild(script)
		}
		const started = Date.now()
		const poll = (): void => {
			if (window.Microsoft && window.Microsoft.Maps) {
				start()
				return
			}
			if (Date.now() - started < 15000) setTimeout(poll, 200)
		}
		poll()
	}

	private async _initMap(): Promise<void> {
		const self = this
		const gen = this._initGen
		const mapEl = this._mapContainer
		if (!mapEl || !window.Microsoft || !window.Microsoft.Maps) return
		if (this._mapReadyGen === gen) return
		this._mapReadyGen = gen
		const mapHost: HTMLDivElement = mapEl

		const Microsoft = window.Microsoft
		const map = new Microsoft.Maps.Map(mapEl, {
			credentials: BING_MAPS_KEY,
			center: new Microsoft.Maps.Location(51.165691, 10.451526),
			zoom: 2,
			disableScrollWheelZoom: false,
			mapTypeId: Microsoft.Maps.MapTypeId.aerial,
			customMapStyle: {
				version: '1.0',
				elements: {
					point: { labelVisible: true, visible: false },
					road: { visible: false, labelVisible: false },
				},
			},
		})

		this._hideBingCredentialsError()
		setTimeout(() => this._hideBingCredentialsError(), 500)
		setTimeout(() => this._hideBingCredentialsError(), 2000)

		const tooltip = new Microsoft.Maps.Infobox(map.getCenter(), {
			visible: false,
			showPointer: false,
			showCloseButton: false,
			offset: new Microsoft.Maps.Point(20, -260),
		})
		tooltip.setMap(map)
		const infobox = new Microsoft.Maps.Infobox(map.getCenter(), {
			visible: false,
		})
		infobox.setMap(map)

		const allSites = await this._getSiteMasterData()
		if (gen !== this._initGen || this._mapContainer !== mapEl) return
		const filters = { sector: 'all', units: [] as string[] }

		const units: string[] = []
		for (let i = 0; i < allSites.length; i++) {
			const bu = allSites[i].BusinessUnit
			if (bu && bu !== '-' && units.indexOf(bu) < 0) units.push(bu)
		}
		units.sort()
		const unitPanel = this.domElement.querySelector('#smd-dynamic-unit')
		const applyBtn = this.domElement.querySelector('#smd-apply-filter')
		if (unitPanel && applyBtn) {
			units.forEach((unit, i) => {
				const row = document.createElement('div')
				row.className = 'panel-item'
				row.innerHTML =
					'<input class="smd-unit-option" type="checkbox" id="smd-unit-' +
					i +
					'" value="' +
					self._escape(unit) +
					'"><label for="smd-unit-' +
					i +
					'">' +
					self._escape(unit) +
					'</label>'
				unitPanel.insertBefore(row, applyBtn)
			})
		}

		function createCircle(fill: string, size?: number, label?: string): string {
			const dim = size || 24
			const c = document.createElement('canvas')
			c.width = dim
			c.height = dim
			const ctx = c.getContext('2d')
			if (ctx) {
				ctx.fillStyle = fill
				ctx.lineWidth = 2
				ctx.strokeStyle = '#0f69af'
				ctx.beginPath()
				ctx.arc(c.width * 0.5, c.height * 0.5, dim * 0.5 - 2, 0, 2 * Math.PI)
				ctx.fill()
				ctx.stroke()
				if (label) {
					ctx.fillStyle = fill === '#ffc832' ? '#333' : '#fff'
					ctx.font =
						'bold ' +
						(dim >= 44 ? 16 : 13) +
						'px Segoe UI, Tahoma, sans-serif'
					ctx.textAlign = 'center'
					ctx.textBaseline = 'middle'
					ctx.fillText(label, c.width * 0.5, c.height * 0.5 + 0.5)
				}
			}
			return c.toDataURL()
		}

		function hideTooltip(): void {
			tooltip.setOptions({ visible: false })
			infobox.setOptions({ visible: false })
		}

		function bindTooltipClose(): void {
			const closeBtn = document.getElementById('closeInfoboxSmd')
			if (closeBtn) closeBtn.addEventListener('click', () => hideTooltip())
		}

		function setTooltipOffset(location: any): void {
			if (location.latitude > 60) {
				tooltip.setOptions({ offset: new Microsoft.Maps.Point(20, -400) })
			} else {
				tooltip.setOptions({ offset: new Microsoft.Maps.Point(20, -100) })
			}
		}

		function clearPushpins(): void {
			for (let i = map.entities.getLength() - 1; i >= 0; i--) {
				const entity = map.entities.get(i)
				if (entity && (entity as any).getLocation) map.entities.removeAt(i)
			}
		}

		function smoothPanTo(location: any): void {
			const animationDuration = 1000
			const frames = 60
			const interval = animationDuration / frames
			const startLocation = map.getCenter()
			const latDelta = location.latitude - startLocation.latitude
			const lonDelta = location.longitude - startLocation.longitude
			let frame = 0
			function easeInOutQuad(t: number): number {
				return t < 0.5 ? 2 * t * t : -1 + (4 - 2 * t) * t
			}
			function animate(): void {
				frame++
				const progress = frame / frames
				map.setView({
					center: new Microsoft.Maps.Location(
						startLocation.latitude + latDelta * easeInOutQuad(progress),
						startLocation.longitude + lonDelta * easeInOutQuad(progress),
					),
				})
				if (frame < frames) setTimeout(animate, interval)
			}
			animate()
		}

		function showDetailTooltip(pushpin: any, delay: number): void {
			const meta = pushpin.metadata as IFormattedSite
			if (!meta) return
			const loc = pushpin.getLocation()
			setTooltipOffset(loc)
			infobox.setOptions({ visible: false })
			const esc = (s: string) => self._escape(s)
			const icon = (v: string) => self._assessmentIcon(v)
			const css = self._sectorColor(meta.Sector).css
			const detailUrl =
				SITE_WEB_URL +
				'/Lists/Site%20Master%20Data/DispForm.aspx?ID=' +
				meta.ID
			const htmlContent = `
        <div class="map-tooltip ${css}">
          <div class="map-tooltip-header">
            <div class="map-tooltip-headline">${esc(self._siteHeadline(meta))}</div>
            <span class="map-tooltip-close" id="closeInfoboxSmd">×</span>
          </div>
          <div class="map-tooltip-items">
            <div class="map-tooltip-block">
              <div class="map-tooltip-item">${esc(meta.LegalEntityName)}</div>
              <div class="map-tooltip-item">${esc(meta.Location)}</div>
            </div>
            <div class="map-tooltip-block">
              <div class="map-tooltip-item"><img src="${icon(meta.ItAssessment)}"><strong>IT Assessment</strong></div>
              <div class="map-tooltip-item"><img src="${icon(meta.OtAssessment)}"><strong>OT Assessment</strong></div>
              <div class="map-tooltip-item"><img src="${icon(meta.DpiCoreImplementation)}"><strong>DPI Core Implementation</strong></div>
              <div class="map-tooltip-item"><img src="${icon(meta.OtSystemIntegration)}"><strong>OT System Integration</strong></div>
            </div>
          </div>
          <div class="map-tooltip-items">
            <div class="map-tooltip-block">
              <div class="map-tooltip-item"><strong>Services: </strong>${esc(meta.Services)}</div>
              <div class="map-tooltip-item"><strong>Employees: </strong>${esc(meta.Employees)}</div>
            </div>
            <div class="map-tooltip-block">
              <div class="map-tooltip-item"><strong>Current Maturity (BPOG):</strong></div>
              <div class="map-tooltip-item">${esc(meta.CurrentMaturity)}</div>
            </div>
          </div>
          <div class="map-tooltip-items">
            <div class="map-tooltip-block">
              <div class="map-tooltip-item"><strong>Target DPI Variant: </strong>${esc(meta.TargetDpi)}</div>
              <div class="map-tooltip-item"><strong>OT Systems integrated: </strong>${esc(meta.OtSystemsIntegrated)}</div>
            </div>
            <div class="map-tooltip-block">
              <div class="map-tooltip-item"><strong>Target Maturity (BPOG):</strong></div>
              <div class="map-tooltip-item">${esc(meta.TargetMaturity)}</div>
            </div>
          </div>
          <div class="map-tooltip-footer">
            <div class="tooltip-legend">
              <div><img src="${ICON_BASE}NA.svg" /> N/A</div>
              <div><img src="${ICON_BASE}Planning.svg" /> Planning</div>
              <div><img src="${ICON_BASE}Ongoing.svg" /> Ongoing</div>
              <div><img src="${ICON_BASE}Completed.svg" /> Completed</div>
            </div>
            <a class="see-all" href="${detailUrl}">See all details</a>
          </div>
        </div>`
			setTimeout(() => {
				tooltip.setOptions({ location: loc, htmlContent, visible: true })
				bindTooltipClose()
			}, delay)
		}

		function showClusterList(location: any, pins: any[], page?: number): void {
			const pageSize = 5
			const items = pins.filter((p: any) => p && p.metadata)
			if (!items.length) return
			const pageCount = Math.max(1, Math.ceil(items.length / pageSize))
			const currentPage = Math.min(Math.max(0, page || 0), pageCount - 1)
			const start = currentPage * pageSize
			const slice = items.slice(start, start + pageSize)
			const esc = (s: string) => self._escape(s)
			const itemsHtml = slice
				.map((pin: any, i: number) => {
					const meta = pin.metadata as IFormattedSite
					const fill = self._sectorColor(meta.Sector).fill
					return `<li>
            <button type="button" class="map-cluster-item" id="smd-cluster-item-${i}">
              <span class="map-cluster-dot" style="background:${fill}"></span>
              <span class="map-cluster-item-text">${esc(self._siteHeadline(meta))}</span>
            </button>
          </li>`
				})
				.join('')
			const pagerHtml =
				pageCount > 1
					? `<div class="map-cluster-pager">
            <button type="button" class="map-cluster-page-btn" id="smd-cluster-prev"${
							currentPage === 0 ? ' disabled' : ''
						}>‹</button>
            <span class="map-cluster-page-label">${currentPage + 1} / ${pageCount}</span>
            <button type="button" class="map-cluster-page-btn" id="smd-cluster-next"${
							currentPage >= pageCount - 1 ? ' disabled' : ''
						}>›</button>
          </div>`
					: ''
			tooltip.setOptions({
				location,
				htmlContent: `
        <div class="map-tooltip map-tooltip-cluster sector-mixed">
          <div class="map-tooltip-header">
            <div class="map-tooltip-headline">${items.length} sites</div>
            <span class="map-tooltip-close" id="closeInfoboxSmd">×</span>
          </div>
          <ul class="map-cluster-list${
						pageCount > 1 ? ' map-cluster-list-paged' : ''
					}">${itemsHtml}</ul>
          ${pagerHtml}
        </div>`,
				visible: true,
			})
			let attempt = 0
			const bindListClicks = (): void => {
				bindTooltipClose()
				const first = document.getElementById('smd-cluster-item-0')
				if (!first && attempt < 12) {
					attempt++
					setTimeout(bindListClicks, 50)
					return
				}
				for (let i = 0; i < slice.length; i++) {
					const el = document.getElementById('smd-cluster-item-' + i)
					if (!el) continue
					el.addEventListener('click', (ev) => {
						ev.preventDefault()
						ev.stopPropagation()
						hideTooltip()
						smoothPanTo(slice[i].getLocation())
						showDetailTooltip(slice[i], 500)
					})
				}
				const prev = document.getElementById('smd-cluster-prev')
				const next = document.getElementById('smd-cluster-next')
				if (prev) {
					prev.addEventListener('click', (ev) => {
						ev.preventDefault()
						ev.stopPropagation()
						if (currentPage > 0)
							showClusterList(location, items, currentPage - 1)
					})
				}
				if (next) {
					next.addEventListener('click', (ev) => {
						ev.preventDefault()
						ev.stopPropagation()
						if (currentPage < pageCount - 1)
							showClusterList(location, items, currentPage + 1)
					})
				}
			}
			bindListClicks()
		}

		function pushpinClicked(e: any): void {
			const pushpin = e.target
			if (pushpin && pushpin.containedPushpins) {
				showClusterList(pushpin.getLocation(), pushpin.containedPushpins)
				return
			}
			if (!pushpin || !pushpin.metadata) return
			smoothPanTo(pushpin.getLocation())
			showDetailTooltip(pushpin, 500)
		}

		function pinIconSize(count: number): number {
			if (count >= 100) return 48
			if (count >= 10) return 42
			if (count >= 2) return 36
			return 24
		}

		interface IOverlapGroup {
			pins: any[]
			loc: any
			pixel: { x: number; y: number }
			count: number
		}

		function clusterByPixelOverlap(sourcePins: any[]): IOverlapGroup[] {
			const width = mapHost.clientWidth || (map.getWidth && map.getWidth()) || 0
			const height =
				mapHost.clientHeight || (map.getHeight && map.getHeight()) || 0
			const canCluster = width >= 100 && height >= 100
			const groups: IOverlapGroup[] = []
			for (let i = 0; i < sourcePins.length; i++) {
				const loc = sourcePins[i].getLocation()
				const pixel = canCluster
					? map.tryLocationToPixel(
							loc,
							Microsoft.Maps.PixelReference.control,
						)
					: null
				const valid =
					pixel && isFinite(pixel.x) && isFinite(pixel.y)
				groups.push({
					pins: [sourcePins[i]],
					loc,
					pixel: valid
						? { x: pixel.x, y: pixel.y }
						: { x: i * 10000, y: i * 10000 },
					count: 1,
				})
			}
			if (!canCluster) return groups

			let minX = Infinity
			let maxX = -Infinity
			let minY = Infinity
			let maxY = -Infinity
			for (let i = 0; i < groups.length; i++) {
				minX = Math.min(minX, groups[i].pixel.x)
				maxX = Math.max(maxX, groups[i].pixel.x)
				minY = Math.min(minY, groups[i].pixel.y)
				maxY = Math.max(maxY, groups[i].pixel.y)
			}
			const spread = Math.max(maxX - minX, maxY - minY)
			if (groups.length >= 8 && spread < Math.min(width, height) * 0.12) {
				return groups.map((g, i) => ({
					pins: g.pins,
					loc: g.loc,
					pixel: { x: i * 10000, y: i * 10000 },
					count: 1,
				}))
			}

			let merged = true
			while (merged) {
				merged = false
				outer: for (let i = 0; i < groups.length; i++) {
					for (let j = i + 1; j < groups.length; j++) {
						const r1 = pinIconSize(groups[i].count) / 2
						const r2 = pinIconSize(groups[j].count) / 2
						const dx = groups[i].pixel.x - groups[j].pixel.x
						const dy = groups[i].pixel.y - groups[j].pixel.y
						if (dx * dx + dy * dy > (r1 + r2) * (r1 + r2)) continue
						const a = groups[i]
						const b = groups[j]
						const n = a.count + b.count
						groups[i] = {
							pins: a.pins.concat(b.pins),
							count: n,
							loc: new Microsoft.Maps.Location(
								(a.loc.latitude * a.count + b.loc.latitude * b.count) / n,
								(a.loc.longitude * a.count + b.loc.longitude * b.count) / n,
							),
							pixel: {
								x: (a.pixel.x * a.count + b.pixel.x * b.count) / n,
								y: (a.pixel.y * a.count + b.pixel.y * b.count) / n,
							},
						}
						groups.splice(j, 1)
						merged = true
						break outer
					}
				}
			}
			return groups
		}

		let sourcePins: any[] = []
		let lastClusterZoom = -1

		function renderOverlapClusters(): void {
			if (gen !== self._initGen) return
			clearPushpins()
			const groups = clusterByPixelOverlap(sourcePins)
			for (let g = 0; g < groups.length; g++) {
				const group = groups[g]
				if (group.count === 1) {
					const src = group.pins[0]
					const meta = src.metadata as IFormattedSite
					const size = pinIconSize(1)
					const pin = new Microsoft.Maps.Pushpin(src.getLocation(), {
						icon: createCircle(self._sectorColor(meta.Sector).fill, size),
						anchor: new Microsoft.Maps.Point(size / 2, size / 2),
					})
					;(pin as any).metadata = meta
					Microsoft.Maps.Events.addHandler(pin, 'click', pushpinClicked)
					map.entities.push(pin)
					continue
				}
				const size = pinIconSize(group.count)
				const pin = new Microsoft.Maps.Pushpin(group.loc, {
					icon: createCircle('#503291', size, String(group.count)),
					anchor: new Microsoft.Maps.Point(size / 2, size / 2),
				})
				;(pin as any).containedPushpins = group.pins
				Microsoft.Maps.Events.addHandler(pin, 'click', pushpinClicked)
				map.entities.push(pin)
			}
		}

		function refreshClustersIfZoomChanged(): void {
			if (gen !== self._initGen) return
			const z = map.getZoom()
			if (Math.abs(z - lastClusterZoom) < 0.01) return
			lastClusterZoom = z
			hideTooltip()
			renderOverlapClusters()
		}

		function applyFiltersAndRender(): void {
			let data = allSites
			if (filters.sector.toLowerCase() !== 'all') {
				data = data.filter((s) =>
					self._sectorsEqual(s.Sector, filters.sector),
				)
			}
			if (filters.units.length) {
				data = data.filter(
					(s) => filters.units.indexOf(s.BusinessUnit) >= 0,
				)
			}
			sourcePins = []
			const pinLocations: any[] = []
			for (let i = 0; i < data.length; i++) {
				const city = data[i]
				const coords = self._normalizeCoords(city.Lattitude, city.Longitude)
				if (!coords) {
					if (city.Lattitude != null || city.Longitude != null) {
						console.warn(
							'[MapSMD] Skip invalid coords',
							city.SiteName || city.Title,
							city.Lattitude,
							city.Longitude,
						)
					}
					continue
				}
				let loc: any
				try {
					loc = new Microsoft.Maps.Location(coords.lat, coords.lng)
				} catch (err) {
					console.warn(
						'[MapSMD] Bing rejected coords',
						city.SiteName || city.Title,
						coords,
						err,
					)
					continue
				}
				pinLocations.push(loc)
				const pin = new Microsoft.Maps.Pushpin(loc)
				;(pin as any).metadata = city
				sourcePins.push(pin)
			}
			console.log(
				'[MapSMD] Render pins',
				sourcePins.length,
				'of',
				data.length,
				'filtered /',
				allSites.length,
				'total | sector=',
				filters.sector,
			)
			self._setStatus(
				'Showing ' +
					sourcePins.length +
					' sites on the map (' +
					data.length +
					' after filters, ' +
					allSites.length +
					' loaded). Filter: ' +
					filters.sector +
					'.',
				sourcePins.length === 0,
			)
			hideTooltip()
			lastClusterZoom = -1
			const showingAll =
				filters.sector.toLowerCase() === 'all' && !filters.units.length
			if (showingAll) {
				map.setView({
					center: new Microsoft.Maps.Location(51.165691, 10.451526),
					zoom: 2,
				})
			} else if (pinLocations.length > 1) {
				try {
					const rect = Microsoft.Maps.LocationRect.fromLocations(pinLocations)
					if (rect && rect.height < 80 && rect.width < 170) {
						map.setView({ bounds: rect, padding: 80 })
					}
				} catch (err) {
					console.warn('[MapSMD] Could not fit bounds for filter', filters.sector, err)
				}
			} else if (pinLocations.length === 1) {
				map.setView({ center: pinLocations[0], zoom: 4 })
			}
			renderOverlapClusters()
			window.setTimeout(() => {
				if (gen !== self._initGen) return
				lastClusterZoom = -1
				renderOverlapClusters()
				lastClusterZoom = map.getZoom()
			}, 400)
			window.setTimeout(() => {
				if (gen !== self._initGen) return
				lastClusterZoom = -1
				renderOverlapClusters()
				lastClusterZoom = map.getZoom()
			}, 1200)
		}

		Microsoft.Maps.Events.addHandler(
			map,
			'viewchangeend',
			refreshClustersIfZoomChanged,
		)
		let firstViewRendered = false
		Microsoft.Maps.Events.addHandler(map, 'viewrendered', () => {
			if (firstViewRendered || !sourcePins.length) return
			firstViewRendered = true
			lastClusterZoom = -1
			renderOverlapClusters()
			lastClusterZoom = map.getZoom()
		})
		if (typeof ResizeObserver !== 'undefined') {
			let resizeTimer: number | undefined
			const ro = new ResizeObserver(() => {
				if (resizeTimer) window.clearTimeout(resizeTimer)
				resizeTimer = window.setTimeout(() => {
					if (gen !== self._initGen || !sourcePins.length) return
					lastClusterZoom = -1
					renderOverlapClusters()
					lastClusterZoom = map.getZoom()
				}, 200)
			})
			ro.observe(mapHost)
		}

		const sectorBtns = this.domElement.querySelectorAll('.sector-filter')
		for (let i = 0; i < sectorBtns.length; i++) {
			sectorBtns[i].addEventListener('click', (ev) => {
				const btn = ev.currentTarget as HTMLElement
				for (let j = 0; j < sectorBtns.length; j++) {
					sectorBtns[j].classList.remove('active-filter')
				}
				btn.classList.add('active-filter')
				filters.sector = (btn.getAttribute('data-id') || 'all').trim()
				applyFiltersAndRender()
			})
		}

		const toggle = this.domElement.querySelector('#smd-unit-toggle')
		const panel = this.domElement.querySelector('#smd-dynamic-unit')
		if (toggle && panel) {
			toggle.addEventListener('click', () => {
				const open = panel.classList.contains('open')
				panel.classList.toggle('open', !open)
				toggle.classList.toggle('active', !open)
			})
		}
		if (applyBtn && panel) {
			applyBtn.addEventListener('click', () => {
				const checked = panel.querySelectorAll(
					'.smd-unit-option:checked',
				) as NodeListOf<HTMLInputElement>
				filters.units = []
				for (let i = 0; i < checked.length; i++) {
					filters.units.push(checked[i].value)
				}
				panel.classList.remove('open')
				if (toggle) toggle.classList.remove('active')
				applyFiltersAndRender()
			})
		}

		applyFiltersAndRender()
	}

	private _escape(s: string): string {
		const div = document.createElement('div')
		div.textContent = s
		return div.innerHTML
	}

	protected onDispose(): void {
		this._mapContainer = null
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0')
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: { description: strings.PropertyPaneDescription },
					groups: [
						{
							groupName: strings.BasicGroupName,
							groupFields: [
								PropertyPaneTextField('description', {
									label: strings.DescriptionFieldLabel,
								}),
							],
						},
					],
				},
			],
		}
	}
}
