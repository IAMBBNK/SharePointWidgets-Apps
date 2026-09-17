import { Version } from '@microsoft/sp-core-library'
import {
	type IPropertyPaneConfiguration,
	PropertyPaneTextField,
} from '@microsoft/sp-property-pane'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'

import * as strings from 'MapPredictiveMaintenanceWebPartStrings'

const BING_MAPS_KEY =
	'AgJMis6__ur7UfvExgEQb3R3EWCwZTTAYvhjRDPtG9ad-JlxHO3WzhzLIc9DHRzw'
const SITE_WEB_URL = 'https://mdigital.sharepoint.com/sites/ITOTCommunityHub'
const LIST_TITLE = 'Predictive Maintenance'
const SITE_MASTER_LIST_TITLE = 'Site Master Data'
const LIST_PATH = '/sites/ITOTCommunityHub/Lists/Predictive Maintenance'
const SITE_MASTER_LIST_PATH = '/sites/ITOTCommunityHub/Lists/Site Master Data'
const MAP_ELEMENT_ID = 'map-itot-pm-map'

export interface IMapPredictiveMaintenanceWebPartProps {
	description: string
}

interface IListField {
	InternalName: string
	Title: string
	TypeAsString: string
	Hidden?: boolean
	ReadOnlyField?: boolean
	FromBaseType?: boolean
}

interface ITooltipFieldSpec {
	key: keyof ITooltipValues
	label: string
	aliases: string[]
	valueOnNextLine?: boolean
}

interface ITooltipValues {
	production: string
	technicalContact: string
	systemSize: string
	startOfOperation: string
	technicalSpecials: string
	phase: string
}

interface IFormattedItem extends ITooltipValues {
	ID: number
	siteName: string
	businessUnit: string
	legalEntityName: string
	location: string
	services: string
	employees: string
	sector: string
	Lattitude?: number
	Longitude?: number
}

interface ISmdSite {
	lat: number
	lng: number
	siteName: string
	businessUnit: string
	legalEntityName: string
	location: string
	services: string
	employees: string
}

interface ISmdIndex {
	byName: Record<string, ISmdSite>
	byId: Record<string, ISmdSite>
	bySiteId: Record<string, ISmdSite>
}

const PHASE_STYLES: Record<string, { fill: string; label: string }> = {
	design: { fill: '#5b6abf', label: 'Design' },
	installation: { fill: '#e87722', label: 'Installation' },
	training: { fill: '#1aa6c1', label: 'Training' },
	operation: { fill: '#2e8b57', label: 'Operation' },
}

const SECTOR_COLORS: Record<string, { fill: string; css: string }> = {
	healthcare: { fill: '#eb3c96', css: 'sector-pink' },
	'health care': { fill: '#eb3c96', css: 'sector-pink' },
	'life science': { fill: '#ffc832', css: 'sector-yellow' },
	electronics: { fill: '#4ca5e9', css: 'sector-blue' },
}

const TOOLTIP_FIELDS: ITooltipFieldSpec[] = [
	{ key: 'production', label: 'Production', aliases: ['production'] },
	{
		key: 'technicalContact',
		label: 'Technical contact person',
		aliases: [
			'technical contact person',
			'technical contact',
			'pm technical contact',
			'contact person',
			'tech contact',
			'technicalcontact',
		],
		valueOnNextLine: true,
	},
	{ key: 'systemSize', label: 'System size', aliases: ['system size'] },
	{
		key: 'startOfOperation',
		label: 'Start of Operation',
		aliases: ['start of operation', 'start ofoperation'],
	},
	{
		key: 'technicalSpecials',
		label: 'Technical specials',
		aliases: ['technical specials'],
	},
	{ key: 'phase', label: 'Phase', aliases: ['phase', 'pm phase'] },
]

declare global {
	interface Window {
		Microsoft?: any
		initMapItotPm?: () => void
	}
}

export default class MapPredictiveMaintenanceWebPart extends BaseClientSideWebPart<IMapPredictiveMaintenanceWebPartProps> {
	private _mapContainer: HTMLDivElement | null = null

	public render(): void {
		this.domElement.innerHTML = ''
		const container = document.createElement('div')
		container.className = 'map-itot-container'
		this._mapContainer = document.createElement('div')
		this._mapContainer.id = MAP_ELEMENT_ID
		this._mapContainer.style.height = '700px'
		this._mapContainer.style.width = '100%'
		container.appendChild(this._mapContainer)
		this._injectStyles()
		this.domElement.appendChild(container)
		this._loadBingAndInit()
	}

	private _injectStyles(): void {
		const style = document.createElement('style')
		style.textContent = `
      .map-itot-container { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
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
      .map-tooltip-header {
        position: relative;
        padding-bottom: 15px;
      }
      .map-tooltip-headline { font-size: 18px; font-weight: 600; margin: 0; }
      .map-tooltip-close {
        position: absolute; top: -10px; right: 0px;
        font-size: 25px; font-weight: 600; cursor: pointer; line-height: 1; color: #666;
      }
      .map-tooltip-body { display: flex; gap: 24px; }
      .map-tooltip-col { flex: 1; min-width: 0; }
      .map-tooltip-col .map-tooltip-line { margin-bottom: 8px; font-size: 14px; }
      .map-tooltip-col .map-tooltip-line strong { display: inline-block; }
      .map-tooltip-footer { padding: 8px 0px; text-align: end;}
      .map-tooltip .see-all {
        cursor: pointer;
        padding: 8px 15px;
        background-color: #0f69af;
        color: white;
        text-decoration: none;
        border-radius: 10px;
        font-size: 12px;
        font-weight: 500;
      }
      .map-tooltip.sector-mixed { border-top: 15px solid #503291; }
      .map-tooltip.map-tooltip-cluster { width: 360px; }
      .map-cluster-list {
        list-style: none;
        margin: 0;
        padding: 0;
      }
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
      .map-phase-bubble {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 999px;
        color: #fff;
        font-size: 12px;
        font-weight: 600;
        line-height: 1.3;
        white-space: nowrap;
      }
      .map-phase-bubble-sm {
        padding: 2px 8px;
        font-size: 11px;
      }
      .map-tooltip-header .map-phase-bubble { margin-top: 8px; }
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
      .map-cluster-page-btn:disabled {
        opacity: 0.35;
        cursor: default;
      }
      .map-cluster-page-label {
        font-size: 13px;
        color: #555;
        font-weight: 600;
      }
    `
		this.domElement.appendChild(style)
	}

	private _hideBingCredentialsError(): void {
		const credsText = 'The specified credentials are invalid'
		const root = document.getElementById(MAP_ELEMENT_ID) || this.domElement
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

	private _fetchOpts(): RequestInit {
		return {
			method: 'GET',
			headers: { accept: 'application/json;odata=verbose' },
			credentials: 'include',
		}
	}

	private _listUrls(path: string, listTitle: string, listPath: string): string[] {
		const titleEnc = encodeURIComponent(listTitle)
		const byTitle =
			SITE_WEB_URL +
			'/_api/web/lists/getbytitle(%27' +
			titleEnc +
			'%27)' +
			path
		const byPath =
			SITE_WEB_URL +
			"/_api/web/GetList('" +
			listPath.replace(/'/g, "''") +
			"')" +
			path
		return [byTitle, byPath]
	}

	private async _getJson(urls: string[]): Promise<any | undefined> {
		for (let i = 0; i < urls.length; i++) {
			try {
				const res = await fetch(urls[i], this._fetchOpts())
				if (res.ok) return await res.json()
				console.warn('[MapPM] Request failed', res.status, urls[i])
			} catch (err) {
				console.warn('[MapPM] Request error', urls[i], err)
			}
		}
		return undefined
	}

	private async _getListFields(): Promise<IListField[]> {
		const json = await this._getJson(
			this._listUrls(
				'/fields?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField,FromBaseType',
				LIST_TITLE,
				LIST_PATH,
			),
		)
		if (!json) {
			console.warn('[MapPM] Fields API failed')
			return []
		}
		return (json.d && json.d.results) || []
	}

	private _normName(s: string): string {
		return (s || '')
			.toLowerCase()
			.replace(/_x0020_/g, ' ')
			.replace(/[^a-z0-9]/g, '')
	}

	private _normSite(s: string): string {
		return (s || '')
			.toLowerCase()
			.replace(/[_.,;:/\\-]+/g, ' ')
			.replace(/\s+/g, ' ')
			.trim()
	}

	private _normSiteId(s: string): string {
		return (s || '').toUpperCase().replace(/\s+/g, '').trim()
	}

	private _sectorColor(sector: string): { fill: string; css: string } {
		const key = (sector || '').toLowerCase().trim()
		if (SECTOR_COLORS[key]) return SECTOR_COLORS[key]
		const names = Object.keys(SECTOR_COLORS)
		for (let i = 0; i < names.length; i++) {
			if (key.indexOf(names[i]) >= 0) return SECTOR_COLORS[names[i]]
		}
		return SECTOR_COLORS.healthcare
	}

	private _phaseStyle(
		phase: string,
	): { fill: string; label: string } | undefined {
		const key = (phase || '').toLowerCase().trim()
		if (!key || key === '-') return undefined
		if (PHASE_STYLES[key]) return PHASE_STYLES[key]
		const names = Object.keys(PHASE_STYLES)
		for (let i = 0; i < names.length; i++) {
			if (key.indexOf(names[i]) >= 0) return PHASE_STYLES[names[i]]
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
				const alias = aliases[a]
				if (title === alias || this._normName(alias) === internal) {
					hits.push(f)
					break
				}
			}
		}
		if (!hits.length) return undefined
		const prefer = (type: string): IListField | undefined =>
			this._find(hits, (f) => f.TypeAsString === type && !f.Hidden)
		return (
			prefer('User') ||
			prefer('UserMulti') ||
			prefer('Lookup') ||
			this._find(hits, (f) => !f.Hidden) ||
			hits[0]
		)
	}

	private _isPersonField(field?: IListField): boolean {
		if (!field) return false
		return (
			field.TypeAsString === 'User' ||
			field.TypeAsString === 'UserMulti' ||
			field.TypeAsString === 'Lookup'
		)
	}

	private _userIdFromValue(raw: unknown): number | undefined {
		if (typeof raw === 'number' && raw > 0) return raw
		if (typeof raw === 'string' && /^\d+$/.test(raw.trim())) {
			const n = parseInt(raw.trim(), 10)
			return n > 0 ? n : undefined
		}
		if (raw && typeof raw === 'object') {
			const obj = raw as { Id?: number; ID?: number }
			if (typeof obj.Id === 'number' && obj.Id > 0) return obj.Id
			if (typeof obj.ID === 'number' && obj.ID > 0) return obj.ID
		}
		return undefined
	}

	private _extractPerson(
		item: Record<string, unknown>,
		field?: IListField,
	): { display: string; userId?: number } {
		if (!field) return { display: '-' }
		const name = field.InternalName
		const raw = item[name]
		const display = this._fieldValue(raw, field.TypeAsString)
		const userId =
			this._userIdFromValue(item[name + 'Id']) || this._userIdFromValue(raw)
		if (display !== '-' && !/^\d+$/.test(display)) return { display, userId }
		return { display: '-', userId }
	}

	private async _resolveUserTitles(
		ids: number[],
	): Promise<Record<number, string>> {
		const out: Record<number, string> = {}
		const unique: number[] = []
		for (let i = 0; i < ids.length; i++) {
			if (ids[i] > 0 && unique.indexOf(ids[i]) < 0) unique.push(ids[i])
		}
		if (!unique.length) return out
		const filter = unique.map((id) => 'Id eq ' + id).join(' or ')
		const json = await this._getJson([
			SITE_WEB_URL +
				'/_api/web/siteusers?$select=Id,Title,Email&$filter=' +
				encodeURIComponent(filter),
		])
		const results: { Id?: number; Title?: string }[] =
			(json && json.d && json.d.results) || []
		for (let i = 0; i < results.length; i++) {
			if (results[i].Id && results[i].Title)
				out[results[i].Id as number] = String(results[i].Title)
		}
		for (let i = 0; i < unique.length; i++) {
			if (out[unique[i]]) continue
			const ujson = await this._getJson([
				SITE_WEB_URL + '/_api/web/getuserbyid(' + unique[i] + ')',
			])
			const u = ujson && ujson.d
			if (u && u.Title) out[unique[i]] = String(u.Title)
		}
		return out
	}

	private _formatDate(raw: string): string | undefined {
		const msMatch = raw.match(/\/Date\((-?\d+)\)\//)
		const d = msMatch
			? new Date(parseInt(msMatch[1], 10))
			: new Date(raw)
		if (isNaN(d.getTime())) return undefined
		const dd = d.getDate() < 10 ? '0' + d.getDate() : String(d.getDate())
		const mm =
			d.getMonth() + 1 < 10
				? '0' + (d.getMonth() + 1)
				: String(d.getMonth() + 1)
		return dd + '.' + mm + '.' + d.getFullYear()
	}

	private _fieldValue(val: unknown, typeAsString?: string): string {
		if (val == null || val === '') return '-'
		if (typeof val === 'string') {
			const t = val.trim()
			if (t === '') return '-'
			if (typeAsString === 'DateTime' || t.indexOf('/Date(') === 0) {
				const formatted = this._formatDate(t)
				if (formatted) return formatted
			}
			return t
		}
		if (typeof val === 'number' || typeof val === 'boolean') return String(val)
		if (typeof val === 'object') {
			const obj = val as {
				Title?: string
				Results?: unknown[]
				results?: unknown[]
				Description?: string
				Url?: string
				Label?: string
				EMail?: string
				Email?: string
			}
			if (obj.Title && String(obj.Title).trim()) return String(obj.Title).trim()
			if (obj.Label && String(obj.Label).trim()) return String(obj.Label).trim()
			if (obj.EMail && String(obj.EMail).trim()) return String(obj.EMail).trim()
			if (obj.Email && String(obj.Email).trim()) return String(obj.Email).trim()
			if (obj.Description && String(obj.Description).trim())
				return String(obj.Description).trim()
			if (obj.Url && String(obj.Url).trim()) return String(obj.Url).trim()
			const nested = obj.Results || obj.results
			if (Array.isArray(nested) && nested.length) {
				return (
					nested
						.map((r) => this._fieldValue(r))
						.filter((s) => s !== '-')
						.join(', ') || '-'
				)
			}
		}
		return '-'
	}

	private _find<T>(arr: T[], pred: (item: T) => boolean): T | undefined {
		for (let i = 0; i < arr.length; i++) {
			if (pred(arr[i])) return arr[i]
		}
		return undefined
	}

	private _parseCoord(raw: unknown): number | undefined {
		if (raw == null || raw === '') return undefined
		if (typeof raw === 'object') {
			const geo = raw as { Latitude?: number; Longitude?: number }
			if (typeof geo.Latitude === 'number') return geo.Latitude
		}
		const n = parseFloat(String(raw))
		return typeof n === 'number' && !isNaN(n) ? n : undefined
	}

	private _lookupId(item: Record<string, unknown>, internalName: string): string {
		const idVal = item[internalName + 'Id']
		if (typeof idVal === 'number' && idVal > 0) return String(idVal)
		if (typeof idVal === 'string' && idVal) return idVal
		const obj = item[internalName] as { Id?: number; ID?: number } | undefined
		if (obj && typeof obj === 'object') {
			if (typeof obj.Id === 'number' && obj.Id > 0) return String(obj.Id)
			if (typeof obj.ID === 'number' && obj.ID > 0) return String(obj.ID)
		}
		return ''
	}

	private _resolveSite(
		item: Record<string, unknown>,
		siteNameField: IListField | undefined,
		siteIdField: IListField | undefined,
		smd: ISmdIndex,
	): ISmdSite | undefined {
		if (siteIdField) {
			const siteId = this._normSiteId(
				this._fieldValue(item[siteIdField.InternalName], siteIdField.TypeAsString),
			)
			if (siteId && smd.bySiteId[siteId]) return smd.bySiteId[siteId]
		}
		const rawSiteId = this._normSiteId(
			this._fieldValue(item.SiteID || item.SiteId || item.Site_x0020_ID),
		)
		if (rawSiteId && smd.bySiteId[rawSiteId]) return smd.bySiteId[rawSiteId]

		const candidates: string[] = []
		const add = (raw: unknown, type?: string): void => {
			const text = this._fieldValue(raw, type)
			if (text && text !== '-') candidates.push(this._normSite(text))
		}

		if (siteNameField) {
			const byId = this._lookupId(item, siteNameField.InternalName)
			if (byId && smd.byId[byId]) return smd.byId[byId]
			add(item[siteNameField.InternalName], siteNameField.TypeAsString)
		}
		add(item.field_2)
		add(item.Title)
		add(item.SiteName)
		add(item.Site_x0020_Name)

		const idFallbacks = [
			this._lookupId(item, 'field_2'),
			this._lookupId(item, 'SiteName'),
			this._lookupId(item, 'Site_x0020_Name'),
			this._lookupId(item, 'Site'),
		]
		for (let i = 0; i < idFallbacks.length; i++) {
			if (idFallbacks[i] && smd.byId[idFallbacks[i]]) {
				return smd.byId[idFallbacks[i]]
			}
		}

		for (let i = 0; i < candidates.length; i++) {
			if (smd.byName[candidates[i]]) return smd.byName[candidates[i]]
		}

		const smdNames = Object.keys(smd.byName)
		for (let i = 0; i < candidates.length; i++) {
			const cand = candidates[i]
			const first = cand.split(' ')[0]
			if (first && smd.byName[first]) return smd.byName[first]
			for (let j = 0; j < smdNames.length; j++) {
				const n = smdNames[j]
				if (cand.indexOf(n) === 0 || n.indexOf(cand) === 0) {
					return smd.byName[n]
				}
			}
		}

		const keys = Object.keys(item)
		for (let i = 0; i < keys.length; i++) {
			const key = keys[i]
			if (key === '__metadata' || key.lastIndexOf('Id') === key.length - 2)
				continue
			add(item[key])
		}
		for (let i = 0; i < candidates.length; i++) {
			if (smd.byName[candidates[i]]) return smd.byName[candidates[i]]
		}
		return undefined
	}

	private async _getSiteMasterCoordMap(): Promise<ISmdIndex> {
		const out: ISmdIndex = { byName: {}, byId: {}, bySiteId: {} }
		const withSelect = this._listUrls(
			'/items?$top=1000&$select=ID,Title,field_2,field_5,field_6,field_7,field_8,field_9,LegalEntityNumber,Lattitude,Longitude',
			SITE_MASTER_LIST_TITLE,
			SITE_MASTER_LIST_PATH,
		)
		const withoutSelect = this._listUrls(
			'/items?$top=1000',
			SITE_MASTER_LIST_TITLE,
			SITE_MASTER_LIST_PATH,
		)
		const json =
			(await this._getJson(withSelect)) || (await this._getJson(withoutSelect))
		if (!json) {
			console.warn('[MapPM] Site Master Data fetch failed')
			return out
		}
		const results: Record<string, unknown>[] = (json.d && json.d.results) || []
		results.forEach((item) => {
			const lat = this._parseCoord(item.Lattitude ?? item.Latitude)
			const lng = this._parseCoord(item.Longitude)
			if (lat == null || lng == null || lat === 0 || lng === 0) return
			const siteName = this._fieldValue(item.field_2)
			const site: ISmdSite = {
				lat,
				lng,
				siteName: siteName !== '-' ? siteName : this._fieldValue(item.Title),
				businessUnit: this._fieldValue(item.field_5),
				legalEntityName: this._fieldValue(item.field_7),
				location: this._fieldValue(item.field_8),
				services: this._fieldValue(item.field_9),
				employees: this._fieldValue(item.field_6),
			}
			const id = item.ID != null ? String(item.ID) : ''
			if (id) out.byId[id] = site
			const names = [item.field_2, item.Title]
			for (let i = 0; i < names.length; i++) {
				const key = this._normSite(this._fieldValue(names[i]))
				if (key && key !== '-') out.byName[key] = site
			}
			const siteIds = [
				item.LegalEntityNumber,
				item.SiteID,
				item.SiteId,
				item.Site_x0020_ID,
				item.Title,
			]
			for (let i = 0; i < siteIds.length; i++) {
				const sid = this._normSiteId(this._fieldValue(siteIds[i]))
				if (sid && sid !== '-') out.bySiteId[sid] = site
			}
		})
		console.log(
			'[MapPM] Site Master Data coords:',
			Object.keys(out.byName).length,
			'names /',
			Object.keys(out.byId).length,
			'ids from',
			results.length,
			'items',
		)
		return out
	}

	private async _getPredictiveMaintenanceData(): Promise<IFormattedItem[]> {
		const fields = await this._getListFields()
		const resolved: { spec: ITooltipFieldSpec; field?: IListField }[] = []
		for (let i = 0; i < TOOLTIP_FIELDS.length; i++) {
			resolved.push({
				spec: TOOLTIP_FIELDS[i],
				field: this._matchField(fields, TOOLTIP_FIELDS[i].aliases),
			})
		}
		for (let i = 0; i < resolved.length; i++) {
			if (resolved[i].spec.key !== 'technicalContact' || resolved[i].field)
				continue
			resolved[i].field = this._find(fields, (f) => {
				const blob =
					(f.Title || '').toLowerCase() +
					' ' +
					this._normName(f.InternalName)
				return blob.indexOf('technical') >= 0 && blob.indexOf('contact') >= 0
			})
		}
		const siteNameField =
			this._matchField(fields, [
				'site name',
				'sitename',
				'site',
				'field_2',
			]) || this._find(fields, (f) => f.InternalName === 'field_2')
		const siteIdField = this._matchField(fields, [
			'site id',
			'siteid',
			'site code',
		])
		const addressField = this._matchField(fields, ['address', 'location'])
		const servicesField = this._matchField(fields, ['services'])
		const employeesField = this._matchField(fields, ['employees'])
		const sectorField = this._matchField(fields, ['sector'])

		const contactField = this._find(
			resolved,
			(r) => r.spec.key === 'technicalContact',
		)?.field

		console.log(
			'[MapPM] Site Name field:',
			siteNameField
				? siteNameField.InternalName +
					' / ' +
					siteNameField.Title +
					' / ' +
					siteNameField.TypeAsString
				: 'not found, will scan item values',
		)
		console.log(
			'[MapPM] Tooltip fields:',
			resolved.map((r) => ({
				label: r.spec.label,
				internal: r.field ? r.field.InternalName : '(not found)',
				type: r.field ? r.field.TypeAsString : '',
			})),
		)

		const personSelect: string[] = []
		const personExpand: string[] = []
		resolved.forEach((r) => {
			if (!this._isPersonField(r.field) || !r.field) return
			const n = r.field.InternalName
			if (personExpand.indexOf(n) < 0) personExpand.push(n)
			personSelect.push(n + '/Id', n + '/Title', n + '/EMail')
		})

		const queries: string[] = []
		if (contactField && this._isPersonField(contactField)) {
			const n = contactField.InternalName
			queries.push(
				'/items?$top=1000&$select=ID,' +
					n +
					'/Id,' +
					n +
					'/Title,' +
					n +
					'/EMail&$expand=' +
					n,
			)
			queries.push(
				'/items?$top=1000&$select=ID,' + n + '/Id,' + n + '/Title&$expand=' + n,
			)
		}
		if (personExpand.length && personSelect.length) {
			queries.push(
				'/items?$top=1000&$select=*,' +
					personSelect.join(',') +
					'&$expand=' +
					personExpand.join(','),
			)
			queries.push('/items?$top=1000&$expand=' + personExpand.join(','))
		}
		queries.push('/items?$top=1000')

		const fetched = await Promise.all([
			this._getJson(this._listUrls(queries[0], LIST_TITLE, LIST_PATH)),
			this._getSiteMasterCoordMap(),
		])
		let json = fetched[0]
		const smd = fetched[1]
		for (let q = 1; !json && q < queries.length; q++) {
			json = await this._getJson(
				this._listUrls(queries[q], LIST_TITLE, LIST_PATH),
			)
		}
		if (!json) {
			console.error('[MapPM] Predictive Maintenance items fetch failed')
			return []
		}

		let results: Record<string, unknown>[] = (json.d && json.d.results) || []
		const contactById: Record<number, { display: string; userId?: number }> = {}
		if (contactField && results.length && results[0] && results[0].ID != null) {
			const onlyContact =
				Object.keys(results[0]).filter(
					(k) => k !== '__metadata' && k !== 'ID' && k !== 'Id',
				).length <= 2
			if (onlyContact) {
				for (let i = 0; i < results.length; i++) {
					const id = Number(results[i].ID)
					if (!id) continue
					contactById[id] = this._extractPerson(results[i], contactField)
				}
				const fullJson = await this._getJson(
					this._listUrls('/items?$top=1000', LIST_TITLE, LIST_PATH),
				)
				if (fullJson && fullJson.d && fullJson.d.results) {
					results = fullJson.d.results
				}
			}
		}
		console.log(
			'[MapPM] PM items:',
			results.length,
			'| first keys:',
			results[0] ? Object.keys(results[0]) : [],
		)

		const prefer = (
			field: IListField | undefined,
			item: Record<string, unknown>,
			fallback: string,
		): string => {
			if (!field) return fallback
			const val = this._fieldValue(item[field.InternalName], field.TypeAsString)
			return val !== '-' ? val : fallback
		}

		const pendingUserIds: { index: number; userId: number }[] = []
		const mapped = results.map((item, index) => {
			const site = this._resolveSite(item, siteNameField, siteIdField, smd)
			const values: ITooltipValues = {
				production: '-',
				technicalContact: '-',
				systemSize: '-',
				startOfOperation: '-',
				technicalSpecials: '-',
				phase: '-',
			}
			resolved.forEach((r) => {
				if (!r.field) return
				if (r.spec.key === 'technicalContact') {
					const fromItem = this._extractPerson(item, r.field)
					const fromExpand = contactById[Number(item.ID)]
					values.technicalContact =
						fromItem.display !== '-'
							? fromItem.display
							: fromExpand && fromExpand.display !== '-'
								? fromExpand.display
								: '-'
					const userId = fromItem.userId || (fromExpand && fromExpand.userId)
					if (values.technicalContact === '-' && userId) {
						pendingUserIds.push({ index, userId })
					}
					return
				}
				values[r.spec.key] = this._fieldValue(
					item[r.field.InternalName],
					r.field.TypeAsString,
				)
			})
			if (values.technicalContact === '-') {
				const keys = Object.keys(item)
				for (let k = 0; k < keys.length; k++) {
					const key = keys[k]
					const lower = key.toLowerCase()
					if (lower.lastIndexOf('id') === lower.length - 2) continue
					if (lower.indexOf('contact') < 0) continue
					const scanned = this._fieldValue(item[key])
					if (scanned !== '-' && !/^\d+$/.test(scanned)) {
						values.technicalContact = scanned
						break
					}
					const scannedId = this._userIdFromValue(item[key + 'Id'] || item[key])
					if (scannedId) pendingUserIds.push({ index, userId: scannedId })
				}
			}
			const pmSiteName = prefer(siteNameField, item, '')
			const smdName = site ? site.siteName : '-'

			return {
				ID: Number(item.ID),
				siteName: pmSiteName || smdName,
				businessUnit: site ? site.businessUnit : '-',
				legalEntityName: site ? site.legalEntityName : '-',
				location: prefer(addressField, item, site ? site.location : '-'),
				services: prefer(servicesField, item, site ? site.services : '-'),
				employees: prefer(employeesField, item, site ? site.employees : '-'),
				sector: prefer(sectorField, item, '-'),
				production: values.production,
				technicalContact: values.technicalContact,
				systemSize: values.systemSize,
				startOfOperation: values.startOfOperation,
				technicalSpecials: values.technicalSpecials,
				phase: values.phase,
				Lattitude: site ? site.lat : undefined,
				Longitude: site ? site.lng : undefined,
			}
		})

		if (pendingUserIds.length) {
			const titles = await this._resolveUserTitles(
				pendingUserIds.map((p) => p.userId),
			)
			for (let i = 0; i < pendingUserIds.length; i++) {
				const row = mapped[pendingUserIds[i].index]
				if (!row || row.technicalContact !== '-') continue
				const title = titles[pendingUserIds[i].userId]
				if (title) row.technicalContact = title
			}
		}
		const matched = mapped.filter(
			(m) => m.Lattitude != null && m.Longitude != null,
		).length
		console.log('[MapPM] Matched coords:', matched, '/', mapped.length)
		return mapped
	}

	private _siteHeadline(meta: IFormattedItem): string {
		const name = meta.siteName && meta.siteName !== '-' ? meta.siteName : ''
		const bu =
			meta.businessUnit && meta.businessUnit !== '-' ? meta.businessUnit : ''
		if (name && bu) return name + ' / ' + bu
		return name || bu || 'Site'
	}

	private _clusterFill(_pins: any[]): { fill: string; css: string } {
		return { fill: '#503291', css: 'sector-mixed' }
	}

	private _loadBingAndInit(): void {
		const self = this
		if (window.Microsoft && window.Microsoft.Maps) {
			self._initMap()
			return
		}
		window.initMapItotPm = function () {
			self._initMap()
		}
		const script = document.createElement('script')
		script.type = 'text/javascript'
		script.src = `https://www.bing.com/api/maps/mapcontrol?key=${BING_MAPS_KEY}&callback=initMapItotPm`
		script.async = true
		script.defer = true
		document.head.appendChild(script)
	}

	private async _initMap(): Promise<void> {
		const self = this
		const mapEl = document.getElementById(MAP_ELEMENT_ID)
		if (!mapEl || typeof window.Microsoft === 'undefined') return

		const Microsoft = window.Microsoft
		const hiddenCityLabels = {
			version: '1.0',
			elements: {
				point: { labelVisible: true, visible: false },
				road: { visible: false, labelVisible: false },
			},
		}

		const map = new Microsoft.Maps.Map(mapEl, {
			credentials: BING_MAPS_KEY,
			center: new Microsoft.Maps.Location(51.165691, 10.451526),
			zoom: 2,
			disableScrollWheelZoom: false,
			mapTypeId: Microsoft.Maps.MapTypeId.aerial,
			customMapStyle: hiddenCityLabels,
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

		const data = await this._getPredictiveMaintenanceData()
		const withCoords = data.filter(
			(c) =>
				c.Lattitude != null &&
				c.Longitude != null &&
				c.Lattitude !== 0 &&
				c.Longitude !== 0,
		)
		if (withCoords.length === 0 && typeof console !== 'undefined') {
			console.warn(
				'[MapPM] No points to display. PM items:',
				data.length,
				', with coords from Site Master Data: 0. Check Site Name matches Site Master Data (field_2).',
			)
		}

		const sectorColor = (sector: string) => self._sectorColor(sector)

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
			const closeBtn = document.getElementById('closeInfoboxPm')
			if (closeBtn) {
				closeBtn.addEventListener('click', () => hideTooltip())
			}
		}

		function setTooltipOffset(location: any, tall?: boolean): void {
			if (location.latitude > 60) {
				tooltip.setOptions({ offset: new Microsoft.Maps.Point(20, -400) })
			} else if (tall) {
				tooltip.setOptions({ offset: new Microsoft.Maps.Point(20, -220) })
			} else {
				tooltip.setOptions({ offset: new Microsoft.Maps.Point(20, -100) })
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
				const currentLatitude =
					startLocation.latitude + latDelta * easeInOutQuad(progress)
				const currentLongitude =
					startLocation.longitude + lonDelta * easeInOutQuad(progress)
				map.setView({
					center: new Microsoft.Maps.Location(
						currentLatitude,
						currentLongitude,
					),
				})
				if (frame < frames) setTimeout(animate, interval)
			}
			animate()
		}

		function showDetailTooltip(pushpin: any, delay: number): void {
			const meta = pushpin.metadata as IFormattedItem
			if (!meta) return
			const pushpinLocation = pushpin.getLocation()
			setTooltipOffset(pushpinLocation)
			infobox.setOptions({ visible: false })
			const esc = (s: string) => self._escape(s)
			const addressHtml = esc(meta.location).replace(/\n/g, '<br>')
			const sectorCss = sectorColor(meta.sector).css
			const phase = self._phaseStyle(meta.phase)
			const phaseHtml = phase
				? `<span class="map-phase-bubble" style="background:${phase.fill}">${esc(
						phase.label,
					)}</span>`
				: ''
			const htmlContent = `
        <div class="map-tooltip ${sectorCss}">
          <div class="map-tooltip-header">
            <div class="map-tooltip-headline">${esc(self._siteHeadline(meta))}</div>
            <span class="map-tooltip-close" id="closeInfoboxPm">×</span>
            ${phaseHtml}
          </div>
          <div class="map-tooltip-body">
            <div class="map-tooltip-col">
              <div class="map-tooltip-line">${addressHtml}</div>
              <div class="map-tooltip-line"><strong>Services:</strong> ${esc(meta.services)}</div>
              <div class="map-tooltip-line"><strong>Employees:</strong> ${esc(meta.employees)}</div>
              <div class="map-tooltip-line"><strong>Sector:</strong> ${esc(meta.sector)}</div>
            </div>
            <div class="map-tooltip-col">
              <div class="map-tooltip-line"><strong>Production:</strong> ${esc(meta.production)}</div>
              <div class="map-tooltip-line"><strong>Technical contact person:</strong></div>
              <div class="map-tooltip-line">${esc(meta.technicalContact)}</div>
              <div class="map-tooltip-line"><strong>System size:</strong> ${esc(meta.systemSize)}</div>
              <div class="map-tooltip-line"><strong>Start of Operation:</strong> ${esc(meta.startOfOperation)}</div>
              <div class="map-tooltip-line"><strong>Technical specials:</strong> ${esc(meta.technicalSpecials)}</div>
            </div>
          </div>
        </div>`
			setTimeout(() => {
				tooltip.setOptions({
					location: pushpinLocation,
					htmlContent,
					visible: true,
				})
				bindTooltipClose()
			}, delay)
		}

		function showClusterList(
			location: any,
			pins: any[],
			page?: number,
		): void {
			const pageSize = 5
			const items = pins.filter((p: any) => p && p.metadata)
			if (!items.length) return
			const pageCount = Math.max(1, Math.ceil(items.length / pageSize))
			const currentPage = Math.min(
				Math.max(0, page || 0),
				pageCount - 1,
			)
			const start = currentPage * pageSize
			const slice = items.slice(start, start + pageSize)
			const esc = (s: string) => self._escape(s)
			const color = self._clusterFill(items)
			const itemsHtml = slice
				.map((pin: any, i: number) => {
					const meta = pin.metadata as IFormattedItem
					const fill = sectorColor(meta.sector).fill
					const phase = self._phaseStyle(meta.phase)
					const phaseHtml = phase
						? `<span class="map-phase-bubble map-phase-bubble-sm" style="background:${phase.fill}">${esc(
								phase.label,
							)}</span>`
						: ''
					return `<li>
            <button type="button" class="map-cluster-item" id="pm-cluster-item-${i}">
              <span class="map-cluster-dot" style="background:${fill}"></span>
              <span class="map-cluster-item-text">${esc(self._siteHeadline(meta))}</span>
              ${phaseHtml}
            </button>
          </li>`
				})
				.join('')
			const pagerHtml =
				pageCount > 1
					? `<div class="map-cluster-pager">
            <button type="button" class="map-cluster-page-btn" id="pm-cluster-prev"${
							currentPage === 0 ? ' disabled' : ''
						}>‹</button>
            <span class="map-cluster-page-label">${currentPage + 1} / ${pageCount}</span>
            <button type="button" class="map-cluster-page-btn" id="pm-cluster-next"${
							currentPage >= pageCount - 1 ? ' disabled' : ''
						}>›</button>
          </div>`
					: ''
			const htmlContent = `
        <div class="map-tooltip map-tooltip-cluster ${color.css}">
          <div class="map-tooltip-header">
            <div class="map-tooltip-headline">${items.length} sites</div>
            <span class="map-tooltip-close" id="closeInfoboxPm">×</span>
          </div>
          <ul class="map-cluster-list${
						pageCount > 1 ? ' map-cluster-list-paged' : ''
					}">${itemsHtml}</ul>
          ${pagerHtml}
        </div>`
			setTooltipOffset(location, false)
			tooltip.setOptions({
				location,
				htmlContent,
				visible: true,
			})
			let attempt = 0
			const bindListClicks = (): void => {
				bindTooltipClose()
				const first = document.getElementById('pm-cluster-item-0')
				if (!first && attempt < 12) {
					attempt++
					setTimeout(bindListClicks, 50)
					return
				}
				for (let i = 0; i < slice.length; i++) {
					const el = document.getElementById('pm-cluster-item-' + i)
					if (!el) continue
					el.addEventListener('click', (ev) => {
						ev.preventDefault()
						ev.stopPropagation()
						hideTooltip()
						smoothPanTo(slice[i].getLocation())
						showDetailTooltip(slice[i], 500)
					})
				}
				const prev = document.getElementById('pm-cluster-prev')
				const next = document.getElementById('pm-cluster-next')
				if (prev) {
					prev.addEventListener('click', (ev) => {
						ev.preventDefault()
						ev.stopPropagation()
						if (currentPage > 0) {
							showClusterList(location, items, currentPage - 1)
						}
					})
				}
				if (next) {
					next.addEventListener('click', (ev) => {
						ev.preventDefault()
						ev.stopPropagation()
						if (currentPage < pageCount - 1) {
							showClusterList(location, items, currentPage + 1)
						}
					})
				}
			}
			bindListClicks()
		}

		function handleClusterClick(cluster: any): void {
			const pins = cluster.containedPushpins || []
			if (!pins.length) return
			showClusterList(cluster.getLocation(), pins)
		}

		function pushpinClicked(e: any): void {
			const pushpin = e.target
			if (pushpin && pushpin.containedPushpins) {
				handleClusterClick(pushpin)
				return
			}
			const meta = pushpin && (pushpin.metadata as IFormattedItem)
			if (!meta) return
			smoothPanTo(pushpin.getLocation())
			showDetailTooltip(pushpin, 500)
		}

		function pinIconSize(count: number): number {
			if (count >= 100) return 48
			if (count >= 10) return 42
			if (count >= 2) return 36
			return 24
		}

		function clearPushpins(): void {
			for (let i = map.entities.getLength() - 1; i >= 0; i--) {
				const entity = map.entities.get(i)
				if (entity && (entity as any).getLocation) {
					map.entities.removeAt(i)
				}
			}
		}

		interface IOverlapGroup {
			pins: any[]
			loc: any
			pixel: { x: number; y: number }
			count: number
		}

		function clusterByPixelOverlap(sourcePins: any[]): IOverlapGroup[] {
			const groups: IOverlapGroup[] = []
			for (let i = 0; i < sourcePins.length; i++) {
				const loc = sourcePins[i].getLocation()
				const pixel =
					map.tryLocationToPixel(
						loc,
						Microsoft.Maps.PixelReference.control,
					) || { x: 0, y: 0 }
				groups.push({
					pins: [sourcePins[i]],
					loc,
					pixel: { x: pixel.x, y: pixel.y },
					count: 1,
				})
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
								(a.loc.latitude * a.count + b.loc.latitude * b.count) /
									n,
								(a.loc.longitude * a.count +
									b.loc.longitude * b.count) /
									n,
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

		function renderOverlapClusters(): void {
			clearPushpins()
			const groups = clusterByPixelOverlap(sourcePins)
			for (let g = 0; g < groups.length; g++) {
				const group = groups[g]
				if (group.count === 1) {
					const src = group.pins[0]
					const meta = src.metadata as IFormattedItem
					const size = pinIconSize(1)
					const pin = new Microsoft.Maps.Pushpin(src.getLocation(), {
						icon: createCircle(sectorColor(meta.sector).fill, size),
						anchor: new Microsoft.Maps.Point(size / 2, size / 2),
					})
					;(pin as any).metadata = meta
					Microsoft.Maps.Events.addHandler(pin, 'click', pushpinClicked)
					map.entities.push(pin)
					continue
				}
				const size = pinIconSize(group.count)
				const color = self._clusterFill(group.pins)
				const pin = new Microsoft.Maps.Pushpin(group.loc, {
					icon: createCircle(color.fill, size, String(group.count)),
					anchor: new Microsoft.Maps.Point(size / 2, size / 2),
				})
				;(pin as any).containedPushpins = group.pins
				Microsoft.Maps.Events.addHandler(pin, 'click', pushpinClicked)
				map.entities.push(pin)
			}
		}

		const pinLocations: any[] = []
		const sourcePins: any[] = []
		for (const city of withCoords) {
			const lat = Number(city.Lattitude)
			const lng = Number(city.Longitude)
			if (isNaN(lat) || isNaN(lng)) continue
			const loc = new Microsoft.Maps.Location(lat, lng)
			pinLocations.push(loc)
			const pin = new Microsoft.Maps.Pushpin(loc)
			;(pin as any).metadata = city
			sourcePins.push(pin)
		}

		let lastClusterZoom = -1
		function refreshClustersIfZoomChanged(): void {
			const z = map.getZoom()
			if (Math.abs(z - lastClusterZoom) < 0.01) return
			lastClusterZoom = z
			hideTooltip()
			renderOverlapClusters()
		}

		Microsoft.Maps.Events.addHandler(
			map,
			'viewchangeend',
			refreshClustersIfZoomChanged,
		)
		renderOverlapClusters()
		lastClusterZoom = map.getZoom()
		if (pinLocations.length > 1) {
			lastClusterZoom = -1
			map.setView({
				bounds: Microsoft.Maps.LocationRect.fromLocations(pinLocations),
				padding: 80,
			})
		}
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
