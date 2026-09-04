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
		aliases: ['technical contact person', 'technical contact'],
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

	private _matchField(
		fields: IListField[],
		aliases: string[],
	): IListField | undefined {
		return this._find(fields, (f) => {
			const title = (f.Title || '').toLowerCase().trim()
			const internal = this._normName(f.InternalName)
			for (let i = 0; i < aliases.length; i++) {
				const alias = aliases[i]
				if (title === alias || this._normName(alias) === internal) return true
			}
			return false
		})
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
				Description?: string
				Url?: string
				Label?: string
			}
			if (obj.Title && String(obj.Title).trim()) return String(obj.Title).trim()
			if (obj.Label && String(obj.Label).trim()) return String(obj.Label).trim()
			if (obj.Description && String(obj.Description).trim())
				return String(obj.Description).trim()
			if (obj.Url && String(obj.Url).trim()) return String(obj.Url).trim()
			if (Array.isArray(obj.Results) && obj.Results.length) {
				return obj.Results.map((r) => this._fieldValue(r))
					.filter((s) => s !== '-')
					.join(', ') || '-'
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

		const expandNames: string[] = []
		const addExpand = (field?: IListField): void => {
			if (!field) return
			if (field.TypeAsString === 'User' || field.TypeAsString === 'Lookup') {
				if (expandNames.indexOf(field.InternalName) < 0)
					expandNames.push(field.InternalName)
			}
		}
		resolved.forEach((r) => addExpand(r.field))
		addExpand(siteNameField)
		addExpand(siteIdField)
		addExpand(addressField)
		addExpand(servicesField)
		addExpand(employeesField)
		addExpand(sectorField)

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
			})),
		)

		const itemQuery = expandNames.length
			? '/items?$top=1000&$expand=' + expandNames.join(',')
			: '/items?$top=1000'
		const fetched = await Promise.all([
			this._getJson(this._listUrls(itemQuery, LIST_TITLE, LIST_PATH)),
			this._getSiteMasterCoordMap(),
		])
		let json = fetched[0]
		const smd = fetched[1]
		if (!json) {
			json = await this._getJson(
				this._listUrls('/items?$top=1000', LIST_TITLE, LIST_PATH),
			)
		}
		if (!json) {
			console.error('[MapPM] Predictive Maintenance items fetch failed')
			return []
		}

		const results: Record<string, unknown>[] = (json.d && json.d.results) || []
		console.log('[MapPM] PM items:', results.length, '| first keys:', results[0] ? Object.keys(results[0]) : [])

		const prefer = (
			field: IListField | undefined,
			item: Record<string, unknown>,
			fallback: string,
		): string => {
			if (!field) return fallback
			const val = this._fieldValue(item[field.InternalName], field.TypeAsString)
			return val !== '-' ? val : fallback
		}

		const mapped = results.map((item) => {
			const site = this._resolveSite(item, siteNameField, siteIdField, smd)
			const values: ITooltipValues = {
				production: '-',
				technicalContact: '-',
				systemSize: '-',
				startOfOperation: '-',
				technicalSpecials: '-',
			}
			resolved.forEach((r) => {
				if (!r.field) return
				values[r.spec.key] = this._fieldValue(
					item[r.field.InternalName],
					r.field.TypeAsString,
				)
			})
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
				Lattitude: site ? site.lat : undefined,
				Longitude: site ? site.lng : undefined,
			}
		})
		const matched = mapped.filter(
			(m) => m.Lattitude != null && m.Longitude != null,
		).length
		console.log('[MapPM] Matched coords:', matched, '/', mapped.length)
		return this._spreadOverlapping(mapped)
	}

	private _spreadOverlapping(items: IFormattedItem[]): IFormattedItem[] {
		const groups: Record<string, IFormattedItem[]> = {}
		for (let i = 0; i < items.length; i++) {
			const item = items[i]
			if (item.Lattitude == null || item.Longitude == null) {
				if (!groups.none) groups.none = []
				groups.none.push(item)
				continue
			}
			const key =
				item.Lattitude.toFixed(5) + ',' + item.Longitude.toFixed(5)
			if (!groups[key]) groups[key] = []
			groups[key].push(item)
		}
		const out: IFormattedItem[] = []
		const keys = Object.keys(groups)
		for (let g = 0; g < keys.length; g++) {
			const list = groups[keys[g]]
			if (keys[g] === 'none' || list.length === 1) {
				for (let i = 0; i < list.length; i++) out.push(list[i])
				continue
			}
			const baseLat = list[0].Lattitude as number
			const baseLng = list[0].Longitude as number
			const radius = 0.00035
			const latRad = (baseLat * Math.PI) / 180
			for (let i = 0; i < list.length; i++) {
				const angle = (2 * Math.PI * i) / list.length - Math.PI / 2
				const copy: IFormattedItem = {
					ID: list[i].ID,
					siteName: list[i].siteName,
					businessUnit: list[i].businessUnit,
					legalEntityName: list[i].legalEntityName,
					location: list[i].location,
					services: list[i].services,
					employees: list[i].employees,
					sector: list[i].sector,
					production: list[i].production,
					technicalContact: list[i].technicalContact,
					systemSize: list[i].systemSize,
					startOfOperation: list[i].startOfOperation,
					technicalSpecials: list[i].technicalSpecials,
					Lattitude: baseLat + radius * Math.cos(angle),
					Longitude:
						baseLng + (radius * Math.sin(angle)) / Math.cos(latRad),
				}
				out.push(copy)
			}
		}
		return out
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

		function createCircle(fill: string): string {
			const c = document.createElement('canvas')
			c.width = 24
			c.height = 24
			const ctx = c.getContext('2d')
			if (ctx) {
				ctx.fillStyle = fill
				ctx.lineWidth = 2
				ctx.strokeStyle = '#0f69af'
				ctx.beginPath()
				ctx.arc(c.width * 0.5, c.height * 0.5, 10, 0, 2 * Math.PI)
				ctx.fill()
				ctx.stroke()
			}
			return c.toDataURL()
		}

		function clearPushpins(): void {
			for (let i = map.entities.getLength() - 1; i >= 0; i--) {
				const entity = map.entities.get(i)
				if (entity && (entity as any).getLocation) {
					map.entities.removeAt(i)
				}
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

		function pushpinClicked(e: any): void {
			const pushpin = e.target
			const meta = (pushpin as any).metadata as IFormattedItem
			if (!meta) return

			const pushpinLocation = pushpin.getLocation()
			smoothPanTo(pushpinLocation)

			if (pushpinLocation.latitude > 60) {
				tooltip.setOptions({ offset: new Microsoft.Maps.Point(20, -400) })
			} else {
				tooltip.setOptions({ offset: new Microsoft.Maps.Point(20, -100) })
			}

			infobox.setOptions({ visible: false })

			const esc = (s: string) => self._escape(s)
			const city = (() => {
				const name = meta.siteName && meta.siteName !== '-' ? meta.siteName : ''
				const bu =
					meta.businessUnit && meta.businessUnit !== '-'
						? meta.businessUnit
						: ''
				if (name && bu) return name + ' / ' + bu
				return name || bu || ''
			})()
			const addressHtml = esc(meta.location).replace(/\n/g, '<br>')
			const sectorCss = sectorColor(meta.sector).css
			const htmlContent = `
        <div class="map-tooltip ${sectorCss}">
          <div class="map-tooltip-header">
            <div class="map-tooltip-headline">${esc(city)}</div>
            <span class="map-tooltip-close" id="closeInfoboxPm">×</span>
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
				const closeBtn = document.getElementById('closeInfoboxPm')
				if (closeBtn) {
					closeBtn.addEventListener('click', () =>
						tooltip.setOptions({ visible: false }),
					)
				}
			}, 500)
		}

		clearPushpins()
		const pinLocations: any[] = []
		for (const city of withCoords) {
			const lat = Number(city.Lattitude)
			const lng = Number(city.Longitude)
			if (isNaN(lat) || isNaN(lng)) continue
			const loc = new Microsoft.Maps.Location(lat, lng)
			pinLocations.push(loc)
			const pin = new Microsoft.Maps.Pushpin(loc, {
				icon: createCircle(sectorColor(city.sector).fill),
			})
			;(pin as any).metadata = city
			Microsoft.Maps.Events.addHandler(pin, 'click', pushpinClicked)
			map.entities.push(pin)
		}
		if (pinLocations.length > 1) {
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
