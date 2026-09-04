import { Version } from '@microsoft/sp-core-library'
import {
	type IPropertyPaneConfiguration,
	PropertyPaneTextField,
} from '@microsoft/sp-property-pane'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'

import * as strings from 'MapItotWebPartStrings'

const BING_MAPS_KEY =
	'AgJMis6__ur7UfvExgEQb3R3EWCwZTTAYvhjRDPtG9ad-JlxHO3WzhzLIc9DHRzw'
/** Same site as in main.html — list is here; workbench context may point elsewhere */
const SITE_MASTER_DATA_WEB_URL =
	'https://mdigital.sharepoint.com/sites/ITOTCommunityHub'

export interface IMapItotWebPartProps {
	description: string
}

interface ISiteMasterItem {
	ID: number
	Title?: string
	Sector?: string | { Results?: string[] }
	field_2?: string
	field_5?: string
	field_6?: string
	field_7?: string
	field_8?: string
	field_9?: string
	LegalEntityNumber?: string
	Lattitude?: number | string
	Longitude?: number | string
	Latitude?: number | string
	[key: string]:
		| string
		| number
		| { Title?: string; Results?: string[] }
		| undefined
}

interface IMESFieldNames {
	system?: string
	version?: string
	template?: string
	projectStatus?: string
	mainContact?: string
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
	Lattitude?: number
	Longitude?: number
	MesSystem: string
	MesVersion: string
	MesTemplate: string
	MesProjectStatus: string
	MesMainContact: string
}

declare global {
	interface Window {
		Microsoft?: any
		initMapItot?: () => void
	}
}

export default class MapItotWebPart extends BaseClientSideWebPart<IMapItotWebPartProps> {
	private _mapContainer: HTMLDivElement | null = null

	public render(): void {
		this.domElement.innerHTML = ''
		const container = document.createElement('div')
		container.className = 'map-itot-container'
		this._mapContainer = document.createElement('div')
		this._mapContainer.id = 'map-itot-map'
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
        font-size: 14px;
        font-weight: 500;
        font-size: 12px;
      }
    `
		this.domElement.appendChild(style)
	}

	private _hideBingCredentialsError(): void {
		const credsText = 'The specified credentials are invalid'
		const root = document.getElementById('map-itot-map') || this.domElement
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

	private async _getListMESFieldNames(): Promise<IMESFieldNames> {
		const url = `${SITE_MASTER_DATA_WEB_URL}/_api/web/lists/getbytitle('Site Master Data')/fields?$select=InternalName,Title`
		const out: IMESFieldNames = {}
		try {
			const res = await fetch(url, {
				method: 'GET',
				headers: { accept: 'application/json;odata=verbose' },
				credentials: 'include',
			})
			if (!res.ok) {
				console.warn(
					'[MapItot] Fields API failed:',
					res.status,
					res.statusText,
					url,
				)
				return out
			}
			const json = await res.json()
			const fields: { InternalName: string; Title: string }[] =
				json.d?.results || []
			const byTitle: Record<string, string> = {}
			fields.forEach((f: { InternalName: string; Title: string }) => {
				const t = f.Title && f.Title.trim()
				if (t) {
					byTitle[t] = f.InternalName
					byTitle[t.toLowerCase()] = f.InternalName
				}
			})
			const find = (...titles: string[]): string | undefined => {
				for (const t of titles) {
					if (!t) continue
					const key = t.trim()
					if (byTitle[key]) return byTitle[key]
					if (byTitle[key.toLowerCase()]) return byTitle[key.toLowerCase()]
				}
				return undefined
			}
			out.system = find('MES System', 'MESSystem')
			out.version = find('MES Version', 'MESVersion')
			out.template = find('MES Template', 'MESTemplate')
			out.projectStatus = find(
				'MES Project Status',
				'MESProjectStatus',
				'Project Running',
			)
			out.mainContact = find('MES Main Contact', 'MESMainContact')
			if (typeof console !== 'undefined' && console.log) {
				console.log(
					'[MapItot] MES field InternalNames from list (used in $select):',
					out,
				)
			}
			return out
		} catch {
			return out
		}
	}

	private async _getSiteMasterData(): Promise<IFormattedSite[]> {
		const listTitle = 'Site Master Data'
		const fetchOpts: RequestInit = {
			method: 'GET',
			headers: { accept: 'application/json;odata=verbose' },
			credentials: 'include',
		}
		const selectBase = [
			'ID',
			'Title',
			'Sector',
			'field_2',
			'field_5',
			'field_6',
			'field_7',
			'field_8',
			'field_9',
			'LegalEntityNumber',
			'Lattitude',
			'Longitude',
		]
		const mesFields = await this._getListMESFieldNames()
		const selectMESNoContact = [
			mesFields.system,
			mesFields.version,
			mesFields.template,
			mesFields.projectStatus,
		].filter((x): x is string => !!x)
		// Person or Group: в $select нужны и поле, и подполя развёрнутой сущности (как в PnP: Employee/Id, Employee/Title, Employee/EMail)
		const contactField = mesFields.mainContact || 'MESMainContact'
		const selectMESWithContact = [
			mesFields.system,
			mesFields.version,
			mesFields.template,
			mesFields.projectStatus,
			contactField,
			contactField + '/Id',
			contactField + '/Title',
			contactField + '/EMail',
		].filter((x): x is string => !!x)
		const selectFullNoContact = selectBase.concat(selectMESNoContact).join(',')
		const selectFullWithContact = selectBase
			.concat(selectMESWithContact)
			.join(',')
		const selectShortNoContact = selectBase
			.concat(['MESSystem', 'MESVersion', 'MESTemplate', 'MESProjectStatus'])
			.join(',')
		const selectShortWithContact = selectBase
			.concat([
				'MESSystem',
				'MESVersion',
				'MESTemplate',
				'MESProjectStatus',
				'MESMainContact',
				'MESMainContact/Id',
				'MESMainContact/Title',
				'MESMainContact/EMail',
			])
			.join(',')
		const selectMinimal = selectBase.join(',')
		const listTitleEnc = encodeURIComponent(listTitle)
		const listPath =
			'/_api/web/lists/getbytitle(%27' + listTitleEnc + '%27)/items'
		const buildUrl = (select: string, expand?: string): string => {
			const params = new URLSearchParams()
			params.set('$top', '1000')
			params.set('$select', select)
			if (expand) params.set('$expand', expand)
			return SITE_MASTER_DATA_WEB_URL + listPath + '?' + params.toString()
		}
		const expandName = mesFields.mainContact || 'MESMainContact'

		// 1) Пробуем с Main Contact + $expand (точное InternalName из API)
		let url = buildUrl(selectFullWithContact, expandName)
		let response: Response
		try {
			if (typeof console !== 'undefined' && console.log) {
				console.log('[MapItot] Try 1: with Main Contact, $expand=', expandName)
			}
			response = await fetch(url, fetchOpts)
			if (response.ok) {
				if (typeof console !== 'undefined' && console.log) {
					console.log(
						'[MapItot] Loaded with Main Contact (expand=',
						expandName,
						')',
					)
				}
			}
		} catch (err) {
			console.error('[MapItot] Fetch failed', err)
			response = new Response(null, { status: 0 })
		}

		if (!response.ok && response.status !== 0) {
			const errBody = await response.text()
			console.warn(
				'[MapItot] Try 1 failed:',
				response.status,
				errBody.slice(0, 300),
			)
			// 2) Короткие имена + $expand MESMainContact
			url = buildUrl(selectShortWithContact, 'MESMainContact')
			response = await fetch(url, fetchOpts)
			if (response.ok && typeof console !== 'undefined' && console.log) {
				console.log(
					'[MapItot] Loaded with Main Contact (short names + expand MESMainContact)',
				)
			}
		}
		if (!response.ok && response.status !== 0) {
			const errBody = await response.text()
			console.warn(
				'[MapItot] Try 2 failed:',
				response.status,
				errBody.slice(0, 200),
			)
			// 3) Короткие имена + $expand с _x0020_
			url = buildUrl(selectShortWithContact, 'MES_x0020_Main_x0020_Contact')
			response = await fetch(url, fetchOpts)
			if (response.ok && typeof console !== 'undefined' && console.log) {
				console.log(
					'[MapItot] Loaded with Main Contact (short names + expand MES_x0020_Main_x0020_Contact)',
				)
			}
		}
		if (!response.ok && response.status !== 0) {
			const errBody = await response.text()
			console.warn(
				'[MapItot] Try 3 failed:',
				response.status,
				errBody.slice(0, 200),
			)
			// 4) Без Main Contact — стабильный вариант (все 4 MES текстовых поля)
			url = buildUrl(selectFullNoContact)
			response = await fetch(url, fetchOpts)
			if (response.ok && typeof console !== 'undefined' && console.log) {
				console.log(
					'[MapItot] Loaded without Main Contact (MES System/Version/Template/Project Status only)',
				)
			}
		}
		if (!response.ok && response.status !== 0) {
			const errBody = await response.text()
			console.warn(
				'[MapItot] Try 4 failed:',
				response.status,
				errBody.slice(0, 200),
			)
			// 5) Короткие имена без контакта
			url = buildUrl(selectShortNoContact)
			response = await fetch(url, fetchOpts)
		}
		if (!response.ok && response.status !== 0) {
			// 6) Минимум — только базовые поля
			url = buildUrl(selectMinimal)
			response = await fetch(url, fetchOpts)
			if (!response.ok) {
				const errText = await response.text()
				console.error(
					'[MapItot] All attempts failed',
					response.status,
					response.statusText,
					errText.slice(0, 300),
				)
				return []
			}
		}
		const json = await response.json()
		const results: ISiteMasterItem[] = json.d?.results || []
		if (typeof console !== 'undefined' && console.log) {
			console.log('[MapItot] --- Data fetch summary ---')
			console.log('[MapItot] Success URL (used):', url)
			console.log(
				'[MapItot] Items from API:',
				results.length,
				'| $select contains MES?',
				url.indexOf('MES') >= 0,
			)
			console.log('[MapItot] MES keys we use:', mesFields)
			console.log(
				'[MapItot] All keys in first item:',
				results[0] ? Object.keys(results[0]) : [],
			)
			const sample = results.slice(0, 3)
			sample.forEach((row: ISiteMasterItem, i: number) => {
				const item = row as Record<string, unknown>
				const mesKeys = Object.keys(item).filter((k) => k.indexOf('MES') >= 0)
				const mesValues: Record<string, unknown> = {}
				mesKeys.forEach((k) => {
					mesValues[k] = item[k]
				})
				console.log(
					`[MapItot] Item ${i} (ID ${item.ID}, ${item.field_2}): MES keys in response =`,
					mesKeys.length ? mesValues : '(none)',
				)
				if (mesFields.projectStatus) {
					console.log(
						`[MapItot] Item ${i} item["${mesFields.projectStatus}"] =`,
						item[mesFields.projectStatus],
					)
				}
			})
			console.log('[MapItot] --- end summary ---')
		}
		const getSector = (item: ISiteMasterItem): string => {
			const s = item.Sector
			if (!s) return '-'
			if (typeof s === 'string') return s.trim()
			if (
				s &&
				typeof s === 'object' &&
				Array.isArray((s as any).Results) &&
				(s as any).Results[0]
			)
				return (s as any).Results[0].trim()
			return '-'
		}
		const getMESText = (
			val: string | { Title?: string; Results?: string[] } | undefined | null,
		): string => {
			if (val == null || val === '') return '-'
			if (typeof val === 'string') {
				const t = val.trim()
				return t === '' ? '-' : t
			}
			if (typeof val === 'object') {
				if (val.Title && String(val.Title).trim())
					return String(val.Title).trim()
				if (Array.isArray(val.Results) && val.Results[0])
					return String(val.Results[0]).trim()
			}
			return '-'
		}
		const formatted: IFormattedSite[] = results.map((item: ISiteMasterItem) => {
			const latRaw = item.Lattitude ?? item.Latitude
			const lngRaw = item.Longitude
			const lat = latRaw != null ? parseFloat(String(latRaw)) : undefined
			const lng = lngRaw != null ? parseFloat(String(lngRaw)) : undefined
			const toMesVal = (
				v: unknown,
			): string | { Title?: string; Results?: string[] } | undefined => {
				if (v == null) return undefined
				if (
					typeof v === 'string' ||
					(typeof v === 'object' && !Array.isArray(v))
				)
					return v as string | { Title?: string; Results?: string[] }
				if (typeof v === 'number') return String(v)
				return undefined
			}
			const mesSystem =
				mesFields.system != null ? toMesVal(item[mesFields.system]) : undefined
			const mesVersion =
				mesFields.version != null
					? toMesVal(item[mesFields.version])
					: undefined
			const mesTemplate =
				mesFields.template != null
					? toMesVal(item[mesFields.template])
					: undefined
			const mesProjectStatus =
				mesFields.projectStatus != null
					? toMesVal(item[mesFields.projectStatus])
					: undefined
			const mesMainContact =
				mesFields.mainContact != null
					? toMesVal(item[mesFields.mainContact])
					: undefined
			return {
				ID: item.ID,
				Title: item.Title || '-',
				Sector: getSector(item),
				SiteName: item.field_2 || '-',
				BusinessUnit: item.field_5 || '-',
				Location: item.field_8 || '-',
				LegalEntityName: item.field_7 || '-',
				LegalEntityCode: item.LegalEntityNumber || '-',
				Services: item.field_9 || '-',
				Employees: item.field_6 || '-',
				Lattitude: typeof lat === 'number' && !isNaN(lat) ? lat : undefined,
				Longitude: typeof lng === 'number' && !isNaN(lng) ? lng : undefined,
				MesSystem: getMESText(
					mesSystem ??
						toMesVal((item as any).MES_x0020_System) ??
						toMesVal((item as any).MESSystem),
				),
				MesVersion: getMESText(
					mesVersion ??
						toMesVal((item as any).MES_x0020_Version) ??
						toMesVal((item as any).MESVersion),
				),
				MesTemplate: getMESText(
					mesTemplate ??
						toMesVal((item as any).MES_x0020_Template) ??
						toMesVal((item as any).MESTemplate),
				),
				MesProjectStatus: getMESText(
					mesProjectStatus ??
						toMesVal((item as any).MES_x0020_Project_x0020_Status) ??
						toMesVal((item as any).MESProjectStatus),
				),
				MesMainContact: getMESText(
					mesMainContact ??
						toMesVal((item as any).MES_x0020_Main_x0020_Contact) ??
						toMesVal((item as any).MESMainContact),
				),
			}
		})
		const healthcare = formatted.filter((s) => {
			const sector = (s.Sector || '').toLowerCase().trim()
			return sector === 'healthcare' || sector.indexOf('healthcare') !== -1
		})
		return healthcare
	}

	private _loadBingAndInit(): void {
		const self = this
		window.initMapItot = function () {
			self._initMap()
		}
		const script = document.createElement('script')
		script.type = 'text/javascript'
		script.src = `https://www.bing.com/api/maps/mapcontrol?key=${BING_MAPS_KEY}&callback=initMapItot`
		script.async = true
		script.defer = true
		document.head.appendChild(script)
	}

	private async _initMap(): Promise<void> {
		const self = this
		const mapEl = document.getElementById('map-itot-map')
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

		const data = await this._getSiteMasterData()
		const withCoords = data.filter(
			(c) =>
				c.Lattitude != null &&
				c.Longitude != null &&
				c.Lattitude !== 0 &&
				c.Longitude !== 0,
		)
		if (withCoords.length === 0 && typeof console !== 'undefined') {
			console.warn(
				'[MapItot] No points to display. Healthcare items:',
				data.length,
				', with valid coordinates: 0. Check list "Site Master Data": Sector=Healthcare and Lattitude/Longitude filled.',
			)
		}

		function createCircle(): string {
			const c = document.createElement('canvas')
			c.width = 24
			c.height = 24
			const ctx = c.getContext('2d')
			if (ctx) {
				ctx.fillStyle = '#eb3c96'
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
			const meta = (pushpin as any).metadata as IFormattedSite
			if (!meta) return

			const pushpinLocation = pushpin.getLocation()
			smoothPanTo(pushpinLocation)

			if (pushpinLocation.latitude > 60) {
				tooltip.setOptions({ offset: new Microsoft.Maps.Point(20, -400) })
			} else {
				tooltip.setOptions({ offset: new Microsoft.Maps.Point(20, -100) })
			}

			infobox.setOptions({ visible: false })

			const detailUrl = `https://mdigital.sharepoint.com/sites/ITOTCommunityHub/Lists/Site%20Master%20Data/DispForm.aspx?ID=${meta.ID}`
			const esc = (s: string) => self._escape(s)
			const htmlContent = `
        <div class="map-tooltip sector-pink">
          <div class="map-tooltip-header">
            <div class="map-tooltip-headline">${esc(meta.SiteName)} / ${esc(meta.BusinessUnit)}</div>
            <span class="map-tooltip-close" id="closeInfobox">×</span>
          </div>
          <div class="map-tooltip-body">
            <div class="map-tooltip-col">
              <div class="map-tooltip-line">${esc(meta.LegalEntityName)}</div>
              <div class="map-tooltip-line">${esc(meta.Location)}</div>
              <div class="map-tooltip-line"><strong>Services:</strong> ${esc(meta.Services)}</div>
              <div class="map-tooltip-line"><strong>Employees:</strong> ${esc(meta.Employees)}</div>
            </div>
            <div class="map-tooltip-col">
              <div class="map-tooltip-line"><strong>MES System:</strong> ${esc(meta.MesSystem)}</div>
              <div class="map-tooltip-line"><strong>Version:</strong> ${esc(meta.MesVersion)}</div>
              <div class="map-tooltip-line"><strong>Template:</strong> ${esc(meta.MesTemplate)}</div>
              <div class="map-tooltip-line"><strong>Project Running:</strong> ${esc(meta.MesProjectStatus)}</div>
              <div class="map-tooltip-line"><strong>Main Contact:</strong></div>
              <div class="map-tooltip-line">${esc(meta.MesMainContact)}</div>
            </div>
          </div>
          <div class="map-tooltip-footer">
            <a class="see-all" href="${detailUrl}">See all details</a>
          </div>
        </div>`

			setTimeout(() => {
				tooltip.setOptions({
					location: pushpinLocation,
					htmlContent,
					visible: true,
				})
				const closeBtn = document.getElementById('closeInfobox')
				if (closeBtn) {
					closeBtn.addEventListener('click', () =>
						tooltip.setOptions({ visible: false }),
					)
				}
			}, 500)
		}

		clearPushpins()
		for (const city of withCoords) {
			const lat = Number(city.Lattitude)
			const lng = Number(city.Longitude)
			if (isNaN(lat) || isNaN(lng)) continue
			const pin = new Microsoft.Maps.Pushpin(
				{ latitude: lat, longitude: lng },
				{ icon: createCircle() },
			)
			;(pin as any).metadata = city
			Microsoft.Maps.Events.addHandler(pin, 'click', pushpinClicked)
			map.entities.push(pin)
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
