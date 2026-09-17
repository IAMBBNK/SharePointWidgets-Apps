import { Version } from '@microsoft/sp-core-library'
import {
	type IPropertyPaneConfiguration,
	PropertyPaneDropdown,
	PropertyPaneTextField,
} from '@microsoft/sp-property-pane'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'

import * as strings from 'ImageCarouselWebPartStrings'

export interface IImageCarouselWebPartProps {
	buttonText: string
	buttonIcon: string
	buttonColor: string
	buttonAlign: string
	visibleCount: string
	slidesJson: string
}

interface ICarouselSlide {
	image: string
	title: string
	description: string
}

const CHEVRON_LEFT =
	'<svg viewBox="0 0 24 24" aria-hidden="true"><polyline points="15 6 9 12 15 18"></polyline></svg>'
const CHEVRON_RIGHT =
	'<svg viewBox="0 0 24 24" aria-hidden="true"><polyline points="9 6 15 12 9 18"></polyline></svg>'

export default class ImageCarouselWebPart extends BaseClientSideWebPart<IImageCarouselWebPartProps> {
	private static _stylesInjected: boolean = false
	private _overlay: HTMLDivElement | undefined
	private _cardsEl: HTMLDivElement | undefined
	private _dotsEl: HTMLDivElement | undefined
	private _prevBtn: HTMLButtonElement | undefined
	private _nextBtn: HTMLButtonElement | undefined
	private _closeBtn: HTMLButtonElement | undefined
	private _isOpen: boolean = false
	private _index: number = 0
	private _direction: 'next' | 'prev' = 'next'
	private _lastFocused: HTMLElement | undefined
	private _bodyOverflow: string = ''

	public render(): void {
		this._injectStyles()
		this._renderTrigger()
		this._ensureOverlay()
		if (this._isOpen) {
			this._renderSlides(false)
		}
	}

	protected onDispose(): void {
		this._close(true)
		if (this._overlay && this._overlay.parentElement) {
			this._overlay.parentElement.removeChild(this._overlay)
		}
		this._overlay = undefined
		super.onDispose()
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
							groupName: strings.ButtonGroupName,
							groupFields: [
								PropertyPaneTextField('buttonText', {
									label: strings.ButtonTextFieldLabel,
								}),
								PropertyPaneTextField('buttonIcon', {
									label: strings.ButtonIconFieldLabel,
								}),
								PropertyPaneTextField('buttonColor', {
									label: strings.ButtonColorFieldLabel,
								}),
								PropertyPaneDropdown('buttonAlign', {
									label: strings.ButtonAlignFieldLabel,
									options: [
										{ key: 'left', text: 'Left' },
										{ key: 'center', text: 'Center' },
										{ key: 'right', text: 'Right' },
									],
								}),
							],
						},
						{
							groupName: strings.CarouselGroupName,
							groupFields: [
								PropertyPaneDropdown('visibleCount', {
									label: strings.VisibleCountFieldLabel,
									options: [
										{ key: '1', text: strings.VisibleCountOne },
										{ key: '3', text: strings.VisibleCountThree },
									],
								}),
								PropertyPaneTextField('slidesJson', {
									label: strings.SlidesJsonFieldLabel,
									description: strings.SlidesJsonDescription,
									multiline: true,
									rows: 14,
								}),
							],
						},
					],
				},
			],
		}
	}

	private _getSlides(): ICarouselSlide[] {
		try {
			const parsed: unknown = JSON.parse(this.properties.slidesJson || '[]')
			if (!Array.isArray(parsed)) {
				return []
			}
			const slides: ICarouselSlide[] = []
			for (let i = 0; i < parsed.length; i++) {
				const item: { image?: unknown; title?: unknown; description?: unknown } =
					parsed[i] as { image?: unknown; title?: unknown; description?: unknown }
				if (!item || !item.image) {
					continue
				}
				slides.push({
					image: String(item.image),
					title: item.title ? String(item.title) : '',
					description: item.description ? String(item.description) : '',
				})
			}
			return slides
		} catch (error) {
			return []
		}
	}

	private _getVisibleCount(slideCount: number): number {
		const preferred: number = this.properties.visibleCount === '1' ? 1 : 3
		const forViewport: number = window.innerWidth < 720 ? 1 : preferred
		return Math.max(1, Math.min(forViewport, slideCount || 1))
	}

	private _renderTrigger(): void {
		const slides: ICarouselSlide[] = this._getSlides()
		const align: string = (this.properties.buttonAlign || 'left').toLowerCase()
		const buttonText: string = this.properties.buttonText || 'View gallery'
		const buttonIcon: string = (this.properties.buttonIcon || '').trim()
		const buttonColor: string = this.properties.buttonColor || '#0f69af'

		this.domElement.className = 'imgc-widget'
		this.domElement.classList.toggle('is-center', align === 'center')
		this.domElement.classList.toggle('is-right', align === 'right')

		if (!slides.length) {
			this.domElement.innerHTML =
				'<div class="imgc-empty">Add slides in the web part properties (image, title, description).</div>'
			return
		}

		const iconHtml: string = buttonIcon
			? `<img class="imgc-trigger-icon" src="${this._escapeAttr(buttonIcon)}" alt="" />`
			: ''

		this.domElement.innerHTML =
			`<button type="button" class="imgc-trigger" style="background:${this._escapeAttr(buttonColor)}">` +
			iconHtml +
			`<span>${this._escapeHtml(buttonText)}</span></button>`

		const trigger: HTMLButtonElement | null =
			this.domElement.querySelector('.imgc-trigger')
		if (trigger) {
			trigger.addEventListener('click', this._open)
		}
	}

	private _ensureOverlay(): void {
		if (this._overlay) {
			return
		}

		const overlay: HTMLDivElement = document.createElement('div')
		overlay.className = 'imgc-overlay'
		overlay.setAttribute('hidden', 'hidden')
		overlay.innerHTML =
			'<button type="button" class="imgc-close" aria-label="Close gallery">&times;</button>' +
			'<div class="imgc-overlay-inner">' +
			'<div class="imgc-stage">' +
			`<button type="button" class="imgc-nav imgc-prev" aria-label="Previous">${CHEVRON_LEFT}</button>` +
			'<div class="imgc-viewport"><div class="imgc-cards"></div></div>' +
			`<button type="button" class="imgc-nav imgc-next" aria-label="Next">${CHEVRON_RIGHT}</button>` +
			'</div>' +
			'<div class="imgc-dots"></div>' +
			'</div>'

		document.body.appendChild(overlay)

		this._overlay = overlay
		this._cardsEl = overlay.querySelector('.imgc-cards') as HTMLDivElement
		this._dotsEl = overlay.querySelector('.imgc-dots') as HTMLDivElement
		this._prevBtn = overlay.querySelector('.imgc-prev') as HTMLButtonElement
		this._nextBtn = overlay.querySelector('.imgc-next') as HTMLButtonElement
		this._closeBtn = overlay.querySelector('.imgc-close') as HTMLButtonElement

		overlay.addEventListener('click', this._onOverlayClick)
		this._closeBtn.addEventListener('click', this._onCloseClick)
		this._prevBtn.addEventListener('click', this._prev)
		this._nextBtn.addEventListener('click', this._next)
	}

	private _open = (): void => {
		if (this._isOpen || !this._overlay || !this._closeBtn) {
			return
		}
		const slides: ICarouselSlide[] = this._getSlides()
		if (!slides.length) {
			return
		}
		this._isOpen = true
		this._index = 0
		this._lastFocused = document.activeElement as HTMLElement
		this._bodyOverflow = document.body.style.overflow
		this._overlay.classList.add('is-open')
		this._overlay.removeAttribute('hidden')
		this._overlay.setAttribute('role', 'dialog')
		this._overlay.setAttribute('aria-modal', 'true')
		this._overlay.setAttribute(
			'aria-label',
			this.properties.buttonText || 'Image gallery'
		)
		document.body.style.overflow = 'hidden'
		this._renderSlides(false)
		this._closeBtn.focus()
		document.addEventListener('keydown', this._onKeydown)
		window.addEventListener('resize', this._onResize)
	}

	private _close = (fromDispose?: boolean): void => {
		if (!this._isOpen) {
			return
		}
		this._isOpen = false
		if (this._overlay) {
			this._overlay.classList.remove('is-open')
			this._overlay.setAttribute('hidden', 'hidden')
		}
		document.body.style.overflow = this._bodyOverflow
		document.removeEventListener('keydown', this._onKeydown)
		window.removeEventListener('resize', this._onResize)
		if (
			!fromDispose &&
			this._lastFocused &&
			typeof this._lastFocused.focus === 'function'
		) {
			this._lastFocused.focus()
		}
	}

	private _onCloseClick = (): void => {
		this._close()
	}

	private _onOverlayClick = (event: MouseEvent): void => {
		if (event.target === this._overlay) {
			this._close()
		}
	}

	private _next = (): void => {
		const slides: ICarouselSlide[] = this._getSlides()
		if (slides.length <= 1) {
			return
		}
		this._direction = 'next'
		this._index = (this._index + 1) % slides.length
		this._renderSlides(true)
	}

	private _prev = (): void => {
		const slides: ICarouselSlide[] = this._getSlides()
		if (slides.length <= 1) {
			return
		}
		this._direction = 'prev'
		this._index = (this._index - 1 + slides.length) % slides.length
		this._renderSlides(true)
	}

	private _onKeydown = (event: KeyboardEvent): void => {
		if (!this._isOpen) {
			return
		}
		if (event.key === 'Escape') {
			this._close()
		} else if (event.key === 'ArrowRight') {
			this._next()
		} else if (event.key === 'ArrowLeft') {
			this._prev()
		}
	}

	private _onResize = (): void => {
		if (this._isOpen) {
			this._renderSlides(false)
		}
	}

	private _renderSlides(animate: boolean): void {
		if (
			!this._overlay ||
			!this._cardsEl ||
			!this._dotsEl ||
			!this._prevBtn ||
			!this._nextBtn
		) {
			return
		}

		const slides: ICarouselSlide[] = this._getSlides()
		if (!slides.length) {
			this._close()
			return
		}
		if (this._index >= slides.length) {
			this._index = 0
		}

		const visibleCount: number = this._getVisibleCount(slides.length)
		const showNav: boolean = slides.length > visibleCount
		const visible: ICarouselSlide[] = []
		for (let i = 0; i < visibleCount; i++) {
			visible.push(slides[(this._index + i) % slides.length])
		}

		this._overlay.classList.toggle('is-single', visibleCount === 1)
		this._prevBtn.classList.toggle('is-hidden', !showNav)
		this._nextBtn.classList.toggle('is-hidden', !showNav)
		this._dotsEl.classList.toggle('is-hidden', !showNav)

		let cardsHtml: string = ''
		for (let i = 0; i < visible.length; i++) {
			const item: ICarouselSlide = visible[i]
			cardsHtml +=
				'<article class="imgc-card">' +
				'<div class="imgc-card-image-wrap">' +
				`<img class="imgc-card-image" src="${this._escapeAttr(item.image)}" alt="${this._escapeAttr(item.title)}" />` +
				'</div>' +
				`<h3 class="imgc-card-title">${this._escapeHtml(item.title)}</h3>` +
				`<p class="imgc-card-desc">${this._escapeHtml(item.description)}</p>` +
				'</article>'
		}
		this._cardsEl.innerHTML = cardsHtml

		if (animate) {
			const animClass: string =
				this._direction === 'prev' ? 'is-animating-prev' : 'is-animating'
			this._cardsEl.classList.remove('is-animating', 'is-animating-prev')
			// Force reflow so the animation restarts.
			// eslint-disable-next-line @typescript-eslint/no-unused-expressions
			this._cardsEl.offsetWidth
			this._cardsEl.classList.add(animClass)
		}

		let dotsHtml: string = ''
		for (let i = 0; i < slides.length; i++) {
			const active: string = i === this._index ? ' is-active' : ''
			dotsHtml += `<button type="button" class="imgc-dot${active}" data-index="${i}" aria-label="Go to slide ${i + 1}"></button>`
		}
		this._dotsEl.innerHTML = dotsHtml

		const dots: NodeListOf<HTMLButtonElement> =
			this._dotsEl.querySelectorAll('.imgc-dot')
		for (let i = 0; i < dots.length; i++) {
			dots[i].addEventListener('click', this._onDotClick)
		}
	}

	private _onDotClick = (event: Event): void => {
		const target: HTMLElement = event.currentTarget as HTMLElement
		const nextIndex: number = parseInt(target.getAttribute('data-index') || '0', 10)
		this._direction = nextIndex < this._index ? 'prev' : 'next'
		this._index = nextIndex
		this._renderSlides(true)
	}

	private _escapeHtml(value: string): string {
		return String(value || '')
			.replace(/&/g, '&amp;')
			.replace(/</g, '&lt;')
			.replace(/>/g, '&gt;')
			.replace(/"/g, '&quot;')
	}

	private _escapeAttr(value: string): string {
		return this._escapeHtml(value).replace(/'/g, '&#39;')
	}

	private _injectStyles(): void {
		if (ImageCarouselWebPart._stylesInjected) {
			return
		}
		ImageCarouselWebPart._stylesInjected = true
		const style: HTMLStyleElement = document.createElement('style')
		style.textContent = `
.imgc-widget,
.imgc-overlay {
  font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
  box-sizing: border-box;
  color: #1f1f1f;
}
.imgc-widget *,
.imgc-widget *::before,
.imgc-widget *::after,
.imgc-overlay *,
.imgc-overlay *::before,
.imgc-overlay *::after { box-sizing: border-box; }
.imgc-widget { display: flex; width: 100%; }
.imgc-widget.is-center { justify-content: center; }
.imgc-widget.is-right { justify-content: flex-end; }
.imgc-trigger {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 10px;
  margin: 0;
  border: none;
  border-radius: 5px;
  padding: 12px 18px;
  background: #0f69af;
  color: #fff;
  font-family: inherit;
  font-size: 15px;
  font-weight: 600;
  line-height: 1.2;
  cursor: pointer;
  box-shadow: 0 1px 2px rgba(15, 105, 175, 0.25);
}
.imgc-trigger:hover,
.imgc-trigger:focus-visible { filter: brightness(1.08); outline: none; }
.imgc-trigger-icon {
  width: 22px;
  height: 22px;
  object-fit: contain;
  flex: 0 0 auto;
  display: block;
}
.imgc-overlay {
  position: fixed;
  inset: 0;
  z-index: 1000000;
  display: none;
  align-items: center;
  justify-content: center;
  padding: 28px 64px 36px;
  background: rgba(120, 120, 120, 0.88);
}
.imgc-overlay.is-open { display: flex; }
.imgc-overlay-inner {
  position: relative;
  display: flex;
  flex-direction: column;
  align-items: center;
  width: min(980px, 100%);
  gap: 18px;
}
.imgc-overlay.is-single .imgc-overlay-inner { width: min(440px, 100%); }
.imgc-close {
  position: absolute;
  top: 16px;
  right: 16px;
  width: 40px;
  height: 40px;
  border: none;
  border-radius: 50%;
  background: rgba(255, 255, 255, 0.92);
  color: #333;
  font-size: 26px;
  line-height: 1;
  cursor: pointer;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.18);
}
.imgc-close:hover,
.imgc-close:focus-visible { background: #fff; outline: none; }
.imgc-stage {
  display: flex;
  align-items: center;
  width: 100%;
  gap: 8px;
}
.imgc-viewport { flex: 1; overflow: hidden; min-width: 0; }
.imgc-cards { display: flex; gap: 16px; width: 100%; }
.imgc-cards.is-animating { animation: imgc-slide 0.32s ease; }
.imgc-cards.is-animating-prev { animation: imgc-slide-prev 0.32s ease; }
@keyframes imgc-slide {
  from { opacity: 0.35; transform: translateX(28px); }
  to { opacity: 1; transform: translateX(0); }
}
@keyframes imgc-slide-prev {
  from { opacity: 0.35; transform: translateX(-28px); }
  to { opacity: 1; transform: translateX(0); }
}
.imgc-card {
  flex: 1 1 0;
  min-width: 0;
  background: #fff;
  border-radius: 12px;
  padding: 12px 12px 16px;
  box-shadow: 0 10px 28px rgba(0, 0, 0, 0.18);
}
.imgc-card-image-wrap {
  width: 100%;
  height: 168px;
  border-radius: 10px;
  overflow: hidden;
  background: #ececec;
}
.imgc-overlay.is-single .imgc-card-image-wrap { height: 260px; }
.imgc-card-image { width: 100%; height: 100%; object-fit: cover; display: block; }
.imgc-card-title {
  margin: 12px 4px 4px;
  font-size: 16px;
  font-weight: 700;
  line-height: 1.3;
  color: #111;
}
.imgc-card-desc {
  margin: 0 4px;
  font-size: 13px;
  line-height: 1.45;
  color: #6b6b6b;
}
.imgc-nav {
  flex: 0 0 auto;
  width: 48px;
  height: 48px;
  border: none;
  border-radius: 50%;
  background: #3a3a3a;
  color: #fff;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  box-shadow: 0 3px 10px rgba(0, 0, 0, 0.28);
  padding: 0;
}
.imgc-nav:hover,
.imgc-nav:focus-visible { background: #2b2b2b; outline: none; }
.imgc-nav svg {
  width: 18px;
  height: 18px;
  fill: none;
  stroke: currentColor;
  stroke-width: 2.4;
  stroke-linecap: round;
  stroke-linejoin: round;
}
.imgc-nav.is-hidden,
.imgc-dots.is-hidden { visibility: hidden; pointer-events: none; }
.imgc-dots {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  min-height: 12px;
}
.imgc-dot {
  width: 9px;
  height: 9px;
  padding: 0;
  border: none;
  border-radius: 50%;
  background: #c8c8c8;
  cursor: pointer;
}
.imgc-dot.is-active { background: #7cb342; }
.imgc-empty { color: #b00020; font-size: 13px; font-weight: 600; }
@media (max-width: 720px) {
  .imgc-overlay { padding: 56px 12px 24px; }
  .imgc-nav { width: 38px; height: 38px; }
  .imgc-card-image-wrap,
  .imgc-overlay.is-single .imgc-card-image-wrap { height: 190px; }
}
`
		document.head.appendChild(style)
	}
}
