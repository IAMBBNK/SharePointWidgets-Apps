import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import * as strings from 'PageViewTrackerApplicationCustomizerStrings';

const LOG_SOURCE: string = 'PageViewTrackerApplicationCustomizer';

/** Site Pages library template type */
const WEB_PAGE_LIBRARY_TEMPLATE: number = 119;

/** PromotedState value for SharePoint news posts */
const NEWS_PROMOTED_STATE: number = 2;

export interface IPageViewTrackerApplicationCustomizerProperties {
  trackingListTitle: string;
  trackingListSiteUrl: string;
  sourceQueryParam: string;
  trackNewsOnly: boolean;
  defaultSource: string;
  dedupeInSession: boolean;
}

interface IPageMetadata {
  isNewsPage: boolean;
  hubSection: string;
}

interface IListItemPayload {
  Title: string;
  PageUrl: string;
  PagePath: string;
  Source: string;
  UserEmail: string;
  UserDisplayName: string;
  ViewedAt: string;
  Referrer: string;
  IsNewsPage: boolean;
  HubSection: string;
  UserAgent: string;
}

export default class PageViewTrackerApplicationCustomizer
  extends BaseApplicationCustomizer<IPageViewTrackerApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.trackPageViewAsync().catch((error: unknown) => {
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(String(error)));
    });

    return Promise.resolve();
  }

  private async trackPageViewAsync(): Promise<void> {
    if (!this.shouldTrackCurrentPage()) {
      return;
    }

    const properties = this.getResolvedProperties();
    const pageUrl: string = window.location.href.split('#')[0];
    const pagePath: string = this.context.pageContext.site.serverRequestPath || window.location.pathname;
    const source: string = this.getTrafficSource(properties.sourceQueryParam, properties.defaultSource);

    if (properties.dedupeInSession && this.hasSessionRecord(pagePath, source)) {
      Log.info(LOG_SOURCE, `Skipped duplicate view for ${pagePath} (${source})`);
      return;
    }

    const metadata: IPageMetadata = await this.getPageMetadata();
    if (properties.trackNewsOnly && !metadata.isNewsPage) {
      return;
    }

    await this.writeTrackingItem({
      Title: this.getPageTitle(),
      PageUrl: pageUrl,
      PagePath: pagePath,
      Source: source,
      UserEmail: this.context.pageContext.user.email || '',
      UserDisplayName: this.context.pageContext.user.displayName || '',
      ViewedAt: new Date().toISOString(),
      Referrer: document.referrer || '',
      IsNewsPage: metadata.isNewsPage,
      HubSection: metadata.hubSection,
      UserAgent: navigator.userAgent || ''
    }, properties);

    if (properties.dedupeInSession) {
      this.markSessionRecord(pagePath, source);
    }
  }

  private getResolvedProperties(): Required<IPageViewTrackerApplicationCustomizerProperties> {
    return {
      trackingListTitle: this.properties.trackingListTitle || 'Page View Tracking',
      trackingListSiteUrl: this.properties.trackingListSiteUrl || this.context.pageContext.web.absoluteUrl,
      sourceQueryParam: this.properties.sourceQueryParam || 'source',
      trackNewsOnly: this.properties.trackNewsOnly === true,
      defaultSource: this.properties.defaultSource || 'direct',
      dedupeInSession: this.properties.dedupeInSession !== false
    };
  }

  private shouldTrackCurrentPage(): boolean {
    if (this.context.pageContext.legacyPageContext?.isWebWelcomePage) {
      return false;
    }

    if (this.isEditMode()) {
      return false;
    }

    const listTemplateType: number | undefined = this.context.pageContext.list?.templateType;
    if (listTemplateType !== WEB_PAGE_LIBRARY_TEMPLATE) {
      return false;
    }

    const pagePath: string = (this.context.pageContext.site.serverRequestPath || '').toLowerCase();
    if (!pagePath || pagePath.indexOf('/sitepages/') === -1) {
      return false;
    }

    return true;
  }

  private isEditMode(): boolean {
    const legacyContext = this.context.pageContext.legacyPageContext;
    if (legacyContext && typeof legacyContext.editMode !== 'undefined') {
      return legacyContext.editMode === true;
    }

    const mode: string = (this.context.pageContext.webUIContext as { ViewMode?: string } | undefined)?.ViewMode || '';
    return mode.toLowerCase() === 'edit';
  }

  private getTrafficSource(queryParamName: string, defaultSource: string): string {
    const params: URLSearchParams = new URLSearchParams(window.location.search);
    const rawValue: string | null = params.get(queryParamName);

    if (!rawValue) {
      return defaultSource;
    }

    const normalized: string = rawValue.trim().toLowerCase();
    return normalized || defaultSource;
  }

  private getPageTitle(): string {
    const listItemTitle: string | undefined = this.context.pageContext.listItem?.title;
    if (listItemTitle) {
      return listItemTitle;
    }

    if (document.title) {
      return document.title.replace(/\s*[|\u2013\u2014-].*$/, '').trim();
    }

    return 'Untitled page';
  }

  private getSessionStorageKey(pagePath: string, source: string): string {
    return `page-view-tracker:${pagePath}:${source}`;
  }

  private hasSessionRecord(pagePath: string, source: string): boolean {
    try {
      return window.sessionStorage.getItem(this.getSessionStorageKey(pagePath, source)) === '1';
    } catch {
      return false;
    }
  }

  private markSessionRecord(pagePath: string, source: string): void {
    try {
      window.sessionStorage.setItem(this.getSessionStorageKey(pagePath, source), '1');
    } catch {
      // Ignore storage failures (privacy mode, blocked storage, etc.)
    }
  }

  private async getPageMetadata(): Promise<IPageMetadata> {
    const listId: string | undefined = this.context.pageContext.list?.id?.toString();
    const itemId: number | undefined = this.context.pageContext.listItem?.id;

    if (!listId || !itemId) {
      return {
        isNewsPage: false,
        hubSection: ''
      };
    }

    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const requestUrl: string =
      `${webUrl}/_api/web/lists(guid'${listId}')/items(${itemId})` +
      `?$select=PromotedState,HubSection`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        requestUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        }
      );

      if (!response.ok) {
        return {
          isNewsPage: false,
          hubSection: ''
        };
      }

      const payload: { PromotedState?: number; HubSection?: string } = await response.json();
      return {
        isNewsPage: payload.PromotedState === NEWS_PROMOTED_STATE,
        hubSection: payload.HubSection || ''
      };
    } catch {
      return {
        isNewsPage: false,
        hubSection: ''
      };
    }
  }

  private async writeTrackingItem(
    item: IListItemPayload,
    properties: Required<IPageViewTrackerApplicationCustomizerProperties>
  ): Promise<void> {
    const webUrl: string = properties.trackingListSiteUrl.replace(/\/$/, '');
    const listTitle: string = encodeURIComponent(properties.trackingListTitle);
    const endpoint: string = `${webUrl}/_api/web/lists/getbytitle('${listTitle}')/items`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: JSON.stringify(item)
      }
    );

    if (!response.ok) {
      const errorBody: string = await response.text();
      throw new Error(`Failed to write tracking item (${response.status}): ${errorBody}`);
    }

    Log.info(LOG_SOURCE, `Tracked view for ${item.PagePath} from ${item.Source}`);
  }
}
