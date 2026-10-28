import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';


export interface IUsageLoggerFredApplicationCustomizerProperties {
  LoggingSiteUrl?: string;
  ListTitle?: string;
  ThrottleMs?: number;
}


const LOG_SOURCE: string = 'UsageLoggerFredApplicationCustomizer';


export default class UsageLoggerFredApplicationCustomizer
  extends BaseApplicationCustomizer<IUsageLoggerFredApplicationCustomizerProperties> {


  // ============================================================
  // INITIALIZATION
  // ============================================================

public async onInit(): Promise<void> {

  try {

    console.log('FRED DOCUMENT LOGGER LOADED');

    // Keep document tracking running
    //this._initializeDocumentTracking();


    // ============================================================
    // DEMO / EXTENSION LOAD RECORD
    // ============================================================

    const targetSite =
      this.properties?.LoggingSiteUrl ||
      this.context.pageContext.web.absoluteUrl;


    await this._addItem(
      targetSite,
      'Document Usage Data',
      {

        DocumentName:
          'TEST - Extension Loaded',

        DocumentUrl:
          location.href,

        LibraryName:
          'Fred Document Tracking',

        SiteUrl:
          this.context.pageContext.site.absoluteUrl,

        WebUrl:
          this.context.pageContext.web.absoluteUrl,

        UserDisplayName:
          this.context.pageContext.user.displayName || '',

        UserEmail:
          this.context.pageContext.user.email || '',

        FileType:
          'test',

        ActionType:
          'ExtensionLoaded',

        SessionId:
          this._ensureSessionId(),

        SourcePageUrl:
          location.href,

        DocumentUniqueId:
          '',

        TimeStamp:
          new Date().toISOString()

      }
    );


    console.log(
      'FRED DEMO DOCUMENT USAGE RECORD CREATED'
    );


    Log.info(
      LOG_SOURCE,
      'Document usage tracking initialized.'
    );


  } catch (error) {

    console.error(
      'FRED Document logger error:',
      error
    );


    Log.warn(
      LOG_SOURCE,
      error as any
    );

  }

}


  // ============================================================
  // CHECK IF CLICKED URL IS A DOCUMENT
  // ============================================================




  // ============================================================
  // LOG DOCUMENT INTERACTION
  // ============================================================




  // ============================================================
  // SESSION ID
  // ============================================================

  private _ensureSessionId(): string {

    let id =
      sessionStorage.getItem(
        'usageSessionId'
      );


    if (!id) {

      id =
        Math.random()
          .toString(36)
          .slice(2)
        +
        Date.now()
          .toString(36);


      sessionStorage.setItem(
        'usageSessionId',
        id
      );

    }


    return id;

  }


  // ============================================================
  // SHAREPOINT LIST WRITER
  // ============================================================

  private async _addItem(
    targetSiteUrl: string,
    listTitle: string,
    payload: any
  ): Promise<void> {

    const endpoint =
      `${targetSiteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items`;


    const resp: SPHttpClientResponse =
      await this.context.spHttpClient.post(

        endpoint,

        SPHttpClient.configurations.v1,

        {

          headers: {

            'Accept':
              'application/json;odata=nometadata',

            'Content-type':
              'application/json;odata=nometadata'

          },

          body:
            JSON.stringify(payload)

        }

      );


    if (!resp.ok) {

      const text =
        await resp.text();


      throw new Error(
        `Document usage log failed: ` +
        `${resp.status} ` +
        `${resp.statusText} - ` +
        `${text}`
      );

    }

  }

}