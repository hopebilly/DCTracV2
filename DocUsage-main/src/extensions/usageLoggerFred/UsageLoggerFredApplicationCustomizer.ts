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

    // Start document tracking
    this._initializeDocumentTracking();

    Log.info(
      LOG_SOURCE,
      'Document usage tracking initialized.'
    );

  } catch (error) {

    Log.warn(
      LOG_SOURCE,
      error as any
    );

  }

}


  // ============================================================
  // DOCUMENT TRACKING
  // ============================================================

  private _initializeDocumentTracking(): void {

    document.addEventListener(
      'click',

      (event: MouseEvent) => {

        try {

          const target =
            event.target as HTMLElement;

          if (!target) {
            return;
          }


          // Find the closest clicked link
const anchor =
  target.closest('a') as HTMLAnchorElement;

if (!anchor || !anchor.href) {
  return;
}

console.log(
  'FRED DOCUMENT CLICK:',
  anchor.href,
  anchor.innerText
);

const documentUrl =
  anchor.href;


          // Only continue when the clicked URL is a document
          if (!this._isDocument(documentUrl)) {
            return;
          }


          this._logDocumentInteraction(
            documentUrl
          ).catch((error) => {

            Log.warn(
              LOG_SOURCE,
              `Document logging failed: ${error}`
            );

          });

        } catch (error) {

          Log.warn(
            LOG_SOURCE,
            `Document click detection failed: ${error}`
          );

        }

      },

      true
    );

  }


  // ============================================================
  // CHECK IF CLICKED URL IS A DOCUMENT
  // ============================================================

private _isDocument(url: string): boolean {
  try {
    const lowerUrl = url.toLowerCase();

    const extensions: string[] = [
      '.pdf',
      '.doc',
      '.docx',
      '.xls',
      '.xlsx',
      '.ppt',
      '.pptx',
      '.csv',
      '.txt',
      '.zip'
    ];

    if (extensions.some(ext => lowerUrl.includes(ext))) {
      return true;
    }

    if (
      lowerUrl.includes('/_layouts/15/doc.aspx') ||
      lowerUrl.includes('/_layouts/15/wopiframe.aspx') ||
      lowerUrl.includes('/:w:/') ||
      lowerUrl.includes('/:x:/') ||
      lowerUrl.includes('/:p:/') ||
      lowerUrl.includes('/:b:/')
    ) {
      return true;
    }

    return false;

  } catch {
    return false;
  }
}


  // ============================================================
  // LOG DOCUMENT INTERACTION
  // ============================================================

  private async _logDocumentInteraction(
    documentUrl: string
  ): Promise<void> {

    try {

      // Remove query string
      const parsedUrl =
  new URL(
    documentUrl,
    window.location.origin
  );

let documentPath =
  parsedUrl.searchParams.get('id') ||
  parsedUrl.searchParams.get('SourceUrl') ||
  parsedUrl.searchParams.get('sourceUrl') ||
  parsedUrl.pathname;

try {
  documentPath =
    decodeURIComponent(documentPath);
} catch {
  // keep original
}

let documentName =
  parsedUrl.searchParams.get('file') || '';

if (!documentName) {
  documentName =
    documentPath.substring(
      documentPath.lastIndexOf('/') + 1
    );
}

try {
  documentName =
    decodeURIComponent(documentName);
} catch {
  // keep original
}

const extensionIndex =
  documentName.lastIndexOf('.');

const fileType =
  extensionIndex > -1
    ? documentName
        .substring(extensionIndex + 1)
        .toLowerCase()
    : '';


// Get library name
const webPath =
  new URL(
    this.context.pageContext.web.absoluteUrl
  ).pathname.replace(/\/$/, '');

let relativePath =
  documentPath;

if (
  relativePath
    .toLowerCase()
    .startsWith(
      webPath.toLowerCase() + '/'
    )
) {
  relativePath =
    relativePath.substring(
      webPath.length + 1
    );
}

const pathParts =
  relativePath
    .split('/')
    .filter(part => part.length > 0);

let libraryName = '';

if (
  pathParts.length > 1 &&
  pathParts[0].toLowerCase() !== '_layouts'
) {
  libraryName =
    pathParts[0];
}


      // ============================================================
      // SHAREPOINT CONTEXT
      // ============================================================

      const siteUrl =
        this.context.pageContext.site.absoluteUrl;


      const webUrl =
        this.context.pageContext.web.absoluteUrl;


      // ============================================================
      // USER INFORMATION
      // ============================================================

      const userDisplayName =
        (
          this.context
            .pageContext
            .user
            .displayName || ''
        ).substring(0, 255);


      const userEmail =
        (
          this.context
            .pageContext
            .user
            .email || ''
        ).substring(0, 255);


      // ============================================================
      // SESSION
      // ============================================================

      const sessionId =
        this._ensureSessionId();


      // Page where the user clicked the document
      const sourcePageUrl =
        location.href.split('?')[0];


      // Central logging site
      const targetSite =
        this.properties?.LoggingSiteUrl ||
        webUrl;


      // Document usage list
const listTitle =
  'Document Usage Data';

      // ============================================================
      // DUPLICATE PROTECTION
      // ============================================================
      //
      // SharePoint can trigger multiple browser events from one click.
      // Ignore the same document for 3 seconds.
      // ============================================================

     const duplicateKey =
  `fredDocument:${documentPath}`;

      const previousClick =
        sessionStorage.getItem(
          duplicateKey
        );


      const currentTime =
        Date.now();


      if (previousClick) {

        const difference =
          currentTime -
          Number(previousClick);


        if (difference < 3000) {

          return;

        }

      }


      sessionStorage.setItem(
        duplicateKey,
        currentTime.toString()
      );


      // ============================================================
      // SAVE TO DOCUMENT USAGE DATA
      // ============================================================

      await this._addItem(

        targetSite,

        listTitle,

        {

          DocumentName:
            documentName.substring(
              0,
              255
            ),

         DocumentUrl:
  documentUrl,

LibraryName:
  libraryName,

          SiteUrl:
            siteUrl,

          WebUrl:
            webUrl,

          UserDisplayName:
            userDisplayName,

          UserEmail:
            userEmail,

          FileType:
            fileType,

          ActionType:
            'DocumentClicked',

          SessionId:
            sessionId,

          SourcePageUrl:
            sourcePageUrl,

          // Can be populated in a later version
          DocumentUniqueId:
            '',

          TimeStamp:
            new Date().toISOString()

        }

      );


      Log.info(
        LOG_SOURCE,
        `Document logged: ${documentName}`
      );

    } catch (error) {

      Log.warn(
        LOG_SOURCE,
        `Document interaction error: ${error}`
      );

    }

  }


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