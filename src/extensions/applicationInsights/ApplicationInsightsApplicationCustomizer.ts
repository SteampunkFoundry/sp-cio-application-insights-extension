import { Log } from '@microsoft/sp-core-library';

import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';


import { ApplicationInsights, DistributedTracingModes } from '@microsoft/applicationinsights-web';
import { ITelemetryItem } from '@microsoft/applicationinsights-web';


import * as strings from 'ApplicationInsightsApplicationCustomizerStrings';

const LOG_SOURCE: string = 'ApplicationInsightsApplicationCustomizer';




export interface IAzureAppInsightsWebPartProps {
  enabled: boolean;
  instrumentationKey: string;
  trackUserId: boolean;
  trackExceptions: boolean;
  cloudRole: string;
  cloudRoleInstance: string;
  excludedDependencyTargets: string;
}

const DEFAULT_PROPERTIES: IAzureAppInsightsWebPartProps = {
  enabled: true,
  instrumentationKey: "1a3f9226-224a-4802-9793-ba64f10437ec",
  trackUserId: true,
  trackExceptions: true,
  cloudRole: "sharepoint-page",
  cloudRoleInstance: "",
  excludedDependencyTargets: 'browser.pipe.aria.microsoft.com\nbusiness.bing.com\nmeasure.office.com\nofficeapps.live.com\noutlook.office365.com\noutlook.office.com'
};

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IApplicationInsightsApplicationCustomizerProperties {
  // This is an example; replace with your own property
  key: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ApplicationInsightsApplicationCustomizer
  extends BaseApplicationCustomizer<IApplicationInsightsApplicationCustomizerProperties> {
  private _appInsights: ApplicationInsights = undefined;
  private _config: IAzureAppInsightsWebPartProps = undefined;
  private _excludedDependencies: string[] = [];
  
  private _appInsightsInitializer = (telemetryItem: ITelemetryItem): boolean | void => {
    if (telemetryItem) {
      if (!!this._config.cloudRole) {
        telemetryItem.tags['ai.cloud.role'] = this._config.cloudRole;
      }
      if (!!this._config.cloudRoleInstance) {
        telemetryItem.tags['ai.cloud.roleInstance'] = this._config.cloudRoleInstance;
      }
    }
  }

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    
    this._config = {
      ...DEFAULT_PROPERTIES,
      ...this.properties
    };

    if (!this._config.enabled) {
      console.info(`[AzureAppInsights Web Part] Tracking disabled. No analytics will be logged.`);
      return;
    }

    if (!!this._config.excludedDependencyTargets) {
      this._excludedDependencies = this._config.excludedDependencyTargets.split('\n').filter(n => n && !!n.trim() && n.trim().length > 5);
    }

    const userId: string = this._config.trackUserId ? this.context.pageContext.user.loginName.replace(/([\|:;=])/g, '') : undefined;

    this._appInsights = new ApplicationInsights({
      config: {
        // Instrumentation key that you obtained from the Azure Portal.
        instrumentationKey: this._config.instrumentationKey,

        // An optional account id, if your app groups users into accounts. No spaces, commas, semicolons, equals, or vertical bars
        accountId: userId,

        // If true, Fetch requests are not autocollected. Default is true
        disableFetchTracking: false,

        // If true, AJAX & Fetch request headers is tracked, default is false.
        enableRequestHeaderTracking: true,

        // If true, AJAX & Fetch request's response headers is tracked, default is false.
        enableResponseHeaderTracking: true,

        // Default false. If true, include response error data text in dependency event on failed AJAX requests.
        enableAjaxErrorStatusText: true,

        // Default false. Flag to enable looking up and including additional browser window.performance timings in the reported ajax (XHR and fetch) reported metrics.
        enableAjaxPerfTracking: true,

        // If true, unhandled promise rejections will be autocollected and reported as a javascript error. When disableExceptionTracking is true (dont track exceptions) the config value will be ignored and unhandled promise rejections will not be reported.
        enableUnhandledPromiseRejectionTracking: true,

        // If true, the SDK will add two headers ('Request-Id' and 'Request-Context') to all CORS requests tocorrelate outgoing AJAX dependencies with corresponding requests on the server side. Default is false
        enableCorsCorrelation: true,

        // If true, exceptions are not autocollected. Default is false.
        disableExceptionTracking: !this._config.trackExceptions,

        // Sets the distributed tracing mode. If AI_AND_W3C mode or W3C mode is set, W3C trace context headers (traceparent/tracestate) will be generated and included in all outgoing requests. AI_AND_W3C is provided for back-compatibility with any legacy Application Insights instrumented services.
        distributedTracingMode: DistributedTracingModes.AI_AND_W3C
      }
    });

    this._appInsights.loadAppInsights();
//    this._appInsights.addTelemetryInitializer(this._appInsightsInitializer);
    this._appInsights.setAuthenticatedUserContext(userId, userId, true);
  //  this._appInsights.trackPageView();

    this._appInsights.trackPageView({
      name: document.title, uri: window.location.href,
      properties: {
        ["CustomProps"]: {
          WebAbsUrl: this.context.pageContext.web.absoluteUrl,
          WebSerUrl: this.context.pageContext.web.serverRelativeUrl,
          WebId: this.context.pageContext.web.id,
          UserTitle: this.context.pageContext.user.displayName,
          UserEmail: this.context.pageContext.user.email,
          UserLoginName: this.context.pageContext.user.loginName
        }
      }
    });
    
  
    return Promise.resolve();
  }
}
