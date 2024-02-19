import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'MmsJsInjectionApplicationCustomizerStrings';

const LOG_SOURCE: string = 'MmsJsInjectionApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IMmsJsInjectionApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}




/** A Custom Action which can be run during execution of a Client Side Application */
export default class MmsJsInjectionApplicationCustomizer
  extends BaseApplicationCustomizer<IMmsJsInjectionApplicationCustomizerProperties> {

    private _externalJsUrl: string = "https://tsmms.sharepoint.com/sites/AnalyticsTest/Style%20Library/custom-script.js";

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    console.log(`MmsJsInjectionApplicationCustomizer.onInit(): Entered.`);

    const scriptTag: HTMLScriptElement = document.createElement("script");
    scriptTag.src = this._externalJsUrl;
    scriptTag.type = "text/javascript";
    document.getElementsByTagName("head")[0].appendChild(scriptTag);

    console.log(`MmsJsInjectionApplicationCustomizer.onInit(): Added script link.`);

    return Promise.resolve();
  }
}
