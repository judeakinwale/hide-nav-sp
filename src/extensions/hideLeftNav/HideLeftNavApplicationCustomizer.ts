import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";
import { Dialog } from "@microsoft/sp-dialog";

import * as strings from "HideLeftNavApplicationCustomizerStrings";

const LOG_SOURCE: string = "HideLeftNavApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHideLeftNavApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
  isWorking: boolean;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HideLeftNavApplicationCustomizer extends BaseApplicationCustomizer<IHideLeftNavApplicationCustomizerProperties> {
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let message: string = this.properties.testMessage;
    if (!message) {
      message = "(No properties were provided.)";
    }
    const workingText = this.properties.isWorking
      ? "Customizer Working!!"
      : "Customizer not Working ;(";

    //  Inject custom CSS to hide the left navigation
    const css: string = `
        #spLeftNav {
          display: none !important;
        }

        #spCommandBar {
          display: none !important;
        }
    `;
    // ? this was adding style tags inside style tags in the header
    // const css: string = `
    //   <style>
    //     #spLeftNav {
    //       display: block !important;
    //     }

    //     #spCommandBar {
    //       display: block !important;
    //     }
    //   </style>
    // `;

    // Append the custom CSS to the head of the page
    const head =
      document.getElementsByTagName("head")[0] || document.documentElement;
    const styleElement = document.createElement("style");
    styleElement.innerHTML = css;
    head.appendChild(styleElement);

    // uncomment this to show an alert once the customizer is loaded
    // Dialog.alert(
    //   `Hello from ${strings.Title}:\n\n${message}. ${workingText}`
    // ).catch(() => {
    //   /* handle error */
    // });

    /**
     * ! the params added to the site / page should look like this
     *
     * ?debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js&loadSPFX=true&customActions={"fb2dd0a1-960e-4a25-bccd-b38dc89b799b"%3A{"location"%3A"ClientSideExtension.ApplicationCustomizer"}}
     * ?debugManifestsFile=https://localhost:4321/temp/manifests.js&loadSPFX=true&customActions={"fb2dd0a1-960e-4a25-bccd-b38dc89b799b":{"location":"ClientSideExtension.ApplicationCustomizer"}}
     *
     * this includes the properties set in the serve.json in config
     * ?debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js&loadSPFX=true&customActions={"fb2dd0a1-960e-4a25-bccd-b38dc89b799b"%3A{"location"%3A"ClientSideExtension.ApplicationCustomizer"%2C"properties"%3A{"testMessage"%3A"Test+message"}}}
     * ?debugManifestsFile=https://localhost:4321/temp/manifests.js&loadSPFX=true&customActions={"fb2dd0a1-960e-4a25-bccd-b38dc89b799b":{"location":"ClientSideExtension.ApplicationCustomizer", "properties":{"testMessage"%3A"Test+message"}}}
     */

    return Promise.resolve();
  }
}

// gulp build
// gulp bundle --ship
// gulp package-solution --ship
// gulp serve
