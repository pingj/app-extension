import { escape } from 'lodash';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderName,
  PlaceholderContent
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'HelloWorldApplicationCustomizerStrings';
import styles from './AppCustomizer.module.scss';

const LOG_SOURCE: string = 'HelloWorldApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHelloWorldApplicationCustomizerProperties {
  // This is an example; replace with your own property
  bottom: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HelloWorldApplicationCustomizer
  extends BaseApplicationCustomizer<IHelloWorldApplicationCustomizerProperties> {
  private _bottomPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Added to handle possible changes on the existence of placeholders.
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    // Call render method for generating the HTML elements.
    this._renderPlaceHolders();

    return Promise.resolve();
  }

  private _renderPlaceHolders() {
    console.log("render place holder");

    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        {
          onDispose: this._onDispose
        }
      );
    }

    // The extension should not assume that the expected placeholder is available.
    if (!this._bottomPlaceholder) {
      console.error('The expected placeholder (Bottom) was not found.');
      return;
    }

    if (this.properties) {
      if (this._bottomPlaceholder.domElement) {
        this._bottomPlaceholder.domElement.innerHTML = `
        <div class="${styles.app}">
          <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.bottom}">
            <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape(this.properties.bottom)}
          </div>
        </div>`;
      }
    }
  }

  private _onDispose() {
    // unmount footer
    console.log('place holder is disposed');
  }
}
