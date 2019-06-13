import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'SpfxExtensionApplicationCustomizerStrings';

import * as $ from 'jquery';

const LOG_SOURCE: string = 'SpfxExtensionApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpfxExtensionApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpfxExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<ISpfxExtensionApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }

    let message2: string = "no placeholders";
    message2 = this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(", ");

    //Dialog.alert(`test`);
    //Dialog.alert(`QueryParam: ${strings.Title}:\n\n${message}`);
    //Dialog.alert(`Title:${strings.Title}    QueryParam:${message}    Available Place Holders:${message2}`);

    alert( $(`QueryParam:${message}    Available Place Holders:${message2}`).val() );

    return Promise.resolve();
  }
}
