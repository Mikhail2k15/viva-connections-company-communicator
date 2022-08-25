import {
  BaseBasicCardView,
  IBasicCardParameters,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'CompanyCommunicatorAuthorAdaptiveCardExtensionStrings';
import { ICompanyCommunicatorAuthorAdaptiveCardExtensionProps, ICompanyCommunicatorAuthorAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../CompanyCommunicatorAuthorAdaptiveCardExtension';

export class CardView extends BaseBasicCardView<ICompanyCommunicatorAuthorAdaptiveCardExtensionProps, ICompanyCommunicatorAuthorAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    if (this.state.messages && this.state.messages.length > 0 
      && this.state.messages[0].status === "Sent") {
      return [
        {
          title: strings.QuickViewButton,
          action: {
            type: 'QuickView',
            parameters: {
              view: QUICK_VIEW_REGISTRY_ID
            }
          }
        }
      ];
    }
  }

  public get data(): IBasicCardParameters {
    let primaryText: string = "loading...";
    if (this.state.messages && this.state.messages.length > 0) {
      primaryText = `The last message sending status: ${this.state.messages[0].formattedStatus}`;
    }
    return {
      primaryText: primaryText, //The last message was successfully delivered",
      title: this.properties.title
    };
  }
}
