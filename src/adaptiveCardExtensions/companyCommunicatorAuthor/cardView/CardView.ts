import {
  BaseBasicCardView,
  IBasicCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'CompanyCommunicatorAuthorAdaptiveCardExtensionStrings';
import { ICompanyCommunicatorAuthorAdaptiveCardExtensionProps, ICompanyCommunicatorAuthorAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../CompanyCommunicatorAuthorAdaptiveCardExtension';

export class CardView extends BaseBasicCardView<ICompanyCommunicatorAuthorAdaptiveCardExtensionProps, ICompanyCommunicatorAuthorAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
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

  public get data(): IBasicCardParameters {
    return {
      primaryText: strings.PrimaryText,
      title: this.properties.title
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'ExternalLink',
      parameters: {
        target: 'https://www.bing.com'
      }
    };
  }
}
