import { ISPFxAdaptiveCard, BaseAdaptiveCardView, IActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'CompanyCommunicatorAdaptiveCardExtensionStrings';
import { IMessage } from '../../../service/messages/IMessage';
import { ICompanyCommunicatorAdaptiveCardExtensionProps, ICompanyCommunicatorAdaptiveCardExtensionState } from '../CompanyCommunicatorAdaptiveCardExtension';

export interface IListQuickViewData {
  items: IMessage[];
}

export class ListQuickView extends BaseAdaptiveCardView<
  ICompanyCommunicatorAdaptiveCardExtensionProps,
  ICompanyCommunicatorAdaptiveCardExtensionState,
  IListQuickViewData
> {
  public get data(): IListQuickViewData {
    return {
      items: this.state.messages
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/ListQuickViewTemplate.json');
  }
}