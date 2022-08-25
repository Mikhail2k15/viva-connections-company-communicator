import { ISPFxAdaptiveCard, BaseAdaptiveCardView } from '@microsoft/sp-adaptive-card-extension-base';
import { Logger, LogLevel } from '@pnp/logging';
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
    const userId = this.context.pageContext.aadInfo.userId._guid;
    this.state.messages.forEach(message => {
      const trackInfo = {
        notificationId: message.id,
        userId: userId,
        quickView: "ListQuickView"
      };
      Logger.log({
        message: "TrackView",
        data: trackInfo,
        level: LogLevel.Info
      });
    });
    
    return {
      items: this.state.messages
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/ListQuickViewTemplate.json');
  }
}