import { BaseAdaptiveCardView, ISPFxAdaptiveCard } from "@microsoft/sp-adaptive-card-extension-base";
import { Logger, LogLevel } from "@pnp/logging";
import { ICompanyCommunicatorAdaptiveCardExtensionProps, ICompanyCommunicatorAdaptiveCardExtensionState } from "../CompanyCommunicatorAdaptiveCardExtension";

export interface IDetailsQuickViewData {
    title: string;
    summary: string;
    imageLink: string;
    author?: string;
    buttonTitle?: string;
    buttonLink?: string;
}

export class DetailsQuickView extends BaseAdaptiveCardView<ICompanyCommunicatorAdaptiveCardExtensionProps,
  ICompanyCommunicatorAdaptiveCardExtensionState,
  IDetailsQuickViewData> {
    public get data(): IDetailsQuickViewData {
        console.log('DetailsQuickView:data()');
        const message = this.state.messages[this.state.currentIndex];
                 
        const trackInfo = {
            notificationId: message.id,
            userId: this.context.pageContext.aadInfo.userId._guid,
            quickView: "DetailsQuickView"
        };
        Logger.log({
          message: "TrackView",
          data: trackInfo,
          level: LogLevel.Info
        });

        return {
          title: message.title,
          summary: message.summary,
          imageLink: message.imageLink,
          author: message.author,
          buttonTitle: message.buttonTitle,
          buttonLink: message.buttonLink,
        };
    }    
    
    public get template(): ISPFxAdaptiveCard {
        const card: ISPFxAdaptiveCard = require('./template/DetailsQuickViewTemplate.json');
        const message = this.state.messages[this.state.currentIndex];
        if (message.buttonLink){
            card.actions = [
            {
                "id": "1",
                "type": "Action.OpenUrl",
                "title": message.buttonTitle,
                "url": message.buttonLink                
            }
        ];
        } else {
            delete card.actions;
        }
        return card;
    }
}