import { BaseAdaptiveCardView, ISPFxAdaptiveCard } from "@microsoft/sp-adaptive-card-extension-base";
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
        const message = this.state.messages[this.state.currentIndex];
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
        let card: ISPFxAdaptiveCard =  require('./template/DetailsQuickViewTemplate.json');
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
    
      /*public onAction(action: IActionArguments): void {
        if (action.type === 'Submit') {
          const { id, op } = action.data;
          switch (id) {
            case 'prev':
            case 'next':
            this.setState({ currentIndex: this.state.currentIndex + op });
            break;
          }
        }
      }*/ 
}