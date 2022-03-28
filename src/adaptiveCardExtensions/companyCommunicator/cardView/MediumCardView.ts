import { BaseBasicCardView, IBasicCardParameters, ICardButton, IExternalLinkCardAction, IQuickViewCardAction } from "@microsoft/sp-adaptive-card-extension-base";
import * as strings from "CompanyCommunicatorAdaptiveCardExtensionStrings";
import { ICompanyCommunicatorAdaptiveCardExtensionProps, ICompanyCommunicatorAdaptiveCardExtensionState, MEDIUM_QUICK_VIEW_REGISTRY_ID } from "../CompanyCommunicatorAdaptiveCardExtension";

export class MediumCardView extends BaseBasicCardView<ICompanyCommunicatorAdaptiveCardExtensionProps, ICompanyCommunicatorAdaptiveCardExtensionState> {
  
  public get data(): IBasicCardParameters {
    // a loading view
    if (this.state.currentIndex < 0) {      
      return {        
        primaryText: strings.Loading,           
      };
    } 
    const primaryText: string = this.state.messages?.length > 0 ? strings.MediumCardWelcomeMessage : strings.NoMessages;
    return {         
        primaryText: primaryText 
    };
  }

  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    if (this.state.messages?.length > 0) {
      return [
        {
          title: strings.SeeAll,
          action: {
            type: "QuickView",
            parameters: {
              view: MEDIUM_QUICK_VIEW_REGISTRY_ID
            }
          }
        }
      ];
    }    
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    if (this.state.messages?.length > 0) {
      return {
          type: 'QuickView',
          parameters: {
              view: MEDIUM_QUICK_VIEW_REGISTRY_ID
          }
      };
    }
  }
}