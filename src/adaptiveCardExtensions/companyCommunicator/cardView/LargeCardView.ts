import {
  BaseImageCardView,
  IImageCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton,
  IActionArguments,
  ISubmitActionArguments
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'CompanyCommunicatorAdaptiveCardExtensionStrings';
import { ICompanyCommunicatorAdaptiveCardExtensionProps, ICompanyCommunicatorAdaptiveCardExtensionState, LARGE_QUICK_VIEW_REGISTRY_ID } from '../CompanyCommunicatorAdaptiveCardExtension';

export class LargeCardView extends BaseImageCardView<ICompanyCommunicatorAdaptiveCardExtensionProps, ICompanyCommunicatorAdaptiveCardExtensionState> {
  
  public get data(): IImageCardParameters {
    // a loading view
    if (this.state.currentIndex < 0) {      
      return {        
        primaryText: strings.Loading,
        imageUrl: this.properties.iconProperty,            
      };
    }

    return {         
      primaryText: this.properties.summary ?  
        this.state.messages[this.state.currentIndex].title + "\n\n" + this.state.messages[this.state.currentIndex].summary 
        : this.state.messages[this.state.currentIndex].title,
      imageUrl: this.properties.image ? this.state.messages[this.state.currentIndex].imageLink : "",
    };
  }

  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    const buttons: ICardButton[] = [];

    if (this.state.currentIndex > -1 && this.state.messages.length > 1){      
      // here we add buttons based on where we are in the paging
      if (this.state.currentIndex > 0) {
        buttons.push({
          id: "prev",
          title: strings.PrevButton,
          action: {
            type: "Submit",
            parameters: {
              id: 'prev',
              op: -1 // Decrement the index
            }
          }
        });
      }

      if (this.state.currentIndex < this.properties.count-1) {
        buttons.push({
          id: "next",
          title: strings.NextButton,
          action: {
            type: "Submit",
            parameters: {
              id: 'next',
              op: 1 // Increment the index
            }            
          }
        });
      }

      return buttons as [ICardButton] | [ICardButton, ICardButton];
    }
  }

  public onAction(action: IActionArguments | ISubmitActionArguments): void {
    let submitAction = action as ISubmitActionArguments;
    if (submitAction) {
      const { id, op } = submitAction.data;
      switch (id) {
        case 'prev':
        case 'next':
        this.setState({ currentIndex: this.state.currentIndex + op });
        break;
      }
    }
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    if (this.state.messages?.length > 0) {
      return {
        type: 'QuickView',
        parameters: {
          view: LARGE_QUICK_VIEW_REGISTRY_ID
        }
      };
    }
  }
}
