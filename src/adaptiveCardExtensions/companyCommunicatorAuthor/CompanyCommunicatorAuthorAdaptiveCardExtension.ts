import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { CompanyCommunicatorAuthorPropertyPane } from './CompanyCommunicatorAuthorPropertyPane';

export interface ICompanyCommunicatorAuthorAdaptiveCardExtensionProps {
  title: string;
}

export interface ICompanyCommunicatorAuthorAdaptiveCardExtensionState {
}

const CARD_VIEW_REGISTRY_ID: string = 'CompanyCommunicatorAuthor_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'CompanyCommunicatorAuthor_QUICK_VIEW';

export default class CompanyCommunicatorAuthorAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ICompanyCommunicatorAuthorAdaptiveCardExtensionProps,
  ICompanyCommunicatorAuthorAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: CompanyCommunicatorAuthorPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = { };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'CompanyCommunicatorAuthor-property-pane'*/
      './CompanyCommunicatorAuthorPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.CompanyCommunicatorAuthorPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }
}
