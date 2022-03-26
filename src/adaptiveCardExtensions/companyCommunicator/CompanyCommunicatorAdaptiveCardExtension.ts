import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { CompanyCommunicatorPropertyPane } from './CompanyCommunicatorPropertyPane';

export interface ICompanyCommunicatorAdaptiveCardExtensionProps {
  title: string;
}

export interface ICompanyCommunicatorAdaptiveCardExtensionState {
}

const CARD_VIEW_REGISTRY_ID: string = 'CompanyCommunicator_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'CompanyCommunicator_QUICK_VIEW';

export default class CompanyCommunicatorAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ICompanyCommunicatorAdaptiveCardExtensionProps,
  ICompanyCommunicatorAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: CompanyCommunicatorPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = { };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'CompanyCommunicator-property-pane'*/
      './CompanyCommunicatorPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.CompanyCommunicatorPropertyPane();
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
