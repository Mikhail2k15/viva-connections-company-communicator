import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { AadHttpClient } from '@microsoft/sp-http';

import { CompanyCommunicatorPropertyPane } from './CompanyCommunicatorPropertyPane';
import { IMessage } from '../../service/messages/IMessage';
import { MessagesService } from '../../service/messages/MessagesService';

import { LargeCardView } from './cardView/LargeCardView';
import { MediumCardView } from './cardView/MediumCardView';
import { ListQuickView } from './quickView/ListQuickView';
import { DetailsQuickView } from './quickView/DetailsQuickView';
import { Logger, LogLevel } from '@pnp/logging';
import { AppInsightsTelemetryTracker } from '../../service/analytics/AppInsightsTelemetryTracker';

export interface ICompanyCommunicatorAdaptiveCardExtensionProps {
  aiKey: string;
  title: string;
  description: string;
  iconProperty: string;
  applicationIdUri: string;
  resourceEndpoint: string;
  count: number;
  image: boolean;
  summary: boolean;
}

export interface ICompanyCommunicatorAdaptiveCardExtensionState {
  currentIndex: number;
  messages: IMessage[];
}

const LARGE_CARD_VIEW_REGISTRY_ID: string = 'CompanyCommunicator_LARGE_CARD_VIEW';
const MEDIUM_CARD_VIEW_REGISTRY_ID: string = 'CompanyCommunicator_MEDIUM_CARD_VIEW';

export const LARGE_QUICK_VIEW_REGISTRY_ID: string = 'CompanyCommunicator_LARGE_QUICK_VIEW';
export const MEDIUM_QUICK_VIEW_REGISTRY_ID: string = 'CompanyCommunicator_MEDIUM_QUICK_VIEW';

export default class CompanyCommunicatorAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ICompanyCommunicatorAdaptiveCardExtensionProps,
  ICompanyCommunicatorAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: CompanyCommunicatorPropertyPane | undefined;
  private aadClient: AadHttpClient;

  public async onInit(): Promise<void> {
    this.state = { 
      currentIndex: -1,
      messages: []
    };
    Logger.activeLogLevel = LogLevel.Verbose;

    if (this.properties.aiKey) {
      Logger.log({
        message: "Try to init AppInsights tracker",
        data: { aiKey: this.properties.aiKey },
        level: LogLevel.Verbose
      });
      let ai = new AppInsightsTelemetryTracker(this.properties.aiKey);         
      ai.trackEvent(this.context.deviceContext);
      
      try{
        
        Logger.subscribe(ai);   
      }
      catch {} 
    }

    this.cardNavigator.register(LARGE_CARD_VIEW_REGISTRY_ID, () => new LargeCardView());
    this.cardNavigator.register(MEDIUM_CARD_VIEW_REGISTRY_ID, () => new MediumCardView());

    this.quickViewNavigator.register(MEDIUM_QUICK_VIEW_REGISTRY_ID, () => new ListQuickView());
    this.quickViewNavigator.register(LARGE_QUICK_VIEW_REGISTRY_ID, () => new DetailsQuickView());

    if (this.properties.applicationIdUri && this.properties.resourceEndpoint) {
      this.aadClient = await this.context.aadHttpClientFactory.getClient(this.properties.applicationIdUri);
      setTimeout(()=> { this.fetchData(this.aadClient, this.properties.resourceEndpoint); }, 500);
    }

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
    return this.cardSize === 'Medium' ? MEDIUM_CARD_VIEW_REGISTRY_ID: LARGE_CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    if (propertyPath === 'applicationIdUri' && newValue !== oldValue) { 
      this.aadClient = await this.context.aadHttpClientFactory.getClient(this.properties.applicationIdUri);
    }

    if (((propertyPath === 'resourceEndpoint') || (propertyPath === 'count')) && newValue !== oldValue){
      if (newValue){
        this.fetchData(this.aadClient, this.properties.resourceEndpoint);
      } else{
        this.setState({messages: []});
      }
    }
  }

  private async fetchData(aadClient: AadHttpClient, resourceEndpoint: string) {
    Logger.log({
      message: "start fetching data",      
      level: LogLevel.Verbose
    });    
    const messagesService = new MessagesService(aadClient, resourceEndpoint);
    const items = await messagesService.getSentMessages();

    const data = items.map(async(item) => {
      const message = await messagesService.getMessage(item.id);
      if (message.allUsers === true) {
        return { 
          title: message.title, 
          id: message.id,
          summary: message.summary, 
          imageLink: message.imageLink, 
          author: message.author, 
          buttonTitle: message.buttonTitle, 
          buttonLink: message.buttonLink
        }; 
      }
    });

    Promise.all(data).then((messages: IMessage[]) => {
      const lastMessages = messages?.length > this.properties.count ? messages.slice(0, this.properties.count) : messages;
      this.setState({
        currentIndex: 0,
        messages: lastMessages
      });
      Logger.log({
        message: "end fetching data",      
        level: LogLevel.Verbose
      });  
     });
  }
}
