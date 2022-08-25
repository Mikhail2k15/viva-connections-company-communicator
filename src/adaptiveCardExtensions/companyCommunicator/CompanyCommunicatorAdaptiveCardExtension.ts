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
  title: string;
  description: string;
  iconProperty: string;
  applicationIdUri: string;
  resourceEndpoint: string;
  aiKey: string;
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
      const ai = new AppInsightsTelemetryTracker(this.properties.aiKey);         
      ai.trackEvent(this.context.deviceContext);
      
      try {
        Logger.subscribe(ai);   
      }
      catch {
        console.log("can't initialize logger");
      }  
    }

    if (this.properties.applicationIdUri && this.properties.resourceEndpoint) {
      this.aadClient = await this.context.aadHttpClientFactory.getClient(this.properties.applicationIdUri);
      setTimeout(async () => { await this.fetchData(this.aadClient, this.properties.resourceEndpoint); }, 500);
    }

    this.cardNavigator.register(LARGE_CARD_VIEW_REGISTRY_ID, () => new LargeCardView());
    this.cardNavigator.register(MEDIUM_CARD_VIEW_REGISTRY_ID, () => new MediumCardView());

    this.quickViewNavigator.register(MEDIUM_QUICK_VIEW_REGISTRY_ID, () => new ListQuickView());
    this.quickViewNavigator.register(LARGE_QUICK_VIEW_REGISTRY_ID, () => new DetailsQuickView());

    return Promise.resolve();
  }

  public get title(): string {
    return this.properties.title;
  }

  protected get iconProperty(): string {
    return this.properties.iconProperty || 'megaphone';
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
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    if (propertyPath === 'applicationIdUri' && newValue !== oldValue) { 
      this.aadClient = await this.context.aadHttpClientFactory.getClient(this.properties.applicationIdUri);
    }

    if (((propertyPath === 'resourceEndpoint') || (propertyPath === 'count')) && newValue !== oldValue){
      if (newValue){
        await this.fetchData(this.aadClient, this.properties.resourceEndpoint);
      } else{
        this.setState({messages: []});
      }
    }
  }

  private async fetchData(aadClient: AadHttpClient, resourceEndpoint: string): Promise<void> {
    Logger.log({
      message: "start fetching data",      
      level: LogLevel.Verbose
    });    
    const messagesService = new MessagesService(aadClient, resourceEndpoint);
    const items = await messagesService.getSentMessages();

    const data = items.map(async(item) => {
      const message = await messagesService.getMessage(item.id);
     
        return { 
          id: message.id,
          allUsers: message.allUsers,
          title: message.title,          
          summary: message.summary, 
          imageLink: message.imageLink, 
          author: message.author, 
          buttonTitle: message.buttonTitle, 
          buttonLink: message.buttonLink
        }; 
      
    });

    await Promise.all(data).then((messages: IMessage[]) => {
      if (messages?.length > 0) {
         console.log(messages);
         const orgWideMessages = messages.filter(m => m.allUsers === true);
         const lastMessages = orgWideMessages?.length > this.properties.count ? orgWideMessages.slice(0, this.properties.count) : orgWideMessages;
         this.setState({
            currentIndex: 0,
            messages: lastMessages
          });
      }
      
      Logger.log({
        message: "end fetching data",      
        level: LogLevel.Verbose
      });  
     });
  }
}
