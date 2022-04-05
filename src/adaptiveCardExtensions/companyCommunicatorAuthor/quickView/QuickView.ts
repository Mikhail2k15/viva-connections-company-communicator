import { ISPFxAdaptiveCard, BaseAdaptiveCardView } from '@microsoft/sp-adaptive-card-extension-base';
import { Logger, LogLevel } from '@pnp/logging';
import * as strings from 'CompanyCommunicatorAuthorAdaptiveCardExtensionStrings';
import AppInsightsAnalyticsService from '../../../service/analytics/AppInsightsAnalyticsService';
import { TimeSpan } from '../../../service/analytics/TimeSpan';
import VivaConnectionsInsights from '../../../service/analytics/VivaConnectionsInsights';
import { IMessageDetails } from '../../../service/messages/IMessage';
import { ICompanyCommunicatorAuthorAdaptiveCardExtensionProps, ICompanyCommunicatorAuthorAdaptiveCardExtensionState } from '../CompanyCommunicatorAuthorAdaptiveCardExtension';

export interface IQuickViewData {
  items: IMessageDetails[];
  subTitle: string;
  title: string;
  monthly: string;
  msteams: string;
  spo: string;
  desktop: string;
  mobile: string;
  web: string;
}

export class QuickView extends BaseAdaptiveCardView<
  ICompanyCommunicatorAuthorAdaptiveCardExtensionProps,
  ICompanyCommunicatorAuthorAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    console.log('QuickView.data');
    
    const monthly: number = this.state.monthly;
    const msteams: number = +this.state.desktop + +this.state.mobile + +this.state.web;    
    const spo: number = this.state.spo;
    const msteamsPercent = (msteams /(+msteams + +spo)) * 100;
    return {
      subTitle: strings.SubTitle,
      title: strings.Title,
      monthly: monthly.toString(),
      msteams: `${msteamsPercent.toFixed(0)} % (${msteams})`,
      spo: `${(100 - msteamsPercent).toFixed(0)} % (${spo})`,
      desktop: this.state.desktop.toString(),
      mobile: this.state.mobile.toString(),
      web: this.state.web.toString(),
      items: this.state.messages 
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}