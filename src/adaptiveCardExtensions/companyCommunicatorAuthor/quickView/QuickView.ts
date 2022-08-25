import { ISPFxAdaptiveCard, BaseAdaptiveCardView } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'CompanyCommunicatorAuthorAdaptiveCardExtensionStrings';
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
    const desktopPercent = (+this.state.desktop /(+msteams)) * 100;
    const mobilePercent = (+this.state.mobile /(+msteams)) * 100;
    const webPercent = (+this.state.web /(+msteams)) * 100;
    return {
      subTitle: strings.SubTitle,
      title: strings.Title,
      monthly: monthly.toString(),
      msteams: `${msteamsPercent.toFixed(0)} %`,
      spo: `${(100 - msteamsPercent).toFixed(0)} %`,
      desktop: `${desktopPercent.toFixed(0)} %`,
      mobile: `${mobilePercent.toFixed(0)} %`,
      web: `${webPercent.toFixed(0)} %`,
      items: this.state.messages 
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}