import { IPropertyPaneConfiguration, PropertyPaneCheckbox, PropertyPaneSlider, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'CompanyCommunicatorAuthorAdaptiveCardExtensionStrings';

export class CompanyCommunicatorAuthorPropertyPane {
  public getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                })
              ]
            },
            {
              groupName: 'Company Communicator',
              groupFields: [
                PropertyPaneTextField('applicationIdUri', {
                  label: 'Application ID URI'
                }),
                PropertyPaneTextField('resourceEndpoint', {
                  label: 'Resource endpoint'
                }),                
                PropertyPaneTextField('userappid', {
                  label: "Company Communicator Teams User App Id"
                }),
                PropertyPaneCheckbox('teamsDeepLink', {
                  text: "Deep Link to Teams"
                })    
              ]
            },
            {
              groupName: strings.AppInsightsFieldsGroupName,
              groupFields: [
                PropertyPaneTextField('aiKey', {
                  label: strings.AppInsightsInstrumentationKeyFieldLabel
                }),
                PropertyPaneTextField('aiAppId', {
                  label: strings.AppInsightsApplicationIDFieldLabel
                }),
                PropertyPaneTextField('aiAppKey', {
                  label: strings.AppInsightsAPIKeyFieldLabel
                }) 
              ]
            }
          ]
        }
      ]
    };
  }
}
