import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'AcesCreateNewItemAdaptiveCardExtensionStrings';

export class AcesCreateNewItemPropertyPane {
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
                }),
                PropertyPaneTextField('subTitle', {
                  label: strings.SubTitle
                }),
                PropertyPaneTextField('siteTitle', {
                  label: "Site Title"
                }),
                PropertyPaneTextField('listTitle', {
                  label: "List Title"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
