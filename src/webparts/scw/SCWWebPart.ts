import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SCWWebPartStrings';
import SCW from './components/SCW';
import { ISCWProps } from './components/ISCWProps';
import { sp } from "@pnp/sp";
import { graph } from "@pnp/graph/presets/all";
import O365Service from '../../services/O365Service';

export interface ISCWWebPartProps {
  listTemplate: string;
  prefLang: string;
}

export default class SCWWebPart extends BaseClientSideWebPart<ISCWWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISCWProps> = React.createElement(
      SCW,
      {
        listTemplate: this.properties.listTemplate,
        context: this.context,
        prefLang: this.properties.prefLang,
      }
    );
    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {
    return super.onInit().then( _ => {
      sp.setup({
        spfxContext: this.context
      });

      O365Service.setup(this.context);
      graph.setup(this.context);
      const r =  async () => {(await graph.groups()).length};
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('listTemplate', {
                  label: strings.ListTemplateFieldLabel
                }),
                PropertyPaneDropdown('prefLang', {
                  label: 'Preferred Language',
                  options: [
                    { key: 'account', text: 'Account' },
                    { key: 'en-us', text: 'English' },
                    { key: 'fr-fr', text: 'Français' }
                  ]}),
              ]
            }
          ]
        }
      ]
    };
  }
}
