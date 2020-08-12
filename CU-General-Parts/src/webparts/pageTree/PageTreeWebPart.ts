import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'PageTreeWebPartStrings';
import PageTree, { IPageTreeProps } from './components/PageTree';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';


export interface IPageTreeWebPartProps {
  rootid: number;
  expandlevels: number;
}

export default class PageTreeWebPart extends BaseClientSideWebPart<IPageTreeWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPageTreeProps> = React.createElement(PageTree, { rootid: this.properties.rootid, context: this.context, expandtolevel: this.properties.expandlevels });

    ReactDom.render(element, this.domElement);
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
                PropertyFieldNumber("rootid", {
                  key: "rootid",
                  label: strings.DescriptionFieldLabel,
                  value: this.properties.rootid,
                  minValue: 1,
                  disabled: false
                }),
                PropertyFieldNumber("expandlevels", {
                  key: "expandlevels",
                  label: strings.ExpandLevel,
                  value: this.properties.expandlevels,
                  minValue: 1,
                  disabled: false
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
