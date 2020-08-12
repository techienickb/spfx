import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'PageBreadcrumbsWebPartStrings';
import PageBreadcrumbs, { IPageBreadcrumbsProps } from './components/PageBreadcrumbs';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';

export interface IPageBreadcrumbsWebPartProps {
  context: WebPartContext;
  rootid: number;
}

export default class PageBreadcrumbsWebPart extends BaseClientSideWebPart<IPageBreadcrumbsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPageBreadcrumbsProps> = React.createElement(
      PageBreadcrumbs,
      {
        context: this.context,
        root: this.properties.rootid
      }
    );

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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
