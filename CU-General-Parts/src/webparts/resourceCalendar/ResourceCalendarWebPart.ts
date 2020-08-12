import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'ResourceCalendarWebPartStrings';
import ResourceCalendar, { IResourceCalendarProps } from './components/ResourceCalendar';
import { PropertyFieldPeoplePicker, PrincipalType, IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';



export interface IResourceCalendarWebPartProps {
  resources: IPropertyFieldGroupOrPerson[];
  mode: string;
}

export default class ResourceCalendarWebPart extends BaseClientSideWebPart<IResourceCalendarWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IResourceCalendarProps> = React.createElement(
      ResourceCalendar,
      {
        context: this.context,
        resources: this.properties.resources,
        mode: this.properties.mode
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
                PropertyFieldPeoplePicker('resources', {
                  label: strings.DescriptionFieldLabel,
                  initialData: this.properties.resources,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users ],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'resources'
                }),
                PropertyPaneDropdown('mode', {
                  options: [{ text: 'Vertical', key: 'vertical' }, { text: 'Horizontal', key: 'horizontal' }],
                  label: "View Mode"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
