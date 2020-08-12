import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { PropertyFieldPeoplePicker, PrincipalType, IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'GroupMembersWebPartStrings';
import GroupMembers, { IGroupMembersProps } from './components/GroupMembers';

export interface IGroupMembersWebPartProps {
  groups: IPropertyFieldGroupOrPerson[];
  ignorePeople: IPropertyFieldGroupOrPerson[];
}

export default class GroupMembersWebPart extends BaseClientSideWebPart<IGroupMembersWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IGroupMembersProps> = React.createElement(GroupMembers, {groups: this.properties.groups, ignorePeople: this.properties.ignorePeople, context: this.context});

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
                PropertyFieldPeoplePicker('groups', {
                  label: "Select a group, please don't select a person",
                  initialData: this.properties.groups,
                  allowDuplicate: false,
                  principalType: [PrincipalType.SharePoint, PrincipalType.Security],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'groups'
                }),
                PropertyFieldPeoplePicker('ignorePeople', {
                  label: "Select any users you wish to ignore",
                  initialData: this.properties.ignorePeople,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'ignorePeople'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
