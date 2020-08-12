import * as React from 'react';
import styles from './GroupMembers.module.scss';
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { MSGraphClient } from '@microsoft/sp-http';
import SPFxPeopleCard, { IPeopleCardProps } from './SPFxPeopleCard';
import { PersonaSize, PersonaInitialsColor, IPersonaSharedProps } from 'office-ui-fabric-react';

export interface IGroupMembersProps {
  groups: IPropertyFieldGroupOrPerson[];
  ignorePeople: IPropertyFieldGroupOrPerson[];
  context: WebPartContext;
}

export default class GroupMembers extends React.Component<IGroupMembersProps, { people: IPeopleCardProps[] }> {
  public state = { people: [] };

  public componentDidMount() {
    this.load(this.props.groups, this.props.ignorePeople);
  }

  public componentWillReceiveProps(nextProps: Readonly<IGroupMembersProps>, nextContext: any): void {
    this.load(nextProps.groups, nextProps.ignorePeople);
  }

  protected load = (groups: IPropertyFieldGroupOrPerson[], ignorePeople: IPropertyFieldGroupOrPerson[]): void => {
    this.setState({ people: [] });
    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      groups.forEach(g => {
        client.api(`/groups/${g.id.split('|')[2]}/members`).get((err, res: any) => {
          if (res !== null) {
            let r: MicrosoftGraph.User[] = res.value;
            if (!ignorePeople) ignorePeople = [];
            let p: IPeopleCardProps[] = r.filter(u => ignorePeople.filter(i => i.email.toLowerCase() === u.mail.toLowerCase()).length === 0).map(r1 => ({ primaryText: r1.displayName, secondaryText: r1.jobTitle, email: r1.userPrincipalName, serviceScope: this.props.context.serviceScope, 
              class: styles.personaCard, size: PersonaSize.regular  }  ));
            this.setState({ people: this.state.people.concat(p) });
          }
        });
      });
    });
  }
  
  public render(): React.ReactElement<IGroupMembersProps> {
    const { people } = this.state;
    if (people.length === 0) return (<div>Select a group or wait for load</div>);
    return (
      <div className={styles.groupMembers}>
        { people.map(p => (<SPFxPeopleCard {...p} />)) }
      </div>
    );
  }
}
