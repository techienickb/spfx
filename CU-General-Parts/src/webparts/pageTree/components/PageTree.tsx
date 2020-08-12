import * as React from 'react';
import styles from './PageTree.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { INavLinkGroup, INavLink, Nav } from 'office-ui-fabric-react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import getParentColumnDefinition from '../../../common/parentcolumn';
import ISitePageItem from '../../../common/ISitePageItem';

export interface IPageTreeProps {
  context: WebPartContext;
  rootid: number;
  expandtolevel: number;
}


export default class PageTree extends React.Component<IPageTreeProps, { items: INavLinkGroup[] }> {
  public state = { items: [] };

  public componentDidMount() {
    const context: WebPartContext = this.props.context;
    context.spHttpClient.get(`${context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Site Pages')/Fields`, SPHttpClient.configurations.v1)
    .then((res: SPHttpClientResponse): Promise<{ value: any; }> => {
      return res.json();
    })
    .then((response): void => {
      if (response.value.filter(i => i.Title === "Parent").length === 0) {
        context.spHttpClient.post(`${context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Site Pages')/Fields/addfield`, SPHttpClient.configurations.v1, { body: JSON.stringify(getParentColumnDefinition(response.value[0])) })
          .catch((error) => {
            alert(error);
            console.error(error);
          }).then(() => {
            this.continueMount();
          });
      } else this.continueMount();
    });
  }

  private continueMount = () => {
    const context: WebPartContext = this.props.context;
    context.spHttpClient.get(`${context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Site Pages')/items?$select=Id,ParentId,FileRef,Title&$orderby=Title`, SPHttpClient.configurations.v1)
    .then((res: SPHttpClientResponse): Promise<{ value: any; }> => {
      return res.json();
    })
    .then((response): void => {
      let items: ISitePageItem[] = response.value;
      this.setState({...this.state, items: [{ links: [this.recurse(items, this.props.rootid, 1)] }]});
    });
  }

  public componentWillReceiveProps(nextProps: Readonly<IPageTreeProps>, nextContext: any): void {
    this.continueMount();
  }

  private recurse = (items: ISitePageItem[], id: number, l: number): INavLink => {
    var item: ISitePageItem = items.filter(i => i.Id == id)[0];
    
    var links: INavLink[] = [];
    links = links.concat(items.filter(i => i.ParentId == id).map(it => this.recurse(items, it.Id, (l+1))));

    return  { name: item.Title, url: item.FileRef, key: item.Id.toString(), links: links, isExpanded: this.props.expandtolevel >= l  };
  }
  
  public render(): React.ReactElement<IPageTreeProps> {
    const { items } = this.state;
    const context: WebPartContext = this.props.context;
    return (<Nav groups={items} selectedKey={context.pageContext.listItem ? context.pageContext.listItem.id.toString() : '6' } />);
  }
}
