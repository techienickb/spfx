import * as React from 'react';
import styles from './PageBreadcrumbs.module.scss';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Breadcrumb, IBreadcrumbItem } from 'office-ui-fabric-react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import getParentColumnDefinition from '../../../common/parentcolumn';
import ISitePageItem from '../../../common/ISitePageItem';
import { unstable_renderSubtreeIntoContainer } from 'react-dom';

export interface IPageBreadcrumbsProps {
  context: WebPartContext;
  root: number;
}

export default class PageBreadcrumbs extends React.Component<IPageBreadcrumbsProps, {items: IBreadcrumbItem[]}> {

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

  public componentWillReceiveProps(nextProps: Readonly<IPageBreadcrumbsProps>, nextContext: any): void {
    this.continueMount();
  }

  private continueMount = () => {
    const context: WebPartContext = this.props.context;
    context.spHttpClient.get(`${context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Site Pages')/items?$select=Id,ParentId,FileRef,Title`, SPHttpClient.configurations.v1)
    .then((res: SPHttpClientResponse): Promise<{ value: any; }> => {
      return res.json();
    })
    .then((response): void => {
      let items: ISitePageItem[] = response.value;
      var bread: IBreadcrumbItem[] = [];
      bread = bread.concat(items.filter(i => i.Id == this.props.root).map(i => ({ text: i.Title, href: i.FileRef, key: i.Id.toString() })));

      bread = bread.concat(this.recurse(items, context.pageContext.listItem ? context.pageContext.listItem.id : 17, context.pageContext.listItem ? context.pageContext.listItem.id : 17));

      this.setState({...this.state, items: bread});
    });
  }

  private recurse = (items: ISitePageItem[], id: number, current: number): IBreadcrumbItem[] => {
    var item: ISitePageItem = items.filter(i => i.Id == id)[0];
    let crumbs: IBreadcrumbItem[] = [{ text: item.Title, href: item.FileRef, key: item.Id.toString(), isCurrentItem: current == id }];
    if (item.ParentId && item.ParentId != this.props.root) return this.recurse(items, item.ParentId, current).concat(crumbs);
    return crumbs;
  }

  public render(): React.ReactElement<IPageBreadcrumbsProps> {
    const {items} = this.state;
    return (<Breadcrumb items={items} maxDisplayedItems={3} />);
  }
}
