import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'WorkingWithListItemWebPartStrings';
import WorkingWithListItem from './components/WorkingWithListItem';
import { IWorkingWithListItemProps } from './components/IWorkingWithListItemProps';

export interface IWorkingWithListItemWebPartProps {
  description: string;
}

export default class WorkingWithListItemWebPart extends BaseClientSideWebPart <IWorkingWithListItemWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IWorkingWithListItemProps> = React.createElement(
      WorkingWithListItem,
      {
        context: this.context,
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
