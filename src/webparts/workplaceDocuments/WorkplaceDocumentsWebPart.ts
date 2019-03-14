import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'WorkplaceDocumentsWebPartStrings';
import WorkplaceDocuments from './components/WorkplaceDocuments';
import { IWorkplaceDocumentsProps } from './components/IWorkplaceDocumentsProps';

export interface IWorkplaceDocumentsWebPartProps {
  description: string;
}

export default class WorkplaceDocumentsWebPart extends BaseClientSideWebPart<IWorkplaceDocumentsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IWorkplaceDocumentsProps > = React.createElement(
      WorkplaceDocuments,
      {
        description: this.properties.description
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
