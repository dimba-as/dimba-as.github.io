import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SpListingsWebPartStrings';
import SpListings from './components/SpListings';
import { ISpListingsProps } from './components/ISpListingsProps';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISpListingsWebPartProps {
  description: string;
}

export default class SpListingsWebPart extends BaseClientSideWebPart<ISpListingsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpListingsProps> = React.createElement(
      SpListings,
      {
        context:this.context
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
