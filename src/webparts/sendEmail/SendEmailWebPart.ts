import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SendEmailWebPartStrings';
import SendEmail from './components/SendEmail';
import { ISendEmailProps } from './components/ISendEmailProps';

export interface ISendEmailWebPartProps {
  description: string;
}

export default class SendEmailWebPart extends BaseClientSideWebPart <ISendEmailWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISendEmailProps> = React.createElement(
      SendEmail,
      {
        userEmail: this.context.pageContext.user.email,
        graph: this.context.msGraphClientFactory
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
