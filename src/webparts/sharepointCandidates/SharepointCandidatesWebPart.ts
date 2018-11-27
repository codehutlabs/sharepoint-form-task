import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SharepointCandidatesWebPartStrings';
import SharepointCandidates from './components/SharepointCandidates';
import { ISharepointCandidatesProps } from './components/ISharepointCandidatesProps';
import IDataProvider from './dataproviders/IDataProvider';
import SharePointDataProvider from './dataproviders/SharePointDataProvider';

export interface ISharepointCandidatesWebPartProps {
  description: string;
}

export default class SharepointCandidatesWebPart extends BaseClientSideWebPart<ISharepointCandidatesWebPartProps> {

  private _dataProvider: IDataProvider;

  protected onInit(): Promise<void> {

    this._dataProvider = new SharePointDataProvider();
    this._dataProvider.webPartContext = this.context;

    return super.onInit();
  }


  public render(): void {
    const element: React.ReactElement<ISharepointCandidatesProps > = React.createElement(
      SharepointCandidates,
      {
        description: this.properties.description,
        dataProvider: this._dataProvider
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
