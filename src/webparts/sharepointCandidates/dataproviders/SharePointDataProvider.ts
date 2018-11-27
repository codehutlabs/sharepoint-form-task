import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import * as lodash from '@microsoft/sp-lodash-subset';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import IDataProvider from '../dataproviders/IDataProvider';


export default class SharePointDataProvider implements IDataProvider {

  private _webPartContext: IWebPartContext;

  public set webPartContext(value: IWebPartContext) {
    this._webPartContext = value;
  }

  public get webPartContext(): IWebPartContext {
    return this._webPartContext;
  }

}
