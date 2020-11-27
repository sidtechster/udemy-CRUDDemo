import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CruDdemoWebPart.module.scss';
import * as strings from 'CruDdemoWebPartStrings';

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IMyListItem } from './loc/IMyListItem';

export interface ICruDdemoWebPartProps {
  description: string;
}

export default class CruDdemoWebPart extends BaseClientSideWebPart<ICruDdemoWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.cruDdemo }">
        
        <p>Enter ID</p><br/>
        <input type="text" id="txtID"/>
        <input type="submit" value="Read details" id="btnRead" />
        <br/><br/>

        <p>Title</p><br/>
        <input type="text" id="txtTitle"/><br/><br/>
        
        <input type="submit" value="Insert item" id="btnSubmit" />
        <input type="submit" value="Update item" id="btnUpdate" />
        <input type="submit" value="Delete item" id="btnDelete" />
        <input type="submit" value="Show all item" id="btnShowAll" />

        <br/>

        <div id="divStatus"></div>

        <div id="spListData"></div>

      </div>`;

      this.bindEvents();
      this.readAllItems();
  }

  private bindEvents(): void {
    this.domElement.querySelector('#btnSubmit').addEventListener('click', () => { this.addListItem(); });
    this.domElement.querySelector('#btnRead').addEventListener('click', () => { this.readListItem(); });
    this.domElement.querySelector('#btnUpdate').addEventListener('click', () => { this.updateListItem(); });
    this.domElement.querySelector('#btnDelete').addEventListener('click', () => { this.deleteListItem(); });
    this.domElement.querySelector('#btnShowAll').addEventListener('click', () => { this.readAllItems(); });
  }

  private readAllItems(): void {
    
    this._getListItems().then(listItems => {
      let html: string = '<h2>Title</h2>';
      html += '<ul>';

      listItems.forEach(listItem => {
        html += `<li>${listItem.Title}</li>`;        
      });
      html += '</ul>';

      const listContainer: Element = this.domElement.querySelector('#spListData');
      listContainer.innerHTML = html;

    })
  }

  private _getListItems(): Promise<IMyListItem[]> {
    
    const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('MySampleList')/items";

    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      })
      .then(json => {
        return json.value;
      }) as Promise<IMyListItem[]>;
  }

  private deleteListItem(): void {
    let id = document.getElementById('txtID')["value"];

    const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('MySampleList')/items(" + id + ")";

    const headers: any = {
      "X-HTTP-Method": "DELETE",
      "IF-MATCH": "*",
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers
    };

    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if(response.status === 204) {
          let message: Element = this.domElement.querySelector('#divStatus');
          message.innerHTML = "List item deleted";
        } else {
          let message: Element = this.domElement.querySelector('#divStatus');
          message.innerHTML = "List item delete failed. " + response.status + " - " + response.statusText;
        }
      });
  }

  private updateListItem(): void {
    let id = document.getElementById('txtID')["value"];
    let title = document.getElementById('txtTitle')["value"];

    const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('MySampleList')/items(" + id + ")";

    const itemBody: any = {
      "Title": title
    };

    const headers: any = {
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*",
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers,
      "body": JSON.stringify(itemBody)
    };

    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if(response.status === 204) {
          let message: Element = this.domElement.querySelector('#divStatus');
          message.innerHTML = "List item updated";
        } else {
          let message: Element = this.domElement.querySelector('#divStatus');
          message.innerHTML = "List item update failed. " + response.status + " - " + response.statusText;
        }
      });
  }

  private readListItem(): void {

    let id = document.getElementById('txtID')["value"];
    this._getListItemByID(id).then(listItem => {
      document.getElementById("txtTitle")["value"] = listItem.Title;
    })
    .catch(error => {
      let message: Element = this.domElement.querySelector("#divStatus");
      message.innerHTML = "Read: Could not fetch details.." + error.message;
    })

  }

  private _getListItemByID(id: string): Promise<IMyListItem> {
    const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('MySampleList')/items?$filter=Id eq " + id;

    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then ((listItems: any) => {
        const untypedItem: any = listItems.value[0];
        const listItem: IMyListItem = untypedItem as IMyListItem;
        return listItem;
      }) as Promise<IMyListItem>
  }

  private addListItem(): void {

    let title = document.getElementById('txtTitle')["value"];
    
    const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('MySampleList')/items";

    const itemBody: any = {
      "Title": title
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "body": JSON.stringify(itemBody)
    };

    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if(response.status === 201) {
          let statusMessage: Element = this.domElement.querySelector('#divStatus');
          statusMessage.innerHTML = "List item created";
          this.clear();
        }
        else {
          let statusMessage: Element = this.domElement.querySelector('#divStatus');
          statusMessage.innerHTML = "Error " + response.status + " - " + response.statusText;
        }
      });
    }
  

  private clear(): void {
    document.getElementById("txtTitle")["value"] = '';
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
