import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
//import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GetSpListItemsWebPart.module.scss';
import * as strings from 'GetSpListItemsWebPartStrings';

import {
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';
/*import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';*/

export interface IGetSpListItemsWebPartProps {
  description: string;
}
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Course_Name: string;
  Description : string;
  Start_Date : number;
  End_Date : number;
  Assigned_User :string;

}
  
export default class GetSpListItemsWebPart extends BaseClientSideWebPart<IGetSpListItemsWebPartProps> {
  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Training Data')/Items`,SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
        return response.json();
        });
    }
    private _renderList(): void {
      this._getListData()
         .catch((response) => {
          
           
      let html: string = '<table border=1 width=100% style="border-collapse: collapse;">';
      html += '<th>Course Name</th> <th>Description</th><th>Start Date</th><th>End Datee</th><th>Assigned User</th>';
      response.value.forEach((item: ISPList) => {
        html += `
        <tr>            
            <td>${item.Course_Name}</td>
            <td>${item.Description}</td>
            <td>${item.Start_Date}</td>
            <td>${item.End_Date}</td>
            <td>${item.Assigned_User}</td>
            
            </tr>
            `;
      });
      html += '</table>';
    
      const listContainer : Element = this.domElement.querySelector('#spListContainer')!;
      listContainer.innerHTML = html;
    });
    
    }
      
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.getSpListItems }">
        <div class="${ styles }">
          <span class="ms-fontColor-neutralSecondary ms-fontSize-16 ms-fontColor-white">Welcome To Training Data List Report!</span>
          <p class="ms-font-l ms-fontColor-white">Site Name:- ${this.context.pageContext.web.title}</p>
          
        </div>
      </div> 
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles}">
          <div>List in table formate :-</div>
          <br>
           <div id="spListContainer" />
        </div>
      </div>`;
      this._renderList();
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