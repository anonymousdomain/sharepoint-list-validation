import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {SPHttpClient,SPHttpClientResponse} from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ListInfoWebPart.module.scss';
import * as strings from 'ListInfoWebPartStrings';

export interface IListInfoWebPartProps {
  description: string;
  listName:string;
}

export default class ListInfoWebPart extends BaseClientSideWebPart<IListInfoWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.listInfo }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

 public validateDscription(input:string):string{

  if(input===null||input.trim().length===0){
    return "description is required";
  }
  if(input.length>=12){
    return "your description should not be longer than 12";
  }
   return "";
 }
 private async validateListName(value:string):Promise<string>{
   if(value===null||value.trim().length===0){
     return "provide the list name"
   }
   try {
     let response=await this.context.spHttpClient.get(
       this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('${escape(value)}')?$select=Id`,
       SPHttpClient.configurations.v1);
       if(response.ok){
         return "";
       }else if(response.status===404){
         return `list '${escape(value)}' doesn't exist in the current site`
       }else{
         return `Error:${response.statusText}. please try agan`
       }
   } catch (error) {
     return error.message;
   }
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
                  label: strings.DescriptionFieldLabel,
                  onGetErrorMessage:this.validateDscription.bind(this)
                }),
                PropertyPaneTextField('listName',{
                  label:strings.ListNameFieldLabel,
                  onGetErrorMessage:this.validateListName.bind(this),
                  deferredValidationTime:1000
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
