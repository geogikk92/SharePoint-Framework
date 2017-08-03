import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import {
  Environment, EnvironmentType
} from '@microsoft/sp-core-library';

import * as strings from 'autoCompleteStrings';
import AutoComplete from './components/AutoComplete';
import { IAutoCompleteProps } from './components/IAutoCompleteProps';
import { IAutoCompleteWebPartProps } from './IAutoCompleteWebPartProps';
import { Web } from "sp-pnp-js";

export default class AutoCompleteWebPart extends BaseClientSideWebPart<IAutoCompleteWebPartProps> {

  public render(): void {

    const element: React.ReactElement<IAutoCompleteProps> = React.createElement(
      AutoComplete,
      {
        description: this.properties.description
      }
    );

    this.getCurrEmpHolidays("Employees");

    ReactDom.render(element, this.domElement);
  }

  private getCurrEmpHolidays(listName: string): void {

    let web = new Web(this.context.pageContext.site.absoluteUrl);
    let html: string = '';
    let filterTxt = `Title eq  '${this.getCurrentUserName()}'`;

    // use odata operators for more efficient queries
    web.lists.getByTitle(listName).items.filter(filterTxt).top(1).get().then((item => {

      html = `<ul>
                    <li>
                        <span class="ms-font-l">${item.Title}</span> 
                         <span class="ms-font-l">${item.LirexEmplFullName}</span>
                    </li>
                </ul>`


      // Use for multiple items
      // web.lists.getByTitle(listName).items.filter(filterTxt).top(1).get().then((items: any[]) => {

      //   items.forEach(item => {
      //     html += `<ul>
      //                 <li>
      //                     <span class="ms-font-l">${item.Title}</span> 
      //                      <span class="ms-font-l">${item.LirexEmplFullName}</span>
      //                 </li>
      //             </ul>`
      //   });

      this.domElement.querySelector('#lists').innerHTML = html;
    });
  }


  private getCurrentUserName(): string {
    let formattedUserName: string = this.context.pageContext.user.loginName
    formattedUserName = formattedUserName.substring(formattedUserName.lastIndexOf("|") + 1, formattedUserName.lastIndexOf("@"));
    return formattedUserName;
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


