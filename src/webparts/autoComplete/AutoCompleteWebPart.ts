import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import {
  Environment, EnvironmentType
} from '@microsoft/sp-core-library';

import * as strings from 'autoCompleteStrings';
import AutoComplete from './components/AutoComplete';
import { IAutoCompleteProps } from './components/IAutoCompleteProps';
import { IAutoCompleteWebPartProps } from './IAutoCompleteWebPartProps';
import { Web } from "sp-pnp-js";
import { List, ListEnsureResult } from "sp-pnp-js";

export default class AutoCompleteWebPart extends BaseClientSideWebPart<IAutoCompleteWebPartProps> {

  public render(): void {

    const element: React.ReactElement<IAutoCompleteProps> = React.createElement(
      AutoComplete,
      {
        description: this.properties.description,
        EmpListTitleProperty: this.properties.EmpListTitleProperty
      }
    );

    this.getCurrEmpData();
    ReactDom.render(element, this.domElement);
  }

  private getCurrEmpData(): void {

    let web = new Web(this.context.pageContext.site.absoluteUrl);
    let listName: string = this.properties.EmpListTitleProperty;
    let filter: string = `Title eq '${this.getCurrentUserNickName()}'`;
    let selectColumns: string = 'LirexEmplFullName, LirexEmplDays, LirexEmplPosition, LirexEmplDepartment, LirexEmplDirection';
    let generatedHtml: string = '';

    web.lists.getByTitle(listName).items.select(selectColumns).filter(filter).top(1).get().then((items: any[]) => {

      items.forEach(item => {
        generatedHtml += `<td>${item.LirexEmplFullName}</td>
                          <td>${item.LirexEmplDays}</td>
                          <td>${item.LirexEmplPosition}</td>
                          <td>${item.LirexEmplDepartment}</td>
                          <td>${item.LirexEmplDirection}</td>`;
      });
      this.domElement.querySelector('#trUserInfo').innerHTML = generatedHtml;
    });
  }

  private getCurrentUserNickName(): string {
    let formattedUserName: string = this.context.pageContext.user.loginName;
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
                PropertyPaneTextField('EmpListTitleProperty', {
                  label: "Име на списък служители",
                  value: "Служители",
                  placeholder: "Въведете име на списък със служители",
                  onGetErrorMessage: this.simpleTextBoxValidationMethod
                }),
                PropertyPaneTextField('description', {
                  label: "Заглавен текст на уеб парта(Описание)",
                  multiline: true,
                  resizable: true
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private simpleTextBoxValidationMethod(value: string): string {
    if (value.length < 2) {
      return "Моля, въведете повече от два символа!";
    } else {
      return "";
    }
  }
}
