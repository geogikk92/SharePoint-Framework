declare interface IAutoCompleteStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'autoCompleteStrings' {
  const strings: IAutoCompleteStrings;
  export = strings;
}
