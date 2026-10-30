declare interface ITimesheetWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  EmailToFieldLabel: string;
  ListTitleFieldLabel: string;
  TemplateLibraryUrlFieldLabel: string;
  TemplateLibraryUrlFieldDescription: string;
}

declare module 'TimesheetWebPartStrings' {
  const strings: ITimesheetWebPartStrings;
  export = strings;
}
