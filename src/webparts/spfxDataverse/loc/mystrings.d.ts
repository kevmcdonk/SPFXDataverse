declare interface ISpfxDataverseWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'SpfxDataverseWebPartStrings' {
  const strings: ISpfxDataverseWebPartStrings;
  export = strings;
}
