// Copyright (c) Microsoft Corporation.
declare interface IGraphTutorialWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'GraphTutorialWebPartStrings' {
  const strings: IGraphTutorialWebPartStrings;
  export = strings;
}
