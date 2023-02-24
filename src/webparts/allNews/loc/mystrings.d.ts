declare interface IAllNewsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  NewsDetailsPageUrlFieldLabel: string;
  EnableLikeFieldLabel: string;
  EnableLikeToggleTrueLabel: string;
  EnableLikeToggleFalseLabel: string;
  DefaultThumbnailLabel:string;
}

declare module 'AllNewsWebPartStrings' {
  const strings: IAllNewsWebPartStrings;
  export = strings;
}
