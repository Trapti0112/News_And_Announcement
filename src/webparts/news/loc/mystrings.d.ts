declare interface INewsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  WebpartTitleFieldLabel: string;
  listNameLabel: string;
  OrderByFieldLabel: string;
  OrderFieldLabel: string;
  NoOfItemsFieldLabel: string;
  ItemsLimitFieldDescription: string;
  EnableLikeFieldLabel: string;
  EnableLikeToggleTrueLabel: string;
  EnableLikeToggleFalseLabel: string;
  MoreNewsPageUrlFieldLabel: string;
  NewsDetailsPageUrlFieldLabel: string;
  ListNameFieldLabel: string;
  selectedListLabel:string;
  DefaultThumbnailLabel:string;
  componentHeight: string;
  emptyData: string;
  speedOfCarouselFieldLabel: string;
  bgColorFieldLabel: string;
  componentHeightFieldLabel: string;
}

declare module 'NewsWebPartStrings' {
  const strings: INewsWebPartStrings;
  export = strings;
}
