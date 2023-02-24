import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface INewsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  webpartTitle: string;
  orderBy: string;
  order: string;
  enableLike: boolean;
  listName: string;
  moreNewsPageUrl: string;
  newsDetailsPageUrl: string;
  SetHeight:string;
  emptyData:string;
  speedOfCarousel: any;
  noOfSlides: any;
  context:WebPartContext;
  itemsLimit: any;
  bgColor: any;
  defaultThumbnail: string;
}
