import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAllNewsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  listName: string;
  newsDetailsPageUrl: string;  
  SetHeightForQuickLinks:any;
  DefaultThumbnail:string;
  context:WebPartContext;
}
