import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IEventsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  context:WebPartContext;

  numberOfMail:string;
  webHeight:any;
}
