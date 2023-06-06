import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMyPlansProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  context:WebPartContext;
  webHeight:any;
  noofPlans:any;
}
