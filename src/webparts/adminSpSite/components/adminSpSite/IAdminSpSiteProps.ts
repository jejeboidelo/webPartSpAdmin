import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAdminSpSiteProps {
  isDarkTheme: boolean;
  environmentMessage: string;
  context:  WebPartContext;
}
