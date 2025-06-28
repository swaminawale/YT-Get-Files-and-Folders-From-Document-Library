import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IGetFilesAndFoldersProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
