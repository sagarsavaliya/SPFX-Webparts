import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IBannerWithSearchProps {
  imageSource: string;
  greetingTextSize: string;
  searchResultType: string;
  imageSettings: string;
  imageUrl: string;
  greetingText: string;
  limitToSite: boolean;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}

export interface ISearchResult {
  title: string;
  description: string;
  author: string;
  modifiedDate: string;
  url: string;
  fileType?: string;
  fileSize?: string;
  icon?: string;
  userImage?: string;
  department?: string;
  officeLocation?: string;
  jobTitle?: string;
  status?: string;
}
