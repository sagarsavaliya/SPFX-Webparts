import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface INewsItem {
  id: string;
  title: string;
  description: string;
  imageUrl: {
    Url: string;
    Description?: string;
  } | undefined;
  publishedDate: Date;
  authorName: string;
  newsUrl: string;
}

export interface INewsSectionProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  layoutType: string;
  maxNewsItems: number;
  showCreateNewsButton: boolean;
  showImages: boolean;
  showAuthor: boolean;
  showPublishedDate: boolean;
}
