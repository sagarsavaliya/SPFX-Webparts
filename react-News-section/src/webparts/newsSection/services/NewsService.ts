import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { INewsItem } from '../components/INewsSectionProps';

export class NewsService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  public async getNews(maxItems: number = 5): Promise<INewsItem[]> {
    try {
      // Get news from the current site
      const newsItems = await this.getNewsFromSite(this.context.pageContext.web.absoluteUrl, maxItems);

      // If we're on a hub site, also try to get news from associated sites
      // Note: Hub site detection and associated sites require additional API calls
      // For now, we'll focus on the current site
      
      // Sort by published date (newest first) and limit to maxItems
      return newsItems
        .sort((a, b) => b.publishedDate.getTime() - a.publishedDate.getTime())
        .slice(0, maxItems);
    } catch (error) {
      console.error('Error fetching news:', error);
      return [];
    }
  }

  private async getNewsFromSite(siteUrl: string, maxItems: number): Promise<INewsItem[]> {
    try {
      // Construct the REST API URL to get news articles from Site Pages library
      const restUrl = `${siteUrl}/_api/web/lists/getbytitle('Site Pages')/items?` +
        `$select=Id,Title,Description,Created,Author/Title,FileRef,BannerImageUrl,PromotedState,FirstPublishedDate&` +
        `$expand=Author&` +
        `$filter=PromotedState eq 2&` +
        `$orderby=FirstPublishedDate desc&` +
        `$top=${maxItems}`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        restUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const data = await response.json();
      const items = data.value || [];

      return items.map((item: {
        Id: number;
        Title: string;
        Description: string;
        Created: string;
        Author?: { Title: string };
        FileRef: string;
        BannerImageUrl?: string | { Url: string; Description?: string };
        FirstPublishedDate?: string;
      }) => ({
        id: item.Id.toString(),
        title: item.Title || 'Untitled',
        description: this.extractDescription(item.Description) || 'No description available',
        imageUrl: this.parseBannerImageUrl(item.BannerImageUrl),
        publishedDate: new Date(item.FirstPublishedDate || item.Created),
        authorName: item.Author?.Title || 'Unknown',
        newsUrl: `${siteUrl}/${item.FileRef}`
      }));
    } catch (error) {
      console.error(`Error fetching news from site ${siteUrl}:`, error);
      return [];
    }
  }

  private extractDescription(description: string): string {
    if (!description) return '';
    
    // Remove HTML tags and get first 200 characters
    const textOnly = description.replace(/<[^>]*>/g, '');
    return textOnly.length > 200 ? textOnly.substring(0, 200) + '...' : textOnly;
  }

  private parseBannerImageUrl(bannerImageUrl: string | { Url: string; Description?: string } | undefined): { Url: string; Description?: string } | undefined {
    if (!bannerImageUrl) return undefined;
    
    // If it's already a string, return it as an object
    if (typeof bannerImageUrl === 'string') {
      // Check if it might be a JSON string
      if (bannerImageUrl.indexOf('{') === 0) {
        try {
          const parsed = JSON.parse(bannerImageUrl);
          return parsed.Url ? { Url: parsed.Url, Description: parsed.Description } : undefined;
        } catch {
          return { Url: bannerImageUrl };
        }
      }
      return { Url: bannerImageUrl };
    }
    
    // If it's already an object with Url property
    if (bannerImageUrl && bannerImageUrl.Url) {
      return { Url: bannerImageUrl.Url, Description: bannerImageUrl.Description };
    }
    
    return undefined;
  }

  public async canUserCreateNews(): Promise<boolean> {
    try {
      // Check if user has contribute permissions to Site Pages library
      const restUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site Pages')/effectivebasepermissions`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        restUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return false;
      }

      const data = await response.json();
      // Check for AddListItems permission (bit 2)
      const permissions = parseInt(data.Low, 10);
      return (permissions & 2) !== 0;
    } catch (error) {
      console.error('Error checking user permissions:', error);
      return false;
    }
  }

  public getCreateNewsUrl(): string {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    
    // Use the correct URL pattern for your SharePoint environment
    // promotedState=1 indicates this is a news post
    return `${siteUrl}/_layouts/15/createpagefromtemplate.aspx?promotedState=1`;
  }
}