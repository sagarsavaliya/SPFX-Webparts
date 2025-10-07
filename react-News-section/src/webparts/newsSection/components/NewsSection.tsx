import * as React from 'react';
import styles from './NewsSection.module.scss';
import type { INewsSectionProps, INewsItem } from './INewsSectionProps';
import { NewsItem } from './NewsItem';
import { NewsService } from '../services/NewsService';

export default class NewsSection extends React.Component<INewsSectionProps, {
  newsItems: INewsItem[];
  loading: boolean;
  error: string | undefined;
  canCreateNews: boolean;
}> {
  private newsService: NewsService;

  constructor(props: INewsSectionProps) {
    super(props);
    
    this.state = {
      newsItems: [],
      loading: true,
      error: undefined,
      canCreateNews: false
    };

    this.newsService = new NewsService(props.context);
  }

  public async componentDidMount(): Promise<void> {
    await this.loadNews();
    await this.checkCreatePermissions();
  }

  public async componentDidUpdate(prevProps: INewsSectionProps): Promise<void> {
    if (prevProps.maxNewsItems !== this.props.maxNewsItems) {
      await this.loadNews();
    }
  }

  private async loadNews(): Promise<void> {
    try {
      this.setState({ loading: true, error: undefined });
      const newsItems = await this.newsService.getNews(this.props.maxNewsItems);
      console.log('Loaded news items:', newsItems);
      this.setState({ newsItems, loading: false });
    } catch (error) {
      console.error('Error loading news:', error);
      this.setState({ 
        error: 'Failed to load news items. Please try again later.',
        loading: false 
      });
    }
  }

  private async checkCreatePermissions(): Promise<void> {
    try {
      const canCreateNews = await this.newsService.canUserCreateNews();
      this.setState({ canCreateNews });
    } catch (error) {
      console.error('Error checking permissions:', error);
      this.setState({ canCreateNews: false });
    }
  }

  private handleCreateNewsClick = (): void => {
    try {
      const createNewsUrl = this.newsService.getCreateNewsUrl();
      // Open in the same window to maintain SharePoint context
      window.location.href = createNewsUrl;
    } catch (error) {
      console.error('Error navigating to create news:', error);
      // Fallback: Navigate to Site Pages library
      const fallbackUrl = `${this.props.context.pageContext.web.absoluteUrl}/SitePages`;
      window.open(fallbackUrl, '_blank');
    }
  };

  public render(): React.ReactElement<INewsSectionProps> {
    const { 
      showCreateNewsButton,
      showImages,
      showAuthor,
      showPublishedDate,
      isDarkTheme
    } = this.props;
    
    const { newsItems, loading, error, canCreateNews } = this.state;

    return (
      <section className={`${styles.newsSection} ${isDarkTheme ? styles.dark : ''}`}>
        {showCreateNewsButton && canCreateNews && (
          <button 
            className={styles.createNewsButton}
            onClick={this.handleCreateNewsClick}
            type="button"
          >
            Create News
          </button>
        )}

        {loading && (
          <div className={styles.loadingMessage}>
            Loading news items...
          </div>
        )}

        {error && (
          <div className={styles.errorMessage}>
            {error}
          </div>
        )}

        {!loading && !error && newsItems.length === 0 && (
          <div className={styles.noNewsMessage}>
            No news items found.
          </div>
        )}

        {!loading && !error && newsItems.length > 0 && (
          <div className={styles.newsList}>
            {newsItems.map((newsItem) => (
              <NewsItem
                key={newsItem.id}
                newsItem={newsItem}
                showImage={showImages}
                showAuthor={showAuthor}
                showPublishedDate={showPublishedDate}
                isDarkTheme={isDarkTheme}
              />
            ))}
          </div>
        )}
      </section>
    );
  }
}
