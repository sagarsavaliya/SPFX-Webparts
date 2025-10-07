import * as React from 'react';
import styles from './BannerWithSearch.module.scss';
import type { IBannerWithSearchProps, ISearchResult } from './IBannerWithSearchProps';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchService } from '../services/SearchService';

export default class BannerWithSearch extends React.Component<IBannerWithSearchProps, {
  searchQuery: string;
  searchResults: ISearchResult[];
  isLoading: boolean;
  showResults: boolean;
}> {
  private searchService: SearchService;
  private searchTimeout: number;

  constructor(props: IBannerWithSearchProps) {
    super(props);
    this.state = {
      searchQuery: '',
      searchResults: [],
      isLoading: false,
      showResults: false
    };
    this.searchService = new SearchService(this.props.context);
  }

  private onSearchChange = (event?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void => {
    const searchValue = newValue || '';
    this.setState({ searchQuery: searchValue });
    
    if (this.searchTimeout) {
      clearTimeout(this.searchTimeout);
    }

    if (searchValue.length >= 3) {
      this.searchTimeout = window.setTimeout(() => {
        this.performSearch(searchValue).catch(error => {
          console.error('Search error:', error);
        });
      }, 500);
    } else {
      this.setState({ showResults: false, searchResults: [] });
    }
  };

  private performSearch = async (query: string): Promise<void> => {
    this.setState({ isLoading: true, showResults: true });
    try {
      const results = await this.searchService.searchAdhocCombined(query, this.props.limitToSite);
      this.setState({ searchResults: results, isLoading: false });
    } catch (error) {
      console.error('Search error:', error);
      this.setState({ isLoading: false });
    }
  };


  private getGreetingTextSize = (): string => {
    switch (this.props.greetingTextSize) {
      case 'Small': return styles.greetingSmall;
      case 'Large': return styles.greetingLarge;
      default: return styles.greetingMedium;
    }
  };

  private getImageStyle = (): React.CSSProperties => {
    const baseStyle: React.CSSProperties = {
      backgroundImage: this.props.imageUrl ? `url('${this.props.imageUrl}')` : 'none',
      backgroundSize: 'cover',
      backgroundPosition: 'center',
      backgroundRepeat: 'no-repeat'
    };

    switch (this.props.imageSettings) {
      case 'ZoomIn':
        return { ...baseStyle, backgroundSize: '120%' };
      case 'ZoomOut':
        return { ...baseStyle, backgroundSize: '80%' };
      case 'Crop':
      default:
        return { ...baseStyle, backgroundSize: 'cover' };
    }
  };

  private renderSearchResult = (result: ISearchResult, index: number): JSX.Element => {
    if (this.props.searchResultType === 'Adhoc' && result.userImage) {
      // People result
      return (
        <div key={index} className={styles.searchResultItem} onClick={() => window.open(result.url, '_blank')}>
          <Persona
            imageUrl={result.userImage}
            text={result.title}
            secondaryText={result.department}
            tertiaryText={result.jobTitle}
            size={PersonaSize.size40}
          />
          <div className={styles.resultDetails}>
            <div className={styles.resultTitle}>{result.title}</div>
            <div className={styles.resultDescription}>{result.description}</div>
            <div className={styles.resultMeta}>
              <span>{result.officeLocation}</span>
              <span className={styles.status}>{result.status}</span>
            </div>
          </div>
        </div>
      );
    } else {
      // Document result
      return (
        <div key={index} className={styles.searchResultItem} onClick={() => window.open(result.url, '_blank')}>
          <div className={styles.resultIcon}>
            <Icon iconName={result.icon || 'Page'} />
          </div>
          <div className={styles.resultDetails}>
            <div className={styles.resultTitle}>{result.title}</div>
            <div className={styles.resultDescription}>{result.description}</div>
            <div className={styles.resultMeta}>
              <span>{result.author}</span>
              <span>{result.modifiedDate}</span>
              {result.fileType && <span>{result.fileType}</span>}
              {result.fileSize && <span>{result.fileSize}</span>}
            </div>
          </div>
        </div>
      );
    }
  };

  public render(): React.ReactElement<IBannerWithSearchProps> {
    const { greetingText, userDisplayName } = this.props;
    const { searchQuery, searchResults, isLoading, showResults } = this.state;

    return (
      <div className={styles.fullWidthBanner} style={this.getImageStyle()}>
        <div className={styles.bannerContent}>
          <div className={`${styles.greetingText} ${this.getGreetingTextSize()}`}>
            {greetingText}, {userDisplayName}
          </div>
          
          <div className={styles.searchContainer}>
            <SearchBox
              placeholder="Search..."
              value={searchQuery}
              onChange={this.onSearchChange}
              className={styles.searchBox}
            />
            
            {showResults && (
              <div className={styles.searchResults}>
                {isLoading ? (
                  <div className={styles.loadingContainer}>
                    <Spinner size={SpinnerSize.medium} />
                    <span>Loading results...</span>
                  </div>
                ) : (
                  <div className={styles.resultsList}>
                    {searchResults.length > 0 ? (
                      searchResults.map((result, index) => this.renderSearchResult(result, index))
                    ) : (
                      <div className={styles.noResults}>No results found</div>
                    )}
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      </div>
    );
  }
}
