import * as React from 'react';
import styles from './NewsItem.module.scss';
import { INewsItem } from './INewsSectionProps';

export interface INewsItemProps {
  newsItem: INewsItem;
  showImage: boolean;
  showAuthor: boolean;
  showPublishedDate: boolean;
  isDarkTheme: boolean;
}

export const NewsItem: React.FC<INewsItemProps> = ({
  newsItem,
  showImage,
  showAuthor,
  showPublishedDate,
  isDarkTheme
}) => {
  const formatDate = (date: Date): string => {
    return date.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  };

  const truncateDescription = (text: string, lines: number = 3): string => {
    const words = text.split(' ');
    const wordsPerLine = 12; // Approximate words per line
    const maxWords = lines * wordsPerLine;
    
    if (words.length <= maxWords) {
      return text;
    }
    
    return words.slice(0, maxWords).join(' ') + '...';
  };

  return (
    <div className={`${styles.newsItem} ${isDarkTheme ? styles.dark : ''}`}>
      {showImage && newsItem.imageUrl && newsItem.imageUrl.Url && (
        <div className={styles.imageContainer}>
          <img 
            src={newsItem.imageUrl.Url} 
            alt={newsItem.title}
            className={styles.newsImage}
            onError={(e) => {
              console.log('Image failed to load:', newsItem.imageUrl?.Url);
              const target = e.target as HTMLImageElement;
              target.style.display = 'none';
            }}
            onLoad={() => {
              console.log('Image loaded successfully:', newsItem.imageUrl?.Url);
            }}
          />
        </div>
      )}
      
      <div className={styles.content}>
        <h3 className={styles.title}>
          <a 
            href={newsItem.newsUrl} 
            target="_blank" 
            rel="noopener noreferrer"
            className={styles.titleLink}
          >
            {newsItem.title}
          </a>
        </h3>
        
        <p className={styles.description}>
          {truncateDescription(newsItem.description)}
        </p>
        
        <div className={styles.metadata}>
          {showPublishedDate && (
            <span className={styles.date}>
              {formatDate(newsItem.publishedDate)}
            </span>
          )}
          
          {showAuthor && (
            <span className={styles.author}>
              {showPublishedDate && ' • '}
              By {newsItem.authorName}
            </span>
          )}
        </div>
      </div>
    </div>
  );
};

