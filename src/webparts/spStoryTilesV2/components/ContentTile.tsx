import * as React from 'react';
import styles from './ContentTile.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { IContentTileProps } from './IContentTileProps';
import { Icon } from '@fluentui/react/lib/Icon';

const ContentTile: React.FC<IContentTileProps> = (props) => {
  const { item } = props;

  const handleTileClick = (): void => {
    // If there's a link URL, navigate to that URL instead of opening modal
    if (item.linkUrl) {
      window.open(item.linkUrl, '_blank');
    }
  };

  return (
    <div 
      className={styles.tileContainer}
      onClick={handleTileClick}
      role="button"
      tabIndex={0}
      onKeyDown={(e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          handleTileClick();
        }
      }}
    >
      <div className={styles.imageContainer} style={{ backgroundImage: `url('${item.imageUrl}')` }}>
      </div>
      <div className={styles.contentContainer}>
        <h3 className={styles.title} title={item.title}>
          {escape(item.title)}
        </h3>
        <p className={styles.description} title={item.description}>
          {escape(item.description.length > 120 ? item.description.substring(0, 120) + '...' : item.description)}
        </p>
        <div className={styles.arrowIcon}>
          <Icon iconName="ChromeBackMirrored" />
        </div>
      </div>
    </div>
  );
};

export default ContentTile;