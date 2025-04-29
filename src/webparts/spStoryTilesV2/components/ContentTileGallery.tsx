import * as React from 'react';
import styles from './ContentTile.module.scss';
import { IContentTileGalleryProps } from './IContentTileGalleryProps';
import { IContentTile } from '../SpStoryTilesV2WebPart';
import ContentTile from './ContentTile';
import ContentModalDialog from './ContentModalDialog';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';

export default class ContentTileGallery extends React.Component<IContentTileGalleryProps, {
  currentPage: number;
  isModalOpen: boolean;
  selectedContent: IContentTile | null;
  containerWidth: number;
  effectiveItemsPerPage: number;
}> {
  
  private _containerRef: React.RefObject<HTMLDivElement>;
  private _resizeObserver: ResizeObserver | null = null;

  constructor(props: IContentTileGalleryProps) {
    super(props);
    this.state = {
      currentPage: 0,
      isModalOpen: false,
      selectedContent: null,
      containerWidth: 0,
      effectiveItemsPerPage: props.itemsPerPage
    };
    this._containerRef = React.createRef();
  }

  public componentDidMount(): void {
    if (this._containerRef.current) {
      // Initialize with current width
      this.setState({ 
        containerWidth: this._containerRef.current.clientWidth
      }, () => {
        this._calculateItemsPerPage();
      });

      // Set up ResizeObserver to monitor container width changes
      this._resizeObserver = new ResizeObserver(entries => {
        for (const entry of entries) {
          if (entry.target === this._containerRef.current) {
            this.setState({ containerWidth: entry.contentRect.width }, () => {
              this._calculateItemsPerPage();
            });
          }
        }
      });
      
      this._resizeObserver.observe(this._containerRef.current);
    }
  }

  public componentWillUnmount(): void {
    // Clean up the ResizeObserver when component unmounts
    if (this._resizeObserver && this._containerRef.current) {
      this._resizeObserver.unobserve(this._containerRef.current);
      this._resizeObserver.disconnect();
    }
  }

  private _calculateItemsPerPage(): void {
    const { containerWidth } = this.state;
    const { itemsPerPage } = this.props;
    
    // More responsive approach to calculate tile width based on container size
    // Base tile width plus margins (adjusted from 180px to account for different screen sizes)
    const baseTileWidth = 160; // Base width of tile
    const tileMargins = 20; // Left + right margins
    const minTileWidth = baseTileWidth + tileMargins; // Minimum width needed for a tile
    
    // Account for container padding (left + right)
    const containerPadding = 40; // 20px on each side
    const availableWidth = Math.max(minTileWidth, containerWidth - containerPadding);
    
    // Determine how many tiles can fit based on available width
    // Adding a small buffer to prevent edge cases
    const buffer = 5;
    const canFit = Math.max(1, Math.floor((availableWidth - buffer) / minTileWidth));
    
    // Calculate effective items per page but don't exceed configured maximum
    let effectiveItemsPerPage = Math.min(canFit, itemsPerPage);
    
    // Ensure we always show at least 1 item
    effectiveItemsPerPage = Math.max(1, effectiveItemsPerPage);
    
    // Update state only if value changed
    if (this.state.effectiveItemsPerPage !== effectiveItemsPerPage) {
      this.setState({ 
        effectiveItemsPerPage,
        // Reset to first page when items per page changes to avoid showing empty pages
        currentPage: 0
      });
    }
  }

  public render(): React.ReactElement<IContentTileGalleryProps> {
    const { 
      contentItems, 
      isLoading, 
      hasTeamsContext,
      isDarkTheme
    } = this.props;

    const { 
      currentPage, 
      isModalOpen, 
      selectedContent, 
      effectiveItemsPerPage 
    } = this.state;
    
    // Calculate total pages based on effectiveItemsPerPage
    const totalPages = Math.ceil(contentItems.length / effectiveItemsPerPage);
    
    // Get items for current page
    const startIndex = currentPage * effectiveItemsPerPage;
    const endIndex = startIndex + effectiveItemsPerPage;
    const currentItems = contentItems.slice(startIndex, endIndex);
    
    // Generate grid class based on effective items per page
    let gridClass = styles.gridOne;
    if (effectiveItemsPerPage === 2) {
      gridClass = styles.gridTwo;
    } else if (effectiveItemsPerPage === 3) {
      gridClass = styles.gridThree;
    } else if (effectiveItemsPerPage === 4) {
      gridClass = styles.gridFour;
    }

    // Handle navigation
    const goToPreviousPage = (): void => {
      if (currentPage > 0) {
        this.setState({ currentPage: currentPage - 1 });
      }
    };

    const goToNextPage = (): void => {
      if (currentPage < totalPages - 1) {
        this.setState({ currentPage: currentPage + 1 });
      }
    };

    const dismissModal = (): void => {
      this.setState({ isModalOpen: false });
    };

    return (
      <div 
        className={`${styles.contentTileGallery} ${hasTeamsContext ? styles.teams : ''}`}
        ref={this._containerRef}
      >
        <div className={styles.container}>          
          {isLoading ? (
            <div className={styles.spinner}>
              <Spinner size={SpinnerSize.large} label="Loading content..." />
            </div>
          ) : contentItems.length === 0 ? (
            <div className={styles.noItems}>
              <p>No content items found. Please check the list configuration in the web part properties.</p>
            </div>
          ) : (
            <div className={styles.carouselContainer}>
              <div className={`${styles.tilesGrid} ${gridClass}`}>
                {currentItems.map((item: IContentTile) => (
                  <ContentTile 
                    key={item.id} 
                    item={item}
                    context={this.props.context}
                  />
                ))}
              </div>
              
              {/* Always show navigation controls if there are multiple pages */}
              {totalPages > 1 && (
                <div className={`${styles.navigationControls} ${styles.alwaysShowNavigation}`}>
                  <button 
                    className={`${styles.navButton} ${currentPage === 0 ? styles.disabled : ''}`}
                    onClick={goToPreviousPage}
                    disabled={currentPage === 0}
                    aria-label="Previous page"
                  >
                    <Icon iconName="ChevronLeftMed" />
                  </button>
                  <div className={styles.pageIndicator}>
                    {`${currentPage + 1} / ${totalPages}`}
                  </div>
                  <button 
                    className={`${styles.navButton} ${currentPage === totalPages - 1 ? styles.disabled : ''}`}
                    onClick={goToNextPage}
                    disabled={currentPage === totalPages - 1}
                    aria-label="Next page"
                  >
                    <Icon iconName="ChevronRightMed" />
                  </button>
                </div>
              )}
            </div>
          )}

          {isModalOpen && selectedContent && (
            <ContentModalDialog 
              isOpen={isModalOpen}
              onDismiss={dismissModal}
              contentData={selectedContent}
              isDarkTheme={isDarkTheme}
            />
          )}
        </div>
      </div>
    );
  }
}