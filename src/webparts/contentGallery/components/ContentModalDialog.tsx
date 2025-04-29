import * as React from 'react';
import styles from './ContentTile.module.scss';
import { IContentModalDialogProps } from './IContentModalDialogProps';
import { Modal } from '@fluentui/react/lib/Modal';
import { IconButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { escape } from '@microsoft/sp-lodash-subset';

const ContentModalDialog: React.FC<IContentModalDialogProps> = (props) => {
  const { isOpen, onDismiss, contentData, isDarkTheme } = props;

  if (!contentData) {
    return null;
  }

  const handleLinkClick = (): void => {
    if (contentData.linkUrl) {
      window.open(contentData.linkUrl, '_blank');
    }
  };

  return (
    <Modal
      isOpen={isOpen}
      onDismiss={onDismiss}
      isBlocking={false}
      containerClassName={`${styles.modalContainer} ${isDarkTheme ? styles.dark : ''}`}
    >
      <div className={styles.modalHeader}>
        <IconButton
          className={styles.closeButton}
          iconProps={{ iconName: 'Cancel' }}
          ariaLabel="Close"
          onClick={onDismiss}
        />
      </div>
      <div className={styles.modalContent}>
        <div className={styles.contentHeader}>
          <h2 className={styles.modalTitle}>{escape(contentData.title)}</h2>
        </div>

        <div className={styles.modalImageContainer}>
          <img src={contentData.imageUrl} alt={contentData.title} className={styles.modalImage} />
        </div>
        
        <div className={styles.modalSection}>
          <p className={styles.sectionText}>{escape(contentData.description)}</p>
        </div>
        
        {contentData.linkUrl && (
          <div className={styles.modalActions}>
            <PrimaryButton 
              text="Open Link" 
              onClick={handleLinkClick} 
              iconProps={{ iconName: 'Link' }}
            />
          </div>
        )}
      </div>
    </Modal>
  );
};

export default ContentModalDialog;