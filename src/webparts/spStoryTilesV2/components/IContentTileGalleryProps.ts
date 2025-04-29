import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IContentTile } from '../SpStoryTilesV2WebPart';

export interface IContentTileGalleryProps {
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  contentItems: IContentTile[];
  isLoading: boolean;
  itemsPerPage: number;
  context: WebPartContext;
}