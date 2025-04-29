import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IContentTile } from '../ContentTileWebPart';

export interface IContentTileGalleryProps {
  webPartTitle: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  contentItems: IContentTile[];
  isLoading: boolean;
  itemsPerPage: number;
  context: WebPartContext;
}