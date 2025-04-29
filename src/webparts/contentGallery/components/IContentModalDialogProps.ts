import { IContentTile } from '../ContentTileWebPart';

export interface IContentModalDialogProps {
  isOpen: boolean;
  onDismiss: () => void;
  contentData: IContentTile | null;
  isDarkTheme: boolean;
}