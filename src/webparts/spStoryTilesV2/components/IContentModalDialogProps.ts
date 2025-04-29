import { IContentTile } from '../SpStoryTilesV2WebPart';

export interface IContentModalDialogProps {
  isOpen: boolean;
  onDismiss: () => void;
  contentData: IContentTile | null;
  isDarkTheme: boolean;
}