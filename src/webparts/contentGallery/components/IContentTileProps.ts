import { IContentTile } from '../ContentTileWebPart';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IContentTileProps {
  item: IContentTile;
  onOpenModal: (item: IContentTile) => void;
  context: WebPartContext;
}