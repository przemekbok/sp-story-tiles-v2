import { IContentTile } from '../SpStoryTilesV2WebPart';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IContentTileProps {
  item: IContentTile;
  context: WebPartContext;
}