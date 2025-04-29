import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneHorizontalRule
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/attachments";
import { IAttachmentInfo } from "@pnp/sp/attachments";
import { IDropdownOption } from '@fluentui/react/lib/Dropdown';

import * as strings from 'UserModalWebPartStrings';
import ContentTileGallery from './components/ContentTileGallery';
import { IContentTileGalleryProps } from './components/IContentTileGalleryProps';
import PropertyPaneListCreationRedirect from './PropertyPaneListCreationRedirect';

export interface IContentTileWebPartProps {
  title: string;
  listName: string;
  itemsPerPage: number;
  descriptionFieldName: string;
  imageFieldName: string;
  linkUrlFieldName: string;
}

export interface IContentTile {
  id: number;
  title: string;
  description: string;
  imageUrl: string;
  linkUrl: string;
}

export default class ContentTileWebPart extends BaseClientSideWebPart<IContentTileWebPartProps> {
  private _sp: SPFI;
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _contentItems: IContentTile[] = [];
  private _isLoading: boolean = true;
  private _availableLists: IDropdownOption[] = []; // Store available lists

  public onInit(): Promise<void> {
    // Initialize PnP JS
    this._sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));
    
    return Promise.all([
      this._getEnvironmentMessage().then(message => {
        this._environmentMessage = message;
      }),
      this._fetchAvailableLists()
    ]).then(() => {
      return Promise.resolve();
    });
  }

  public render(): void {
    this._isLoading = true;
    this._fetchContentFromList().then(() => {
      this._isLoading = false;
      this._renderWebPart();
    }).catch(error => {
      console.error("Error fetching content:", error);
      this._isLoading = false;
      this._renderWebPart();
    });
  }

  private _renderWebPart(): void {
    const element: React.ReactElement<IContentTileGalleryProps> = React.createElement(
      ContentTileGallery,
      {
        webPartTitle: this.properties.title || 'Content Gallery',
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        contentItems: this._contentItems,
        isLoading: this._isLoading,
        itemsPerPage: this.properties.itemsPerPage || 4,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private async _fetchAvailableLists(): Promise<void> {
    try {
      // Get all non-hidden lists from the site
      const lists = await this._sp.web.lists
        .filter("Hidden eq false and BaseTemplate eq 100")
        .select("Title, Id")
        .orderBy("Title")
        (); // Changed from .get() to () for PnP JS v3+
      
      // Convert lists to dropdown options
      this._availableLists = lists.map(list => ({
        key: list.Title,
        text: list.Title
      }));
      
      // Add default empty option
      this._availableLists.unshift({
        key: '',
        text: '- Select a list -'
      });

    } catch (error) {
      console.error("Error fetching SharePoint lists:", error);
      this._availableLists = [{
        key: '',
        text: '- Error loading lists -'
      }];
    }
  }

  private async _fetchContentFromList(): Promise<void> {
    if (!this.properties.listName) {
      this._contentItems = []; // No list selected, empty the array
      return;
    }

    try {
      // Set field names - with defaults if not provided
      const descFieldName = this.properties.descriptionFieldName || 'Description';
      const imageFieldName = this.properties.imageFieldName || 'Image';
      const linkUrlFieldName = this.properties.linkUrlFieldName || 'LinkUrl';

      // Get items from the list
      const items = await this._sp.web.lists.getByTitle(this.properties.listName).items.select(
        'ID',
        'Title',
        descFieldName,
        imageFieldName,
        linkUrlFieldName
      )();

      // Process items
      const processedItems: IContentTile[] = await Promise.all(
        items.map(async (item: any) => {
          let imageUrl = '';
          
          const a = this._sp.web.lists.getByTitle(this.properties.listName)
            .items
            .getById(item.ID);
          
          const itemAttachments: IAttachmentInfo[] = await a.attachmentFiles();
          
          try {
            const imageName = JSON.parse(item[imageFieldName]).fileName;
            const imageAttachment = itemAttachments.filter(attachemnt => attachemnt.FileName === imageName)[0];
            
            imageUrl = imageAttachment.ServerRelativeUrl;
          } catch (error) {
            console.warn(`Getting item image url failed ${item.ID}:`, error);
          }
          
          // Convert relative URLs to absolute URLs
          if (imageUrl && imageUrl.startsWith('/')) {
            const tenantUrl = this.context.pageContext.web.absoluteUrl.replace(this.context.pageContext.web.serverRelativeUrl,'');
            imageUrl = `${tenantUrl}${imageUrl}`;
          }
          
          // Use default image if no image was found
          if (!imageUrl) {
            imageUrl = require('./assets/welcome-light.png');
          }
          
          return {
            id: item.ID,
            title: item.Title || 'Untitled',
            description: item[descFieldName] || '',
            imageUrl: imageUrl,
            linkUrl: item[linkUrlFieldName] || ''
          };
        })
      );
      
      this._contentItems = processedItems;
    } catch (error) {
      console.error("Error fetching items from SharePoint list:", error);
      this._contentItems = [];
    }
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Web Part Title',
                  value: 'Content Gallery'
                }),
                PropertyPaneDropdown('listName', {
                  label: 'SharePoint List',
                  options: this._availableLists,
                  selectedKey: this.properties.listName
                }),
                new PropertyPaneListCreationRedirect(this.context),
                PropertyPaneHorizontalRule(),
                PropertyPaneDropdown('itemsPerPage', {
                  label: 'Tiles Per View',
                  options: [
                    { key: 1, text: '1' },
                    { key: 2, text: '2' },
                    { key: 3, text: '3' },
                    { key: 4, text: '4' }
                  ],
                  selectedKey: this.properties.itemsPerPage || 4
                }),
                PropertyPaneTextField('descriptionFieldName', {
                  label: 'Description Field Name',
                  description: 'Default: Description'
                }),
                PropertyPaneTextField('imageFieldName', {
                  label: 'Image Field Name',
                  description: 'Default: Image'
                }),
                PropertyPaneTextField('linkUrlFieldName', {
                  label: 'Link URL Field Name',
                  description: 'Default: LinkUrl'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}