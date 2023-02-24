import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'NewsWebPartStrings';
import News from './components/News';
import { INewsProps } from './components/INewsProps';

export interface INewsWebPartProps {
  description: string;
  webpartTitle: string;
  orderBy: string;
  order: string;
  enableLike: boolean;
  listName: string;
  moreNewsPageUrl: string;
  newsDetailsPageUrl: string;
  SetHeight:string;
  emptyData:string;
  speedOfCarousel: any;
  noOfSlides: any;
  itemsLimit: any;
  bgColor: any;
  defaultThumbnail: string;
}

export default class NewsWebPart extends BaseClientSideWebPart<INewsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<INewsProps> = React.createElement(
      News,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        webpartTitle: this.properties.webpartTitle,
        orderBy: this.properties.orderBy,
        order: this.properties.order,
        noOfSlides: this.properties.noOfSlides,
        enableLike: this.properties.enableLike,
        context: this.context,
        listName: this.properties.listName,
        moreNewsPageUrl: this.properties.moreNewsPageUrl,
        newsDetailsPageUrl: this.properties.newsDetailsPageUrl,
        speedOfCarousel: this.properties.speedOfCarousel,
        SetHeight: this.properties.SetHeight,
        emptyData: this.properties.emptyData,
        itemsLimit: this.properties.itemsLimit,
        bgColor: this.properties.bgColor,
        defaultThumbnail: this.properties.defaultThumbnail
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
  
  protected onAfterPropertyPaneChangesApplied(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let orderByOptions: IPropertyPaneDropdownOption[] = [
      {
        key: "Created",
        text: "Created Date"
      },
      {
        key: "Modified",
        text: "Modified Date"
      },
      {
        key: "PublishDate",
        text: "Publish Date"
      },
      {
        key: "Title",
        text: "Title"
      }
    ];

    let orderOptions: IPropertyPaneDropdownOption[] = [
      {
        key: "Ascending",
        text: "Ascending"
      },
      {
        key: "Descending",
        text: "Descending"
      }
    ];

    let slideNumberOptions: IPropertyPaneDropdownOption[] = [
      {
        key: 3,
        text: "3"
      },
      {
        key: 6,
        text: "6"
      },
      {
        key: 9,
        text: "9"
      },
      {
        key: 12,
        text: "12"
      }
    ];
    return {
      pages: [
        {
          // header: {
          //   description: strings.PropertyPaneDescription
          // },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('webpartTitle', {
                  label: strings.WebpartTitleFieldLabel
                }),
                PropertyPaneTextField('listName', {
                  label: strings.listNameLabel
                }),
                PropertyPaneDropdown('orderBy', {
                  label: strings.OrderByFieldLabel,
                  options: orderByOptions
                }),
                PropertyPaneDropdown('order', {
                  label: strings.OrderFieldLabel,
                  options: orderOptions
                }),
                PropertyPaneDropdown('itemsLimit', {
                  label: strings.NoOfItemsFieldLabel,
                  options: slideNumberOptions
                }),
                PropertyPaneTextField('emptyData', {
                  label: strings.emptyData
                }),
                PropertyPaneSlider('speedOfCarousel', {
                  label: strings.speedOfCarouselFieldLabel,
                  min: 1, 
                  max: 10,
                  value: 5 
                }),
                PropertyPaneTextField('SetHeight', {
                  label: strings.componentHeightFieldLabel
                }),
                PropertyPaneTextField('bgColor', {
                  label: strings.bgColorFieldLabel
                }),
                PropertyPaneTextField('defaultThumbnail', {
                  label: strings.DefaultThumbnailLabel
                }),
                PropertyPaneTextField('moreNewsPageUrl', {
                  label: strings.MoreNewsPageUrlFieldLabel
                }),
                PropertyPaneTextField('newsDetailsPageUrl', {
                  label: strings.NewsDetailsPageUrlFieldLabel
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
