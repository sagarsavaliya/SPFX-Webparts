import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldFilePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';

import * as strings from 'BannerWithSearchWebPartStrings';
import BannerWithSearch from './components/BannerWithSearch';
import { IBannerWithSearchProps } from './components/IBannerWithSearchProps';

export interface IBannerWithSearchWebPartProps {
  imageSource: string;
  greetingTextSize: string;
  searchResultType: string;
  imageSettings: string;
  imageUrl: string;
  greetingText: string;
  limitToSite: boolean;
  imagePickerResult?: any;
}

export default class BannerWithSearchWebPart extends BaseClientSideWebPart<IBannerWithSearchWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IBannerWithSearchProps> = React.createElement(
      BannerWithSearch,
      {
        imageSource: this.properties.imageSource,
        greetingTextSize: this.properties.greetingTextSize,
        searchResultType: this.properties.searchResultType,
        imageSettings: this.properties.imageSettings,
        imageUrl: this.properties.imageUrl,
        greetingText: this.properties.greetingText,
        limitToSite: this.properties.limitToSite,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
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
                PropertyPaneDropdown('imageSource', {
                  label: strings.ImageSourceFieldLabel,
                  options: [
                    { key: 'SharePoint', text: 'SharePoint' },
                    { key: 'Url', text: 'URL' }
                  ]
                }),
                ...(this.properties.imageSource === 'SharePoint' ? [PropertyFieldFilePicker('imageUrl', {
                  context: this.context as any,
                  label: strings.ImagePickerFieldLabel,
                  buttonLabel: strings.ImagePickerButtonLabel,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'imageFilePicker',
                  accepts: ['.png', '.jpg', '.jpeg', '.gif'],
                  filePickerResult: this.properties.imagePickerResult,
                  onChanged: (filePickerResult: any) => {
                    this.properties.imagePickerResult = filePickerResult;
                    this.properties.imageUrl = filePickerResult?.fileAbsoluteUrl || this.properties.imageUrl;
                    this.render();
                  },
                  onSave: (filePickerResult: any) => {
                    this.properties.imagePickerResult = filePickerResult;
                    this.properties.imageUrl = filePickerResult?.fileAbsoluteUrl || this.properties.imageUrl;
                    this.render();
                  },
                  hideWebSearchTab: true,
                  hideRecentTab: false,
                  hideLinkUploadTab: true,
                  hideOneDriveTab: false,
                  hideStockImages: true
                })] : []),
                ...(this.properties.imageSource === 'Url' ? [PropertyPaneTextField('imageUrl', {
                  label: strings.ImageUrlFieldLabel,
                  multiline: false,
                  resizable: false
                })] : []),
                PropertyPaneDropdown('imageSettings', {
                  label: strings.ImageSettingsFieldLabel,
                  options: [
                    { key: 'Crop', text: 'Crop' },
                    { key: 'ZoomIn', text: 'Zoom In' },
                    { key: 'ZoomOut', text: 'Zoom Out' }
                  ]
                }),
                PropertyPaneDropdown('greetingTextSize', {
                  label: strings.GreetingTextSizeFieldLabel,
                  options: [
                    { key: 'Small', text: 'Small' },
                    { key: 'Medium', text: 'Medium' },
                    { key: 'Large', text: 'Large' }
                  ]
                }),
                PropertyPaneTextField('greetingText', {
                  label: strings.GreetingTextFieldLabel,
                  multiline: false,
                  resizable: false
                }),
                PropertyPaneDropdown('searchResultType', {
                  label: strings.SearchResultTypeFieldLabel,
                  options: [
                    { key: 'SharePoint', text: 'SharePoint Search Page' },
                    { key: 'Adhoc', text: 'Adhoc Result' }
                  ]
                }),
                PropertyPaneToggle('limitToSite', {
                  label: strings.LimitToSiteFieldLabel,
                  onText: strings.OnText,
                  offText: strings.OffText
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
