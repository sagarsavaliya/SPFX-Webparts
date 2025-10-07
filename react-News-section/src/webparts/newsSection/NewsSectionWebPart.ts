import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'NewsSectionWebPartStrings';
import NewsSection from './components/NewsSection';
import { INewsSectionProps } from './components/INewsSectionProps';

export interface INewsSectionWebPartProps {
  description: string;
  layoutType: string;
  maxNewsItems: number;
  showCreateNewsButton: boolean;
  showImages: boolean;
  showAuthor: boolean;
  showPublishedDate: boolean;
}

export default class NewsSectionWebPart extends BaseClientSideWebPart<INewsSectionWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<INewsSectionProps> = React.createElement(
      NewsSection,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        layoutType: this.properties.layoutType || 'list',
        maxNewsItems: this.properties.maxNewsItems || 5,
        showCreateNewsButton: this.properties.showCreateNewsButton !== false,
        showImages: this.properties.showImages !== false,
        showAuthor: this.properties.showAuthor !== false,
        showPublishedDate: this.properties.showPublishedDate !== false
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown('layoutType', {
                  label: 'Layout Type',
                  options: [
                    { key: 'list', text: 'List' },
                    { key: 'hero', text: 'Hero' },
                    { key: 'carousel', text: 'Carousel' }
                  ],
                  selectedKey: this.properties.layoutType || 'list'
                }),
                PropertyPaneSlider('maxNewsItems', {
                  label: 'Maximum News Items',
                  min: 1,
                  max: 20,
                  value: this.properties.maxNewsItems || 5,
                  showValue: true,
                  step: 1
                })
              ]
            },
            {
              groupName: 'Display Options',
              groupFields: [
                PropertyPaneToggle('showCreateNewsButton', {
                  label: 'Show Create News Button',
                  checked: this.properties.showCreateNewsButton !== false
                }),
                PropertyPaneToggle('showImages', {
                  label: 'Show Images',
                  checked: this.properties.showImages !== false
                }),
                PropertyPaneToggle('showAuthor', {
                  label: 'Show Author',
                  checked: this.properties.showAuthor !== false
                }),
                PropertyPaneToggle('showPublishedDate', {
                  label: 'Show Published Date',
                  checked: this.properties.showPublishedDate !== false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
