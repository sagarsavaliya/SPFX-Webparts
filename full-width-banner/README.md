# Full Width Banner WebPart

A modern SharePoint Framework (SPFx) webpart that provides a full-width banner with search functionality for SharePoint communication sites.

## Summary

This webpart creates a beautiful full-width banner with customizable background images, greeting text, and integrated search functionality. It supports both SharePoint search and custom adhoc search with real-time results display.

## Features

- **Full-width banner** with customizable background image
- **Greeting text** with user's display name and configurable size
- **Search functionality** with real-time results
- **Property pane configuration** for easy customization
- **Responsive design** that works on all devices
- **SharePoint Search API integration**
- **People search** with profile information
- **Document search** with file details
- **Loading states** and smooth animations

## Configuration Options

### Image Settings
- **Image Source**: Choose between SharePoint library or URL
- **Image Settings**: Crop, Zoom In, or Zoom Out
- **Image URL**: Direct URL to background image

### Text Settings
- **Greeting Text Size**: Small, Medium, or Large
- **Greeting Text**: Customizable greeting message

### Search Settings
- **Search Result Type**: SharePoint Search Page or Adhoc Results
- **Search Results**: Real-time search with loading indicators

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.21.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

- Node.js (v16.x or v18.x)
- npm or yarn
- SharePoint Framework development environment

## Installation

1. Clone this repository
2. Run `npm install` to install dependencies
3. Run `gulp serve` to start the development server
4. Add the webpart to your SharePoint page

## Building for Production

1. Run `gulp build` to build the project
2. Run `gulp bundle --ship` to create the production bundle
3. Run `gulp package-solution --ship` to create the .sppkg file
4. Upload the .sppkg file to your SharePoint App Catalog

## Usage

1. Add the webpart to your SharePoint page
2. Configure the properties in the property pane:
   - Set your background image source and URL
   - Choose greeting text size and content
   - Select search result type
   - Configure image display settings
3. Save and publish the page

## Browser Support

- Chrome (latest)
- Firefox (latest)
- Edge (latest)
- Safari (latest)

## Dependencies

- SharePoint Framework 1.21.1
- React 17.0.1
- Fluent UI React 8.x
- TypeScript 4.7.4

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | January 2024     | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
