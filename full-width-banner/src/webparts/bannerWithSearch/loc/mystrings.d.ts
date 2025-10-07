declare interface IBannerWithSearchWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ImageSourceFieldLabel: string;
  ImageUrlFieldLabel: string;
  ImagePickerFieldLabel: string;
  ImagePickerButtonLabel: string;
  GreetingTextSizeFieldLabel: string;
  GreetingTextFieldLabel: string;
  SearchResultTypeFieldLabel: string;
  ImageSettingsFieldLabel: string;
  LimitToSiteFieldLabel: string;
  OnText: string;
  OffText: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'BannerWithSearchWebPartStrings' {
  const strings: IBannerWithSearchWebPartStrings;
  export = strings;
}
