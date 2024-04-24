declare interface IAttendancePortalWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'AttendancePortalWebPartStrings' {
  const strings: IAttendancePortalWebPartStrings;
  export = strings;
}
