declare interface IResourceCalendarWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  calStrings: {
    days: string[];
    months: string[];
    shortMonths: string[];
    shortDays: string[];
    goToToday: string;
    weekNumberFormatString: string;
    prevMonthAriaLabel: string;
    nextMonthAriaLabel: string;
    prevYearAriaLabel: string;
    nextYearAriaLabel: string;
    prevYearRangeAriaLabel: string;
    nextYearRangeAriaLabel: string;
    closeButtonAriaLabel: string;
  }
}

declare module 'ResourceCalendarWebPartStrings' {
  const strings: IResourceCalendarWebPartStrings;
  export = strings;
}
