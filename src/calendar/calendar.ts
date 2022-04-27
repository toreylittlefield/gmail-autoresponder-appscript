import { calendarEventsMap } from '../global/maps';
import { getSingleUserPropValue } from '../properties-service/properties-service';
import { getSheetByName, writeEventsToCalendarSheet } from '../sheets/sheets';
import {
  getCurrentMonthAndNextMonthDates,
  getDomainFromEmailAddress,
  getPhoneNumbersFromString,
  initialGlobalMap,
} from '../utils/utils';
import { CALENDAR_EVENTS_SHEET_NAME } from '../variables/publicvariables';

export function deleteEventsWithNoAttendingGuests() {
  const userSelectedCalendar = getSingleUserPropValue('currentCalendarName');
  if (!userSelectedCalendar)
    throw Error(
      `No user selected calendar found in user configuration, call ${deleteEventsWithNoAttendingGuests.name}`
    );
  const [currentCalendar] = CalendarApp.getOwnedCalendarsByName(userSelectedCalendar);
  const { currentDate, nextDate } = getCurrentMonthAndNextMonthDates();
  const events = currentCalendar.getEvents(currentDate, nextDate);
  events.forEach((event) => {
    const guestList = event.getGuestList(false);
    const allGuestDeclined = guestList.every((guest) => {
      const status = guest.getGuestStatus();
      if (status === CalendarApp.GuestStatus.NO) {
        return true;
      }
      return false;
    });
    if (allGuestDeclined) {
      event.deleteEvent();
    }
  });
}

function getPersonNameCompanyNameAndGuestPhoneNumberFromCalendarDescription(eventDescription: string) {
  const guestPhoneNumber = getPhoneNumbersFromString(eventDescription);
  let companyName: RegExpMatchArray | null | string = eventDescription.match(/(?<=Company Name[\D\s\w]*\n)(.*)/gim);
  if (companyName) {
    companyName = companyName[0].trim();
  } else {
    companyName = '';
  }
  let guestName: RegExpMatchArray | null | string = eventDescription.match(/(?<=Booked By[\D\s\w]*\n)([\w\s]+\n)/gim);
  if (guestName) {
    guestName = guestName[0].trim();
  } else {
    guestName = '';
  }
  return { guestPhoneNumber, companyName, guestName };
}

function getFirstGuest(guestList: GoogleAppsScript.Calendar.EventGuest[]) {
  const list = guestList.slice(0, 1).map((guest) => {
    const { getEmail, getGuestStatus, getName } = guest;

    let guestStatus: 'YES' | 'NO' | 'INVITED' | 'MAYBE' | 'OWNER' | '' = '';
    const statusEnum = getGuestStatus();
    if (statusEnum === CalendarApp.GuestStatus.YES) {
      guestStatus = 'YES';
    }

    if (statusEnum === CalendarApp.GuestStatus.NO) {
      guestStatus = 'NO';
    }
    if (statusEnum === CalendarApp.GuestStatus.INVITED) {
      guestStatus = 'INVITED';
    }
    if (statusEnum === CalendarApp.GuestStatus.MAYBE) {
      guestStatus = 'MAYBE';
    }
    if (statusEnum === CalendarApp.GuestStatus.OWNER) {
      guestStatus = 'OWNER';
    }
    const guestDomain = getDomainFromEmailAddress(getEmail());
    return { guestEmail: getEmail(), guestStatus: guestStatus, guestName: getName(), guestDomain: guestDomain };
  });

  if (list.length > 0) {
    return list[0];
  }
  return { guestEmail: '', guestList: '', guestStatus: '', guestName: '', guestDomain: '' };
}

function normalizeCalendarEvents(
  event: GoogleAppsScript.Calendar.CalendarEvent,
  calendarId: string,
  userEmail?: string
) {
  const {
    getDescription,
    getDateCreated,
    getLastUpdated,
    getTitle,
    getStartTime,
    getEndTime,
    getId,
    getGuestList,
    removeGuest,
  } = event;
  const guestList = getGuestList(false);

  const { guestName, guestStatus, guestEmail, guestDomain } = getFirstGuest(guestList);

  const eventDescription = getDescription();
  const {
    companyName,
    guestPhoneNumber,
    guestName: guestNameFromDescription,
  } = getPersonNameCompanyNameAndGuestPhoneNumberFromCalendarDescription(eventDescription);

  userEmail && removeGuest(userEmail);

  const splitEventId = getId().split('@');
  const eventURL =
    'https://www.google.com/calendar/event?eid=' +
    Utilities.base64Encode(splitEventId[0] + ' ' + calendarId).replace('==', '');

  return {
    eventId: getId(),
    eventLastUpdated: getLastUpdated(),
    eventDateCreated: getDateCreated(),
    eventStartTime: getStartTime(),
    eventEndTime: getEndTime(),
    eventTitle: getTitle(),
    eventDescription: eventDescription,
    companyName,
    guestDomain,
    guestName: guestNameFromDescription || guestName,
    guestPhoneNumber,
    guestEmail,
    guestStatus,
    numberOfGuest: guestList.length,
    eventURL,
  };
}

export function getAllCalendarEventsFor30Days() {
  const userSelectedCalendar = getSingleUserPropValue('currentCalendarName');
  const userEmail = Session.getActiveUser().getEmail();
  if (!userSelectedCalendar)
    throw Error(
      `No user selected calendar found in user configuration, call ${deleteEventsWithNoAttendingGuests.name}`
    );
  const [calendar] = CalendarApp.getOwnedCalendarsByName(userSelectedCalendar);
  const { currentDate, nextDate } = getCurrentMonthAndNextMonthDates();

  const events = calendar.getEvents(currentDate, nextDate);
  const list = events.map((event) => normalizeCalendarEvents(event, userEmail));
  return list;
}

function getUserCalendar() {
  const userSelectedCalendar = getSingleUserPropValue('currentCalendarName');
  if (!userSelectedCalendar) throw Error(`No user selected calendar prop set, ${getSingleCalendarEventById.name}`);
  const [calendar] = CalendarApp.getCalendarsByName(userSelectedCalendar);
  return { calendar, calendarId: calendar.getId() };
}

export function getSingleCalendarEventById(
  calendar: GoogleAppsScript.Calendar.Calendar,
  calendarId: string,
  eventId: string
) {
  const event = calendar.getEventById(eventId);
  return normalizeCalendarEvents(event, calendarId);
}

export function getUserCalendarsAndCurrentCalendar() {
  const calendars = CalendarApp.getAllOwnedCalendars();
  const listOfOwnerCalendarNames = calendars.map((calendar) => {
    return calendar.getName();
  });
  const currentCalendar = getSingleUserPropValue('currentCalendarName');
  return { currentCalendar, listOfOwnerCalendarNames };
}

export type ValidRowToWriteInCalendarSheet = [
  CalendarEventId: string,
  EventLastUpdatedTime: GoogleAppsScript.Base.Date,
  EventCreatedTime: GoogleAppsScript.Base.Date,
  EventStartTime: GoogleAppsScript.Base.Date,
  EventEndTime: GoogleAppsScript.Base.Date,
  EventTitle: string,
  EventDescription: string,
  GuestCompanyName: string,
  GuestDomain: string,
  GuestName: string,
  GuestEmail: string,
  GuestPhoneNumber: string,
  GuestStatus: string,
  NumberOfGuests: number,
  EventURL: string
];

export function extractCalendarDataForCalendarSheet() {
  initialGlobalMap('calendarEventsMap');

  updateExistingCalendarsInSheet();

  const eventsForNext30Days = getAllCalendarEventsFor30Days();

  const validRowsToWriteInCalendarSheet: ValidRowToWriteInCalendarSheet[] = [];

  eventsForNext30Days.forEach(
    ({
      companyName,
      eventDateCreated,
      eventDescription,
      eventEndTime,
      eventId,
      eventLastUpdated,
      eventStartTime,
      eventTitle,
      guestDomain,
      guestEmail,
      guestName,
      guestPhoneNumber,
      guestStatus,
      numberOfGuest,
      eventURL,
    }) => {
      if (calendarEventsMap.has(eventId)) return;

      validRowsToWriteInCalendarSheet.push([
        eventId,
        eventLastUpdated,
        eventDateCreated,
        eventStartTime,
        eventEndTime,
        eventTitle,
        eventDescription,
        companyName,
        guestDomain,
        guestName,
        guestPhoneNumber,
        guestEmail,
        guestStatus,
        numberOfGuest,
        eventURL,
      ]);
    }
  );
  writeEventsToCalendarSheet(validRowsToWriteInCalendarSheet);
}

function updateExistingCalendarsInSheet() {
  const calendarSheet = getSheetByName(CALENDAR_EVENTS_SHEET_NAME);
  if (!calendarSheet)
    throw Error(`Cannot find ${CALENDAR_EVENTS_SHEET_NAME} in ${updateExistingCalendarsInSheet.name}`);
  const { calendar, calendarId } = getUserCalendar();
  calendarEventsMap.forEach(({ rowNumber }, eventIdInMap) => {
    const {
      companyName,
      eventDateCreated,
      eventDescription,
      eventEndTime,
      eventId,
      eventLastUpdated,
      eventStartTime,
      eventTitle,
      guestDomain,
      guestEmail,
      guestName,
      guestPhoneNumber,
      guestStatus,
      numberOfGuest,
      eventURL,
    } = getSingleCalendarEventById(calendar, calendarId, eventIdInMap);
    const validRowInCalendarSheet: ValidRowToWriteInCalendarSheet = [
      eventId,
      eventLastUpdated,
      eventDateCreated,
      eventStartTime,
      eventEndTime,
      eventTitle,
      eventDescription,
      companyName,
      guestDomain,
      guestName,
      guestPhoneNumber,
      guestEmail,
      guestStatus,
      numberOfGuest,
      eventURL,
    ];
    calendarSheet.getRange(rowNumber, 1, 1, validRowInCalendarSheet.length).setValues([validRowInCalendarSheet]);
  });
}
