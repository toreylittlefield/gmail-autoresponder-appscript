import { getSingleUserPropValue } from '../properties-service/properties-service';
import { getCurrentMonthAndNextMonthDates } from '../utils/utils';

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

function normalizeCalendarEvents(event: GoogleAppsScript.Calendar.CalendarEvent, userEmail?: string) {
  const {
    getDescription,
    getDateCreated,
    getTitle,
    getStartTime,
    getEndTime,
    getId,
    getGuestList,
    getEventSeries,
    removeGuest,
  } = event;
  const guests = getGuestList(false);

  userEmail && removeGuest(userEmail);

  const guestList = guests.map((guest) => {
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

    return { guestEmail: getEmail(), guestStatus: guestStatus, guestName: getName() };
  });
  const { getId: seriesId, getDescription: seriesDescription, getTitle: seriesTitle } = getEventSeries();
  const eventObject = {
    eventId: getId(),
    eventDateCreated: getDateCreated(),
    eventStartTime: getStartTime(),
    eventEndTime: getEndTime(),
    eventTitle: getTitle(),
    eventDescription: getDescription(),
    guestList: guestList,
    eventSeriesId: seriesId(),
    eventSeriesDescription: seriesDescription(),
    eventSeriesTitle: seriesTitle(),
  };
  return eventObject;
}

export function getAllCalendarEvents() {
  const userSelectedCalendar = getSingleUserPropValue('currentCalendarName');
  const userEmail = Session.getActiveUser().getEmail();
  if (!userSelectedCalendar)
    throw Error(
      `No user selected calendar found in user configuration, call ${deleteEventsWithNoAttendingGuests.name}`
    );
  const [calendar] = CalendarApp.getOwnedCalendarsByName(userSelectedCalendar);
  const { currentDate, nextDate } = getCurrentMonthAndNextMonthDates();

  const events = calendar.getEvents(currentDate, nextDate);
  return events.map((event) => normalizeCalendarEvents(event, userEmail));
}

export function getSingleCalendarEventById(eventId: string) {
  const userSelectedCalendar = getSingleUserPropValue('currentCalendarName');
  if (!userSelectedCalendar) throw Error(`No user selected calendar prop set, ${getSingleCalendarEventById.name}`);
  const [calendar] = CalendarApp.getCalendarsByName(userSelectedCalendar);
  const event = calendar.getEventById(eventId);
  return normalizeCalendarEvents(event);
}

export function getUserCalendarsAndCurrentCalendar() {
  const calendars = CalendarApp.getAllOwnedCalendars();
  const listOfOwnerCalendarNames = calendars.map((calendar) => {
    return calendar.getName();
  });
  const currentCalendar = getSingleUserPropValue('currentCalendarName');
  return { currentCalendar, listOfOwnerCalendarNames };
}
