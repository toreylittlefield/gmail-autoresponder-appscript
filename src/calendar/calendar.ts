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

export function getUserCalendarsAndCurrentCalendar() {
  const calendars = CalendarApp.getAllOwnedCalendars();
  const listOfOwnerCalendarNames = calendars.map((calendar) => {
    return calendar.getName();
  });
  const currentCalendar = getSingleUserPropValue('currentCalendarName');
  return { currentCalendar, listOfOwnerCalendarNames };
}
