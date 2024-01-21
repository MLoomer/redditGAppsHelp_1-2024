const SHEETNAME = 'main';
const CALNAME = 'test';//'SHAREDCALENDERNAME';
const RANGE = 'A1:AB';

function onOpen(e) {
  setUpOption();
}

function setUpOption() {
  SpreadsheetApp.getUi()
    .createMenu('Event Creater')
    .addItem('Run', 'main')
    .addToUi();
}

function main() {
  let sheet = new Spreadsheet(SHEETNAME);
  let calendar = new Calendar(CALNAME);
  sheet.eventRunner(calendar);
}


class Spreadsheet {
  constructor(name) {
    this.SS = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = this.SS.getSheetByName(name);
    if (!this.sheet) {
      throw new Error('No Matching sheet to name ' + name);
    }
    this.data = this.sheet.getRange(RANGE);
    this.values = this.data.getValues();
    this.headers = this.values.shift();
  }

  updateSheetStatus(status, index) {
    this.sheet.getRange(index + 2, 1).setValue(status);
  }

  updateSheetEvent(id, index) {
    let range = this.sheet.getRange(index + 2, 28);
    let value = range.getDisplayValues();
    if (value != "") {
      range.setValue(value + "," + id);
    } else {
      range.setValue(id);
    }
  }

  eventRunner(calendar) {
    for (let [index, row] of this.values.entries()) {
      let status = row[0];
      if (status == "") {
        //Do nothing, end function if any blank rows
        console.log("break");
        break;
      }
      if (status == "Yes") {
        //Do nothing, move to next row
        continue;
      }
      if (status == "No" || status == "Modify") {
        this.eventWorker(calendar, index, row);
      }
    }
  }

  eventWorker(calendar, index, row) {
    let eventFields = 3; //number of columns linked to 1 event
    let [status, , , , email, username, , reference, title, numOfAttends, , , colorSelector, , , ...meetingData] = row;
    let meetingIDs = meetingData.pop(); //last element contains the IDs
    for (let i = 0, j = 0; i < meetingData.length; i += eventFields, j++) {
      if (meetingData[i] != "") {
        let reminderID;
        let [start, end, times] = meetingData.slice(i, i + eventFields);
        let color;
        switch (colorSelector.toLowerCase()) {
          case "yes":
            color = "8";
            break;
          case "no":
            color = "3";
            break;
          default:
            color = null;
            break;
        }
        try {
          start = new Date(Date.parse(start));
          end = new Date(Date.parse(end));
        } catch (e) {
          throw new Error("Could not parse dates into time: " + start + " " + end);
        }
        let eventObj = {
          start: start,
          end: end,
          times: times,
          participants: email,
          title: title,
          reference: reference,
          username: username,
          color: color,
        };
        if (status == "No") {
          let eventID = calendar.addEvent(eventObj);
          reminderID = calendar.addReminderEvent(eventObj);
          this.updateSheetEvent(eventID + ":" + reminderID, index);
          this.updateSheetStatus('Yes', index);

        }
        if (status == "Modify") {
          let eventIDs = meetingIDs.split(',')[j];
          eventObj.pEventID = eventIDs.split(":")[0];
          eventObj.sEventID = eventIDs.split(":")[1];
          calendar.modifyEvent(eventIDPrimary, eventObj);
          calendar.modifyReminderEvent(eventIDSecondary, eventObj);
          this.updateSheetStatus('Yes', index);
        }
      }
    }

  }
}

class Calendar {
  constructor(calendarName) {
    this.calendar;
    const calendars = CalendarApp.getCalendarsByName(calendarName);
    if (!calendars) {
      throw new Error('No Matching calendar to name ' + calendarName);
    }
    if (calendars.length > 1) {
      throw new Error('Too many matching calendar names.')
    }
    this.calendar = calendars[0];
  }

  eventCreater(eventObj) {
    let event = this.calendar.createEvent(eventObj.eventTitle, eventObj.time.start, eventObj.time.end, eventObj.options)
    if (eventObj.color) {
      event.setColor(eventObj.color);
    }
    return event.getId();
  }

  addEvent(eventObj) {
    eventObj.eventTitle = `${eventObj.title}, (${eventObj.reference})`;
    eventObj.time = this.dateAndTimeParse(eventObj.start, eventObj.end, eventObj.times);
    eventObj.options = {
      description: eventObj.username,
      sendInvites: true,
      guests: eventObj.email,
    }
    return this.eventCreater(eventObj);
  }

  addReminderEvent(eventObj) {
    let remObj = { ...eventObj }
    remObj.eventTitle = `${eventObj.title}, (${eventObj.reference}) - 48Hr Reminder`;
    remObj.start = new Date(new Date(eventObj.start).setDate(eventObj.start.getDate() - 2));
    remObj.end = new Date(new Date(eventObj.end).setDate(eventObj.end.getDate() - 2));
    remObj.time = this.dateAndTimeParse(remObj.start, remObj.end, remObj.times);
    remObj.color = null;
    remObj.options = {
      description: eventObj.username,
    }
    return this.eventCreater(remObj);
  }

  modifyEventCreator(event, eventObj) {
    event.setTitle(eventObj.eventTitle);
    event.setDescription(eventObj.username);
    event.setTime(eventObj.time.start, eventObj.time.end);
  }

  modifyEvent(id, eventObj) {
    let event = this.calendar.getEventById(id);
    eventObj.eventTitle = `${eventObj.title}, (${eventObj.reference})`;
    eventObj.time = this.dateAndTimeParse(eventObj.start, eventObj.end, eventObj.times);
    this.modifyEventCreator(event, eventObj);
  }

  modifyReminderEvent(id, eventObj) {
    let remObj = { ...eventObj }
    let event = this.calendar.getEventById(id);
    remObj.eventTitle = `${eventObj.title}, (${eventObj.reference})`;
    remObj.start = new Date(new Date(eventObj.start).setDate(eventObj.start.getDate() - 2));
    remObj.end = new Date(new Date(eventObj.end).setDate(eventObj.end.getDate() - 2));
    remObj.time = this.dateAndTimeParse(remObj.start, remObj.end, remObj.times);
    this.modifyEventCreator(event, remObj);
  }

  dateAndTimeParse(start, end, times) {
    const regex = /\s*-\s*/;
    const timeArr = times.split(regex);
    if (timeArr.length != 2) {
      throw new Error(`Could not split time into start and end time for field ${times}. Alteration needed.`)
    }
    let startDate = this.setDateTime(start, timeArr[0]);
    let endDate = this.setDateTime(end, timeArr[1]);
    if (isNaN(startDate) || isNaN(endDate)) {
      throw new Error("Datetime was not compatible, check for date time of " + start + " " + end + " " + times)
    }
    return { start: startDate, end: endDate }
  }

  setDateTime(date, time) {
    let index = time.indexOf(":");
    let index2 = time.indexOf(" ");

    let hours = time.substring(0, index);
    let minutes = time.substring(index + 1, index2);

    let mer = time.substring(index2 + 1, time.length);
    if (mer.toLowerCase() == "pm") {
      hours = parseInt(hours) + 12;
    }
    date.setHours(hours);
    date.setMinutes(minutes);
    date.setSeconds("00");
    return date;
  }
}