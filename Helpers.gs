/**
 * Formats the date and time according to the format specified in the configuration.
 *
 * @param {string} date The date to be formatted.
 * @return {string} The formatted date string.
 */
function formatDate(date) {
  const year = date.slice(0,4);
  const month = date.slice(5,7);
  const day = date.slice(8,10);
  let formattedDate;

  if (dateFormat == "YYYY/MM/DD") {
    formattedDate = year + "/" + month + "/" + day
  }
  else if (dateFormat == "DD/MM/YYYY") {
    formattedDate = day + "/" + month + "/" + year
  }
  else if (dateFormat == "MM/DD/YYYY") {
    formattedDate = month + "/" + day + "/" + year
  }
  else if (dateFormat == "YYYY-MM-DD") {
    formattedDate = year + "-" + month + "-" + day
  }
  else if (dateFormat == "DD-MM-YYYY") {
    formattedDate = day + "-" + month + "-" + year
  }
  else if (dateFormat == "MM-DD-YYYY") {
    formattedDate = month + "-" + day + "-" + year
  }
  else if (dateFormat == "YYYY.MM.DD") {
    formattedDate = year + "." + month + "." + day
  }
  else if (dateFormat == "DD.MM.YYYY") {
    formattedDate = day + "." + month + "." + year
  }
  else if (dateFormat == "MM.DD.YYYY") {
    formattedDate = month + "." + day + "." + year
  }

  if (date.length < 11) {
    return formattedDate
  }

  const time = date.slice(11,16)
  const timeZone = date.slice(19)

  return formattedDate + " at " + time + " (UTC" + (timeZone == "Z" ? "": timeZone) + ")"
}


/**
 * Takes an intended frequency in minutes and adjusts it to be the closest
 * acceptable value to use Google "everyMinutes" trigger setting (i.e. one of
 * the following values: 1, 5, 10, 15, 30).
 *
 * @param {?integer} The manually set frequency that the user intends to set.
 * @return {integer} The closest valid value to the intended frequency setting. Defaulting to 15 if no valid input is provided.
 */
function getValidTriggerFrequency(origFrequency) {
  if (!origFrequency > 0) {
    Logger.log("No valid frequency specified. Defaulting to 15 minutes.");
    return 15;
  }

  // Limit the original frequency to 1440
  origFrequency = Math.min(origFrequency, 1440);

  var acceptableValues = [5, 10, 15, 30].concat(
    Array.from({ length: 24 }, (_, i) => (i + 1) * 60)
  ); // [5, 10, 15, 30, 60, 120, ..., 1440]

  // Find the smallest acceptable value greater than or equal to the original frequency
  var roundedUpValue = acceptableValues.find(value => value >= origFrequency);

  Logger.log(
    "Intended frequency = " + origFrequency + ", Adjusted frequency = " + roundedUpValue
  );
  return roundedUpValue;
}

String.prototype.includes = function(phrase){
  return this.indexOf(phrase) > -1;
}

/**
 * Takes an array of ICS calendars and target Google calendars and combines them
 *
 * @param {Array.string} calendarMap - User-defined calendar map
 * @return {Array.string} Condensed calendar map
 */
function condenseCalendarMap(calendarMap){
  var result = [];
  for (var mapping of calendarMap){
    var index = -1;
    for (var i = 0; i < result.length; i++){
      if (result[i][0] == mapping[1]){
        index = i;
        break;
      }
    }

    if (index > -1)
      result[index][1].push([mapping[0],mapping[2]]);
    else
      result.push([ mapping[1], [[mapping[0],mapping[2]]] ]);
  }

  return result;
}

/**
 * Removes all triggers for the script's 'startSync' and 'install' function.
 */
function deleteAllTriggers(){
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++){
    if (["startSync","install","main"].includes(triggers[i].getHandlerFunction())){
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Fetches events from the specified source Google Calendars.
 *
 * @param {Array} sourceCalendarConfigs - Array with source calendar configs [[calendarId, colorId], ...]
 * @return {Array} Array of event objects with colorId and source calendar name
 */
function fetchSourceCalendars(sourceCalendarConfigs){
  var result = [];

  for (var source of sourceCalendarConfigs){
    var sourceCalendarIdOrName = source[0];
    var colorId = source[1];

    try {
      // Find the source calendar by ID or name
      var sourceCalendar = findSourceCalendar(sourceCalendarIdOrName);
      if (sourceCalendar == null) {
        Logger.log("[ERROR] Could not find source calendar: " + sourceCalendarIdOrName);
        reportOverallFailure = true;
        continue;
      }

      var sourceCalendarId = sourceCalendar.id;
      var sourceCalendarName = sourceCalendar.summaryOverride || sourceCalendar.summary;
      Logger.log("Fetching events from source calendar: " + sourceCalendarName + " (" + sourceCalendarId + ")");

      // Fetch all events from the source calendar
      var events = [];
      var eventList = callWithBackoff(function(){
        return Calendar.Events.list(sourceCalendarId, {
          showDeleted: false,
          singleEvents: false, // Get recurring event definitions
          maxResults: 2500
        });
      }, defaultMaxRetries);

      if (eventList != null && eventList.items) {
        events = events.concat(eventList.items);

        // Loop until we received all events
        while (typeof eventList.nextPageToken !== 'undefined') {
          eventList = callWithBackoff(function(){
            return Calendar.Events.list(sourceCalendarId, {
              showDeleted: false,
              singleEvents: false,
              maxResults: 2500,
              pageToken: eventList.nextPageToken
            });
          }, defaultMaxRetries);

          if (eventList != null && eventList.items) {
            events = events.concat(eventList.items);
          }
        }
      }

      Logger.log("Fetched " + events.length + " events from " + sourceCalendarName);

      // Add colorId and source calendar name to each event
      events.forEach(function(event){
        event.sourceColorId = colorId;
        event.sourceCalendarName = sourceCalendarName;
        result.push(event);
      });

    } catch (e) {
      Logger.log("[ERROR] Failed to fetch from calendar " + sourceCalendarIdOrName + ": " + e);
      reportOverallFailure = true;
    }
  }

  return result;
}

/**
 * Finds a Google Calendar by ID or name from the user's calendar list.
 *
 * @param {string} calendarIdOrName - The calendar ID or name to search for
 * @return {?Calendar} The calendar if found, null otherwise
 */
function findSourceCalendar(calendarIdOrName) {
  var calendarList = Calendar.CalendarList.list({showHidden: true, maxResults: 250}).items;

  // First try to find by exact ID match
  for (var cal of calendarList) {
    if (cal.id == calendarIdOrName) {
      return cal;
    }
  }

  // Then try to find by name match
  for (var cal of calendarList) {
    var calName = cal.summaryOverride || cal.summary;
    if (calName == calendarIdOrName) {
      return cal;
    }
  }

  return null;
}

/**
 * Gets the user's Google Calendar with the specified name.
 * A new Calendar will be created if the user does not have a Calendar with the specified name.
 *
 * @param {string} targetCalendarName - The name of the calendar to return
 * @return {Calendar} The calendar retrieved or created
 */
function setupTargetCalendar(targetCalendarName){
  var targetCalendar = Calendar.CalendarList.list({showHidden: true, maxResults: 250}).items.filter(function(cal) {
    return ((cal.summaryOverride || cal.summary) == targetCalendarName) &&
                (cal.accessRole == "owner" || cal.accessRole == "writer");
  })[0];

  if(targetCalendar == null){
    Logger.log("Creating Calendar: " + targetCalendarName);
    targetCalendar = Calendar.newCalendar();
    targetCalendar.summary = targetCalendarName;
    targetCalendar.description = "Created by GAS";
    targetCalendar.timeZone = Calendar.Settings.get("timezone").value;
    targetCalendar = Calendar.Calendars.insert(targetCalendar);
  }

  return targetCalendar;
}

/**
 * Processes events fetched from source Google Calendars.
 * Filters out cancelled events and populates sourceEventsIds.
 *
 * @param {Array} events - Array with all events from source calendars
 * @return {Array} Array with filtered events ready for processing
 */
function processSourceEvents(events){
  var result = [];

  // Filter out cancelled events
  result = events.filter(function(event){
    return event.status !== "cancelled";
  });

  // Apply user-defined filters
  result = filterSourceEvents(result);

  // Populate sourceEventsIds for change tracking
  result.forEach(function(event){
    // Use iCalUID if available, otherwise use the event ID
    var eventUid = event.iCalUID || event.id;

    // For recurring event instances, append the original start time
    if (event.recurringEventId) {
      var originalStart = event.originalStartTime;
      var recId = originalStart.dateTime || originalStart.date;
      sourceEventsIds.push(eventUid + "_" + recId);
    } else {
      sourceEventsIds.push(eventUid);
    }
  });

  return result;
}

/**
 * Applies filters to source events based on filters defined in filters.gs
 *
 * @param {Array} events - Array with all events from the source calendars
 * @return {Array} Array with filtered events
 */
function filterSourceEvents(events){
  Logger.log(`Applying ${filters.length} filters on ${events.length} events.`);

  for (var filter of filters){
    filter.parameter = filter.parameter.toLowerCase();
    events = events.filter(function(event){
      try {
        if (["dtstart", "dtend", "start", "end"].includes(filter.parameter)){
          var referenceDate = new Date();
          referenceDate.setDate(referenceDate.getDate() + (filter.offset || 0));

          var eventTime;
          if (filter.parameter === "dtstart" || filter.parameter === "start") {
            eventTime = new Date(event.start.dateTime || event.start.date);
          } else {
            eventTime = new Date(event.end.dateTime || event.end.date);
          }

          switch (filter.comparison){
            case ">":
              return ((eventTime > referenceDate) ^ (filter.type == "exclude"));
            case "<":
              return ((eventTime < referenceDate) ^ (filter.type == "exclude"));
            default:
              return true;
          }
        }
        else {
          // Handle text-based filters (summary, description, location, etc.)
          var fieldValue = "";
          switch(filter.parameter) {
            case "summary":
              fieldValue = event.summary || "";
              break;
            case "description":
              fieldValue = event.description || "";
              break;
            case "location":
              fieldValue = event.location || "";
              break;
            default:
              fieldValue = event[filter.parameter] || "";
          }

          var regexString = `${(["equals", "begins with"].includes(filter.comparison)) ? "^" : ""}(${filter.criterias.join("|")})${(filter.comparison == "equals") ? "$" : ""}`;
          var regex = new RegExp(regexString);
          return regex.test(fieldValue) ^ (filter.type == "exclude");
        }
      } catch(e) {
        Logger.log("Filter error: " + e);
        return (filter.type == "exclude");
      }
    });
  }

  Logger.log(`${events.length} events left after filtering.`);
  return events;
}

/**
 * Creates a Google Calendar event and inserts it to the target calendar.
 *
 * @param {Object} event - The source Google Calendar event to process
 * @param {string} calendarTz - The timezone of the target calendar
 */
function processEvent(event, calendarTz){
  //------------------------ Create the event object ------------------------
  var newEvent = createEvent(event, calendarTz);
  if (newEvent == null)
    return;

  var index = calendarEventsIds.indexOf(newEvent.extendedProperties.private["id"]);
  var needsUpdate = index > -1;

  //------------------------ Save instance overrides ------------------------
  //----------- To make sure the parent event is actually created -----------
  if (event.recurringEventId){
    Logger.log("Saving event instance for later: " + newEvent.recurringEventId);
    recurringEvents.push(newEvent);
    return;
  }
  else{
    //------------------------ Send event object to gcal ------------------------
    if (needsUpdate){
      if (modifyExistingEvents){
        oldEvent = calendarEvents[index]
        Logger.log("Updating existing event " + newEvent.extendedProperties.private["id"]);
        try{
          newEvent = callWithBackoff(function(){
            return Calendar.Events.update(newEvent, targetCalendarId, calendarEvents[index].id);
          }, defaultMaxRetries);
        }
        catch (e){
          Logger.log(`Operation failed with error "${e}"`);
          reportOverallFailure = true;
        }
        if (newEvent != null && emailSummary){
          modifiedEvents.push([[oldEvent.summary, newEvent.summary, oldEvent.start.date||oldEvent.start.dateTime, newEvent.start.date||newEvent.start.dateTime, oldEvent.end.date||oldEvent.end.dateTime, newEvent.end.date||newEvent.end.dateTime, oldEvent.location, newEvent.location, oldEvent.description, newEvent.description], targetCalendarName]);
        }
      }
    }
    else{
      if (addEventsToCalendar){
        Logger.log("Adding new event " + newEvent.extendedProperties.private["id"]);
        try{
          newEvent = callWithBackoff(function(){
            return Calendar.Events.insert(newEvent, targetCalendarId);
          }, defaultMaxRetries);
        }
        catch (e){
          Logger.log(`Operation failed with error "${e}"`);
          reportOverallFailure = true;
        }
        if (newEvent != null && emailSummary){
          addedEvents.push([[newEvent.summary, newEvent.start.date||newEvent.start.dateTime, newEvent.end.date||newEvent.end.dateTime, newEvent.location, newEvent.description], targetCalendarName]);
        }
      }
    }
  }
}

/**
 * Creates a Google Calendar Event based on the source Google Calendar event.
 * Will return null if the event has not changed since the last sync.
 *
 * @param {Object} event - The source Google Calendar event to process
 * @param {string} calendarTz - The timezone of the target calendar
 * @return {?Calendar.Event} The Calendar.Event that will be added to the target calendar
 */
function createEvent(event, calendarTz){
  // Create a digest for change detection (exclude fields that change on sync)
  var eventForDigest = {
    summary: event.summary,
    description: event.description,
    location: event.location,
    start: event.start,
    end: event.end,
    recurrence: event.recurrence,
    status: event.status,
    transparency: event.transparency,
    visibility: event.visibility
  };
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(eventForDigest), Utilities.Charset.UTF_8).toString();

  // Use iCalUID if available, otherwise use event ID
  var eventUid = event.iCalUID || event.id;

  if(calendarEventsMD5s.indexOf(digest) >= 0){
    Logger.log("Skipping unchanged Event " + eventUid);
    return null;
  }

  var newEvent = {};

  // Handle all-day vs timed events
  if (event.start.date) {
    // All-day event
    newEvent.start = { date: event.start.date };
    newEvent.end = { date: event.end.date };
  } else {
    // Timed event
    newEvent.start = {
      dateTime: event.start.dateTime,
      timeZone: event.start.timeZone || calendarTz
    };
    newEvent.end = {
      dateTime: event.end.dateTime,
      timeZone: event.end.timeZone || calendarTz
    };
  }

  // Copy attendees if enabled
  if (addAttendees && event.attendees && event.attendees.length > 0) {
    newEvent.attendees = event.attendees.map(function(att) {
      return {
        email: att.email,
        displayName: att.displayName,
        responseStatus: att.responseStatus
      };
    });
  }

  // Copy status
  if (event.status) {
    var status = event.status.toLowerCase();
    if (["confirmed", "tentative", "cancelled"].indexOf(status) > -1) {
      newEvent.status = status;
    }
  }

  // Copy source URL if present
  if (event.source && event.source.url) {
    newEvent.source = {
      url: event.source.url,
      title: event.source.title || 'link'
    };
  }

  // Handle summary/title
  if (descriptionAsTitles && event.description) {
    newEvent.summary = event.description;
  } else {
    newEvent.summary = event.summary || "(No title)";
  }

  // Handle organizer
  if (event.organizer) {
    newEvent.organizer = {
      displayName: event.organizer.displayName,
      email: event.organizer.email
    };

    if (addOrganizerToTitle && event.organizer.displayName) {
      newEvent.summary = event.organizer.displayName + ": " + newEvent.summary;
    }
  }

  // Add source calendar name to title if enabled
  if (addCalToTitle && event.sourceCalendarName) {
    newEvent.summary = "(" + event.sourceCalendarName + ") " + newEvent.summary;
  }

  // Copy description
  if (event.description) {
    newEvent.description = event.description;
  }

  // Copy location
  if (event.location) {
    newEvent.location = event.location;
  }

  // Handle visibility
  var validVisibilityValues = ["default", "public", "private", "confidential"];
  if (overrideVisibility && validVisibilityValues.includes(overrideVisibility.toLowerCase())) {
    newEvent.visibility = overrideVisibility.toLowerCase();
  } else if (event.visibility && validVisibilityValues.includes(event.visibility.toLowerCase())) {
    newEvent.visibility = event.visibility.toLowerCase();
  }

  // Copy transparency
  if (event.transparency) {
    var transparency = event.transparency.toLowerCase();
    if (["opaque", "transparent"].indexOf(transparency) > -1) {
      newEvent.transparency = transparency;
    }
  }

  // Handle reminders
  var isAllDay = !!event.start.date;
  if (isAllDay) {
    if (0 <= defaultAllDayReminder && defaultAllDayReminder <= 40320) {
      newEvent.reminders = { 'useDefault': false, 'overrides': [{'method': 'popup', 'minutes': defaultAllDayReminder}] };
    } else {
      newEvent.reminders = { 'useDefault': false, 'overrides': [] };
    }
  } else {
    newEvent.reminders = { 'useDefault': true, 'overrides': [] };
  }

  // Handle reminders based on addAlerts setting
  switch (addAlerts) {
    case "yes":
      // Copy reminders from source event if they exist
      if (event.reminders && event.reminders.overrides && event.reminders.overrides.length > 0) {
        var overrides = [];
        for (var i = 0; i < Math.min(event.reminders.overrides.length, 5); i++) {
          var reminder = event.reminders.overrides[i];
          if (0 <= reminder.minutes && reminder.minutes <= 40320) {
            overrides.push({'method': reminder.method || 'popup', 'minutes': reminder.minutes});
          }
        }
        if (overrides.length > 0) {
          newEvent.reminders = { 'useDefault': false, 'overrides': overrides };
        }
      }
      break;
    case "no":
      newEvent.reminders = { 'useDefault': false, 'overrides': [] };
      break;
    default:
    case "default":
      newEvent.reminders = { 'useDefault': true, 'overrides': [] };
      break;
  }

  // Copy recurrence rules for recurring events
  if (event.recurrence && event.recurrence.length > 0) {
    newEvent.recurrence = event.recurrence;
  }

  // Set extended properties for tracking
  newEvent.extendedProperties = { private: { MD5: digest, fromGAS: "true", id: eventUid } };

  // Handle recurring event instances
  if (event.recurringEventId) {
    var originalStart = event.originalStartTime;
    newEvent.recurringEventId = originalStart.dateTime || originalStart.date;
    newEvent.extendedProperties.private['rec-id'] = eventUid + "_" + newEvent.recurringEventId;
  }

  // Handle color - use source color or override color
  if (event.sourceColorId) {
    var colorID = event.sourceColorId;
    if (Object.keys(CalendarApp.EventColor).includes(colorID)) {
      newEvent.colorId = CalendarApp.EventColor[colorID];
    } else if (Object.values(CalendarApp.EventColor).includes(colorID)) {
      newEvent.colorId = colorID;
    }
  } else if (event.colorId) {
    newEvent.colorId = event.colorId;
  }

  return newEvent;
}

/**
 * Patches an existing event instance with the provided Calendar.Event.
 * The instance that needs to be updated is identified by the recurrence-id of the provided event.
 *
 * @param {Calendar.Event} recEvent - The event instance to process
 */
function processEventInstance(recEvent){
  Logger.log("ID: " + recEvent.extendedProperties.private["id"] + " | Date: "+ recEvent.recurringEventId);

  var eventInstanceToPatch = callWithBackoff(function(){
    return Calendar.Events.list(targetCalendarId,
      { singleEvents : true,
        privateExtendedProperty : "fromGAS=true",
        privateExtendedProperty : "rec-id=" + recEvent.extendedProperties.private["id"] + "_" + recEvent.recurringEventId
      }).items;
  }, defaultMaxRetries);

  if (eventInstanceToPatch == null || eventInstanceToPatch.length == 0){
    if (recEvent.recurringEventId.length == 10){
      recEvent.recurringEventId += "T00:00:00Z";
    }
    else if (recEvent.recurringEventId.substr(-1) !== "Z"){
      recEvent.recurringEventId += "Z";
    }
    eventInstanceToPatch = callWithBackoff(function(){
       return Calendar.Events.list(targetCalendarId,
        { singleEvents : true,
          orderBy : "startTime",
          maxResults: 1,
          timeMin : recEvent.recurringEventId,
          privateExtendedProperty : "fromGAS=true",
          privateExtendedProperty : "id=" + recEvent.extendedProperties.private["id"]
        }).items;
    }, defaultMaxRetries);
  }

  if (eventInstanceToPatch !== null && eventInstanceToPatch.length == 1){
    if (modifyExistingEvents){
      Logger.log("Updating existing event instance");
      callWithBackoff(function(){
        Calendar.Events.update(recEvent, targetCalendarId, eventInstanceToPatch[0].id);
      }, defaultMaxRetries);
    }
  }
  else{
    if (addEventsToCalendar){
      Logger.log("No Instance matched, adding as new event!");
      callWithBackoff(function(){
        Calendar.Events.insert(recEvent, targetCalendarId);
      }, defaultMaxRetries);
    }
  }
}

/**
 * Deletes all events from the target calendar that no longer exist in the source calendars.
 * If removePastEventsFromCalendar is set to false, events that have taken place will not be removed.
 */
function processEventCleanup(){
  for (var i = 0; i < calendarEvents.length; i++){
      var currentID = calendarEventsIds[i];
      var feedIndex = sourceEventsIds.indexOf(currentID);

      if(feedIndex  == -1                                             // Event is no longer in source
        && calendarEvents[i].recurringEventId == null                 // And it's not a recurring event
        && (                                                          // And one of:
          removePastEventsFromCalendar                                // We want to remove past events
          || new Date(calendarEvents[i].start.dateTime) > new Date()  // Or the event is in the future
          || new Date(calendarEvents[i].start.date) > new Date()      // (2 different ways event start can be stored)
        )
      )
      {
        Logger.log("Deleting old event " + currentID);
        try{
          callWithBackoff(function(){
            Calendar.Events.remove(targetCalendarId, calendarEvents[i].id);
          }, defaultMaxRetries);
        }
        catch (e){
          Logger.log(`Operation failed with error "${e}"`);
          reportOverallFailure = true;
        }

        if (emailSummary){
          removedEvents.push([[calendarEvents[i].summary, calendarEvents[i].start.date||calendarEvents[i].start.dateTime, calendarEvents[i].end.date||calendarEvents[i].end.dateTime, calendarEvents[i].location, calendarEvents[i].description], targetCalendarName]);
        }
      }
    }
}

/**
* Sends an email summary with added/modified/deleted events.
*/
function sendSummary() {
  var subject;
  var body;

  var subject = `${customEmailSubject ? customEmailSubject : "GAS-Calendar-Sync Execution Summary"}: ${addedEvents.length} new, ${modifiedEvents.length} modified, ${removedEvents.length} deleted`;
  addedEvents = condenseCalendarMap(addedEvents);
  modifiedEvents = condenseCalendarMap(modifiedEvents);
  removedEvents = condenseCalendarMap(removedEvents);

  body = "GAS-Calendar-Sync made the following changes to your calendar:<br/>";
  for (var tgtCal of addedEvents){
    body += `<br/>${tgtCal[0]}: ${tgtCal[1].length} added events<br/><ul>`;
    for (var addedEvent of tgtCal[1]){
      body += "<li>"
        + "Name: " + addedEvent[0][0] + "<br/>"
        + "Start: " + formatDate(addedEvent[0][1]) + "<br/>"
        + "End: " + formatDate(addedEvent[0][2]) + "<br/>"
        + (addedEvent[0][3] ? ("Location: " + addedEvent[0][3] + "<br/>") : "")
        + (addedEvent[0][4] ? ("Description: " + addedEvent[0][4] + "<br/>") : "")
        + "</li>";
    }
    body += "</ul>";
  }

  for (var tgtCal of modifiedEvents){
    body += `<br/>${tgtCal[0]}: ${tgtCal[1].length} modified events<br/><ul>`;
    for (var modifiedEvent of tgtCal[1]){
      body += "<li>"
        + (modifiedEvent[0][0] != modifiedEvent[0][1] ? ("<del>Name: " + modifiedEvent[0][0] + "</del><br/>") : "")
        + "Name: " + modifiedEvent[0][1] + "<br/>"
        + (modifiedEvent[0][2] != modifiedEvent[0][3] ? ("<del>Start: " + formatDate(modifiedEvent[0][2]) + "</del><br/>") : "")
        + " Start: " + formatDate(modifiedEvent[0][3]) + "<br/>"
        + (modifiedEvent[0][4] != modifiedEvent[0][5] ? ("<del>End: " + formatDate(modifiedEvent[0][4]) + "</del><br/>") : "")
        + " End: " + formatDate(modifiedEvent[0][5]) + "<br/>"
        + (modifiedEvent[0][6] != modifiedEvent[0][7] ? ("<del>Location: " + (modifiedEvent[0][6] ? modifiedEvent[0][6] : "") + "</del><br/>") : "")
        + (modifiedEvent[0][7] ? (" Location: " + modifiedEvent[0][7] + "<br/>") : "")
        + (modifiedEvent[0][8] != modifiedEvent[0][9] ? ("<del>Description: " + (modifiedEvent[0][8] ? modifiedEvent[0][8] : "") + "</del><br/>") : "")
        + (modifiedEvent[0][9] ? (" Description: " + modifiedEvent[0][9] + "<br/>") : "")
        + "</li>";
    }
    body += "</ul>";
  }

  for (var tgtCal of removedEvents){
    body += `<br/>${tgtCal[0]}: ${tgtCal[1].length} removed events<br/><ul>`;
    for (var removedEvent of tgtCal[1]){
      body += "<li>"
        + "<del>Name: " + removedEvent[0][0] + "</del><br/>"
        + "<del>Start: " + formatDate(removedEvent[0][1]) + "</del><br/>"
        + "<del>End: " + formatDate(removedEvent[0][2]) + "</del><br/>"
        + (removedEvent[0][3] ? ("<del>Location: " + removedEvent[0][3] + "</del><br/>") : "")
        + (removedEvent[0][4] ? ("<del>Description: " + removedEvent[0][4] + "</del><br/>") : "")
        + "</li>";
    }
    body += "</ul>";
  }

  body += "<br/><br/>GAS-Calendar-Sync - Google Calendar to Google Calendar Sync";
  var message = {
    to: email,
    subject: subject,
    htmlBody: body,
    name: "GAS-Calendar-Sync"
  };

  MailApp.sendEmail(message);
}

/**
 * Runs the specified function with exponential backoff and returns the result.
 * Will return null if the function did not succeed afterall.
 *
 * @param {function} func - The function that should be executed
 * @param {Number} maxRetries - How many times the function should try if it fails
 * @return {?Calendar.Event} The Calendar.Event that was added in the calendar, null if func did not complete successfully
 */
var backoffRecoverableErrors = [
  "service invoked too many times in a short time",
  "rate limit exceeded",
  "internal error",
  "http error 403", // forbidden
  "http error 408", // request timeout
  "http error 423", // locked
  "http error 500", // internal server error
  "http error 503", // service unavailable
  "http error 504"  // gateway timeout
];
function callWithBackoff(func, maxRetries) {
  var tries = 0;
  var result;
  while ( tries <= maxRetries ) {
    tries++;
    try{
      result = func();
      return result;
    }
    catch(err){
      err = err.message  || err;
      if ( err.includes("is not a function")  || !backoffRecoverableErrors.some(function(e){
              return err.toLowerCase().includes(e);
            }) ) {
        throw err;
      } else if ( tries > maxRetries) {
        Logger.log(`Error, giving up after trying ${maxRetries} times [${err}]`);
        return null;
      } else {
        Logger.log( "Error, Retrying... [" + err  +"]");
        Utilities.sleep (Math.pow(2,tries)*100) +
                            (Math.round(Math.random() * 100));
      }
    }
  }
  return null;
}

