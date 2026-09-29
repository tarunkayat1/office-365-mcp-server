/**
 * Consolidated Calendar module
 * Reduces from 5 tools to 1 tool with operation parameters
 */

const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');
const { safeTool } = require('../utils/errors');
const { enforceCalendarPolicy, recordSideEffect } = require('../policy');

/**
 * Check if a datetime string already has a timezone indicator (Z, +offset, or -offset).
 * Prevents double-appending Z to strings that already carry tz info.
 */
function hasTzIndicator(dt) {
  if (!dt) return false;
  if (dt.endsWith('Z')) return true;
  // Check for +HH:MM or -HH:MM offset at end (e.g. "...T15:00:00-05:00")
  return /[+-]\d{2}:\d{2}$/.test(dt);
}

const VALID_RECURRENCE_PATTERN_TYPES = new Set([
  "daily",
  "weekly",
  "absoluteMonthly",
  "relativeMonthly",
  "absoluteYearly",
  "relativeYearly",
]);

const VALID_RECURRENCE_RANGE_TYPES = new Set(["endDate", "noEnd", "numbered"]);

function isNonEmptyString(value) {
  return typeof value === "string" && value.trim().length > 0;
}

function isDateOnlyString(value) {
  return isNonEmptyString(value) && /^\d{4}-\d{2}-\d{2}$/.test(value.trim());
}

function omitEmptyObjectFields(record) {
  return Object.fromEntries(
    Object.entries(record).filter(([, value]) => {
      if (value === undefined || value === null) return false;
      if (typeof value === "string") return value.trim().length > 0;
      if (Array.isArray(value)) return value.length > 0;
      return true;
    }),
  );
}

/** Build Graph recurrence payload only when pattern and range are complete. */
function buildGraphRecurrence(recurrence) {
  if (!recurrence || typeof recurrence !== "object") {
    return null;
  }

  const pattern = recurrence.pattern;
  const range = recurrence.range;
  if (!pattern?.type || !range?.type) {
    return null;
  }

  if (
    !VALID_RECURRENCE_PATTERN_TYPES.has(pattern.type) ||
    !VALID_RECURRENCE_RANGE_TYPES.has(range.type)
  ) {
    return null;
  }

  if (!isDateOnlyString(range.startDate)) {
    return null;
  }

  if (range.type === "endDate" && !isDateOnlyString(range.endDate)) {
    return null;
  }

  if (range.type === "numbered" && !(Number(range.numberOfOccurrences) > 0)) {
    return null;
  }

  if (
    pattern.type === "weekly" &&
    (!Array.isArray(pattern.daysOfWeek) || pattern.daysOfWeek.length === 0)
  ) {
    return null;
  }

  const graphPattern = omitEmptyObjectFields({
    type: pattern.type,
    interval: Number(pattern.interval) > 0 ? Number(pattern.interval) : 1,
    daysOfWeek: pattern.daysOfWeek,
    dayOfMonth: pattern.dayOfMonth,
    month: pattern.month,
    firstDayOfWeek: pattern.firstDayOfWeek || "sunday",
    index: pattern.index,
  });

  const graphRange = omitEmptyObjectFields({
    type: range.type,
    startDate: range.startDate.trim(),
    endDate: isDateOnlyString(range.endDate) ? range.endDate.trim() : undefined,
    numberOfOccurrences:
      Number(range.numberOfOccurrences) > 0
        ? Number(range.numberOfOccurrences)
        : undefined,
    recurrenceTimeZone: isNonEmptyString(range.recurrenceTimeZone)
      ? range.recurrenceTimeZone.trim()
      : config.getMsTimezone(),
  });

  return {
    pattern: graphPattern,
    range: graphRange,
  };
}

/**
 * Unified calendar handler for all calendar operations
 */
async function handleCalendar(args) {
  const { operation, ...params } = args;
  
  if (!operation) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: operation. Valid operations are: list, create, get, update, delete" 
      }]
    };
  }

  const policyError = await enforceCalendarPolicy({ operation, ...params });
  if (policyError) {
    return {
      content: [{
        type: "text",
        text: policyError
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    let result;
    switch (operation) {
      case 'list':
        result = await listCalendarEvents(accessToken, params);
        break;
      case 'create':
        result = await createCalendarEvent(accessToken, params);
        break;
      case 'get':
        result = await getCalendarEvent(accessToken, params);
        break;
      case 'update':
        result = await updateCalendarEvent(accessToken, params);
        break;
      case 'delete':
        result = await deleteCalendarEvent(accessToken, params);
        break;
      default:
        return {
          content: [{ 
            type: "text", 
            text: `Invalid operation: ${operation}. Valid operations are: list, create, get, update, delete` 
          }]
        };
    }

    if (!['list', 'get'].includes(operation)) {
      await recordSideEffect('calendar.write', 'success', params.eventId || null, {
        operation,
        subject: params.subject || null
      });
    }

    return result;
  } catch (error) {
    if (!['list', 'get'].includes(operation)) {
      await recordSideEffect('calendar.write', 'failed', params.eventId || null, {
        operation,
        error: error.message
      });
    }
    console.error(`Error in calendar ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error in calendar ${operation}: ${error.message}` }]
    };
  }
}

// Implementation functions (existing logic from original files)

async function listCalendarEvents(accessToken, params) {
  const { startDateTime, endDateTime, maxResults = 10 } = params;

  console.error(`Calendar list request with startDateTime: ${startDateTime}, endDateTime: ${endDateTime}`);

  // For a bounded date range, use calendarView as the authoritative source.
  // calendarView expands recurring-series occurrences; the /events collection
  // only returns the master series, so a recurring meeting whose series began
  // earlier never matches a same-day start/end filter (this is exactly why a
  // recurring standup was missed and minutes fell back to a single recipient).
  // Return the calendarView result directly — an empty result is the correct
  // answer for that window. Only fall back to /events if calendarView errors.
  if (startDateTime && endDateTime) {
    const formattedStartDateTime = hasTzIndicator(startDateTime) ? startDateTime : `${startDateTime}Z`;
    const formattedEndDateTime = hasTzIndicator(endDateTime) ? endDateTime : `${endDateTime}Z`;

    try {
      console.error('Using calendarView endpoint (expands recurring occurrences) for ranged list');
      const viewResponse = await callGraphAPI(
        accessToken,
        'GET',
        'me/calendar/calendarView',
        null,
        {
          $select: config.CALENDAR_SELECT_FIELDS,
          $orderby: 'start/dateTime',
          $top: maxResults,
          startDateTime: formattedStartDateTime,
          endDateTime: formattedEndDateTime
        }
      );

      console.error(`calendarView returned ${viewResponse.value ? viewResponse.value.length : 0} events`);
      return formatCalendarResponse(viewResponse);
    } catch (viewError) {
      console.error('Error using calendarView approach, falling back to regular query:', viewError);
    }
  }

  // No bounded range (or calendarView failed): query the events collection.
  const queryParams = {
    $select: config.CALENDAR_SELECT_FIELDS,
    $orderby: 'start/dateTime',
    $top: maxResults
  };

  if (startDateTime || endDateTime) {
    const filters = [];
    if (startDateTime) {
      const formattedStartDateTime = hasTzIndicator(startDateTime) ? startDateTime : `${startDateTime}Z`;
      filters.push(`start/dateTime ge '${formattedStartDateTime}'`);
    }
    if (endDateTime) {
      const formattedEndDateTime = hasTzIndicator(endDateTime) ? endDateTime : `${endDateTime}Z`;
      filters.push(`end/dateTime le '${formattedEndDateTime}'`);
    }
    queryParams.$filter = filters.join(' and ');
  }

  const response = await callGraphAPI(
    accessToken,
    'GET',
    'me/calendar/events',
    null,
    queryParams
  );

  return formatCalendarResponse(response);
}

/**
 * Helper function to format calendar response
 */
function formatCalendarResponse(response) {
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No calendar events found." }]
    };
  }
  
  const eventsList = response.value.map((event, index) => {
    const startTime = config.formatDateTime(event.start.dateTime, event.start.timeZone);
    const endTime = config.formatDateTime(event.end.dateTime, event.end.timeZone);
    return `${index + 1}. ${event.subject}\n   Start: ${startTime}\n   End: ${endTime}\n   Location: ${event.location.displayName || 'N/A'}\n   ID: ${event.id}\n`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} calendar events:\n\n${eventsList}` 
    }]
  }
}

async function createCalendarEvent(accessToken, params) {
  const {
    subject,
    content,
    start: startParam,
    end: endParam,
    startDateTime,
    endDateTime,
    location,
    attendees,
    isOnlineMeeting,
    recurrence,
  } = params;
  const start = startParam || startDateTime;
  const end = endParam || endDateTime;

  if (!subject || !start || !end) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters: subject, start, and end" 
      }]
    };
  }
  
  const event = {
    subject: subject,
    body: {
      contentType: "HTML",
      content: content || ""
    },
    start: {
      dateTime: start,
      timeZone: config.getMsTimezone()
    },
    end: {
      dateTime: end,
      timeZone: config.getMsTimezone()
    }
  };
  
  if (location) {
    event.location = { displayName: location };
  }
  
  if (attendees && attendees.length > 0) {
    event.attendees = attendees.map(email => ({
      emailAddress: { address: email },
      type: "required"
    }));
  }
  
  if (isOnlineMeeting) {
    event.isOnlineMeeting = true;
    event.onlineMeetingProvider = "teamsForBusiness";
  }
  
  const graphRecurrence = buildGraphRecurrence(recurrence);
  if (graphRecurrence) {
    event.recurrence = graphRecurrence;
  } else if (recurrence && typeof recurrence === "object") {
    console.error(
      "Skipping invalid recurrence payload for one-time calendar create",
    );
  }

  const response = await callGraphAPI(
    accessToken,
    'POST',
    'me/calendar/events',
    event
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Calendar event created successfully!\nEvent ID: ${response.id}` 
    }]
  };
}

async function getCalendarEvent(accessToken, params) {
  const { eventId } = params;
  
  if (!eventId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: eventId" 
      }]
    };
  }
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    `me/calendar/events/${eventId}`,
    null,
    {
      $select: config.CALENDAR_SELECT_FIELDS
    }
  );
  
  const startTime = config.formatDateTime(response.start.dateTime, response.start.timeZone);
  const endTime = config.formatDateTime(response.end.dateTime, response.end.timeZone);

  let eventDetails = `Subject: ${response.subject}\n`;
  eventDetails += `Start: ${startTime}\n`;
  eventDetails += `End: ${endTime}\n`;
  eventDetails += `Location: ${response.location?.displayName || 'N/A'}\n`;
  eventDetails += `Online Meeting: ${response.isOnlineMeeting ? 'Yes' : 'No'}\n`;
  
  if (response.attendees && response.attendees.length > 0) {
    eventDetails += `Attendees: ${response.attendees.map(a => a.emailAddress.address).join(', ')}\n`;
  }
  
  if (response.body?.content) {
    eventDetails += `\nDescription:\n${response.body.content}`;
  }
  
  return {
    content: [{ type: "text", text: eventDetails }]
  };
}

async function updateCalendarEvent(accessToken, params) {
  const { eventId, ...updateFields } = params;
  
  if (!eventId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: eventId" 
      }]
    };
  }
  
  const allowedFields = ['subject', 'location', 'body', 'start', 'end', 'isOnlineMeeting'];
  const update = {};
  
  // Build update object with allowed fields
  for (const [key, value] of Object.entries(updateFields)) {
    if (allowedFields.includes(key)) {
      if (key === 'location') {
        update.location = { displayName: value };
      } else if (key === 'body') {
        update.body = { contentType: "HTML", content: value };
      } else if (key === 'start' || key === 'end') {
        update[key] = { dateTime: value, timeZone: config.getMsTimezone() };
      } else {
        update[key] = value;
      }
    }
  }
  
  if (Object.keys(update).length === 0) {
    return {
      content: [{ 
        type: "text", 
        text: "No valid fields to update. Valid fields: subject, location, body, start, end, isOnlineMeeting" 
      }]
    };
  }
  
  await callGraphAPI(
    accessToken,
    'PATCH',
    `me/calendar/events/${eventId}`,
    update
  );
  
  return {
    content: [{ type: "text", text: "Calendar event updated successfully!" }]
  };
}

async function deleteCalendarEvent(accessToken, params) {
  const { eventId, sendCancellations = true } = params;
  
  if (!eventId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: eventId" 
      }]
    };
  }
  
  await callGraphAPI(
    accessToken,
    'DELETE',
    `me/calendar/events/${eventId}`,
    null,
    sendCancellations ? { sendCancellations: 'true' } : {}
  );
  
  return {
    content: [{ type: "text", text: "Calendar event deleted successfully!" }]
  };
}

// Export consolidated tool
const calendarTools = [
  {
    name: "calendar",
    description: "Manage calendar events: list, create, get, update, or delete",
    inputSchema: {
      type: "object",
      properties: {
        operation: { 
          type: "string", 
          enum: ["list", "create", "get", "update", "delete"],
          description: "The operation to perform" 
        },
        // List parameters
        startDateTime: { type: "string", description: "Start date/time filter in ISO format (for list)" },
        endDateTime: { type: "string", description: "End date/time filter in ISO format (for list)" },
        maxResults: { type: "number", description: "Maximum number of results (default: 10)" },
        // Create parameters
        subject: { type: "string", description: "Event subject (for create/update)" },
        content: { type: "string", description: "Event body content in HTML (for create/update)" },
        start: { type: "string", description: "Start date/time in ISO format (for create/update)" },
        end: { type: "string", description: "End date/time in ISO format (for create/update)" },
        location: { type: "string", description: "Event location (for create/update)" },
        attendees: { 
          type: "array", 
          items: { type: "string" },
          description: "Attendee email addresses (for create)" 
        },
        isOnlineMeeting: { type: "boolean", description: "Create as Teams meeting (for create/update)" },
        // Recurrence parameters (for create)
        recurrence: {
          type: "object",
          description:
            "Optional recurrence for repeating events only. Omit entirely for one-time events.",
          properties: {
            pattern: {
              type: "object",
              properties: {
                type: { 
                  type: "string", 
                  enum: ["daily", "weekly", "absoluteMonthly", "relativeMonthly", "absoluteYearly", "relativeYearly"],
                  description: "Recurrence pattern type" 
                },
                interval: { type: "number", description: "Interval between occurrences" },
                daysOfWeek: { 
                  type: "array", 
                  items: { type: "string" },
                  description: "Days of week for weekly pattern (e.g., ['monday', 'wednesday'])" 
                },
                dayOfMonth: { type: "number", description: "Day of month for absoluteMonthly pattern" },
                month: { type: "number", description: "Month for yearly patterns (1-12)" },
                firstDayOfWeek: { type: "string", description: "First day of week (default: sunday)" },
                index: { 
                  type: "string",
                  enum: ["first", "second", "third", "fourth", "last"],
                  description: "Week index for relativeMonthly pattern" 
                }
              }
            },
            range: {
              type: "object",
              properties: {
                type: { 
                  type: "string", 
                  enum: ["endDate", "noEnd", "numbered"],
                  description: "Recurrence range type" 
                },
                startDate: { type: "string", description: "Start date in YYYY-MM-DD format" },
                endDate: { type: "string", description: "End date in YYYY-MM-DD format (for endDate type)" },
                numberOfOccurrences: { type: "number", description: "Number of occurrences (for numbered type)" },
                recurrenceTimeZone: { type: "string", description: "Time zone for recurrence (default: UTC)" }
              }
            }
          }
        },
        // Get/Update/Delete parameters
        eventId: { type: "string", description: "Event ID (for get/update/delete)" },
        // Delete parameters
        sendCancellations: { type: "boolean", description: "Send cancellation notices (default: true)" }
      },
      required: ["operation"]
    },
    handler: safeTool('calendar', handleCalendar)
  }
];

module.exports = { calendarTools, buildGraphRecurrence };
