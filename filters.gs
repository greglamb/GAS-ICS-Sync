/*
Filters for calendar events based on Google Calendar event properties.
Define each filter with the following structure and add them to the var filters array:
{
  parameter: "property",      // Event property to filter by (e.g., "summary", "description", "location", "start", "end").
  type: "include/exclude",    // Whether to include or exclude events matching the criteria.
  comparison: "method",       // Comparison method: "equals", "begins with", "contains", "regex", "<", ">".
                              // Note: "<", ">" only apply for date/time properties (start, end).
  criterias: ["values"],      // Array of values or patterns for comparison.
  offset: number              // (Optional) For date/time properties, specify an offset in days from today.
}
*/
var filters = [];

/* Examples:
var filters = [
  {
    parameter: "summary",       // Exclude events whose summary starts with "Pending:" or contains "cancelled".
    type: "exclude",
    comparison: "regex",
    criterias: ["^Pending:", "cancelled"]
  },
  {
    parameter: "location",      // Include only events at "Conference Room A".
    type: "include",
    comparison: "equals",
    criterias: ["Conference Room A"]
  },
  {
    parameter: "end",           // Include only future events (end date is after today).
    type: "include",
    comparison: ">",
    offset: 0
  },
  {
    parameter: "start",         // Exclude events starting more than 14 days from now.
    type: "exclude",
    comparison: ">",
    offset: 14
  }
];
*/