const { buildGraphRecurrence } = require("../calendar");

describe("buildGraphRecurrence", () => {
  it("returns null for empty or partial recurrence objects", () => {
    expect(buildGraphRecurrence(undefined)).toBeNull();
    expect(buildGraphRecurrence(null)).toBeNull();
    expect(buildGraphRecurrence({})).toBeNull();
    expect(buildGraphRecurrence({ pattern: {}, range: {} })).toBeNull();
    expect(
      buildGraphRecurrence({
        pattern: { type: "weekly" },
        range: {},
      }),
    ).toBeNull();
    expect(
      buildGraphRecurrence({
        pattern: {
          type: "daily",
          interval: 1,
          daysOfWeek: [],
          dayOfMonth: 0,
          month: 0,
          firstDayOfWeek: "sunday",
          index: "first",
        },
        range: {
          type: "endDate",
          startDate: "",
          endDate: "",
          numberOfOccurrences: 0,
          recurrenceTimeZone: "",
        },
      }),
    ).toBeNull();
  });

  it("returns Graph recurrence when pattern and range types are set", () => {
    const result = buildGraphRecurrence({
      pattern: {
        type: "weekly",
        interval: 2,
        daysOfWeek: ["monday"],
      },
      range: {
        type: "endDate",
        startDate: "2026-09-28",
        endDate: "2026-12-31",
      },
    });

    expect(result).toEqual({
      pattern: {
        type: "weekly",
        interval: 2,
        daysOfWeek: ["monday"],
        firstDayOfWeek: "sunday",
      },
      range: {
        type: "endDate",
        startDate: "2026-09-28",
        endDate: "2026-12-31",
        recurrenceTimeZone: "Eastern Standard Time",
      },
    });
  });
});
