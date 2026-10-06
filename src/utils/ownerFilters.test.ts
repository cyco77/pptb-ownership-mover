import { describe, expect, it } from "vitest";
import { filterSystemUsers, filterTeams } from "./ownerFilters";
import type { SystemUser } from "../types/systemUser";
import type { Team } from "../types/team";

const users: SystemUser[] = [
  {
    systemuserid: "user-1",
    fullname: "Ada Lovelace",
    domainname: "ada@example.com",
    isdisabled: false,
    applicationid: null,
    businessunitid: { businessunitid: "bu-1", name: "Research" },
  },
  {
    systemuserid: "app-1",
    fullname: "Background Service",
    domainname: "service@example.com",
    isdisabled: true,
    applicationid: "application-1",
    businessunitid: { businessunitid: "bu-2", name: "Operations" },
  },
];

const baseUserFilters = {
  status: "all" as const,
  userType: "all" as const,
  businessUnitId: "all",
  text: "",
};

describe("filterSystemUsers", () => {
  it("applies status and user type filters together", () => {
    expect(
      filterSystemUsers(users, {
        ...baseUserFilters,
        status: "disabled",
        userType: "applications",
      }).map((user) => user.systemuserid),
    ).toEqual(["app-1"]);
  });

  it("searches names, domain names, and business units case-insensitively", () => {
    expect(
      filterSystemUsers(users, { ...baseUserFilters, text: "RESEARCH" }).map(
        (user) => user.systemuserid,
      ),
    ).toEqual(["user-1"]);
    expect(
      filterSystemUsers(users, { ...baseUserFilters, text: "ADA@EXAMPLE" }),
    ).toHaveLength(1);
    expect(
      filterSystemUsers(users, { ...baseUserFilters, text: "unknown" }),
    ).toHaveLength(0);
  });

  it("trims search text and applies the business unit filter", () => {
    expect(
      filterSystemUsers(users, {
        ...baseUserFilters,
        businessUnitId: "bu-2",
        text: " Service ",
      }).map((user) => user.systemuserid),
    ).toEqual(["app-1"]);
  });

  it("returns all users when no filters are active", () => {
    expect(filterSystemUsers(users, baseUserFilters)).toEqual(users);
  });

  it("retains regular users when only the user type filter is active", () => {
    expect(
      filterSystemUsers(users, {
        ...baseUserFilters,
        userType: "users",
      }).map((user) => user.systemuserid),
    ).toEqual(["user-1"]);
  });
});

describe("filterTeams", () => {
  const teams: Team[] = [
    {
      teamid: "team-1",
      name: "Platform Admins",
      teamtype: 0,
      isdefault: false,
      businessunitid: { businessunitid: "bu-1", name: "Research" },
    },
    {
      teamid: "team-2",
      name: "Support",
      teamtype: 1,
      isdefault: true,
    },
  ];

  it("filters by business unit and searches team or business unit names", () => {
    expect(
      filterTeams(teams, { businessUnitId: "bu-1", text: "research" }).map(
        (team) => team.teamid,
      ),
    ).toEqual(["team-1"]);
  });

  it("includes teams without a business unit when no business unit is selected", () => {
    expect(filterTeams(teams, { businessUnitId: "all", text: "support" })).toEqual([
      teams[1],
    ]);
  });
});
