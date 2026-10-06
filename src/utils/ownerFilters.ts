import type { SystemUser } from "../types/systemUser";
import type { Team } from "../types/team";

export type OwnerFilters = {
  status: "all" | "enabled" | "disabled";
  userType: "all" | "users" | "applications";
  businessUnitId: string;
  text: string;
};

export const filterSystemUsers = (
  users: SystemUser[],
  filters: OwnerFilters,
): SystemUser[] => {
  const searchTerm = filters.text.trim().toLowerCase();

  return users.filter((user) => {
    if (
      filters.status !== "all" &&
      user.isdisabled !== (filters.status === "disabled")
    ) {
      return false;
    }

    if (
      filters.userType !== "all" &&
      Boolean(user.applicationid) !== (filters.userType === "applications")
    ) {
      return false;
    }

    if (
      filters.businessUnitId !== "all" &&
      user.businessunitid?.businessunitid !== filters.businessUnitId
    ) {
      return false;
    }

    return (
      !searchTerm ||
      [user.fullname, user.domainname, user.businessunitid?.name].some(
        (value) => value?.toLowerCase().includes(searchTerm),
      )
    );
  });
};

export const filterTeams = (
  teams: Team[],
  filters: Pick<OwnerFilters, "businessUnitId" | "text">,
): Team[] => {
  const searchTerm = filters.text.trim().toLowerCase();

  return teams.filter((team) => {
    if (
      filters.businessUnitId !== "all" &&
      team.businessunitid?.businessunitid !== filters.businessUnitId
    ) {
      return false;
    }

    return (
      !searchTerm ||
      [team.name, team.businessunitid?.name].some((value) =>
        value?.toLowerCase().includes(searchTerm),
      )
    );
  });
};
