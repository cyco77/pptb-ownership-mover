import React from "react";
import type {
  JSXElement,
  OptionOnSelectData,
  SelectionEvents,
} from "@fluentui/react-components";
import {
  Dropdown,
  makeStyles,
  Option,
  useId,
  SearchBox,
  SearchBoxChangeEvent,
} from "@fluentui/react-components";
import { SystemUser } from "../types/systemUser";
import { Team } from "../types/team";

export interface IFilterProps {
  entityType: "systemuser" | "team";
  systemUsers: SystemUser[];
  teams: Team[];
  statusFilter?: "all" | "enabled" | "disabled";
  userTypeFilter?: "all" | "users" | "applications";
  businessUnitFilter?: string;
  textFilter?: string;
  onEntityTypeChanged: (entityType: "systemuser" | "team") => void;
  onTextFilterChanged: (searchText: string) => void;
  onStatusFilterChanged?: (status: "all" | "enabled" | "disabled") => void;
  onUserTypeFilterChanged?: (
    userType: "all" | "users" | "applications"
  ) => void;
  onBusinessUnitFilterChanged?: (businessUnitId: string) => void;
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    gap: "20px",
    alignItems: "flex-end",
    flexWrap: "wrap",
  },
  field: {
    display: "grid",
    justifyItems: "start",
    gap: "2px",
  },
  dropdown: {
    minWidth: "200px",
  },
  dropdownSmall: {
    minWidth: "150px",
  },
  searchInput: {
    minWidth: "350px",
  },
});

export const Filter = (props: IFilterProps): JSXElement => {
  const entityTypeDropdownId = useId("entity-type-dropdown");
  const searchInputId = useId("search-input");
  const statusDropdownId = useId("status-dropdown");
  const userTypeDropdownId = useId("user-type-dropdown");
  const businessUnitDropdownId = useId("business-unit-dropdown");

  const onEntityTypeSelect = (
    _event: SelectionEvents,
    data: OptionOnSelectData
  ) => {
    props.onEntityTypeChanged(data.optionValue as "systemuser" | "team");
  };

  const onTextFilterChange = (
    _event: SearchBoxChangeEvent,
    data: { value: string }
  ) => {
    props.onTextFilterChanged(data.value);
  };

  const onStatusSelect = (
    _event: SelectionEvents,
    data: OptionOnSelectData
  ) => {
    props.onStatusFilterChanged?.(
      data.optionValue as "all" | "enabled" | "disabled"
    );
  };

  const onUserTypeSelect = (
    _event: SelectionEvents,
    data: OptionOnSelectData
  ) => {
    props.onUserTypeFilterChanged?.(
      data.optionValue as "all" | "users" | "applications"
    );
  };

  const onBusinessUnitSelect = (
    _event: SelectionEvents,
    data: OptionOnSelectData
  ) => {
    props.onBusinessUnitFilterChanged?.(data.optionValue ?? "all");
  };

  // Extract unique business units from system users and teams
  const businessUnits = React.useMemo(() => {
    const uniqueUnits = new Map<string, string>();
    props.systemUsers.forEach((user) => {
      if (user.businessunitid) {
        uniqueUnits.set(
          user.businessunitid.businessunitid,
          user.businessunitid.name
        );
      }
    });
    props.teams.forEach((team) => {
      if (team.businessunitid) {
        uniqueUnits.set(
          team.businessunitid.businessunitid,
          team.businessunitid.name
        );
      }
    });
    return Array.from(uniqueUnits.entries())
      .map(([id, name]) => ({ id, name }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [props.systemUsers, props.teams]);

  const styles = useStyles();

  return (
    <div className={styles.root}>
      <div className={styles.field}>
        <label htmlFor={entityTypeDropdownId}>Entity Type</label>
        <Dropdown
          id={entityTypeDropdownId}
          placeholder="Select entity type"
          onOptionSelect={onEntityTypeSelect}
          className={styles.dropdownSmall}
          value={props.entityType === "systemuser" ? "System Users" : "Teams"}
          selectedOptions={[props.entityType]}
        >
          <Option value="systemuser" text="System Users">
            System Users
          </Option>
          <Option value="team" text="Teams">
            Teams
          </Option>
        </Dropdown>
      </div>

      {props.entityType === "systemuser" && (
        <>
          <div className={styles.field}>
            <label htmlFor={statusDropdownId}>Status</label>
            <Dropdown
              id={statusDropdownId}
              placeholder="Enabled"
              onOptionSelect={onStatusSelect}
              className={styles.dropdownSmall}
              value={
                props.statusFilter === "all"
                  ? "All"
                  : props.statusFilter === "enabled"
                  ? "Enabled"
                  : "Disabled"
              }
              selectedOptions={[props.statusFilter ?? "enabled"]}
            >
              <Option value="all" text="All">
                All
              </Option>
              <Option value="enabled" text="Enabled">
                Enabled
              </Option>
              <Option value="disabled" text="Disabled">
                Disabled
              </Option>
            </Dropdown>
          </div>
          <div className={styles.field}>
            <label htmlFor={userTypeDropdownId}>User Type</label>
            <Dropdown
              id={userTypeDropdownId}
              placeholder="All"
              onOptionSelect={onUserTypeSelect}
              className={styles.dropdownSmall}
              value={
                props.userTypeFilter === "all"
                  ? "All"
                  : props.userTypeFilter === "users"
                  ? "Users"
                  : "Applications"
              }
              selectedOptions={[props.userTypeFilter ?? "all"]}
            >
              <Option value="all" text="All">
                All
              </Option>
              <Option value="users" text="Users">
                Users
              </Option>
              <Option value="applications" text="Applications">
                Applications
              </Option>
            </Dropdown>
          </div>
        </>
      )}

      <div className={styles.field}>
        <label htmlFor={businessUnitDropdownId}>Business Unit</label>
        <Dropdown
          id={businessUnitDropdownId}
          placeholder="All"
          onOptionSelect={onBusinessUnitSelect}
          className={styles.dropdown}
          value={
            props.businessUnitFilter === "all"
              ? "All"
              : businessUnits.find((bu) => bu.id === props.businessUnitFilter)
                  ?.name ?? "All"
          }
          selectedOptions={[props.businessUnitFilter ?? "all"]}
        >
          <Option value="all" text="All">
            All
          </Option>
          {businessUnits.map((bu) => (
            <Option key={bu.id} value={bu.id} text={bu.name}>
              {bu.name}
            </Option>
          ))}
        </Dropdown>
      </div>

      <div className={styles.field}>
        <label htmlFor={searchInputId}>Search</label>
        <SearchBox
          id={searchInputId}
          placeholder="Search by name..."
          value={props.textFilter ?? ""}
          onChange={onTextFilterChange}
          className={styles.searchInput}
        />
      </div>
    </div>
  );
};
