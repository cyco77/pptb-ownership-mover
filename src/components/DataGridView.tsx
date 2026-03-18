import React, { useMemo } from "react";
import {
  DataGrid,
  DataGridBody,
  DataGridCell,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridRow,
  makeStyles,
  TableColumnDefinition,
  createTableColumn,
  type DataGridProps,
  type OnSelectionChangeData,
} from "@fluentui/react-components";
import { SystemUser } from "../types/systemUser";
import { Team } from "../types/team";

interface IDataGridViewProps {
  entityType: "systemuser" | "team";
  systemUsers: SystemUser[];
  teams: Team[];
  selectedIds: string[];
  onSelectionChange: (ids: string[]) => void;
}

const useStyles = makeStyles({
  container: {
    height: "100%",
    display: "flex",
    flexDirection: "column",
    overflow: "hidden",
  },
  dataGrid: {
    height: "100%",
    display: "flex",
    flexDirection: "column",
    overflow: "hidden",
  },
  dataGridBody: {
    overflowY: "auto",
    flex: 1,
  },
  cellStyles: {
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    maxWidth: "100%",
  },
  headerCell: {
    fontWeight: 700,
  },
});

export const DataGridView: React.FC<IDataGridViewProps> = ({
  entityType,
  systemUsers,
  teams,
  selectedIds,
  onSelectionChange,
}) => {
  const styles = useStyles();

  const selectedItems = useMemo(() => {
    return new Set<string>(selectedIds);
  }, [selectedIds]);

  const handleSelectionChange: DataGridProps["onSelectionChange"] = (
    _e,
    data: OnSelectionChangeData,
  ) => {
    const selected = Array.from(data.selectedItems).map((id) => String(id));
    onSelectionChange(selected);
  };

  const systemUserColumns: TableColumnDefinition<SystemUser>[] = [
    createTableColumn<SystemUser>({
      columnId: "fullname",
      compare: (a, b) => a.fullname.localeCompare(b.fullname),
      renderHeaderCell: () => "Full Name",
      renderCell: (item) => (
        <span className={styles.cellStyles} title={item.fullname}>
          {item.fullname}
        </span>
      ),
    }),
    createTableColumn<SystemUser>({
      columnId: "domainname",
      compare: (a, b) => a.domainname.localeCompare(b.domainname),
      renderHeaderCell: () => "Domain Name",
      renderCell: (item) => (
        <span className={styles.cellStyles} title={item.domainname}>
          {item.domainname}
        </span>
      ),
    }),
    createTableColumn<SystemUser>({
      columnId: "businessunit",
      compare: (a, b) =>
        (a.businessunitid?.name ?? "").localeCompare(
          b.businessunitid?.name ?? "",
        ),
      renderHeaderCell: () => "Business Unit",
      renderCell: (item) => (
        <span
          className={styles.cellStyles}
          title={item.businessunitid?.name ?? "N/A"}
        >
          {item.businessunitid?.name ?? "N/A"}
        </span>
      ),
    }),
    createTableColumn<SystemUser>({
      columnId: "status",
      compare: (a, b) =>
        a.isdisabled === b.isdisabled ? 0 : a.isdisabled ? 1 : -1,
      renderHeaderCell: () => "Status",
      renderCell: (item) => (
        <span
          className={styles.cellStyles}
          title={item.isdisabled ? "Disabled" : "Enabled"}
        >
          {item.isdisabled ? "Disabled" : "Enabled"}
        </span>
      ),
    }),
  ];

  const teamColumns: TableColumnDefinition<Team>[] = [
    createTableColumn<Team>({
      columnId: "teamname",
      compare: (a, b) => a.name.localeCompare(b.name),
      renderHeaderCell: () => "Team Name",
      renderCell: (item) => (
        <div className={styles.cellStyles} title={item.name}>
          {item.name}
        </div>
      ),
    }),
    createTableColumn<Team>({
      columnId: "teamtype",
      compare: (a, b) => a.teamtype - b.teamtype,
      renderHeaderCell: () => "Team Type",
      renderCell: (item) => {
        const typeText =
          item.teamtype === 0
            ? "Owner"
            : item.teamtype === 1
              ? "Access"
              : "Other";
        return (
          <div className={styles.cellStyles} title={typeText}>
            {typeText}
          </div>
        );
      },
    }),
    createTableColumn<Team>({
      columnId: "businessunit",
      compare: (a, b) =>
        (a.businessunitid?.name ?? "").localeCompare(
          b.businessunitid?.name ?? "",
        ),
      renderHeaderCell: () => "Business Unit",
      renderCell: (item) => {
        const buName = item.businessunitid?.name ?? "N/A";
        return (
          <div className={styles.cellStyles} title={buName}>
            {buName}
          </div>
        );
      },
    }),
    createTableColumn<Team>({
      columnId: "isdefault",
      compare: (a, b) =>
        a.isdefault === b.isdefault ? 0 : a.isdefault ? -1 : 1,
      renderHeaderCell: () => "Default Team",
      renderCell: (item) => {
        const defaultText = item.isdefault ? "Yes" : "No";
        return (
          <div className={styles.cellStyles} title={defaultText}>
            {defaultText}
          </div>
        );
      },
    }),
  ];

  const columnSizingOptionsSystemusers = {
    fullname: {
      minWidth: 300,
      defaultWidth: 300,
    },
    domainname: {
      minWidth: 300,
      defaultWidth: 300,
    },
    businessunit: {
      minWidth: 250,
      defaultWidth: 250,
    },
    status: {
      minWidth: 200,
      defaultWidth: 200,
    },
  };

  if (entityType === "systemuser") {
    return (
      <div className={styles.container}>
        <DataGrid
          key="systemuser-grid"
          items={systemUsers}
          columns={systemUserColumns}
          sortable
          defaultSortState={{
            sortColumn: "fullname",
            sortDirection: "ascending",
          }}
          resizableColumns
          columnSizingOptions={columnSizingOptionsSystemusers}
          className={styles.dataGrid}
          size="small"
          selectionMode="multiselect"
          selectedItems={selectedItems}
          onSelectionChange={handleSelectionChange}
          getRowId={(item) => item.systemuserid}
        >
          <DataGridHeader>
            <DataGridRow>
              {({ renderHeaderCell }) => (
                <DataGridHeaderCell className={styles.headerCell}>
                  {renderHeaderCell()}
                </DataGridHeaderCell>
              )}
            </DataGridRow>
          </DataGridHeader>
          <DataGridBody<SystemUser> className={styles.dataGridBody}>
            {({ item, rowId }) => (
              <DataGridRow<SystemUser> key={rowId}>
                {({ renderCell }) => (
                  <DataGridCell>{renderCell(item)}</DataGridCell>
                )}
              </DataGridRow>
            )}
          </DataGridBody>
        </DataGrid>
      </div>
    );
  }

  const columnSizingOptionsTeams = {
    teamname: {
      minWidth: 300,
      defaultWidth: 300,
    },
    teamtype: {
      minWidth: 100,
      defaultWidth: 100,
    },
    businessunit: {
      minWidth: 250,
      defaultWidth: 250,
    },
    isdefault: {
      minWidth: 200,
      defaultWidth: 200,
    },
  };

  return (
    <div className={styles.container}>
      <DataGrid
        key="team-grid"
        items={teams}
        columns={teamColumns}
        sortable
        resizableColumns
        columnSizingOptions={columnSizingOptionsTeams}
        defaultSortState={{
          sortColumn: "teamname",
          sortDirection: "ascending",
        }}
        className={styles.dataGrid}
        size="small"
        selectionMode="multiselect"
        selectedItems={selectedItems}
        onSelectionChange={handleSelectionChange}
        getRowId={(item) => item.teamid}
      >
        <DataGridHeader>
          <DataGridRow>
            {({ renderHeaderCell }) => (
              <DataGridHeaderCell className={styles.headerCell}>
                {renderHeaderCell()}
              </DataGridHeaderCell>
            )}
          </DataGridRow>
        </DataGridHeader>
        <DataGridBody<Team> className={styles.dataGridBody}>
          {({ item, rowId }) => (
            <DataGridRow<Team> key={rowId}>
              {({ renderCell }) => (
                <DataGridCell>{renderCell(item)}</DataGridCell>
              )}
            </DataGridRow>
          )}
        </DataGridBody>
      </DataGrid>
    </div>
  );
};
