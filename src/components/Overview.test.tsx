// @vitest-environment jsdom

import { act, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { Overview } from "./Overview";
import type { OwnershipAnalysisResult } from "../types/ownership";

const mocks = vi.hoisted(() => ({
  loadSystemUsers: vi.fn(),
  loadTeams: vi.fn(),
  loadOwnershipCountsForOwners: vi.fn(),
  reassignOwnedRecordsForUser: vi.fn(),
  gridProps: undefined as any,
  filterProps: undefined as any,
  drawerProps: undefined as any,
  showNotification: vi.fn(),
}));

vi.mock("../services/dataverseService", () => ({
  loadSystemUsers: mocks.loadSystemUsers,
  loadTeams: mocks.loadTeams,
  loadOwnershipCountsForOwners: mocks.loadOwnershipCountsForOwners,
  reassignOwnedRecordsForUser: mocks.reassignOwnedRecordsForUser,
}));

vi.mock("./DataGridView", () => ({
  DataGridView: (props: any) => {
    mocks.gridProps = props;
    return (
      <button
        data-testid="data-grid"
        onClick={() => props.onSelectionChange(["user-1"])}
      >
        Select owner
      </button>
    );
  },
}));

vi.mock("./Filter", () => ({
  Filter: (props: any) => {
    mocks.filterProps = props;
    return <div data-testid="filter" />;
  },
}));

vi.mock("./OwnershipResultsDrawer", () => ({
  OwnershipResultsDrawer: (props: any) => {
    mocks.drawerProps = props;
    return <div data-testid="ownership-drawer" />;
  },
}));

const analysisResult: OwnershipAnalysisResult = {
  scannedEntities: 1,
  analyzedEntities: 1,
  failedEntities: 0,
  failedEntityDetails: [],
  users: [
    {
      userId: "user-1",
      totalOwnedRecords: 2,
      entitiesWithRecords: 1,
      entityCounts: [
        {
          entityLogicalName: "account",
          entityDisplayName: "Account",
          entitySetName: "accounts",
          primaryIdAttribute: "accountid",
          recordCount: 2,
        },
      ],
    },
  ],
};

const users = [
  {
    systemuserid: "user-1",
    fullname: "Ada Lovelace",
    domainname: "ada@example.com",
    isdisabled: false,
    applicationid: null,
  },
  {
    systemuserid: "app-1",
    fullname: "Service Principal",
    domainname: "service@example.com",
    isdisabled: true,
    applicationid: "app-1",
  },
];
const teams = [
  {
    teamid: "team-1",
    name: "Platform",
    teamtype: 0,
    isdefault: true,
  },
];

beforeEach(() => {
  vi.clearAllMocks();
  mocks.loadSystemUsers.mockResolvedValue(users);
  mocks.loadTeams.mockResolvedValue(teams);
  mocks.loadOwnershipCountsForOwners.mockResolvedValue(analysisResult);
  mocks.reassignOwnedRecordsForUser.mockResolvedValue({
    sourceOwnerId: "user-1",
    sourceOwnerType: "systemuser",
    targetOwnerId: "team-1",
    targetOwnerType: "team",
    processedEntities: 1,
    reassignedRecords: 2,
    failedRecords: 0,
    entityResults: [],
  });
  Object.assign(window, {
    toolboxAPI: { utils: { showNotification: mocks.showNotification } },
  });
});

describe("Overview", () => {
  it("loads and maps owners into the drawer and renders users", async () => {
    render(<Overview connection={{} as any} />);

    await waitFor(() => expect(mocks.gridProps).toBeDefined());
    expect(mocks.gridProps.systemUsers).toEqual([users[0]]);
    expect(mocks.drawerProps.allSystemUsers).toEqual([
      expect.objectContaining({ userId: "user-1", isApplication: false, isDisabled: false }),
      expect.objectContaining({ userId: "app-1", isApplication: true, isDisabled: true }),
    ]);
  });

  it("loads teams and filters by team text and business unit", async () => {
    render(<Overview connection={{} as any} />);
    await waitFor(() => expect(mocks.filterProps).toBeDefined());

    act(() => {
      mocks.filterProps.onEntityTypeChanged("team");
      mocks.filterProps.onTextFilterChanged("platform");
    });
    await waitFor(() => expect(mocks.gridProps.entityType).toBe("team"));
    expect(mocks.gridProps.teams.map((team: any) => team.teamid)).toEqual([
      "team-1",
    ]);

    act(() => mocks.filterProps.onBusinessUnitFilterChanged("missing-bu"));
    await waitFor(() => expect(screen.getByText("No teams found.")).toBeInTheDocument());
  });

  it("shows a load error notification and clears loading state", async () => {
    mocks.loadTeams.mockRejectedValueOnce(new Error("offline"));
    vi.spyOn(console, "error").mockImplementation(() => undefined);
    render(<Overview connection={{} as any} />);

    expect(await screen.findByText("No system users found.")).toBeInTheDocument();
    expect(mocks.showNotification).toHaveBeenCalledWith(
      expect.objectContaining({ title: "Error Loading Data", type: "error" }),
    );
  });

  it("skips loading without a connection and resets filters/selections on owner type changes", async () => {
    const { rerender } = render(<Overview connection={null} />);
    expect(mocks.loadSystemUsers).not.toHaveBeenCalled();

    rerender(<Overview connection={{} as any} />);
    await waitFor(() => expect(mocks.filterProps).toBeDefined());
    fireEvent.click(screen.getByText("Select owner"));
    expect(mocks.gridProps.selectedIds).toEqual(["user-1"]);

    act(() => {
      mocks.filterProps.onTextFilterChanged("Ada");
      mocks.filterProps.onStatusFilterChanged("disabled");
      mocks.filterProps.onUserTypeFilterChanged("applications");
      mocks.filterProps.onBusinessUnitFilterChanged("bu-2");
    });
    act(() => mocks.filterProps.onEntityTypeChanged("team"));

    await waitFor(() => expect(mocks.filterProps.entityType).toBe("team"));
    expect(mocks.filterProps.textFilter).toBe("");
    expect(mocks.filterProps.statusFilter).toBe("enabled");
    expect(mocks.filterProps.userTypeFilter).toBe("all");
    expect(mocks.filterProps.businessUnitFilter).toBe("all");
    expect(mocks.gridProps.selectedIds).toEqual([]);
    expect(screen.queryByText("No teams found.")).not.toBeInTheDocument();

    act(() => mocks.filterProps.onBusinessUnitFilterChanged("unknown"));
    await waitFor(() => expect(screen.getByText("No teams found.")).toBeInTheDocument());
  });

  it("analyzes selected owners and delegates only selected entities for assignment", async () => {
    mocks.showNotification.mockResolvedValue(undefined);
    render(<Overview connection={{} as any} />);
    await waitFor(() => expect(mocks.gridProps).toBeDefined());

    fireEvent.click(screen.getByText("Select owner"));
    fireEvent.click(screen.getByRole("button", { name: "Analyze Ownership" }));

    await waitFor(() => expect(mocks.drawerProps.result).toEqual(analysisResult));
    expect(mocks.loadOwnershipCountsForOwners).toHaveBeenCalledWith(
      ["user-1"],
      expect.any(Function),
    );
    expect(mocks.drawerProps.users).toEqual([
      expect.objectContaining({ userId: "user-1", userName: "Ada Lovelace" }),
    ]);

    await expect(
      mocks.drawerProps.onAssignRecords("user-1", "team", "team-1", ["account"]),
    ).resolves.toMatchObject({ reassignedRecords: 2 });
    expect(mocks.reassignOwnedRecordsForUser).toHaveBeenCalledWith(
      "user-1",
      "systemuser",
      "team-1",
      "team",
      [analysisResult.users[0].entityCounts[0]],
      undefined,
    );
    expect(mocks.showNotification).toHaveBeenCalledWith(
      expect.objectContaining({ title: "Ownership Assignment Completed" }),
    );
  });

  it("does not open an ownership analysis when there is no selected owner", async () => {
    render(<Overview connection={{} as any} />);
    await waitFor(() => expect(mocks.gridProps).toBeDefined());

    fireEvent.click(screen.getByRole("button", { name: "Analyze Ownership" }));

    expect(mocks.loadOwnershipCountsForOwners).not.toHaveBeenCalled();
    expect(mocks.drawerProps.open).toBe(false);
  });

  it("reports analysis errors and rejects assignment for unknown owners", async () => {
    mocks.loadOwnershipCountsForOwners.mockRejectedValueOnce(new Error("metadata unavailable"));
    vi.spyOn(console, "error").mockImplementation(() => undefined);
    render(<Overview connection={{} as any} />);
    await waitFor(() => expect(mocks.gridProps).toBeDefined());
    fireEvent.click(screen.getByText("Select owner"));
    fireEvent.click(screen.getByRole("button", { name: "Analyze Ownership" }));

    await waitFor(() =>
      expect(mocks.showNotification).toHaveBeenCalledWith(
        expect.objectContaining({ title: "Ownership Analysis Failed" }),
      ),
    );
    await expect(
      mocks.drawerProps.onAssignRecords("missing-user", "team", "team-1", []),
    ).rejects.toThrow("No ownership data found");
  });
});
