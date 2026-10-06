// @vitest-environment jsdom

import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { OwnershipResultsDrawer } from "./OwnershipResultsDrawer";
import type {
  OwnershipAnalysisProgress,
  OwnershipAnalysisResult,
} from "../types/ownership";

beforeEach(() => {
  vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => undefined);
  vi.stubGlobal("URL", {
    createObjectURL: vi.fn(() => "blob:csv"),
    revokeObjectURL: vi.fn(),
  });
});

vi.mock("@fluentui/react-components", async (importOriginal) => {
  const actual = await importOriginal<typeof import("@fluentui/react-components")>();
  const React = await import("react");
  const Option = ({ children, value, text, onSelect }: any) => (
    <div
      role="option"
      aria-label={text}
      data-value={value}
      onClick={() => onSelect?.(value)}
    >
      {children}
    </div>
  );
  return {
    ...actual,
    OverlayDrawer: ({ open, children }: any) =>
      open ? <div role="dialog">{children}</div> : null,
    DrawerBody: ({ children }: any) => <div>{children}</div>,
    Accordion: ({ children }: any) => <div>{children}</div>,
    AccordionItem: ({ children }: any) => <section>{children}</section>,
    AccordionHeader: ({ children }: any) => <div>{children}</div>,
    AccordionPanel: ({ children }: any) => <div>{children}</div>,
    Option,
    Dropdown: ({ children, value, placeholder, selectedOptions, onOptionSelect }: any) => (
      <div>
        <button type="button" role="combobox" onClick={() => undefined}>
          {value || placeholder}
        </button>
        {React.Children.map(children, (child: any) =>
          React.isValidElement(child)
            ? React.cloneElement(child, {
                onSelect: (optionValue: string) =>
                  onOptionSelect?.(new Event("change"), { optionValue }),
                selected: selectedOptions?.includes((child.props as any).value),
              } as any)
            : child,
        )}
      </div>
    ),
  };
});

const result: OwnershipAnalysisResult = {
  scannedEntities: 2,
  analyzedEntities: 1,
  failedEntities: 1,
  failedEntityDetails: [
    {
      entityLogicalName: "contact",
      entityDisplayName: "Contact",
      message: "Permission denied",
    },
  ],
  users: [
    {
      userId: "source-user",
      totalOwnedRecords: 3,
      entitiesWithRecords: 1,
      entityCounts: [
        {
          entityLogicalName: "account",
          entityDisplayName: "Account",
          entitySetName: "accounts",
          primaryIdAttribute: "accountid",
          recordCount: 3,
        },
      ],
    },
  ],
};

const ownerPool = [
  {
    userId: "source-user",
    userName: "Source User",
    domainName: "source@example.com",
    isApplication: false,
    isDisabled: false,
  },
  {
    userId: "target-user",
    userName: "Active User",
    domainName: "active@example.com",
    isApplication: false,
    isDisabled: false,
  },
  {
    userId: "disabled-user",
    userName: "Disabled User",
    domainName: "disabled@example.com",
    isApplication: false,
    isDisabled: true,
  },
  {
    userId: "application-user",
    userName: "Automation",
    isApplication: true,
    isDisabled: true,
  },
];
const teams = [{ userId: "team-1", userName: "Operations Team" }];

const assignmentResult = (failedRecords = 0) => ({
  sourceOwnerId: "source-user",
  sourceOwnerType: "systemuser" as const,
  targetOwnerId: "target-user",
  targetOwnerType: "systemuser" as const,
  processedEntities: 1,
  reassignedRecords: 3 - failedRecords,
  failedRecords,
  entityResults: [
    {
      entityLogicalName: "account",
      entityDisplayName: "Account",
      reassignedRecords: 3 - failedRecords,
      failedRecords,
      assignedRecordIds: failedRecords ? ["record-1"] : ["record-1", "record-2", "record-3"],
      failedRecordDetails: failedRecords
        ? [{ recordId: "record-4", error: "Assignment denied" }]
        : [],
    },
  ],
});

const makeProps = (overrides: Record<string, unknown> = {}) => ({
  open: true,
  isLoading: false,
  users: [
    {
      userId: "source-user",
      userName: "Source User",
      domainName: "source@example.com",
    },
  ],
  sourceOwnerType: "systemuser" as const,
  allSystemUsers: ownerPool,
  allTeams: teams,
  result,
  progress: null,
  onAssignRecords: vi.fn().mockResolvedValue(assignmentResult()),
  onOpenChange: vi.fn(),
  ...overrides,
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("OwnershipResultsDrawer", () => {
  it("shows analysis progress, recent entities, and failed entities while loading", () => {
    const progress: OwnershipAnalysisProgress = {
      totalEntities: 4,
      processedEntities: 2,
      analyzedEntities: 1,
      failedEntities: 1,
      currentEntityLogicalName: "contact",
      currentEntityDisplayName: "Contact",
      recentEntities: [
        {
          entityLogicalName: "account",
          entityDisplayName: "Account",
          status: "analyzed",
          recordsFound: 5,
        },
        {
          entityLogicalName: "contact",
          entityDisplayName: "Contact",
          status: "failed",
          recordsFound: 0,
        },
      ],
    };
    render(<OwnershipResultsDrawer {...makeProps({ isLoading: true, progress })} />);

    expect(screen.getByText("Processing 2 of 4 entities")).toBeInTheDocument();
    expect(screen.getByText("Contact (contact)")).toBeInTheDocument();
    expect(screen.getByText("Analysis log")).toBeInTheDocument();
  });

  it("shows result summaries, entity selection, target kinds, and active status badges", async () => {
    const user = userEvent.setup();
    render(<OwnershipResultsDrawer {...makeProps()} />);

    expect(screen.getByText("Scanned Entities")).toBeInTheDocument();
    expect(screen.getByText("Successfully Analyzed")).toBeInTheDocument();
    expect(screen.getByText("Failed Entities")).toBeInTheDocument();
    expect(screen.getByText("Permission denied")).toBeInTheDocument();
    expect(screen.getByText("Selected entities for assignment: 1")).toBeInTheDocument();

    await waitFor(() => expect(screen.getAllByRole("combobox")).toHaveLength(2));
    const targetTypeDropdown = screen.getAllByRole("combobox")[0];
    await user.click(targetTypeDropdown);
    await user.click(await screen.findByRole("option", { name: "Target: Application" }));
    await user.click(screen.getAllByRole("combobox")[1]);

    const option = await screen.findByRole("option", { name: /Automation/ });
    expect(option).toHaveTextContent("Inactive");
    expect(option).toHaveTextContent("Application");

    fireEvent.click(screen.getAllByRole("combobox")[0]);
    fireEvent.click(screen.getByRole("option", { name: "Target: Team" }));
    expect(screen.getAllByRole("combobox")[1]).toHaveTextContent("Operations Team");
  });

  it("assigns selected records, displays progress/results, and exports assignment CSVs", async () => {
    const user = userEvent.setup();
    const onAssignRecords = vi.fn(async (_id, _type, _target, _entities, onProgress) => {
      onProgress?.({
        assignedRecords: 1,
        failedRecords: 0,
        totalRecords: 3,
        currentEntity: "Account",
      });
      return assignmentResult(1);
    });
    render(<OwnershipResultsDrawer {...makeProps({ onAssignRecords })} />);

    await waitFor(() => expect(screen.getAllByRole("combobox")).toHaveLength(2));
    await user.click(screen.getByRole("button", { name: "Assign selected records" }));

    await waitFor(() => expect(onAssignRecords).toHaveBeenCalledOnce());
    expect(onAssignRecords).toHaveBeenCalledWith(
      "source-user",
      "systemuser",
      "target-user",
      ["account"],
      expect.any(Function),
    );
    expect(await screen.findByText(/Last assignment result: reassigned/)).toBeInTheDocument();

    await user.click(screen.getByRole("button", { name: "Download Assignment Errors (CSV)" }));
    await user.click(screen.getByRole("button", { name: "Download Assignment Summary (CSV)" }));
    expect(screen.getByRole("button", { name: "Download Assignment Errors (CSV)" })).toBeInTheDocument();
  });

  it("handles duplicate assignment clicks and clears errors after retry succeeds", async () => {
    const onAssignRecords = vi
      .fn()
      .mockRejectedValueOnce(new Error("retry me"))
      .mockResolvedValueOnce(assignmentResult());
    render(<OwnershipResultsDrawer {...makeProps({ onAssignRecords })} />);

    const assignButton = screen.getByRole("button", { name: "Assign selected records" });
    fireEvent.click(assignButton);
    await waitFor(() => expect(screen.getByRole("alert")).toHaveTextContent("retry me"));

    fireEvent.click(assignButton);
    await waitFor(() => expect(onAssignRecords).toHaveBeenCalledTimes(2));
    expect(screen.queryByRole("alert")).not.toBeInTheDocument();
  });

  it("does not assign without a selected target or entity and closes on request", async () => {
    const user = userEvent.setup();
    const onAssignRecords = vi.fn();
    const onOpenChange = vi.fn();
    const noCandidates = makeProps({
      allSystemUsers: [],
      allTeams: [],
      onAssignRecords,
      onOpenChange,
    });
    render(<OwnershipResultsDrawer {...noCandidates} />);

    const assignButton = screen.getByRole("button", { name: "Assign selected records" });
    expect(assignButton).toBeDisabled();
    await user.click(screen.getByRole("button", { name: "Close ownership results" }));
    expect(onOpenChange).toHaveBeenCalledWith(false);
    expect(onAssignRecords).not.toHaveBeenCalled();
  });

  it("renders the empty result state and resets transient data when closed", () => {
    const props = makeProps({
      open: false,
      result: { ...result, users: [], failedEntityDetails: [] },
    });
    const { rerender } = render(<OwnershipResultsDrawer {...props} />);
    rerender(
      <OwnershipResultsDrawer
        {...props}
        open
        result={{ ...result, users: [], failedEntityDetails: [] }}
      />,
    );

    expect(screen.getByText("Download Complete Analysis CSV")).toBeInTheDocument();
    expect(screen.queryByText("Selected entities for assignment: 1")).not.toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Close ownership results" }));
  });

  it("supports empty target selections, no result data, and closing through the drawer event", () => {
    const props = makeProps({
      result: null,
      allSystemUsers: [],
      allTeams: [],
      sourceOwnerType: "team",
    });
    render(<OwnershipResultsDrawer {...props} />);

    expect(screen.queryByText("Scanned Entities")).not.toBeInTheDocument();
    expect(screen.queryByRole("combobox")).not.toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Close ownership results" }));
    expect(props.onOpenChange).toHaveBeenCalledWith(false);
  });

  it("shows assignment errors and clears the assigning state when assignment rejects", async () => {
    const onAssignRecords = vi.fn().mockRejectedValue(new Error("assignment failed"));
    const props = makeProps({ onAssignRecords });
    render(<OwnershipResultsDrawer {...props} />);

    fireEvent.click(screen.getByRole("button", { name: "Assign selected records" }));
    await waitFor(() => expect(onAssignRecords).toHaveBeenCalledOnce());
    expect(await screen.findByRole("alert")).toHaveTextContent(
      "Assignment failed: assignment failed",
    );
    await waitFor(() =>
      expect(screen.getByRole("button", { name: "Assign selected records" })).toBeEnabled(),
    );
  });

  it("does nothing when all selected entities are removed and no target was selected", () => {
    const onAssignRecords = vi.fn();
    const props = makeProps({
      allSystemUsers: [],
      selectedEntities: [],
      onAssignRecords,
    });
    render(<OwnershipResultsDrawer {...props} />);

    expect(screen.getByRole("button", { name: "Assign selected records" })).toBeDisabled();
    expect(onAssignRecords).not.toHaveBeenCalled();
  });

  it("renders owners with no known display name and no owned entities", () => {
    const emptyResult = {
      ...result,
      failedEntityDetails: [],
      users: [
        {
          userId: "unresolved-id",
          totalOwnedRecords: 0,
          entitiesWithRecords: 0,
          entityCounts: [],
        },
      ],
    };
    render(
      <OwnershipResultsDrawer
        {...makeProps({
          result: emptyResult,
          users: [],
          allSystemUsers: [],
          sourceOwnerType: "team",
        })}
      />,
    );

    expect(screen.getByText("unresolved-id")).toBeInTheDocument();
    expect(screen.getByText("No owned records found for this owner.")).toBeInTheDocument();
    expect(screen.getAllByRole("combobox")[0]).toHaveTextContent("Target: User");
    expect(screen.getAllByRole("combobox")[1]).toHaveTextContent("Select target");
  });
});
