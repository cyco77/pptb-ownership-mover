// @vitest-environment jsdom

import { fireEvent, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, expect, it, vi } from "vitest";
import { Filter } from "./Filter";
import type { SystemUser } from "../types/systemUser";
import type { Team } from "../types/team";

const systemUsers: SystemUser[] = [
  {
    systemuserid: "user-1",
    fullname: "Ada Lovelace",
    domainname: "ada@example.com",
    isdisabled: false,
    applicationid: null,
    businessunitid: { businessunitid: "bu-1", name: "Research" },
  },
];
const teams: Team[] = [
  {
    teamid: "team-1",
    name: "Platform",
    teamtype: 0,
    isdefault: true,
    businessunitid: { businessunitid: "bu-1", name: "Research" },
  },
  {
    teamid: "team-2",
    name: "Support",
    teamtype: 1,
    isdefault: false,
    businessunitid: { businessunitid: "bu-2", name: "Support Unit" },
  },
];

const makeProps = (overrides: Partial<React.ComponentProps<typeof Filter>> = {}) => ({
  entityType: "systemuser" as const,
  systemUsers,
  teams,
  statusFilter: "enabled" as const,
  userTypeFilter: "all" as const,
  businessUnitFilter: "all",
  textFilter: "",
  onEntityTypeChanged: vi.fn(),
  onTextFilterChanged: vi.fn(),
  onStatusFilterChanged: vi.fn(),
  onUserTypeFilterChanged: vi.fn(),
  onBusinessUnitFilterChanged: vi.fn(),
  ...overrides,
});

describe("Filter", () => {
  it("renders user-only filters and unique sorted business units", () => {
    render(<Filter {...makeProps()} />);

    expect(screen.getByText("Status")).toBeInTheDocument();
    expect(screen.getByText("User Type")).toBeInTheDocument();
    expect(screen.getByRole("combobox", { name: "Business Unit" })).toHaveTextContent(
      "All",
    );
  });

  it("hides user-specific filters for teams", () => {
    render(<Filter {...makeProps({ entityType: "team" })} />);

    expect(screen.queryByLabelText("Status")).not.toBeInTheDocument();
    expect(screen.queryByLabelText("User Type")).not.toBeInTheDocument();
  });

  it("forwards search text changes", () => {
    const props = makeProps();
    render(<Filter {...props} />);

    fireEvent.change(screen.getByPlaceholderText("Search by name..."), {
      target: { value: "ada" },
    });

    expect(props.onTextFilterChanged).toHaveBeenCalledWith("ada");
  });

  it("forwards entity, status, user type, and business unit selections", async () => {
    const user = userEvent.setup();
    const props = makeProps();
    render(<Filter {...props} />);

    await user.click(screen.getByLabelText("Entity Type"));
    await user.click(await screen.findByRole("option", { name: "Teams" }));
    expect(props.onEntityTypeChanged).toHaveBeenCalledWith("team");

    await user.click(screen.getByLabelText("Status"));
    await user.click(await screen.findByRole("option", { name: "Disabled" }));
    expect(props.onStatusFilterChanged).toHaveBeenCalledWith("disabled");

    await user.click(screen.getByLabelText("User Type"));
    await user.click(await screen.findByRole("option", { name: "Applications" }));
    expect(props.onUserTypeFilterChanged).toHaveBeenCalledWith("applications");

    await user.click(screen.getByLabelText("Business Unit"));
    await user.click(await screen.findByRole("option", { name: "Support Unit" }));
    expect(props.onBusinessUnitFilterChanged).toHaveBeenCalledWith("bu-2");
  });

  it("falls back to all business units when the selected ID is not present", () => {
    render(<Filter {...makeProps({ businessUnitFilter: "missing" })} />);

    expect(screen.getByLabelText("Business Unit")).toHaveTextContent(
      "All",
    );
  });

  it("supports omitted optional callbacks and missing business units", async () => {
    const user = userEvent.setup();
    render(
      <Filter
        {...makeProps({
          onStatusFilterChanged: undefined,
          onUserTypeFilterChanged: undefined,
          onBusinessUnitFilterChanged: undefined,
          systemUsers: [],
          teams: [],
          statusFilter: "all",
          userTypeFilter: "users",
        })}
      />,
    );

    await user.click(screen.getByLabelText("Status"));
    await user.click(await screen.findByRole("option", { name: "Enabled" }));
    await user.click(screen.getByLabelText("User Type"));
    await user.click(await screen.findByRole("option", { name: "Applications" }));
    await user.click(screen.getByLabelText("Business Unit"));
    await user.click(await screen.findByRole("option", { name: "All" }));

    expect(screen.getByRole("combobox", { name: "Business Unit" })).toHaveTextContent("All");
  });

  it("falls back to default status and user type labels", () => {
    render(
      <Filter
        {...makeProps({
          statusFilter: undefined,
          userTypeFilter: undefined,
          businessUnitFilter: undefined,
        })}
      />,
    );

    expect(screen.getByLabelText("Status")).toHaveTextContent("Disabled");
    expect(screen.getByLabelText("User Type")).toHaveTextContent("Applications");
    expect(screen.getByLabelText("Business Unit")).toHaveTextContent("All");
  });
});
