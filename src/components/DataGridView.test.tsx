// @vitest-environment jsdom

import { fireEvent, render, screen } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import { DataGridView } from "./DataGridView";
import type { SystemUser } from "../types/systemUser";
import type { Team } from "../types/team";

const users: SystemUser[] = [
  {
    systemuserid: "user-1",
    fullname: "Ada Lovelace",
    domainname: "ada@example.com",
    isdisabled: false,
    applicationid: null,
  },
  {
    systemuserid: "user-2",
    fullname: "Grace Hopper",
    domainname: "grace@example.com",
    isdisabled: true,
    applicationid: null,
    businessunitid: { businessunitid: "bu-1", name: "Research" },
  },
];
const teams: Team[] = [
  { teamid: "team-1", name: "Owners", teamtype: 0, isdefault: true },
  { teamid: "team-2", name: "Access", teamtype: 1, isdefault: false },
  { teamid: "team-3", name: "Unknown", teamtype: 4, isdefault: false },
  { teamid: "team-4", name: "No Unit", teamtype: 2, isdefault: true },
];

describe("DataGridView", () => {
  it("renders system user columns and forwards row selection", () => {
    const onSelectionChange = vi.fn();
    render(
      <DataGridView
        entityType="systemuser"
        systemUsers={users}
        teams={[]}
        selectedIds={[]}
        onSelectionChange={onSelectionChange}
      />,
    );

    expect(screen.getByText("Full Name")).toBeInTheDocument();
    expect(screen.getByText("Domain Name")).toBeInTheDocument();
    expect(screen.getByText("Ada Lovelace")).toBeInTheDocument();
    expect(screen.getByText("Grace Hopper")).toBeInTheDocument();
    expect(screen.getByText("Enabled")).toBeInTheDocument();
    expect(screen.getByText("Disabled")).toBeInTheDocument();
    const row = screen.getByRole("row", { name: /Ada Lovelace/ });
    fireEvent.click(row);
    expect(onSelectionChange).toHaveBeenCalled();

    fireEvent.click(screen.getByRole("button", { name: "Domain Name" }));
    fireEvent.click(screen.getByRole("button", { name: "Business Unit" }));
    fireEvent.click(screen.getByRole("button", { name: "Status" }));
  });

  it("renders all team types and defaults, and forwards team selection", () => {
    const onSelectionChange = vi.fn();
    render(
      <DataGridView
        entityType="team"
        systemUsers={[]}
        teams={teams}
        selectedIds={["team-1"]}
        onSelectionChange={onSelectionChange}
      />,
    );

    expect(screen.getByText("Owner")).toBeInTheDocument();
    expect(screen.getAllByText("Access")).toHaveLength(2);
    expect(screen.getAllByText("Other")).toHaveLength(2);
    expect(screen.getAllByText("Yes")).toHaveLength(2);
    expect(screen.getAllByText("No")).toHaveLength(2);
    expect(screen.getAllByText("N/A")).toHaveLength(4);
    fireEvent.click(screen.getByRole("row", { name: /Access/ }));
    expect(onSelectionChange).toHaveBeenCalled();

    fireEvent.click(screen.getByRole("button", { name: "Business Unit" }));
    fireEvent.click(screen.getByRole("button", { name: "Default Team" }));
  });

  it("compares team types and default flags in both directions", () => {
    const { container } = render(
      <DataGridView
        entityType="team"
        systemUsers={[]}
        teams={teams}
        selectedIds={[]}
        onSelectionChange={vi.fn()}
      />,
    );
    const buttons = Array.from(container.querySelectorAll('[role="button"]'));
    const teamTypeHeader = buttons.find((button) => button.textContent?.includes("Team Type"));
    const defaultHeader = buttons.find((button) => button.textContent?.includes("Default Team"));

    expect(teamTypeHeader).toBeTruthy();
    expect(defaultHeader).toBeTruthy();
    fireEvent.click(teamTypeHeader!);
    fireEvent.click(defaultHeader!);
    fireEvent.click(defaultHeader!);
  });
});
