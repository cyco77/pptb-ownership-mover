// @vitest-environment jsdom

import { act, render, screen, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import { teamsDarkTheme, teamsLightTheme } from "@fluentui/react-components";

const mocks = vi.hoisted(() => ({
  refreshConnection: vi.fn(),
  eventCallback: undefined as ((event: string, data: unknown) => void) | undefined,
  getCurrentTheme: vi.fn(),
  log: vi.fn(),
  info: vi.fn(),
}));

vi.mock("./hooks/useConnection", () => ({
  useConnection: () => ({ connection: { id: "connection-1" }, refreshConnection: mocks.refreshConnection }),
}));

vi.mock("./hooks/useToolboxEvents", () => ({
  useToolboxEvents: (callback: (event: string, data: unknown) => void) => {
    mocks.eventCallback = callback;
  },
}));

vi.mock("./services/loggerService", () => ({
  logger: { info: mocks.info },
}));

vi.mock("./components/Overview", () => ({
  Overview: ({ connection }: any) => <div data-testid="overview">{connection?.id}</div>,
}));

beforeEach(() => {
  vi.clearAllMocks();
  mocks.getCurrentTheme.mockResolvedValue("dark");
  Object.assign(window, {
    toolboxAPI: { utils: { getCurrentTheme: mocks.getCurrentTheme } },
  });
  vi.spyOn(console, "log").mockImplementation(() => undefined);
});

describe("App", () => {
  it("loads initial theme and connection, and ignores terminal events", async () => {
    const { container } = render(<App />);

    expect(screen.getByText("Ownership Mover")).toBeInTheDocument();
    expect(screen.getByTestId("overview")).toHaveTextContent("connection-1");
    await waitFor(() => expect(mocks.getCurrentTheme).toHaveBeenCalledOnce());
    expect(mocks.info).toHaveBeenCalledWith("Initialized");
    expect(container.querySelector("[data-theme='teamsDarkTheme']")).toBeNull();

    for (const event of ["terminal:output", "terminal:command:completed", "terminal:error", "unknown:event"]) {
      act(() => mocks.eventCallback?.(event, {}));
    }
    expect(mocks.refreshConnection).not.toHaveBeenCalled();
  });

  it.each(["connection:updated", "connection:created", "connection:deleted"])(
    "refreshes the active connection on %s",
    (event) => {
      render(<App />);
      act(() => mocks.eventCallback?.(event, {}));
      expect(mocks.refreshConnection).toHaveBeenCalledOnce();
    },
  );

  it("switches Fluent UI themes for settings events", async () => {
    mocks.getCurrentTheme.mockResolvedValue("light");
    const { container } = render(<App />);
    await waitFor(() => expect(mocks.info).toHaveBeenCalledWith("Initialized"));
    act(() => mocks.eventCallback?.("settings:updated", {}));
    await waitFor(() => expect(mocks.info).toHaveBeenCalledWith("Theme updated:light"));
    expect(container.textContent).toContain("Ownership Mover");
    expect(teamsLightTheme).toBeDefined();
    expect(teamsDarkTheme).toBeDefined();
  });
});
