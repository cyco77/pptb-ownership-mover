// @vitest-environment jsdom

import { renderHook } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import { useToolboxEvents } from "./useToolboxEvents";

describe("useToolboxEvents", () => {
  it("forwards toolbox event name and payload data to the callback", () => {
    let eventHandler: ((event: unknown, payload: any) => void) | undefined;
    const on = vi.fn((handler) => {
      eventHandler = handler;
    });
    Object.assign(window, { toolboxAPI: { events: { on } } });
    const log = vi.spyOn(console, "log").mockImplementation(() => undefined);
    const callback = vi.fn();

    renderHook(() => useToolboxEvents(callback));
    eventHandler?.({}, { event: "connection:updated", data: { id: "c1" } });

    expect(on).toHaveBeenCalledOnce();
    expect(callback).toHaveBeenCalledWith("connection:updated", { id: "c1" });
    expect(log).toHaveBeenCalledWith(
      "Toolbox event received: connection:updated",
      { id: "c1" },
    );
  });
});
