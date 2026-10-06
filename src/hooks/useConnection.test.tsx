// @vitest-environment jsdom

import { act, renderHook, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { useConnection } from "./useConnection";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("useConnection", () => {
  it("loads the active connection and can refresh it", async () => {
    const getActiveConnection = vi
      .fn()
      .mockResolvedValueOnce({ id: "connection-1" })
      .mockResolvedValueOnce({ id: "connection-2" });
    Object.assign(window, {
      toolboxAPI: { connections: { getActiveConnection } },
    });

    const { result } = renderHook(() => useConnection());
    await waitFor(() => expect(result.current.isLoading).toBe(false));
    expect(result.current.connection).toEqual({ id: "connection-1" });

    await act(async () => result.current.refreshConnection());
    expect(result.current.connection).toEqual({ id: "connection-2" });
    expect(getActiveConnection).toHaveBeenCalledTimes(2);
  });

  it("stops loading after a connection lookup fails", async () => {
    const getActiveConnection = vi.fn().mockRejectedValue(new Error("offline"));
    Object.assign(window, {
      toolboxAPI: { connections: { getActiveConnection } },
    });
    vi.spyOn(console, "error").mockImplementation(() => undefined);

    const { result } = renderHook(() => useConnection());

    await waitFor(() => expect(result.current.isLoading).toBe(false));
    expect(result.current.connection).toBeNull();
  });
});
