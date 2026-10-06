// @vitest-environment node

import { afterEach, describe, expect, it, vi } from "vitest";
import { createCsvTimestamp, downloadCsv, toCsvLine } from "./csv";

describe("CSV utilities", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("quotes values, doubles embedded quotes, and converts nullish values to empty cells", () => {
    expect(toCsvLine(["plain", 'has "quotes"', null, undefined, 42])).toBe(
      '"plain","has ""quotes""","","","42"',
    );
  });

  it("creates a filesystem-safe UTC timestamp", () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-01-02T03:04:05.678Z"));
    expect(createCsvTimestamp()).toBe("2026-01-02T03-04-05-678Z");
    vi.useRealTimers();
  });

  it("does not create a download for an empty CSV", () => {
    const createObjectURL = vi.spyOn(URL, "createObjectURL");
    downloadCsv([], "empty.csv");
    expect(createObjectURL).not.toHaveBeenCalled();
  });

  it("downloads a UTF-8 BOM CSV and releases the temporary URL", () => {
    const click = vi.fn();
    const anchor = { href: "", download: "", click };
    const appendChild = vi.fn();
    const removeChild = vi.fn();
    const createObjectURL = vi.fn((_blob: Blob) => "blob:csv");
    const revokeObjectURL = vi.fn();
    vi.stubGlobal("document", {
      createElement: vi.fn(() => anchor),
      body: { appendChild, removeChild },
    });
    vi.stubGlobal("URL", {
      createObjectURL,
      revokeObjectURL,
    });

    downloadCsv(['"name"'], "records.csv");

    expect(createObjectURL).toHaveBeenCalledWith(
      expect.objectContaining({ type: "text/csv;charset=utf-8;" }),
    );
    const blob = createObjectURL.mock.calls[0]?.[0];
    expect(blob).toBeInstanceOf(Blob);
    expect(blob?.size).toBeGreaterThan(0);
    expect(anchor).toMatchObject({ href: "blob:csv", download: "records.csv" });
    expect(appendChild).toHaveBeenCalledWith(anchor);
    expect(click).toHaveBeenCalledOnce();
    expect(removeChild).toHaveBeenCalledWith(anchor);
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:csv");
  });
});
