import { afterEach, describe, expect, it, vi } from "vitest";
import { logger } from "./loggerService";

describe("logger", () => {
  afterEach(() => {
    logger.setLogCallback(null as never);
    vi.restoreAllMocks();
  });

  it("writes typed messages to the console without a callback", () => {
    const log = vi.spyOn(console, "log").mockImplementation(() => undefined);

    logger.info("started");
    logger.success("saved");
    logger.warning("retry");
    logger.error("failed");

    expect(log.mock.calls).toEqual([
      ["[INFO] started"],
      ["[SUCCESS] saved"],
      ["[WARNING] retry"],
      ["[ERROR] failed"],
    ]);
  });

  it("forwards messages to the configured callback", () => {
    const callback = vi.fn();
    logger.setLogCallback(callback);

    logger.log("custom", "warning");

    expect(callback).toHaveBeenCalledWith("custom", "warning");
  });
});
