// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  createRoot: vi.fn(),
  render: vi.fn(),
  app: vi.fn(() => <div>Application</div>),
}));

vi.mock("react-dom/client", () => ({
  createRoot: mocks.createRoot,
}));
vi.mock("./App", () => ({ default: mocks.app }));

beforeEach(() => {
  vi.resetModules();
  document.body.innerHTML = "";
  mocks.createRoot.mockReturnValue({ render: mocks.render });
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("application bootstrap", () => {
  it("creates a React root once when the root element exists", async () => {
    document.body.innerHTML = '<div id="root"></div>';

    await import("./main");

    const root = document.getElementById("root");
    expect(root).toHaveAttribute("data-reactroot-initialized", "true");
    expect(mocks.createRoot).toHaveBeenCalledWith(root);
    expect(mocks.render).toHaveBeenCalledOnce();
  });

  it("does not mount twice on an already initialized root", async () => {
    document.body.innerHTML =
      '<div id="root" data-reactroot-initialized="true"></div>';

    await import("./main");

    expect(mocks.createRoot).not.toHaveBeenCalled();
    expect(mocks.render).not.toHaveBeenCalled();
  });

  it("logs an error when the root element is missing", async () => {
    const error = vi.spyOn(console, "error").mockImplementation(() => undefined);

    await import("./main");

    expect(error).toHaveBeenCalledWith(
      "Root element not found. Make sure the HTML contains <div id=\"root\"></div>",
    );
    expect(mocks.createRoot).not.toHaveBeenCalled();
  });
});
