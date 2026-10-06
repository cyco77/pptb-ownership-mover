import "@testing-library/jest-dom/vitest";
import { cleanup } from "@testing-library/react";
import { afterEach } from "vitest";

afterEach(cleanup);

if (typeof window !== "undefined" && typeof NodeFilter === "undefined") {
  Object.defineProperty(globalThis, "NodeFilter", {
    configurable: true,
    value: window.NodeFilter,
  });
}
