import { describe, expect, it } from "vitest";
import {
  buildOwnedRecordsQuery,
  chunkArray,
  getEntityDisplayName,
  isUserAssignableEntity,
  normalizeDataverseErrorMessage,
  normalizeOwnershipType,
  sanitizeGuid,
} from "./dataverseUtils";

describe("Dataverse utilities", () => {
  it("normalizes ownership types from strings, numbers, and enum objects", () => {
    expect(normalizeOwnershipType("UserOwned")).toBe("userowned");
    expect(normalizeOwnershipType(4)).toBe("4");
    expect(normalizeOwnershipType({ Value: "TeamOwned" })).toBe("teamowned");
    expect(normalizeOwnershipType({ Value: 0 })).toBe("0");
    expect(normalizeOwnershipType({ Value: true })).toBe("");
    expect(normalizeOwnershipType(null)).toBe("");
  });

  it("recognizes user- and team-owned entities and rejects other ownership types", () => {
    expect(isUserAssignableEntity({ OwnershipType: "UserOwned" })).toBe(true);
    expect(isUserAssignableEntity({ OwnershipType: { Value: "TeamOwned" } })).toBe(
      true,
    );
    expect(isUserAssignableEntity({ OwnershipType: 0 })).toBe(true);
    expect(isUserAssignableEntity({ OwnershipType: 4 })).toBe(true);
    expect(isUserAssignableEntity({ OwnershipType: "OrganizationOwned" })).toBe(
      false,
    );
  });

  it("uses a localized entity label or falls back to the logical name", () => {
    expect(
      getEntityDisplayName({
        LogicalName: "account",
        DisplayName: { LocalizedLabels: [{ Label: "Account" }] },
      }),
    ).toBe("Account");
    expect(getEntityDisplayName({ LogicalName: "contact" })).toBe("contact");
  });

  it("normalizes wrapped Dataverse errors and preserves unrelated errors", () => {
    expect(
      normalizeDataverseErrorMessage(
        new Error("Error invoking remote method 'dataverse.queryData':: denied"),
      ),
    ).toBe(": denied");
    expect(normalizeDataverseErrorMessage("network down")).toBe("network down");
  });

  it("sanitizes IDs and builds owner queries with optional selected columns", () => {
    expect(sanitizeGuid("{abc-123}")).toBe("abc-123");
    expect(buildOwnedRecordsQuery("accounts", "accountid", "{abc-123}")).toBe(
      "accounts?$select=accountid&$filter=_ownerid_value eq abc-123",
    );
    expect(buildOwnedRecordsQuery("accounts", "", "abc-123")).toBe(
      "accounts?$filter=_ownerid_value eq abc-123",
    );
  });

  it("chunks arrays without losing order and handles non-positive sizes", () => {
    expect(chunkArray([1, 2, 3, 4, 5], 2)).toEqual([[1, 2], [3, 4], [5]]);
    expect(chunkArray([1, 2], 0)).toEqual([[1, 2]]);
    expect(chunkArray([], 2)).toEqual([]);
  });
});
