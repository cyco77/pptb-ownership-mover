import { afterEach, describe, expect, it, vi } from "vitest";
import {
  loadOwnershipCountsForOwners,
  loadSystemUsers,
  loadTeams,
  reassignOwnedRecordsForUser,
} from "./dataverseService";
import { logger } from "./loggerService";

const makeDataverseApi = () => ({
  queryData: vi.fn(),
  getAllEntitiesMetadata: vi.fn(),
  updateMultiple: vi.fn(),
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("Dataverse service", () => {
  it("maps users and follows absolute OData paging links", async () => {
    const api = makeDataverseApi();
    api.queryData
      .mockResolvedValueOnce({
        value: [
          {
            systemuserid: "user-1",
            fullname: "Ada",
            domainname: "ada@example.com",
            isdisabled: false,
            businessunitid: { businessunitid: "bu-1", name: "Research" },
          },
        ],
        "@odata.nextLink":
          "https://example.test/api/data/v9.2/systemusers?$skiptoken=next",
      })
      .mockResolvedValueOnce({
        value: [
          {
            systemuserid: "app-1",
            fullname: "Service",
            domainname: "service@example.com",
            isdisabled: true,
            applicationid: "application-1",
          },
        ],
      });
    vi.stubGlobal("window", { dataverseAPI: api });

    await expect(loadSystemUsers()).resolves.toEqual([
      {
        systemuserid: "user-1",
        fullname: "Ada",
        domainname: "ada@example.com",
        isdisabled: false,
        applicationid: null,
        businessunitid: { businessunitid: "bu-1", name: "Research" },
      },
      {
        systemuserid: "app-1",
        fullname: "Service",
        domainname: "service@example.com",
        isdisabled: true,
        applicationid: "application-1",
        businessunitid: undefined,
      },
    ]);
    expect(api.queryData).toHaveBeenNthCalledWith(
      2,
      "systemusers?$skiptoken=next",
    );
  });

  it("handles a user without a business unit", async () => {
    const api = makeDataverseApi();
    api.queryData.mockResolvedValue({
      value: [
        {
          systemuserid: "user-no-bu",
          fullname: "No Unit",
          domainname: "no-unit@example.com",
          isdisabled: false,
        },
      ],
    });
    vi.stubGlobal("window", { dataverseAPI: api });

    await expect(loadSystemUsers()).resolves.toEqual([
      expect.objectContaining({
        systemuserid: "user-no-bu",
        applicationid: null,
        businessunitid: undefined,
      }),
    ]);
  });

  it("maps teams and returns an empty collection for empty Dataverse pages", async () => {
    const api = makeDataverseApi();
    api.queryData
      .mockResolvedValueOnce({
        value: [
          {
            teamid: "team-1",
            name: "Platform",
            teamtype: 0,
            isdefault: true,
            businessunitid: { businessunitid: "bu-1", name: "Research" },
          },
        ],
      })
      .mockResolvedValueOnce({ value: [] });
    vi.stubGlobal("window", { dataverseAPI: api });

    await expect(loadTeams()).resolves.toEqual([
      {
        teamid: "team-1",
        name: "Platform",
        teamtype: 0,
        isdefault: true,
        businessunitid: { businessunitid: "bu-1", name: "Research" },
      },
    ]);
    expect(api.queryData).toHaveBeenCalledTimes(1);
  });

  it("normalizes absolute URLs that do not include a versioned API prefix", async () => {
    const api = makeDataverseApi();
    api.queryData
      .mockResolvedValueOnce({
        value: [{ systemuserid: "user-1", fullname: "Ada" }],
        "@odata.nextLink": "https://example.test/custom/path?next=1",
      })
      .mockResolvedValueOnce({ value: [] });
    vi.stubGlobal("window", { dataverseAPI: api });

    await loadSystemUsers();

    expect(api.queryData).toHaveBeenNthCalledWith(2, "/custom/path?next=1");
  });

  it("returns an empty ownership result without querying Dataverse for no owners", async () => {
    const api = makeDataverseApi();
    vi.stubGlobal("window", { dataverseAPI: api });

    await expect(loadOwnershipCountsForOwners([])).resolves.toEqual({
      scannedEntities: 0,
      analyzedEntities: 0,
      failedEntities: 0,
      failedEntityDetails: [],
      users: [],
    });
    expect(api.getAllEntitiesMetadata).not.toHaveBeenCalled();
  });

  it("counts owned records for supported entities and reports progress", async () => {
    const api = makeDataverseApi();
    api.getAllEntitiesMetadata.mockResolvedValue({
      value: [
        {
          LogicalName: "account",
          DisplayName: { LocalizedLabels: [{ Label: "Account" }] },
          EntitySetName: "accounts",
          PrimaryIdAttribute: "accountid",
          OwnershipType: "UserOwned",
        },
        {
          LogicalName: "contact",
          DisplayName: { LocalizedLabels: [] },
          EntitySetName: "contacts",
          PrimaryIdAttribute: "contactid",
          OwnershipType: "UserOwned",
        },
        {
          LogicalName: "organizationconfig",
          EntitySetName: "organizationconfigs",
          OwnershipType: "OrganizationOwned",
        },
      ],
    });
    api.queryData.mockImplementation(async (query: string) => {
      if (query.includes("accounts")) {
        return {
          value: query.includes("owner-a") ? [{}, {}] : [{}],
        };
      }
      return { value: [] };
    });
    vi.stubGlobal("window", { dataverseAPI: api });
    const onProgress = vi.fn();

    const result = await loadOwnershipCountsForOwners(
      ["owner-a", "owner-b"],
      onProgress,
    );

    expect(result.scannedEntities).toBe(2);
    expect(result.analyzedEntities).toBe(2);
    expect(result.failedEntities).toBe(0);
    expect(result.users).toEqual([
      {
        userId: "owner-a",
        entityCounts: [
          expect.objectContaining({
            entityLogicalName: "account",
            entityDisplayName: "Account",
            recordCount: 2,
          }),
        ],
        entitiesWithRecords: 1,
        totalOwnedRecords: 2,
      },
      {
        userId: "owner-b",
        entityCounts: [
          expect.objectContaining({
            entityLogicalName: "account",
            recordCount: 1,
          }),
        ],
        entitiesWithRecords: 1,
        totalOwnedRecords: 1,
      },
    ]);
    expect(onProgress).toHaveBeenCalled();
  });

  it("skips metadata without names or entity sets and handles partial owner query failures", async () => {
    const api = makeDataverseApi();
    api.getAllEntitiesMetadata.mockResolvedValue({
      value: [
        {
          LogicalName: "missing-set",
          OwnershipType: "UserOwned",
        },
        {
          EntitySetName: "nameless",
          OwnershipType: "UserOwned",
        },
        {
          LogicalName: "account",
          EntitySetName: "accounts",
          PrimaryIdAttribute: "accountid",
          OwnershipType: "UserOwned",
        },
      ],
    });
    api.queryData.mockImplementation(async (query: string) => {
      if (query.includes("owner-b")) {
        throw new Error("owner unavailable");
      }
      return { value: [{ accountid: "account-1" }] };
    });
    vi.stubGlobal("window", { dataverseAPI: api });
    vi.spyOn(logger, "warning").mockImplementation(() => undefined);

    const result = await loadOwnershipCountsForOwners(["owner-a", "owner-b"]);

    expect(result).toMatchObject({
      scannedEntities: 2,
      analyzedEntities: 1,
      failedEntities: 0,
      users: [
        { userId: "owner-a", totalOwnedRecords: 1 },
        { userId: "owner-b", totalOwnedRecords: 0 },
      ],
    });
    expect(logger.warning).toHaveBeenCalledWith(
      expect.stringContaining("Skipping account for owner-b"),
    );
  });

  it("reports an entity failure when all owner queries fail", async () => {
    const api = makeDataverseApi();
    api.getAllEntitiesMetadata.mockResolvedValue({
      value: [
        {
          LogicalName: "account",
          EntitySetName: "accounts",
          PrimaryIdAttribute: "accountid",
          OwnershipType: "UserOwned",
        },
      ],
    });
    api.queryData.mockRejectedValue(
      new Error("Error invoking remote method 'dataverse.queryData': denied"),
    );
    vi.stubGlobal("window", { dataverseAPI: api });
    vi.spyOn(logger, "warning").mockImplementation(() => undefined);
    const onProgress = vi.fn();

    const result = await loadOwnershipCountsForOwners(["owner-a"], onProgress);

    expect(result).toMatchObject({
      scannedEntities: 1,
      analyzedEntities: 0,
      failedEntities: 1,
      failedEntityDetails: [
        {
          entityLogicalName: "account",
          message: " denied",
        },
      ],
    });
    expect(onProgress).toHaveBeenLastCalledWith(
      expect.objectContaining({
        processedEntities: 1,
        failedEntities: 1,
        recentEntities: [
          expect.objectContaining({ status: "failed", recordsFound: 0 }),
        ],
      }),
    );
  });

  it("uses a fallback message when a query fails without an error message", async () => {
    const api = makeDataverseApi();
    api.getAllEntitiesMetadata.mockResolvedValue({
      value: [
        {
          LogicalName: "account",
          EntitySetName: "accounts",
          OwnershipType: "UserOwned",
        },
      ],
    });
    api.queryData.mockRejectedValue("");
    vi.stubGlobal("window", { dataverseAPI: api });
    vi.spyOn(logger, "warning").mockImplementation(() => undefined);

    const result = await loadOwnershipCountsForOwners(["owner-a"]);

    expect(result.failedEntityDetails[0]?.message).toBe(
      "Unknown error while counting records.",
    );
  });

  it("keeps only the 300 most recent analysis entries", async () => {
    const api = makeDataverseApi();
    api.getAllEntitiesMetadata.mockResolvedValue({
      value: Array.from({ length: 302 }, (_, index) => ({
        LogicalName: `entity${index}`,
        EntitySetName: `entities${index}`,
        PrimaryIdAttribute: "id",
        OwnershipType: "UserOwned",
      })),
    });
    api.queryData.mockResolvedValue({ value: [] });
    vi.stubGlobal("window", { dataverseAPI: api });
    const progress: any[] = [];

    await loadOwnershipCountsForOwners(["owner-1"], (entry) => progress.push(entry));

    const finalProgress = progress.at(-1);
    expect(finalProgress.recentEntities).toHaveLength(300);
    expect(finalProgress.recentEntities).not.toContainEqual(
      expect.objectContaining({ entityLogicalName: "entity0" }),
    );
  });

  it("returns successfully reassigned record IDs", async () => {
    const api = makeDataverseApi();
    api.queryData.mockResolvedValue({ value: [{ accountid: "record-1" }] });
    api.updateMultiple.mockResolvedValue(undefined);
    vi.stubGlobal("window", { dataverseAPI: api });

    const result = await reassignOwnedRecordsForUser(
      "owner-a",
      "systemuser",
      "user-b",
      "systemuser",
      [
        {
          entityLogicalName: "account",
          entityDisplayName: "Account",
          entitySetName: "accounts",
          primaryIdAttribute: "accountid",
          recordCount: 1,
        },
      ],
    );

    expect(result).toMatchObject({
      reassignedRecords: 1,
      failedRecords: 0,
      entityResults: [
        { assignedRecordIds: ["record-1"], failedRecordDetails: [] },
      ],
    });
    expect(api.updateMultiple).toHaveBeenCalledWith(
      "account",
      [expect.objectContaining({ "ownerid@odata.bind": "/systemusers(user-b)" })],
    );
  });

  it("reassigns records to a team and reports batch failures per record", async () => {
    const api = makeDataverseApi();
    api.queryData.mockResolvedValue({
      value: [{ accountid: "record-1" }, { accountid: "record-2" }],
    });
    api.updateMultiple.mockRejectedValue(new Error("assignment denied"));
    vi.stubGlobal("window", { dataverseAPI: api });
    vi.spyOn(logger, "warning").mockImplementation(() => undefined);
    const onProgress = vi.fn();

    const result = await reassignOwnedRecordsForUser(
      "{source-id}",
      "systemuser",
      "{team-id}",
      "team",
      [
        {
          entityLogicalName: "account",
          entityDisplayName: "Account",
          entitySetName: "accounts",
          primaryIdAttribute: "accountid",
          recordCount: 2,
        },
      ],
      onProgress,
    );

    expect(api.queryData).toHaveBeenCalledWith(
      "accounts?$select=accountid&$filter=_ownerid_value eq source-id",
    );
    expect(api.updateMultiple).toHaveBeenCalledWith(
      "account",
      [
        expect.objectContaining({
          accountid: "record-1",
          "ownerid@odata.bind": "/teams(team-id)",
        }),
        expect.objectContaining({
          accountid: "record-2",
          "ownerid@odata.bind": "/teams(team-id)",
        }),
      ],
    );
    expect(result).toMatchObject({
      reassignedRecords: 0,
      failedRecords: 2,
      processedEntities: 1,
      entityResults: [
        {
          failedRecordDetails: [
            { recordId: "record-1", error: "assignment denied" },
            { recordId: "record-2", error: "assignment denied" },
          ],
        },
      ],
    });
    expect(onProgress).toHaveBeenCalledWith({
      assignedRecords: 0,
      failedRecords: 2,
      totalRecords: 2,
      currentEntity: "Account",
    });
  });

  it("handles missing IDs, records without string IDs, and record-load failures", async () => {
    const api = makeDataverseApi();
    api.queryData
      .mockResolvedValueOnce({ value: [{ accountid: 42 }, { accountid: null }] })
      .mockRejectedValueOnce(
        new Error("Error invoking remote method 'dataverse.queryData': offline"),
      );
    vi.stubGlobal("window", { dataverseAPI: api });
    vi.spyOn(logger, "warning").mockImplementation(() => undefined);

    const result = await reassignOwnedRecordsForUser(
      "source-user",
      "systemuser",
      "{target-user}",
      "systemuser",
      [
        {
          entityLogicalName: "account",
          entityDisplayName: "Account",
          entitySetName: "accounts",
          primaryIdAttribute: "accountid",
          recordCount: 2,
        },
        {
          entityLogicalName: "contact",
          entityDisplayName: "Contact",
          entitySetName: "contacts",
          primaryIdAttribute: "contactid",
          recordCount: 4,
        },
        {
          entityLogicalName: "invalid",
          entityDisplayName: "Invalid",
          entitySetName: "",
          primaryIdAttribute: "id",
          recordCount: 1,
        },
      ],
    );

    expect(api.updateMultiple).not.toHaveBeenCalled();
    expect(result).toMatchObject({
      processedEntities: 2,
      reassignedRecords: 0,
      failedRecords: 4,
      entityResults: [
        { entityLogicalName: "account", failedRecords: 0, assignedRecordIds: [] },
        {
          entityLogicalName: "contact",
          failedRecords: 4,
          failedRecordDetails: [
            { recordId: "(all)", error: " offline" },
          ],
        },
      ],
    });
    expect(api.queryData.mock.calls[0]?.[0]).toContain("accountid");
  });

  it("does not call the progress callback if no entity has record IDs", async () => {
    const api = makeDataverseApi();
    api.queryData.mockResolvedValue({ value: [] });
    vi.stubGlobal("window", { dataverseAPI: api });
    const onProgress = vi.fn();

    const result = await reassignOwnedRecordsForUser(
      "source",
      "systemuser",
      "target",
      "systemuser",
      [
        {
          entityLogicalName: "account",
          entityDisplayName: "Account",
          entitySetName: "accounts",
          primaryIdAttribute: "accountid",
          recordCount: 1,
        },
      ],
      onProgress,
    );

    expect(result.reassignedRecords).toBe(0);
    expect(onProgress).not.toHaveBeenCalled();
  });
});
