// @vitest-environment node

import { afterEach, describe, expect, it, vi } from "vitest";
import {
  downloadAssignmentErrorsCsv,
  downloadAssignmentSummaryCsv,
  downloadCompleteOwnershipAnalysisCsv,
  downloadOwnershipAnalysisSummaryCsv,
} from "./ownershipCsvExportService";
import type {
  OwnershipAnalysisResult,
  OwnershipAssignmentResult,
} from "../types/ownership";

const assignment: OwnershipAssignmentResult & { assignedAt: string } = {
  assignedAt: "2026-01-01T00:00:00.000Z",
  sourceOwnerId: "source-id",
  sourceOwnerType: "systemuser",
  targetOwnerId: "target-id",
  targetOwnerType: "team",
  processedEntities: 1,
  reassignedRecords: 1,
  failedRecords: 1,
  entityResults: [
    {
      entityLogicalName: "account",
      entityDisplayName: "Account",
      reassignedRecords: 1,
      failedRecords: 1,
      assignedRecordIds: ["record-1"],
      failedRecordDetails: [
        { recordId: "record-2", error: 'Access denied: "policy"' },
      ],
    },
  ],
};

const analysis: OwnershipAnalysisResult = {
  scannedEntities: 2,
  analyzedEntities: 2,
  failedEntities: 0,
  failedEntityDetails: [],
  users: [
    {
      userId: "owner-1",
      totalOwnedRecords: 3,
      entitiesWithRecords: 1,
      entityCounts: [
        {
          entityLogicalName: "account",
          entityDisplayName: "Account",
          entitySetName: "accounts",
          primaryIdAttribute: "accountid",
          recordCount: 3,
        },
      ],
    },
    {
      userId: "owner-2",
      totalOwnedRecords: 0,
      entitiesWithRecords: 0,
      entityCounts: [],
    },
  ],
};

const captureCsvDownload = () => {
  const click = vi.fn();
  const anchor = { href: "", download: "", click };
  const createObjectURL = vi.fn((_blob: Blob) => "blob:csv");
  const revokeObjectURL = vi.fn();
  vi.stubGlobal("document", {
    createElement: vi.fn(() => anchor),
    body: { appendChild: vi.fn(), removeChild: vi.fn() },
  });
  vi.stubGlobal("URL", { createObjectURL, revokeObjectURL });
  return { anchor, createObjectURL, click, revokeObjectURL };
};

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("ownership CSV exports", () => {
  it("exports successful, failed, and empty assignment rows with resolved owners", async () => {
    const download = captureCsvDownload();
    const resolveOwnerName = vi.fn((id: string) => `${id}-name`);
    const resolveOwnerDomainName = vi.fn((id: string) => `${id}@example.com`);

    downloadAssignmentSummaryCsv({
      assignmentHistory: [
        assignment,
        {
          ...assignment,
          entityResults: [
            {
              ...assignment.entityResults[0],
              reassignedRecords: 0,
              failedRecords: 0,
              assignedRecordIds: [],
              failedRecordDetails: [],
            },
          ],
        },
      ],
      resolveOwnerName,
      resolveOwnerDomainName,
    });

    const blob = download.createObjectURL.mock.calls[0]?.[0];
    const content = await blob?.text();
    expect(download.anchor.download).toMatch(/^ownership-assignment-summary-/);
    expect(content).toContain('"Assigned At"');
    expect(Array.from(new Uint8Array(await blob!.arrayBuffer()).slice(0, 3))).toEqual([
      0xef, 0xbb, 0xbf,
    ]);
    expect(content).toContain('"record-1","Assigned"');
    expect(content).toContain('"record-2","Failed","Access denied: ""policy"""');
    expect(content).toContain('"","No records"');
    expect(resolveOwnerName).toHaveBeenCalledWith("source-id", "systemuser");
    expect(resolveOwnerDomainName).toHaveBeenCalledWith("target-id", "team");
    expect(download.click).toHaveBeenCalledOnce();
    expect(download.revokeObjectURL).toHaveBeenCalledWith("blob:csv");
  });

  it("uses empty owner domains when no domain resolver is provided", async () => {
    const download = captureCsvDownload();
    downloadAssignmentSummaryCsv({
      assignmentHistory: [assignment],
      resolveOwnerName: (id) => id,
    });

    const content = await download.createObjectURL.mock.calls[0]?.[0].text();
    expect(content).toContain('"source-id","","team","target-id",""');
  });

  it("exports only failed record details", async () => {
    const download = captureCsvDownload();
    downloadAssignmentErrorsCsv({
      assignmentHistory: [assignment],
      resolveOwnerName: (id) => id,
    });

    const content = await download.createObjectURL.mock.calls[0]?.[0].text();
    expect(download.anchor.download).toMatch(/^ownership-assignment-errors-/);
    expect(content).toContain('"record-2","Access denied: ""policy"""');
    expect(content).not.toContain("record-1");
  });

  it("does not create an assignment summary download when history is empty", () => {
    const download = captureCsvDownload();
    downloadAssignmentSummaryCsv({
      assignmentHistory: [],
      resolveOwnerName: (id) => id,
    });
    expect(download.createObjectURL).not.toHaveBeenCalled();
  });

  it("exports selected entity rows and skips an empty selection", async () => {
    const download = captureCsvDownload();
    downloadOwnershipAnalysisSummaryCsv({
      sourceOwnerType: "systemuser",
      sourceOwnerId: "owner-1",
      sourceOwnerName: "Ada",
      sourceOwnerDomainName: "ada@example.com",
      entityRows: analysis.users[0].entityCounts,
    });
    const content = await download.createObjectURL.mock.calls[0]?.[0].text();
    expect(download.anchor.download).toMatch(/^ownership-analysis-summary-owner-1-/);
    expect(content).toContain('"account","3"');

    downloadOwnershipAnalysisSummaryCsv({
      sourceOwnerType: "systemuser",
      sourceOwnerId: "owner-1",
      sourceOwnerName: "Ada",
      entityRows: [],
    });
    expect(download.createObjectURL).toHaveBeenCalledOnce();
  });

  it("exports full analysis including owners with zero records and skips missing analysis", async () => {
    const download = captureCsvDownload();
    downloadCompleteOwnershipAnalysisCsv({
      result: analysis,
      users: [{ userId: "owner-1", userName: "Ada", domainName: "ada@example.com" }],
      sourceOwnerType: "systemuser",
    });

    const content = await download.createObjectURL.mock.calls[0]?.[0].text();
    expect(download.anchor.download).toMatch(/^ownership-complete-analysis-/);
    expect(content).toContain('"owner-1","Ada"');
    expect(content).toContain('"owner-2","owner-2"');
    expect(content).toContain('"0","0","","","0"');

    downloadCompleteOwnershipAnalysisCsv({
      result: null,
      users: [],
      sourceOwnerType: "team",
    });
    expect(download.createObjectURL).toHaveBeenCalledOnce();
  });

  it("exports a complete analysis with an empty user list and no entity rows", async () => {
    const download = captureCsvDownload();
    downloadCompleteOwnershipAnalysisCsv({
      result: {
        ...analysis,
        users: [
          {
            userId: "owner-empty",
            totalOwnedRecords: 0,
            entitiesWithRecords: 0,
            entityCounts: [],
          },
        ],
      },
      users: [],
      sourceOwnerType: "team",
    });

    const content = await download.createObjectURL.mock.calls[0]?.[0].text();
    expect(content).toContain('"team","owner-empty","owner-empty"');
    expect(content).toContain('"0","0","","","0"');
  });
});
