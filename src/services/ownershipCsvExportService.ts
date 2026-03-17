import {
  OwnershipAnalysisResult,
  OwnershipAssignmentResult,
  OwnershipTargetType,
} from "../types/ownership";
import { createCsvTimestamp, downloadCsv, toCsvLine } from "../utils/csv";

export type OwnershipOwnerView = {
  userId: string;
  userName: string;
};

export type OwnershipAssignmentHistoryEntry = OwnershipAssignmentResult & {
  assignedAt: string;
};

type AssignmentSummaryParams = {
  assignmentHistory: OwnershipAssignmentHistoryEntry[];
  resolveOwnerName: (ownerId: string, ownerType: OwnershipTargetType) => string;
};

export const downloadAssignmentSummaryCsv = ({
  assignmentHistory,
  resolveOwnerName,
}: AssignmentSummaryParams): void => {
  if (assignmentHistory.length === 0) {
    return;
  }

  const lines: string[] = [
    toCsvLine([
      "Assigned At",
      "Source Type",
      "Source Owner",
      "Target Type",
      "Target Owner",
      "Entity Display Name",
      "Entity Logical Name",
      "Reassigned Records",
      "Failed Records",
    ]),
  ];

  assignmentHistory.forEach((assignment) => {
    const sourceOwnerName = resolveOwnerName(
      assignment.sourceOwnerId,
      assignment.sourceOwnerType,
    );
    const targetOwnerName = resolveOwnerName(
      assignment.targetOwnerId,
      assignment.targetOwnerType,
    );

    assignment.entityResults.forEach((entity) => {
      lines.push(
        toCsvLine([
          assignment.assignedAt,
          assignment.sourceOwnerType,
          sourceOwnerName,
          assignment.targetOwnerType,
          targetOwnerName,
          entity.entityDisplayName,
          entity.entityLogicalName,
          entity.reassignedRecords,
          entity.failedRecords,
        ]),
      );
    });
  });

  downloadCsv(
    lines,
    `ownership-assignment-summary-${createCsvTimestamp()}.csv`,
  );
};

type AnalysisSummaryParams = {
  sourceOwnerType: OwnershipTargetType;
  sourceOwnerId: string;
  sourceOwnerName: string;
  entityRows: OwnershipAnalysisResult["users"][number]["entityCounts"];
};

export const downloadOwnershipAnalysisSummaryCsv = ({
  sourceOwnerType,
  sourceOwnerId,
  sourceOwnerName,
  entityRows,
}: AnalysisSummaryParams): void => {
  if (entityRows.length === 0) {
    return;
  }

  const lines: string[] = [
    toCsvLine([
      "Source Owner Type",
      "Source Owner Id",
      "Source Owner Name",
      "Entity Display Name",
      "Entity Logical Name",
      "Record Count",
    ]),
  ];

  entityRows.forEach((row) => {
    lines.push(
      toCsvLine([
        sourceOwnerType,
        sourceOwnerId,
        sourceOwnerName,
        row.entityDisplayName,
        row.entityLogicalName,
        row.recordCount,
      ]),
    );
  });

  downloadCsv(
    lines,
    `ownership-analysis-summary-${sourceOwnerId}-${createCsvTimestamp()}.csv`,
  );
};

type CompleteAnalysisParams = {
  result: OwnershipAnalysisResult | null;
  users: OwnershipOwnerView[];
  sourceOwnerType: OwnershipTargetType;
};

export const downloadCompleteOwnershipAnalysisCsv = ({
  result,
  users,
  sourceOwnerType,
}: CompleteAnalysisParams): void => {
  if (!result || result.users.length === 0) {
    return;
  }

  const lines: string[] = [
    toCsvLine([
      "Source Owner Type",
      "Source Owner Id",
      "Source Owner Name",
      "Total Owned Records",
      "Entities With Records",
      "Entity Display Name",
      "Entity Logical Name",
      "Record Count",
    ]),
  ];

  result.users.forEach((ownerSummary) => {
    const ownerName =
      users.find((user) => user.userId === ownerSummary.userId)?.userName ??
      ownerSummary.userId;

    if (ownerSummary.entityCounts.length === 0) {
      lines.push(
        toCsvLine([
          sourceOwnerType,
          ownerSummary.userId,
          ownerName,
          ownerSummary.totalOwnedRecords,
          ownerSummary.entitiesWithRecords,
          "",
          "",
          0,
        ]),
      );
      return;
    }

    ownerSummary.entityCounts.forEach((entity) => {
      lines.push(
        toCsvLine([
          sourceOwnerType,
          ownerSummary.userId,
          ownerName,
          ownerSummary.totalOwnedRecords,
          ownerSummary.entitiesWithRecords,
          entity.entityDisplayName,
          entity.entityLogicalName,
          entity.recordCount,
        ]),
      );
    });
  });

  downloadCsv(lines, `ownership-complete-analysis-${createCsvTimestamp()}.csv`);
};
