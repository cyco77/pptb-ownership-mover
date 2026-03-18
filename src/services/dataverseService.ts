import { SystemUser } from "../types/systemUser";
import { Team } from "../types/team";
import {
  OwnershipAssignmentEntityResult,
  OwnershipAssignmentResult,
  OwnershipAnalysisProgress,
  EntityOwnershipCount,
  OwnershipEntityFailure,
  OwnershipAnalysisResult,
  OwnershipTargetType,
  UserOwnershipSummary,
} from "../types/ownership";
import { logger } from "./loggerService";

export const loadSystemUsers = async (): Promise<SystemUser[]> => {
  let url =
    "systemusers?$select=systemuserid,fullname,domainname,isdisabled,applicationid&$expand=businessunitid($select=businessunitid,name)&$orderby=fullname";

  const allRecords = await loadAllData(url);

  return allRecords.map((record: any) => ({
    systemuserid: record.systemuserid,
    fullname: record.fullname,
    domainname: record.domainname,
    isdisabled: record.isdisabled,
    applicationid: record.applicationid ?? null,
    businessunitid: record.businessunitid
      ? {
          businessunitid: record.businessunitid.businessunitid,
          name: record.businessunitid.name,
        }
      : undefined,
  }));
};

export const loadTeams = async (): Promise<Team[]> => {
  let url =
    "teams?$select=teamid,name,teamtype,isdefault&$expand=businessunitid($select=businessunitid,name)&$orderby=name";

  const allRecords = await loadAllData(url);

  return allRecords.map((record: any) => ({
    teamid: record.teamid,
    name: record.name,
    teamtype: record.teamtype,
    isdefault: record.isdefault,
    businessunitid: record.businessunitid
      ? {
          businessunitid: record.businessunitid.businessunitid,
          name: record.businessunitid.name,
        }
      : undefined,
  }));
};

const loadAllData = async (fullUrl: string) => {
  const allRecords = [];

  while (fullUrl) {
    logger.info(`Fetching data from URL: ${fullUrl}`);

    let relativePath = fullUrl;

    if (fullUrl.startsWith("http")) {
      const url = new URL(fullUrl);
      const apiRegex = /^\/api\/data\/v\d+\.\d+\//;
      relativePath = url.pathname.replace(apiRegex, "") + url.search;
    }

    logger.info(`Cleaned URL: ${relativePath}`);

    const response = await window.dataverseAPI.queryData(relativePath);

    // Add the current page of results
    allRecords.push(...response.value);

    // Check for paging link
    fullUrl = (response as any)["@odata.nextLink"] || null;
  }

  console.log(`Total records fetched: ${allRecords.length}`, allRecords);

  return allRecords;
};

const normalizeOwnershipType = (value: unknown): string => {
  if (typeof value === "string") {
    return value.toLowerCase();
  }

  if (typeof value === "number") {
    return String(value);
  }

  if (value && typeof value === "object") {
    const typedValue = value as Record<string, unknown>;
    const candidate = typedValue.Value;
    if (typeof candidate === "number") {
      return String(candidate);
    }

    if (typeof candidate === "string") {
      return candidate.toLowerCase();
    }
  }

  return "";
};

const isUserAssignableEntity = (entity: Record<string, unknown>): boolean => {
  const ownershipType = normalizeOwnershipType(entity.OwnershipType);

  return (
    ownershipType.includes("userowned") ||
    ownershipType.includes("teamowned") ||
    ownershipType === "0" ||
    ownershipType === "4"
  );
};

const getEntityDisplayName = (entity: Record<string, unknown>): string => {
  const logicalName = String(entity.LogicalName ?? "");
  const displayName = entity.DisplayName as
    | {
        LocalizedLabels?: Array<{ Label?: string }>;
      }
    | undefined;

  const localized = displayName?.LocalizedLabels?.find(
    (label) => !!label.Label,
  )?.Label;

  return localized ?? logicalName;
};

const DATAVERSE_QUERY_ERROR_PREFIX =
  "Error invoking remote method 'dataverse.queryData':";

const normalizeDataverseErrorMessage = (error: unknown): string => {
  const message = error instanceof Error ? error.message : String(error);

  if (message.startsWith(DATAVERSE_QUERY_ERROR_PREFIX)) {
    return message.slice(DATAVERSE_QUERY_ERROR_PREFIX.length);
  }

  return message;
};

const countOwnedRecordsForUser = async (
  entitySetName: string,
  primaryIdAttribute: string,
  systemUserId: string,
): Promise<number> => {
  const sanitizedUserId = systemUserId.replace(/[{}]/g, "");
  const selectClause = primaryIdAttribute
    ? `$select=${primaryIdAttribute}&`
    : "";
  const query = `${entitySetName}?${selectClause}$filter=_ownerid_value eq ${sanitizedUserId}`;
  const records = await loadAllData(query);

  return records.length;
};

export const loadOwnershipCountsForOwners = async (
  ownerIds: string[],
  onProgress?: (progress: OwnershipAnalysisProgress) => void,
): Promise<OwnershipAnalysisResult> => {
  if (ownerIds.length === 0) {
    return {
      scannedEntities: 0,
      analyzedEntities: 0,
      failedEntities: 0,
      failedEntityDetails: [],
      users: [],
    };
  }

  const metadata = await window.dataverseAPI.getAllEntitiesMetadata([
    "LogicalName",
    "DisplayName",
    "EntitySetName",
    "PrimaryIdAttribute",
    "OwnershipType",
  ]);

  const entityDefinitions = metadata.value.filter((entity) => {
    const entitySetName = entity.EntitySetName;
    return typeof entitySetName === "string" && isUserAssignableEntity(entity);
  });

  const userEntityCounts = new Map<string, EntityOwnershipCount[]>();
  ownerIds.forEach((id) => userEntityCounts.set(id, []));

  let processedEntities = 0;
  let analyzedEntities = 0;
  let failedEntities = 0;
  const failedEntityDetails: OwnershipEntityFailure[] = [];
  const recentEntities: OwnershipAnalysisProgress["recentEntities"] = [];
  const maxParallelEntities = 10;

  const emitProgress = (
    currentEntityLogicalName: string,
    currentEntityDisplayName: string,
  ) => {
    if (!onProgress) {
      return;
    }

    onProgress({
      totalEntities: entityDefinitions.length,
      processedEntities,
      analyzedEntities,
      failedEntities,
      currentEntityLogicalName,
      currentEntityDisplayName,
      recentEntities: [...recentEntities],
    });
  };

  let nextEntityIndex = 0;

  const processSingleEntity = async (entity: Record<string, unknown>) => {
    const logicalName = String(entity.LogicalName ?? "");
    const entitySetName = String(entity.EntitySetName ?? "");
    const entityDisplayName = getEntityDisplayName(entity);

    if (!logicalName || !entitySetName) {
      return;
    }

    emitProgress(logicalName, entityDisplayName);

    const primaryIdAttribute = String(entity.PrimaryIdAttribute ?? "");

    const perUserCounts = await Promise.all(
      ownerIds.map(async (userId) => {
        try {
          const count = await countOwnedRecordsForUser(
            entitySetName,
            primaryIdAttribute,
            userId,
          );
          return { userId, count, success: true, errorMessage: "" };
        } catch (error) {
          const errorMessage = normalizeDataverseErrorMessage(error);
          logger.warning(
            `Skipping ${logicalName} for ${userId}: ${errorMessage}`,
          );
          return { userId, count: 0, success: false, errorMessage };
        }
      }),
    );

    if (!perUserCounts.some((item) => item.success)) {
      const representativeError =
        perUserCounts.find((item) => item.errorMessage)?.errorMessage ??
        "Unknown error while counting records.";
      failedEntities += 1;
      processedEntities += 1;
      failedEntityDetails.push({
        entityLogicalName: logicalName,
        entityDisplayName,
        message: representativeError,
      });
      recentEntities.unshift({
        entityLogicalName: logicalName,
        entityDisplayName,
        status: "failed",
        recordsFound: 0,
      });
      if (recentEntities.length > 300) {
        recentEntities.pop();
      }
      emitProgress(logicalName, entityDisplayName);
      return;
    }

    analyzedEntities += 1;
    processedEntities += 1;

    perUserCounts.forEach(({ userId, count }) => {
      if (count <= 0) {
        return;
      }

      const existing = userEntityCounts.get(userId) ?? [];
      existing.push({
        entityLogicalName: logicalName,
        entityDisplayName,
        entitySetName,
        primaryIdAttribute: String(entity.PrimaryIdAttribute ?? ""),
        recordCount: count,
      });
      userEntityCounts.set(userId, existing);
    });

    const recordsFound = perUserCounts.reduce(
      (sum, item) => sum + (item.count > 0 ? item.count : 0),
      0,
    );
    recentEntities.unshift({
      entityLogicalName: logicalName,
      entityDisplayName,
      status: "analyzed",
      recordsFound,
    });
    if (recentEntities.length > 300) {
      recentEntities.pop();
    }
    emitProgress(logicalName, entityDisplayName);
  };

  const workerCount = Math.min(maxParallelEntities, entityDefinitions.length);
  await Promise.all(
    Array.from({ length: workerCount }, async () => {
      while (true) {
        const currentIndex = nextEntityIndex;
        nextEntityIndex += 1;

        if (currentIndex >= entityDefinitions.length) {
          break;
        }

        await processSingleEntity(entityDefinitions[currentIndex]);
      }
    }),
  );

  const users: UserOwnershipSummary[] = ownerIds.map((userId: string) => {
    const entityCounts = (userEntityCounts.get(userId) ?? []).sort(
      (a, b) => b.recordCount - a.recordCount,
    );
    const totalOwnedRecords = entityCounts.reduce(
      (sum, item) => sum + item.recordCount,
      0,
    );

    return {
      userId,
      entityCounts,
      entitiesWithRecords: entityCounts.length,
      totalOwnedRecords,
    };
  });

  return {
    scannedEntities: entityDefinitions.length,
    analyzedEntities,
    failedEntities,
    failedEntityDetails,
    users,
  };
};

const loadOwnedRecordIdsForEntity = async (
  entitySetName: string,
  primaryIdAttribute: string,
  systemUserId: string,
): Promise<string[]> => {
  const sanitizedUserId = systemUserId.replace(/[{}]/g, "");
  const query = `${entitySetName}?$select=${primaryIdAttribute}&$filter=_ownerid_value eq ${sanitizedUserId}`;
  const records = await loadAllData(query);

  return records
    .map((record: Record<string, unknown>) => record[primaryIdAttribute])
    .filter((value): value is string => typeof value === "string");
};

const chunkArray = <T>(items: T[], chunkSize: number): T[][] => {
  if (chunkSize <= 0) {
    return [items];
  }

  const chunks: T[][] = [];
  for (let index = 0; index < items.length; index += chunkSize) {
    chunks.push(items.slice(index, index + chunkSize));
  }

  return chunks;
};

export const reassignOwnedRecordsForUser = async (
  sourceOwnerId: string,
  sourceOwnerType: OwnershipTargetType,
  targetOwnerId: string,
  targetOwnerType: OwnershipTargetType,
  entityCounts: EntityOwnershipCount[],
  onProgress?: (progress: {
    assignedRecords: number;
    failedRecords: number;
    totalRecords: number;
    currentEntity: string;
  }) => void,
): Promise<OwnershipAssignmentResult> => {
  const entityResults: OwnershipAssignmentEntityResult[] = [];
  let reassignedRecords = 0;
  let failedRecords = 0;
  const totalRecords = entityCounts.reduce((sum, e) => sum + e.recordCount, 0);
  const bulkUpdateChunkSize = 200;
  const ownerBinding = `/${
    targetOwnerType === "team" ? "teams" : "systemusers"
  }(${targetOwnerId.replace(/[{}]/g, "")})`;

  for (const entity of entityCounts) {
    if (!entity.primaryIdAttribute || !entity.entitySetName) {
      continue;
    }

    let entityReassigned = 0;
    let entityFailed = 0;
    const assignedRecordIds: string[] = [];
    const entityFailedDetails: Array<{ recordId: string; error: string }> = [];

    try {
      const recordIds = await loadOwnedRecordIdsForEntity(
        entity.entitySetName,
        entity.primaryIdAttribute,
        sourceOwnerId,
      );

      for (const recordIdBatch of chunkArray(recordIds, bulkUpdateChunkSize)) {
        const updateRecords = recordIdBatch.map((recordId) => ({
          [entity.primaryIdAttribute]: recordId,
          "@odata.type": `Microsoft.Dynamics.CRM.${entity.entityLogicalName}`,
          "ownerid@odata.bind": ownerBinding,
        }));

        try {
          await window.dataverseAPI.updateMultiple(
            entity.entityLogicalName,
            updateRecords,
          );
          entityReassigned += recordIdBatch.length;
          assignedRecordIds.push(...recordIdBatch);
        } catch (error) {
          const errorMessage = normalizeDataverseErrorMessage(error);
          entityFailed += recordIdBatch.length;
          recordIdBatch.forEach((recordId) => {
            entityFailedDetails.push({
              recordId,
              error: errorMessage,
            });
          });
          logger.warning(
            `Failed to bulk reassign ${entity.entityLogicalName} batch (${recordIdBatch.length} records): ${errorMessage}`,
          );
        }

        onProgress?.({
          assignedRecords: reassignedRecords + entityReassigned,
          failedRecords: failedRecords + entityFailed,
          totalRecords,
          currentEntity: entity.entityDisplayName,
        });
      }
    } catch (error) {
      entityFailed += entity.recordCount;
      const errorMessage = normalizeDataverseErrorMessage(error);
      entityFailedDetails.push({
        recordId: "(all)",
        error: errorMessage,
      });
      logger.warning(
        `Failed to load records for ${entity.entityLogicalName}: ${errorMessage}`,
      );
    }

    reassignedRecords += entityReassigned;
    failedRecords += entityFailed;
    entityResults.push({
      entityLogicalName: entity.entityLogicalName,
      entityDisplayName: entity.entityDisplayName,
      reassignedRecords: entityReassigned,
      failedRecords: entityFailed,
      assignedRecordIds,
      failedRecordDetails: entityFailedDetails,
    });
  }

  return {
    sourceOwnerId,
    sourceOwnerType,
    targetOwnerId,
    targetOwnerType,
    processedEntities: entityResults.length,
    reassignedRecords,
    failedRecords,
    entityResults,
  };
};
