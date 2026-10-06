export const normalizeOwnershipType = (value: unknown): string => {
  if (typeof value === "string") {
    return value.toLowerCase();
  }

  if (typeof value === "number") {
    return String(value);
  }

  if (value && typeof value === "object") {
    const candidate = (value as Record<string, unknown>).Value;
    if (typeof candidate === "number") {
      return String(candidate);
    }
    if (typeof candidate === "string") {
      return candidate.toLowerCase();
    }
  }

  return "";
};

export const isUserAssignableEntity = (
  entity: Record<string, unknown>,
): boolean => {
  const ownershipType = normalizeOwnershipType(entity.OwnershipType);

  return (
    ownershipType.includes("userowned") ||
    ownershipType.includes("teamowned") ||
    ownershipType === "0" ||
    ownershipType === "4"
  );
};

export const getEntityDisplayName = (
  entity: Record<string, unknown>,
): string => {
  const logicalName = String(entity.LogicalName ?? "");
  const displayName = entity.DisplayName as
    | { LocalizedLabels?: Array<{ Label?: string }> }
    | undefined;

  return (
    displayName?.LocalizedLabels?.find((label) => !!label.Label)?.Label ??
    logicalName
  );
};

const DATAVERSE_QUERY_ERROR_PREFIX =
  "Error invoking remote method 'dataverse.queryData':";

export const normalizeDataverseErrorMessage = (error: unknown): string => {
  const message = error instanceof Error ? error.message : String(error);

  return message.startsWith(DATAVERSE_QUERY_ERROR_PREFIX)
    ? message.slice(DATAVERSE_QUERY_ERROR_PREFIX.length)
    : message;
};

export const sanitizeGuid = (id: string): string => id.replace(/[{}]/g, "");

export const buildOwnedRecordsQuery = (
  entitySetName: string,
  primaryIdAttribute: string,
  ownerId: string,
): string => {
  const selectClause = primaryIdAttribute
    ? `$select=${primaryIdAttribute}&`
    : "";
  return `${entitySetName}?${selectClause}$filter=_ownerid_value eq ${sanitizeGuid(ownerId)}`;
};

export const chunkArray = <T>(items: T[], chunkSize: number): T[][] => {
  if (chunkSize <= 0) {
    return [items];
  }

  const chunks: T[][] = [];
  for (let index = 0; index < items.length; index += chunkSize) {
    chunks.push(items.slice(index, index + chunkSize));
  }
  return chunks;
};
