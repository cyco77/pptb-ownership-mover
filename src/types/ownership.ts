export type EntityOwnershipCount = {
  entityLogicalName: string;
  entityDisplayName: string;
  entitySetName: string;
  primaryIdAttribute: string;
  recordCount: number;
};

export type OwnershipTargetType = "systemuser" | "team";

export type OwnershipEntityFailure = {
  entityLogicalName: string;
  entityDisplayName: string;
  message: string;
};

export type UserOwnershipSummary = {
  userId: string;
  totalOwnedRecords: number;
  entitiesWithRecords: number;
  entityCounts: EntityOwnershipCount[];
};

export type OwnershipAnalysisResult = {
  scannedEntities: number;
  analyzedEntities: number;
  failedEntities: number;
  failedEntityDetails: OwnershipEntityFailure[];
  users: UserOwnershipSummary[];
};

export type OwnershipProgressEntity = {
  entityLogicalName: string;
  entityDisplayName: string;
  status: "analyzed" | "failed";
  recordsFound: number;
};

export type OwnershipAnalysisProgress = {
  totalEntities: number;
  processedEntities: number;
  analyzedEntities: number;
  failedEntities: number;
  currentEntityLogicalName: string;
  currentEntityDisplayName: string;
  recentEntities: OwnershipProgressEntity[];
};

export type OwnershipAssignmentEntityResult = {
  entityLogicalName: string;
  entityDisplayName: string;
  reassignedRecords: number;
  failedRecords: number;
};

export type OwnershipAssignmentResult = {
  sourceOwnerId: string;
  sourceOwnerType: OwnershipTargetType;
  targetOwnerId: string;
  targetOwnerType: OwnershipTargetType;
  processedEntities: number;
  reassignedRecords: number;
  failedRecords: number;
  entityResults: OwnershipAssignmentEntityResult[];
};
