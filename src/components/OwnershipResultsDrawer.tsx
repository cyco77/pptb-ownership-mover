import React, { useEffect, useMemo, useState } from "react";
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Badge,
  Body1,
  Button,
  Caption1,
  createTableColumn,
  DataGrid,
  DataGridBody,
  DataGridCell,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridRow,
  Dropdown,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  makeStyles,
  Option,
  OptionOnSelectData,
  SelectionEvents,
  OnSelectionChangeData,
  OverlayDrawer,
  ProgressBar,
  Spinner,
  TableColumnDefinition,
  Text,
  tokens,
} from "@fluentui/react-components";
import {
  DismissRegular,
  CheckmarkCircleRegular,
  ErrorCircleRegular,
  DatabaseSearchRegular,
} from "@fluentui/react-icons";
import {
  OwnershipAssignmentResult,
  OwnershipAnalysisProgress,
  OwnershipAnalysisResult,
  OwnershipAssignmentProgress,
  OwnershipTargetType,
} from "../types/ownership";
import {
  downloadAssignmentErrorsCsv,
  downloadAssignmentSummaryCsv,
  downloadCompleteOwnershipAnalysisCsv,
  type OwnershipAssignmentHistoryEntry,
} from "../services/ownershipCsvExportService";

type OwnershipTargetSelection = "user" | "application" | "team";

type OwnershipOwnerView = {
  userId: string;
  userName: string;
  domainName?: string;
  isApplication?: boolean;
  isDisabled?: boolean;
};

interface IOwnershipResultsDrawerProps {
  open: boolean;
  isLoading: boolean;
  users: OwnershipOwnerView[];
  sourceOwnerType: OwnershipTargetType;
  allSystemUsers: OwnershipOwnerView[];
  allTeams: OwnershipOwnerView[];
  result: OwnershipAnalysisResult | null;
  progress: OwnershipAnalysisProgress | null;
  onAssignRecords: (
    sourceOwnerId: string,
    targetOwnerType: OwnershipTargetType,
    targetOwnerId: string,
    selectedEntityLogicalNames: string[],
    onProgress?: (progress: OwnershipAssignmentProgress) => void,
  ) => Promise<OwnershipAssignmentResult>;
  onOpenChange: (open: boolean) => void;
}

const useStyles = makeStyles({
  drawerRoot: {
    width: "80vw",
    maxWidth: "80vw",
  },
  drawerBody: {
    height: "100%",
    minHeight: 0,
    display: "flex",
    flexDirection: "column",
    overflow: "hidden",
  },
  summarySection: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))",
    gap: tokens.spacingVerticalM,
    marginBottom: tokens.spacingVerticalL,
  },
  statCard: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalM,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: `0 1px 3px rgba(0, 0, 0, 0.08)`,
  },
  statCardHeader: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  statIcon: {
    fontSize: "24px",
    display: "flex",
    alignItems: "center",
  },
  statLabel: {
    fontSize: "12px",
    color: tokens.colorNeutralForeground3,
    fontWeight: 500,
    letterSpacing: "0.5px",
  },
  statValue: {
    fontSize: "32px",
    fontWeight: 700,
    color: tokens.colorNeutralForeground1,
    lineHeight: "1",
  },
  userSection: {
    marginTop: tokens.spacingVerticalS,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  userMetrics: {
    display: "flex",
    gap: tokens.spacingHorizontalL,
    flexWrap: "wrap",
  },
  tableContainer: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    overflow: "hidden",
    maxHeight: "260px",
    overflowY: "auto",
  },
  dataGrid: {
    width: "100%",
  },
  emptyState: {
    color: tokens.colorNeutralForeground3,
  },
  loadingContainer: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    alignItems: "stretch",
    justifyContent: "flex-start",
    flex: 1,
    minHeight: 0,
  },
  resultContainer: {
    flex: 1,
    minHeight: 0,
    overflowY: "auto",
    paddingRight: tokens.spacingHorizontalXS,
  },
  loadingHeader: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  progressBlock: {
    width: "calc(100% - 10px)",
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  loadingLogContainer: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingHorizontalM,
    flex: 1,
    minHeight: 0,
    overflowY: "auto",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  progressStats: {
    display: "flex",
    flexWrap: "wrap",
    gap: tokens.spacingHorizontalL,
  },
  recentList: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  recentItem: {
    display: "flex",
    justifyContent: "space-between",
    gap: tokens.spacingHorizontalM,
  },
  stickyHeaderCell: {
    position: "sticky",
    top: 0,
    zIndex: 2,
    backgroundColor: tokens.colorNeutralBackground1,
    fontWeight: 700,
  },
  controlRow: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    alignItems: "center",
    flexWrap: "wrap",
  },
  userSelect: {
    minWidth: "500px",
  },
  dropdownListbox: {
    zIndex: 20,
  },
  targetDropdownListbox: {
    zIndex: 20,
    width: "max-content",
    minWidth: "min(360px, calc(100vw - 32px))",
    maxWidth: "min(640px, calc(100vw - 32px))",
  },
  targetOption: {
    display: "flex",
    alignItems: "center",
    width: "100%",
    minWidth: 0,
    gap: tokens.spacingHorizontalM,
  },
  targetOptionName: {
    flex: 1,
    minWidth: 0,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  assignmentResult: {
    color: tokens.colorNeutralForeground3,
  },
  entitySelectionInfo: {
    color: tokens.colorNeutralForeground3,
  },
  summaryActions: {
    display: "flex",
    justifyContent: "flex-end",
    marginBottom: tokens.spacingVerticalS,
  },
  bottomActions: {
    display: "flex",
    justifyContent: "space-between",
    gap: tokens.spacingHorizontalM,
    marginTop: tokens.spacingVerticalM,
    paddingTop: tokens.spacingVerticalS,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  errorItem: {
    padding: `${tokens.spacingVerticalXS} 0`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXXS,
  },
});

export const OwnershipResultsDrawer: React.FC<IOwnershipResultsDrawerProps> = ({
  open,
  isLoading,
  users,
  sourceOwnerType,
  allSystemUsers,
  allTeams,
  result,
  progress,
  onAssignRecords,
  onOpenChange,
}) => {
  const styles = useStyles();
  const [targetTypeBySource, setTargetTypeBySource] = useState<
    Record<string, OwnershipTargetSelection>
  >({});
  const [targetIdBySource, setTargetIdBySource] = useState<
    Record<string, string>
  >({});
  const [selectedEntitiesBySource, setSelectedEntitiesBySource] = useState<
    Record<string, string[]>
  >({});
  const [assigningBySource, setAssigningBySource] = useState<
    Record<string, boolean>
  >({});
  const [assignmentProgressBySource, setAssignmentProgressBySource] = useState<
    Record<string, OwnershipAssignmentProgress>
  >({});
  const [assignmentErrorBySource, setAssignmentErrorBySource] = useState<
    Record<string, string>
  >({});
  const [assignmentBySource, setAssignmentBySource] = useState<
    Record<string, OwnershipAssignmentResult>
  >({});
  const [assignmentHistory, setAssignmentHistory] = useState<
    OwnershipAssignmentHistoryEntry[]
  >([]);

  const entityColumns: TableColumnDefinition<
    OwnershipAnalysisResult["users"][number]["entityCounts"][number]
  >[] = [
    createTableColumn({
      columnId: "entityDisplayName",
      renderHeaderCell: () => "Entity",
      renderCell: (item) => item.entityDisplayName,
      compare: (a, b) => a.entityDisplayName.localeCompare(b.entityDisplayName),
    }),
    createTableColumn({
      columnId: "entityLogicalName",
      renderHeaderCell: () => "Logical Name",
      renderCell: (item) => item.entityLogicalName,
      compare: (a, b) => a.entityLogicalName.localeCompare(b.entityLogicalName),
    }),
    createTableColumn({
      columnId: "recordCount",
      renderHeaderCell: () => "Record Count",
      renderCell: (item) => item.recordCount,
      compare: (a, b) => b.recordCount - a.recordCount,
    }),
  ];

  const getTargetsByType = useMemo(
    () =>
      (
        targetType: OwnershipTargetSelection,
        sourceOwnerId: string,
      ): OwnershipOwnerView[] => {
        const pool =
          targetType === "team"
            ? allTeams
            : allSystemUsers.filter(
                (candidate) => candidate.isApplication === (targetType === "application"),
              );
        const targetOwnerType: OwnershipTargetType =
          targetType === "team" ? "team" : "systemuser";
        if (targetOwnerType === sourceOwnerType) {
          return pool.filter((candidate) => candidate.userId !== sourceOwnerId);
        }
        return pool;
      },
    [allSystemUsers, allTeams, sourceOwnerType],
  );

  useEffect(() => {
    if (!open) {
      setTargetTypeBySource({});
      setTargetIdBySource({});
      setSelectedEntitiesBySource({});
      setAssigningBySource({});
      setAssignmentProgressBySource({});
      setAssignmentErrorBySource({});
      setAssignmentBySource({});
      setAssignmentHistory([]);
    }
  }, [open]);

  useEffect(() => {
    if (!result) {
      return;
    }

    const defaultTypes: Record<string, OwnershipTargetSelection> = {};
    const defaultIds: Record<string, string> = {};
    const defaultSelections: Record<string, string[]> = {};

    result.users.forEach((ownerSummary) => {
      const ownerId = ownerSummary.userId;
      defaultSelections[ownerId] = ownerSummary.entityCounts.map(
        (entry) => entry.entityLogicalName,
      );

      const preferredType: OwnershipTargetSelection = "user";
      defaultTypes[ownerId] = preferredType;
      defaultIds[ownerId] =
        getTargetsByType(preferredType, ownerId)[0]?.userId ?? "";
    });

    setTargetTypeBySource((current) => ({
      ...defaultTypes,
      ...current,
    }));
    setTargetIdBySource((current) => ({
      ...defaultIds,
      ...current,
    }));
    setSelectedEntitiesBySource((current) => ({
      ...defaultSelections,
      ...current,
    }));
  }, [result, sourceOwnerType, getTargetsByType]);

  const handleAssignAll = async (sourceOwnerId: string) => {
    const targetSelection = targetTypeBySource[sourceOwnerId] ?? "user";
    const targetType: OwnershipTargetType =
      targetSelection === "team" ? "team" : "systemuser";
    const targetId = targetIdBySource[sourceOwnerId] ?? "";
    const selectedEntities = selectedEntitiesBySource[sourceOwnerId] ?? [];

    if (!targetId || selectedEntities.length === 0) {
      return;
    }

    setAssigningBySource((current) => ({
      ...current,
      [sourceOwnerId]: true,
    }));
    setAssignmentErrorBySource((current) => ({
      ...current,
      [sourceOwnerId]: "",
    }));

    try {
      const assignmentResult = await onAssignRecords(
        sourceOwnerId,
        targetType,
        targetId,
        selectedEntities,
        (progress) =>
          setAssignmentProgressBySource((current) => ({
            ...current,
            [sourceOwnerId]: progress,
          })),
      );

      setAssignmentHistory((current) => [
        {
          ...assignmentResult,
          assignedAt: new Date().toISOString(),
        },
        ...current,
      ]);

      setAssignmentBySource((current) => ({
        ...current,
        [sourceOwnerId]: assignmentResult,
      }));
    } catch (error) {
      setAssignmentErrorBySource((current) => ({
        ...current,
        [sourceOwnerId]:
          error instanceof Error ? error.message : String(error),
      }));
    } finally {
      setAssigningBySource((current) => ({
        ...current,
        [sourceOwnerId]: false,
      }));
    }
  };

  const resolveOwnerName = (
    ownerId: string,
    ownerType: OwnershipTargetType,
  ): string => {
    const pool = ownerType === "team" ? allTeams : allSystemUsers;
    return pool.find((item) => item.userId === ownerId)?.userName ?? ownerId;
  };

  const resolveOwnerDomainName = (
    ownerId: string,
    ownerType: OwnershipTargetType,
  ): string => {
    if (ownerType === "team") {
      return "";
    }

    return (
      allSystemUsers.find((item) => item.userId === ownerId)?.domainName ?? ""
    );
  };

  const hasAssignmentErrors = assignmentHistory.some((a) =>
    a.entityResults.some((e) => e.failedRecordDetails.length > 0),
  );

  const downloadAssignmentSummary = () =>
    downloadAssignmentSummaryCsv({
      assignmentHistory,
      resolveOwnerName,
      resolveOwnerDomainName,
    });

  const downloadAssignmentErrors = () =>
    downloadAssignmentErrorsCsv({
      assignmentHistory,
      resolveOwnerName,
      resolveOwnerDomainName,
    });

  const downloadCompleteAnalysisSummary = () =>
    downloadCompleteOwnershipAnalysisCsv({
      result,
      users,
      sourceOwnerType,
    });

  return (
    <OverlayDrawer
      position="end"
      size="full"
      className={styles.drawerRoot}
      open={open}
      modalType="modal"
      onOpenChange={(_, data) => onOpenChange(data.open)}
    >
      <DrawerHeader>
        <DrawerHeaderTitle
          action={
            <Button
              appearance="subtle"
              aria-label="Close ownership results"
              icon={<DismissRegular />}
              onClick={() => onOpenChange(false)}
            />
          }
        >
          Ownership Analysis Results
        </DrawerHeaderTitle>
      </DrawerHeader>
      <DrawerBody className={styles.drawerBody}>
        {isLoading && (
          <div className={styles.loadingContainer}>
            <div className={styles.loadingHeader}>
              <Spinner label="Analyzing Dataverse entities..." size="medium" />
              <Caption1>
                Scanning owner-based entities and counting records by owner.
              </Caption1>
            </div>

            {progress && (
              <>
                <div className={styles.progressBlock}>
                  <Body1>
                    Processing {progress.processedEntities} of{" "}
                    {progress.totalEntities} entities
                  </Body1>
                  <ProgressBar
                    value={
                      progress.totalEntities > 0
                        ? progress.processedEntities / progress.totalEntities
                        : 0
                    }
                  />
                  <div className={styles.progressStats}>
                    <Caption1>
                      Analyzed: <strong>{progress.analyzedEntities}</strong>
                    </Caption1>
                    <Caption1>
                      Failed: <strong>{progress.failedEntities}</strong>
                    </Caption1>
                  </div>
                  <Caption1>
                    Current entity: {progress.currentEntityDisplayName} (
                    {progress.currentEntityLogicalName})
                  </Caption1>
                </div>

                <div className={styles.loadingLogContainer}>
                  <div className={styles.recentList}>
                    <Caption1>Analysis log</Caption1>
                    {progress.recentEntities.map((entry, index) => (
                      <div
                        key={`${entry.entityLogicalName}-${index}`}
                        className={styles.recentItem}
                      >
                        <Caption1>
                          {entry.entityDisplayName} ({entry.entityLogicalName})
                        </Caption1>
                        <Caption1>
                          {entry.status === "failed"
                            ? "Failed"
                            : `${entry.recordsFound} records`}
                        </Caption1>
                      </div>
                    ))}
                  </div>
                </div>
              </>
            )}
          </div>
        )}

        {!isLoading && result && (
          <div className={styles.resultContainer}>
            <div className={styles.summarySection}>
              <div className={styles.statCard}>
                <div className={styles.statCardHeader}>
                  <div
                    className={styles.statIcon}
                    style={{ color: tokens.colorBrandForeground1 }}
                  >
                    <DatabaseSearchRegular />
                  </div>
                  <span className={styles.statLabel}>Scanned Entities</span>
                </div>
                <div className={styles.statValue}>{result.scannedEntities}</div>
              </div>

              <div className={styles.statCard}>
                <div className={styles.statCardHeader}>
                  <div
                    className={styles.statIcon}
                    style={{ color: tokens.colorStatusSuccessForeground1 }}
                  >
                    <CheckmarkCircleRegular />
                  </div>
                  <span className={styles.statLabel}>
                    Successfully Analyzed
                  </span>
                </div>
                <div className={styles.statValue}>
                  {result.analyzedEntities}
                </div>
              </div>

              <div className={styles.statCard}>
                <div className={styles.statCardHeader}>
                  <div
                    className={styles.statIcon}
                    style={{ color: tokens.colorStatusDangerForeground1 }}
                  >
                    <ErrorCircleRegular />
                  </div>
                  <span className={styles.statLabel}>Failed Entities</span>
                </div>
                <div className={styles.statValue}>{result.failedEntities}</div>
              </div>
            </div>

            {result.failedEntityDetails.length > 0 && (
              <Accordion collapsible>
                <AccordionItem value="failed-entities">
                  <AccordionHeader>
                    Failed Entity Details ({result.failedEntityDetails.length})
                  </AccordionHeader>
                  <AccordionPanel>
                    {result.failedEntityDetails.map((entry) => (
                      <div
                        key={entry.entityLogicalName}
                        className={styles.errorItem}
                      >
                        <Text weight="semibold">
                          {entry.entityDisplayName} ({entry.entityLogicalName})
                        </Text>
                        <Caption1>{entry.message}</Caption1>
                      </div>
                    ))}
                  </AccordionPanel>
                </AccordionItem>
              </Accordion>
            )}

            <Accordion
              collapsible
              multiple
              defaultOpenItems={
                result.users.length > 0 ? [result.users[0].userId] : []
              }
            >
              {result.users.map((userSummary) => {
                const owner = users.find(
                  (item) => item.userId === userSummary.userId,
                );
                const visibleRows = userSummary.entityCounts.slice(0, 100);
                const selectedEntities =
                  selectedEntitiesBySource[userSummary.userId] ?? [];
                const selectedItems = new Set<string>(selectedEntities);
                const targetType =
                  targetTypeBySource[userSummary.userId] ?? "user";
                const targetCandidates = getTargetsByType(
                  targetType,
                  userSummary.userId,
                );
                const targetOwnerId =
                  targetIdBySource[userSummary.userId] ?? "";
                const selectedTargetOwnerName =
                  targetCandidates.find(
                    (candidate) => candidate.userId === targetOwnerId,
                  )?.userName ?? "";
                const formatTargetLabel = (candidate: OwnershipOwnerView) => {
                  const details = [
                    targetType === "application"
                      ? "Application"
                      : candidate.domainName,
                  ].filter(Boolean);
                  return details.length > 0
                    ? `${candidate.userName} (${details.join(", ")})`
                    : candidate.userName;
                };
                const assignment = assignmentBySource[userSummary.userId];

                return (
                  <AccordionItem
                    key={userSummary.userId}
                    value={userSummary.userId}
                  >
                    <AccordionHeader>
                      {owner?.userName ?? userSummary.userId}
                      {owner?.domainName && ` (${owner.domainName})`}
                    </AccordionHeader>
                    <AccordionPanel>
                      <section className={styles.userSection}>
                        <div className={styles.userMetrics}>
                          <Body1>
                            Total owned records:{" "}
                            <strong>{userSummary.totalOwnedRecords}</strong>
                          </Body1>
                          <Body1>
                            Entities with records:{" "}
                            <strong>{userSummary.entitiesWithRecords}</strong>
                          </Body1>
                        </div>

                        <Caption1 className={styles.entitySelectionInfo}>
                          Selected entities for assignment:{" "}
                          {selectedEntities.length}
                        </Caption1>

                        {visibleRows.length === 0 ? (
                          <Text className={styles.emptyState}>
                            No owned records found for this owner.
                          </Text>
                        ) : (
                          <div className={styles.tableContainer}>
                            <DataGrid
                              items={visibleRows}
                              columns={entityColumns}
                              sortable
                              defaultSortState={{
                                sortColumn: "recordCount",
                                sortDirection: "descending",
                              }}
                              selectionMode="multiselect"
                              selectedItems={selectedItems}
                              getRowId={(item) => item.entityLogicalName}
                              className={styles.dataGrid}
                              onSelectionChange={(
                                _e,
                                data: OnSelectionChangeData,
                              ) =>
                                setSelectedEntitiesBySource((current) => ({
                                  ...current,
                                  [userSummary.userId]: Array.from(
                                    data.selectedItems,
                                  ).map((id) => String(id)),
                                }))
                              }
                            >
                              <DataGridHeader>
                                <DataGridRow>
                                  {({ renderHeaderCell }) => (
                                    <DataGridHeaderCell
                                      className={styles.stickyHeaderCell}
                                    >
                                      {renderHeaderCell()}
                                    </DataGridHeaderCell>
                                  )}
                                </DataGridRow>
                              </DataGridHeader>
                              <DataGridBody>
                                {({ item, rowId }) => (
                                  <DataGridRow key={rowId}>
                                    {({ renderCell }) => (
                                      <DataGridCell>
                                        {renderCell(item)}
                                      </DataGridCell>
                                    )}
                                  </DataGridRow>
                                )}
                              </DataGridBody>
                            </DataGrid>
                          </div>
                        )}

                        <div className={styles.controlRow}>
                          <Dropdown
                            className={styles.userSelect}
                            inlinePopup
                            listbox={{ className: styles.dropdownListbox }}
                            value={
                              targetType === "team"
                                ? "Target: Team"
                                : targetType === "application"
                                  ? "Target: Application"
                                  : "Target: User"
                            }
                            selectedOptions={[targetType]}
                            onOptionSelect={(
                              _event: SelectionEvents,
                              data: OptionOnSelectData,
                            ) => {
                              const nextType = data.optionValue as
                                | OwnershipTargetSelection
                                | undefined;
                              if (!nextType) {
                                return;
                              }
                              const nextCandidates = getTargetsByType(
                                nextType,
                                userSummary.userId,
                              );
                              setTargetTypeBySource((current) => ({
                                ...current,
                                [userSummary.userId]: nextType,
                              }));
                              setTargetIdBySource((current) => ({
                                ...current,
                                [userSummary.userId]:
                                  nextCandidates[0]?.userId ?? "",
                              }));
                            }}
                          >
                            <Option value="user" text="Target: User">
                              Target: User
                            </Option>
                            <Option
                              value="application"
                              text="Target: Application"
                            >
                              Target: Application
                            </Option>
                            <Option value="team" text="Target: Team">
                              Target: Team
                            </Option>
                          </Dropdown>

                          <Dropdown
                            className={styles.userSelect}
                            inlinePopup
                            positioning={{
                              matchTargetSize: undefined,
                              autoSize: "width",
                            }}
                            listbox={{
                              className: styles.targetDropdownListbox,
                            }}
                            placeholder="Select target"
                            value={selectedTargetOwnerName}
                            selectedOptions={
                              targetOwnerId ? [targetOwnerId] : []
                            }
                            onOptionSelect={(
                              _event: SelectionEvents,
                              data: OptionOnSelectData,
                            ) =>
                              setTargetIdBySource((current) => ({
                                ...current,
                                [userSummary.userId]: data.optionValue ?? "",
                              }))
                            }
                          >
                            {targetCandidates.map((candidate) => (
                              <Option
                                key={candidate.userId}
                                value={candidate.userId}
                                text={formatTargetLabel(candidate)}
                              >
                                <span className={styles.targetOption}>
                                  <span className={styles.targetOptionName}>
                                    {formatTargetLabel(candidate)}
                                  </span>
                                  {candidate.isDisabled !== undefined && (
                                    <Badge
                                      appearance="tint"
                                      color={
                                        candidate.isDisabled
                                          ? "danger"
                                          : "success"
                                      }
                                      size="small"
                                    >
                                      {candidate.isDisabled
                                        ? "Inactive"
                                        : "Active"}
                                    </Badge>
                                  )}
                                </span>
                              </Option>
                            ))}
                          </Dropdown>

                          <Button
                            appearance="primary"
                            disabled={
                              !targetOwnerId ||
                              (assigningBySource[userSummary.userId] ??
                                false) ||
                              userSummary.totalOwnedRecords === 0 ||
                              selectedEntities.length === 0
                            }
                            onClick={() => handleAssignAll(userSummary.userId)}
                          >
                            {assigningBySource[userSummary.userId]
                              ? "Assigning..."
                              : "Assign selected records"}
                          </Button>
                        </div>

                        {assigningBySource[userSummary.userId] &&
                          (() => {
                            const ap =
                              assignmentProgressBySource[userSummary.userId];
                            const total = ap?.totalRecords ?? 0;
                            const done =
                              (ap?.assignedRecords ?? 0) +
                              (ap?.failedRecords ?? 0);
                            return (
                              <div
                                style={{
                                  display: "flex",
                                  flexDirection: "column",
                                  gap: tokens.spacingVerticalXS,
                                }}
                              >
                                <ProgressBar
                                  value={total > 0 ? done / total : undefined}
                                />
                                <Caption1>
                                  {ap
                                    ? `Assigning records: ${done} / ${total} — ${ap.currentEntity}`
                                    : "Preparing assignment..."}
                                </Caption1>
                              </div>
                            );
                          })()}

                        {assignment && (
                          <Caption1 className={styles.assignmentResult}>
                            Last assignment result: reassigned{" "}
                            {assignment.reassignedRecords}, failed{" "}
                            {assignment.failedRecords}.
                          </Caption1>
                        )}
                        {assignmentErrorBySource[userSummary.userId] && (
                          <Caption1 role="alert">
                            Assignment failed: {assignmentErrorBySource[userSummary.userId]}
                          </Caption1>
                        )}
                      </section>
                    </AccordionPanel>
                  </AccordionItem>
                );
              })}
            </Accordion>

            <div className={styles.bottomActions}>
              <Button
                appearance="secondary"
                onClick={downloadCompleteAnalysisSummary}
              >
                Download Complete Analysis CSV
              </Button>
              {assignmentHistory.length > 0 && (
                <div
                  style={{ display: "flex", gap: tokens.spacingHorizontalM }}
                >
                  {hasAssignmentErrors && (
                    <Button
                      appearance="secondary"
                      onClick={downloadAssignmentErrors}
                    >
                      Download Assignment Errors (CSV)
                    </Button>
                  )}
                  <Button
                    appearance="secondary"
                    onClick={downloadAssignmentSummary}
                  >
                    Download Assignment Summary (CSV)
                  </Button>
                </div>
              )}
            </div>
          </div>
        )}
      </DrawerBody>
    </OverlayDrawer>
  );
};
