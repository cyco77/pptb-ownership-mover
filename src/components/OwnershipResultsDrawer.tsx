import React, { useEffect, useMemo, useState } from "react";
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
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
import { DismissRegular } from "@fluentui/react-icons";
import {
  OwnershipAssignmentResult,
  OwnershipAnalysisProgress,
  OwnershipAnalysisResult,
  OwnershipTargetType,
} from "../types/ownership";

type OwnershipOwnerView = {
  userId: string;
  userName: string;
  userEmail?: string;
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
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXXS,
    marginBottom: tokens.spacingVerticalM,
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
  },
  controlRow: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    alignItems: "center",
    flexWrap: "wrap",
  },
  userSelect: {
    minWidth: "320px",
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    "& button": {
      whiteSpace: "nowrap",
      overflow: "hidden",
      textOverflow: "ellipsis",
    },
  },
  dropdownListbox: {
    zIndex: 20,
    "& [role='option']": {
      minHeight: "32px",
      display: "flex",
      alignItems: "center",
      paddingTop: tokens.spacingVerticalXS,
      paddingBottom: tokens.spacingVerticalXS,
      whiteSpace: "nowrap",
      overflow: "hidden",
      textOverflow: "ellipsis",
    },
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
    justifyContent: "flex-end",
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
    Record<string, OwnershipTargetType>
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
  const [assignmentBySource, setAssignmentBySource] = useState<
    Record<string, OwnershipAssignmentResult>
  >({});
  const [assignmentHistory, setAssignmentHistory] = useState<
    Array<OwnershipAssignmentResult & { assignedAt: string }>
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
        targetType: OwnershipTargetType,
        sourceOwnerId: string,
      ): OwnershipOwnerView[] => {
        const pool = targetType === "team" ? allTeams : allSystemUsers;
        if (targetType === sourceOwnerType) {
          return pool.filter((candidate) => candidate.userId !== sourceOwnerId);
        }
        return pool;
      },
    [allSystemUsers, allTeams, sourceOwnerType],
  );

  useEffect(() => {
    if (!result) {
      return;
    }

    const defaultTypes: Record<string, OwnershipTargetType> = {};
    const defaultIds: Record<string, string> = {};
    const defaultSelections: Record<string, string[]> = {};

    result.users.forEach((ownerSummary) => {
      const ownerId = ownerSummary.userId;
      defaultSelections[ownerId] = ownerSummary.entityCounts.map(
        (entry) => entry.entityLogicalName,
      );

      const preferredType: OwnershipTargetType = "systemuser";
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
    const targetType = targetTypeBySource[sourceOwnerId] ?? "systemuser";
    const targetId = targetIdBySource[sourceOwnerId] ?? "";
    const selectedEntities = selectedEntitiesBySource[sourceOwnerId] ?? [];

    if (!targetId || selectedEntities.length === 0) {
      return;
    }

    setAssigningBySource((current) => ({
      ...current,
      [sourceOwnerId]: true,
    }));

    try {
      const assignmentResult = await onAssignRecords(
        sourceOwnerId,
        targetType,
        targetId,
        selectedEntities,
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

  const downloadAssignmentSummary = () => {
    if (assignmentHistory.length === 0) {
      return;
    }

    const lines: string[] = [];
    lines.push(
      [
        "Assigned At",
        "Source Type",
        "Source Owner",
        "Target Type",
        "Target Owner",
        "Entity Display Name",
        "Entity Logical Name",
        "Reassigned Records",
        "Failed Records",
      ]
        .map((item) => `"${item}"`)
        .join(","),
    );

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
          [
            assignment.assignedAt,
            assignment.sourceOwnerType,
            sourceOwnerName,
            assignment.targetOwnerType,
            targetOwnerName,
            entity.entityDisplayName,
            entity.entityLogicalName,
            String(entity.reassignedRecords),
            String(entity.failedRecords),
          ]
            .map((item) => `"${String(item).replace(/"/g, '""')}"`)
            .join(","),
        );
      });
    });

    const blob = new Blob(["\uFEFF" + lines.join("\r\n")], {
      type: "text/csv;charset=utf-8;",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    link.href = url;
    link.download = `ownership-assignment-summary-${timestamp}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const downloadOwnershipAnalysisSummary = (
    sourceOwnerId: string,
    sourceOwnerName: string,
    entityRows: OwnershipAnalysisResult["users"][number]["entityCounts"],
  ) => {
    if (entityRows.length === 0) {
      return;
    }

    const csvLines: string[] = [];
    csvLines.push(
      [
        "Source Owner Type",
        "Source Owner Id",
        "Source Owner Name",
        "Entity Display Name",
        "Entity Logical Name",
        "Record Count",
      ]
        .map((item) => `"${item}"`)
        .join(","),
    );

    entityRows.forEach((row) => {
      csvLines.push(
        [
          sourceOwnerType,
          sourceOwnerId,
          sourceOwnerName,
          row.entityDisplayName,
          row.entityLogicalName,
          String(row.recordCount),
        ]
          .map((item) => `"${String(item).replace(/"/g, '""')}"`)
          .join(","),
      );
    });

    const blob = new Blob(["\uFEFF" + csvLines.join("\r\n")], {
      type: "text/csv;charset=utf-8;",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    link.href = url;
    link.download = `ownership-analysis-summary-${sourceOwnerId}-${timestamp}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const downloadCompleteAnalysisSummary = () => {
    if (!result || result.users.length === 0) {
      return;
    }

    const csvLines: string[] = [];
    csvLines.push(
      [
        "Source Owner Type",
        "Source Owner Id",
        "Source Owner Name",
        "Total Owned Records",
        "Entities With Records",
        "Entity Display Name",
        "Entity Logical Name",
        "Record Count",
      ]
        .map((item) => `"${item}"`)
        .join(","),
    );

    result.users.forEach((ownerSummary) => {
      const ownerName =
        users.find((user) => user.userId === ownerSummary.userId)?.userName ??
        ownerSummary.userId;

      if (ownerSummary.entityCounts.length === 0) {
        csvLines.push(
          [
            sourceOwnerType,
            ownerSummary.userId,
            ownerName,
            String(ownerSummary.totalOwnedRecords),
            String(ownerSummary.entitiesWithRecords),
            "",
            "",
            "0",
          ]
            .map((item) => `"${String(item).replace(/"/g, '""')}"`)
            .join(","),
        );
        return;
      }

      ownerSummary.entityCounts.forEach((entity) => {
        csvLines.push(
          [
            sourceOwnerType,
            ownerSummary.userId,
            ownerName,
            String(ownerSummary.totalOwnedRecords),
            String(ownerSummary.entitiesWithRecords),
            entity.entityDisplayName,
            entity.entityLogicalName,
            String(entity.recordCount),
          ]
            .map((item) => `"${String(item).replace(/"/g, '""')}"`)
            .join(","),
        );
      });
    });

    const blob = new Blob(["\uFEFF" + csvLines.join("\r\n")], {
      type: "text/csv;charset=utf-8;",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    link.href = url;
    link.download = `ownership-complete-analysis-${timestamp}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

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
            {assignmentHistory.length > 0 && (
              <div className={styles.summaryActions}>
                <Button
                  appearance="secondary"
                  onClick={downloadAssignmentSummary}
                >
                  Download Assignment Summary (CSV)
                </Button>
              </div>
            )}

            <div className={styles.summarySection}>
              <Body1>
                Scanned entities: <strong>{result.scannedEntities}</strong>
              </Body1>
              <Body1>
                Successfully analyzed:{" "}
                <strong>{result.analyzedEntities}</strong>
              </Body1>
              <Body1>
                Failed entities: <strong>{result.failedEntities}</strong>
              </Body1>
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
                  targetTypeBySource[userSummary.userId] ?? "systemuser";
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
                  const isSystemUserTarget = targetType === "systemuser";
                  if (isSystemUserTarget && candidate.userEmail) {
                    return `${candidate.userName} (${candidate.userEmail})`;
                  }
                  return candidate.userName;
                };
                const assignment = assignmentBySource[userSummary.userId];

                return (
                  <AccordionItem
                    key={userSummary.userId}
                    value={userSummary.userId}
                  >
                    <AccordionHeader>
                      {owner?.userName ?? userSummary.userId} -{" "}
                      {userSummary.totalOwnedRecords} records in{" "}
                      {userSummary.entitiesWithRecords} entities
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

                        <div className={styles.controlRow}>
                          <Dropdown
                            className={styles.userSelect}
                            inlinePopup
                            listbox={{ className: styles.dropdownListbox }}
                            value={
                              targetType === "team"
                                ? "Target: Team"
                                : "Target: User"
                            }
                            selectedOptions={[targetType]}
                            onOptionSelect={(
                              _event: SelectionEvents,
                              data: OptionOnSelectData,
                            ) => {
                              const nextType =
                                data.optionValue as OwnershipTargetType;
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
                            <Option value="systemuser" text="Target: User">
                              Target: User
                            </Option>
                            <Option value="team" text="Target: Team">
                              Target: Team
                            </Option>
                          </Dropdown>

                          <Dropdown
                            className={styles.userSelect}
                            inlinePopup
                            listbox={{ className: styles.dropdownListbox }}
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
                                {formatTargetLabel(candidate)}
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

                          <Button
                            appearance="secondary"
                            disabled={visibleRows.length === 0}
                            onClick={() =>
                              downloadOwnershipAnalysisSummary(
                                userSummary.userId,
                                owner?.userName ?? userSummary.userId,
                                visibleRows,
                              )
                            }
                          >
                            Download Analysis CSV
                          </Button>
                        </div>

                        <Caption1 className={styles.entitySelectionInfo}>
                          Selected entities for assignment:{" "}
                          {selectedEntities.length}
                        </Caption1>

                        {assignment && (
                          <Caption1 className={styles.assignmentResult}>
                            Last assignment result: reassigned{" "}
                            {assignment.reassignedRecords}, failed{" "}
                            {assignment.failedRecords}.
                          </Caption1>
                        )}

                        {visibleRows.length === 0 ? (
                          <Text className={styles.emptyState}>
                            No owned records found for this owner.
                          </Text>
                        ) : (
                          <div className={styles.tableContainer}>
                            <DataGrid
                              items={visibleRows}
                              columns={entityColumns}
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
            </div>
          </div>
        )}
      </DrawerBody>
    </OverlayDrawer>
  );
};
