import React, { useState, useCallback, useEffect, useMemo } from "react";
import {
  loadSystemUsers,
  loadTeams,
  loadOwnershipCountsForOwners,
  reassignOwnedRecordsForUser,
} from "../services/dataverseService";
import { SystemUser } from "../types/systemUser";
import { Team } from "../types/team";
import {
  OwnershipAssignmentResult,
  OwnershipAnalysisProgress,
  OwnershipAnalysisResult,
  OwnershipTargetType,
} from "../types/ownership";
import { Filter } from "./Filter";
import { DataGridView } from "./DataGridView";
import { makeStyles, Spinner, Text, Button } from "@fluentui/react-components";
import { logger } from "../services/loggerService";
import { OwnershipResultsDrawer } from "./OwnershipResultsDrawer";

interface IOverviewProps {
  connection: ToolBoxAPI.DataverseConnection | null;
}

export const Overview: React.FC<IOverviewProps> = ({ connection }) => {
  const [systemUsers, setSystemUsers] = useState<SystemUser[]>([]);
  const [teams, setTeams] = useState<Team[]>([]);
  const [entityType, setEntityType] = useState<"systemuser" | "team">(
    "systemuser",
  );
  const [textFilter, setTextFilter] = useState<string>("");
  const [statusFilter, setStatusFilter] = useState<
    "all" | "enabled" | "disabled"
  >("enabled");
  const [userTypeFilter, setUserTypeFilter] = useState<
    "all" | "users" | "applications"
  >("users");
  const [businessUnitFilter, setBusinessUnitFilter] = useState<string>("all");
  const [isLoading, setIsLoading] = useState(false);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [isOwnershipDrawerOpen, setIsOwnershipDrawerOpen] = useState(false);
  const [isLoadingOwnership, setIsLoadingOwnership] = useState(false);
  const [ownershipSourceType, setOwnershipSourceType] =
    useState<OwnershipTargetType>("systemuser");
  const [ownershipSourceIds, setOwnershipSourceIds] = useState<string[]>([]);
  const [ownershipResult, setOwnershipResult] =
    useState<OwnershipAnalysisResult | null>(null);
  const [ownershipProgress, setOwnershipProgress] =
    useState<OwnershipAnalysisProgress | null>(null);

  const useStyles = makeStyles({
    overviewRoot: {
      height: "100%",
      display: "flex",
      flexDirection: "column",
      gap: "16px",
      overflow: "hidden",
    },
    filterSection: {
      flexShrink: 0,
      display: "flex",
      justifyContent: "space-between",
      alignItems: "flex-end",
      gap: "16px",
    },
    buttonContainer: {
      display: "flex",
      gap: "8px",
      alignItems: "flex-end",
    },
    loadingContainer: {
      display: "flex",
      justifyContent: "center",
      alignItems: "center",
      padding: "40px",
      flexDirection: "column",
      gap: "16px",
    },
    dataSection: {
      flex: 1,
      display: "flex",
      flexDirection: "row",
      overflow: "hidden",
      minHeight: 0,
      gap: "0px",
    },
    gridSection: {
      flex: 1,
      display: "flex",
      flexDirection: "column",
      overflow: "hidden",
      minHeight: 0,
    },
    gridSectionDisabled: {
      pointerEvents: "none",
      userSelect: "none",
    },
  });

  const styles = useStyles();

  useEffect(() => {
    const initialize = async () => {
      if (!connection) {
        return;
      }
      await loadData();
    };

    initialize();
  }, [connection]);

  const loadData = useCallback(async () => {
    try {
      setIsLoading(true);
      const [users, teamsData] = await Promise.all([
        loadSystemUsers(),
        loadTeams(),
      ]);
      setSystemUsers(users);
      setTeams(teamsData);
      logger.info(
        `Fetched ${users.length} system users and ${teamsData.length} teams`,
      );
    } catch (error) {
      logger.error(`Error loading data: ${(error as Error).message}`);
      await window.toolboxAPI.utils.showNotification({
        title: "Error Loading Data",
        body: `Failed to load data from Dataverse: ${(error as Error).message}`,
        type: "error",
      });
    } finally {
      setIsLoading(false);
    }
  }, [connection]);

  const handleSelectionChange = useCallback((ids: string[]) => {
    setSelectedIds(ids);
  }, []);

  const handleAnalyzeOwnership = useCallback(async () => {
    if (selectedIds.length === 0) {
      return;
    }

    setIsOwnershipDrawerOpen(true);
    setIsLoadingOwnership(true);
    setOwnershipSourceType(entityType);
    setOwnershipSourceIds(selectedIds);
    setOwnershipResult(null);
    setOwnershipProgress(null);

    try {
      const result = await loadOwnershipCountsForOwners(
        selectedIds,
        setOwnershipProgress,
      );
      setOwnershipResult(result);
      logger.info(
        `Ownership analysis completed for ${selectedIds.length} owner(s). Scanned ${result.scannedEntities} entities.`,
      );
    } catch (error) {
      logger.error(
        `Error loading ownership analysis: ${(error as Error).message}`,
      );
      await window.toolboxAPI.utils.showNotification({
        title: "Ownership Analysis Failed",
        body: `Failed to analyze user ownership: ${(error as Error).message}`,
        type: "error",
      });
    } finally {
      setIsLoadingOwnership(false);
    }
  }, [selectedIds, entityType]);

  const handleAssignOwnership = useCallback(
    async (
      sourceOwnerId: string,
      targetOwnerType: OwnershipTargetType,
      targetOwnerId: string,
      selectedEntityLogicalNames: string[],
    ): Promise<OwnershipAssignmentResult> => {
      const sourceSummary = ownershipResult?.users.find(
        (user) => user.userId === sourceOwnerId,
      );

      if (!sourceSummary) {
        throw new Error("No ownership data found for selected source owner.");
      }

      const selectedEntitySet = new Set(selectedEntityLogicalNames);
      const selectedEntityCounts = sourceSummary.entityCounts.filter((entity) =>
        selectedEntitySet.has(entity.entityLogicalName),
      );

      const assignmentResult = await reassignOwnedRecordsForUser(
        sourceOwnerId,
        ownershipSourceType,
        targetOwnerId,
        targetOwnerType,
        selectedEntityCounts,
      );

      await window.toolboxAPI.utils.showNotification({
        title: "Ownership Assignment Completed",
        body: `Reassigned ${assignmentResult.reassignedRecords} record(s). Failed: ${assignmentResult.failedRecords}.`,
        type: assignmentResult.failedRecords > 0 ? "warning" : "success",
      });

      // Refresh analysis after reassignment to keep drawer data current.
      const refreshedResult = await loadOwnershipCountsForOwners(
        ownershipSourceIds,
        setOwnershipProgress,
      );
      setOwnershipResult(refreshedResult);

      return assignmentResult;
    },
    [ownershipResult, ownershipSourceIds, ownershipSourceType],
  );

  const filteredSystemUsers = useMemo(() => {
    let result = systemUsers;

    // Apply status filter
    if (statusFilter !== "all") {
      result = result.filter((user) => {
        if (statusFilter === "enabled") {
          return !user.isdisabled;
        } else {
          return user.isdisabled;
        }
      });
    }

    // Apply user type filter
    if (userTypeFilter !== "all") {
      result = result.filter((user) => {
        if (userTypeFilter === "users") {
          return !user.applicationid; // Regular users have no applicationid
        } else {
          return !!user.applicationid; // Application users have applicationid
        }
      });
    }

    // Apply business unit filter
    if (businessUnitFilter !== "all") {
      result = result.filter(
        (user) => user.businessunitid?.businessunitid === businessUnitFilter,
      );
    }

    // Apply text filter
    if (textFilter) {
      const searchTerm = textFilter.toLowerCase();
      result = result.filter((user) => {
        return (
          user.fullname?.toLowerCase().includes(searchTerm) ||
          user.domainname?.toLowerCase().includes(searchTerm) ||
          user.businessunitid?.name?.toLowerCase().includes(searchTerm)
        );
      });
    }

    return result;
  }, [
    systemUsers,
    textFilter,
    statusFilter,
    userTypeFilter,
    businessUnitFilter,
  ]);

  const filteredTeams = useMemo(() => {
    let result = teams;

    // Apply business unit filter
    if (businessUnitFilter !== "all") {
      result = result.filter(
        (team) => team.businessunitid?.businessunitid === businessUnitFilter,
      );
    }

    // Apply text filter
    if (textFilter) {
      const searchTerm = textFilter.toLowerCase();
      result = result.filter((team) => {
        return (
          team.name?.toLowerCase().includes(searchTerm) ||
          team.businessunitid?.name?.toLowerCase().includes(searchTerm)
        );
      });
    }

    return result;
  }, [teams, textFilter, businessUnitFilter]);

  return (
    <div className={styles.overviewRoot}>
      {isLoading ? (
        <div className={styles.loadingContainer}>
          <Spinner label="Loading data..." size="large" />
        </div>
      ) : (
        <>
          <div className={styles.filterSection}>
            <Filter
              entityType={entityType}
              systemUsers={systemUsers}
              teams={teams}
              statusFilter={statusFilter}
              userTypeFilter={userTypeFilter}
              businessUnitFilter={businessUnitFilter}
              textFilter={textFilter}
              onEntityTypeChanged={(type: "systemuser" | "team") => {
                logger.info(`Entity type changed to: ${type}`);
                setEntityType(type);
                setTextFilter("");
                setStatusFilter("enabled");
                setUserTypeFilter("all");
                setBusinessUnitFilter("all");
                setSelectedIds([]);
              }}
              onTextFilterChanged={(searchText: string) => {
                setTextFilter(searchText);
              }}
              onStatusFilterChanged={(
                status: "all" | "enabled" | "disabled",
              ) => {
                logger.info(`Status filter changed to: ${status}`);
                setStatusFilter(status);
              }}
              onUserTypeFilterChanged={(
                userType: "all" | "users" | "applications",
              ) => {
                logger.info(`User type filter changed to: ${userType}`);
                setUserTypeFilter(userType);
              }}
              onBusinessUnitFilterChanged={(businessUnitId: string) => {
                logger.info(
                  `Business unit filter changed to: ${businessUnitId}`,
                );
                setBusinessUnitFilter(businessUnitId);
              }}
            />
            {(entityType === "systemuser" || entityType === "team") && (
              <div className={styles.buttonContainer}>
                <Button
                  appearance="primary"
                  onClick={handleAnalyzeOwnership}
                  disabled={isLoadingOwnership || selectedIds.length === 0}
                >
                  {isLoadingOwnership
                    ? "Analyzing ownership..."
                    : "Analyze Ownership"}
                </Button>
              </div>
            )}
          </div>

          <div className={styles.dataSection}>
            <div
              className={`${styles.gridSection} ${
                isOwnershipDrawerOpen ? styles.gridSectionDisabled : ""
              }`}
              aria-hidden={isOwnershipDrawerOpen}
            >
              {entityType === "systemuser" &&
                filteredSystemUsers.length === 0 && (
                  <Text>No system users found.</Text>
                )}
              {entityType === "team" && filteredTeams.length === 0 && (
                <Text>No teams found.</Text>
              )}
              {((entityType === "systemuser" &&
                filteredSystemUsers.length > 0) ||
                (entityType === "team" && filteredTeams.length > 0)) && (
                <DataGridView
                  entityType={entityType}
                  systemUsers={filteredSystemUsers}
                  teams={filteredTeams}
                  selectedIds={selectedIds}
                  onSelectionChange={handleSelectionChange}
                />
              )}
            </div>
          </div>

          <OwnershipResultsDrawer
            open={isOwnershipDrawerOpen}
            isLoading={isLoadingOwnership}
            result={ownershipResult}
            progress={ownershipProgress}
            sourceOwnerType={ownershipSourceType}
            allSystemUsers={systemUsers.map((user) => ({
              userId: user.systemuserid,
              userName: user.fullname,
              userEmail: user.domainname,
            }))}
            allTeams={teams.map((team) => ({
              userId: team.teamid,
              userName: team.name,
            }))}
            users={ownershipSourceIds.map((id) => ({
              userId: id,
              userName:
                ownershipSourceType === "team"
                  ? (teams.find((team) => team.teamid === id)?.name ?? id)
                  : (systemUsers.find((user) => user.systemuserid === id)
                      ?.fullname ?? id),
            }))}
            onAssignRecords={handleAssignOwnership}
            onOpenChange={setIsOwnershipDrawerOpen}
          />
        </>
      )}
    </div>
  );
};
