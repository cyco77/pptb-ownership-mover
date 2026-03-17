import { useCallback, useEffect, useState } from "react";
import { Overview } from "./components/Overview";
import { useConnection } from "./hooks/useConnection";
import { useToolboxEvents } from "./hooks/useToolboxEvents";
import { logger } from "./services/loggerService";
import {
  FluentProvider,
  Theme,
  teamsLightTheme,
  teamsDarkTheme,
  makeStyles,
  tokens,
  Title3,
  Text,
} from "@fluentui/react-components";
import iconImage from "../icon/user-team-security_small.png";

const useStyles = makeStyles({
  container: {
    backgroundColor: tokens.colorNeutralBackground1,
    height: "100vh",
    display: "flex",
    flexDirection: "column",
    overflow: "hidden",
  },
  header: {
    padding: tokens.spacingVerticalL,
    paddingBottom: tokens.spacingVerticalS,
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
    flexShrink: 0,
  },
  headerTitle: {
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-start",
    gap: tokens.spacingVerticalXXS,
  },
  headerTitleRow: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
  },
  headerIcon: {
    height: "50px",
    objectFit: "contain",
  },
  subtitle: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase300,
  },
  content: {
    padding: tokens.spacingVerticalL,
    flex: 1,
    overflow: "hidden",
    minHeight: 0,
  },
});

function App() {
  const { connection, refreshConnection } = useConnection();

  const [theme, setTheme] = useState<Theme>(teamsDarkTheme);
  const styles = useStyles();
  // Handle platform events
  const handleEvent = useCallback(
    (event: string, _data: any) => {
      console.log(`Received event: ${event}`);
      switch (event) {
        case "connection:updated":
        case "connection:created":
          refreshConnection();
          break;

        case "connection:deleted":
          refreshConnection();
          break;

        case "terminal:output":
        case "terminal:command:completed":
        case "terminal:error":
          // Terminal events handled by dedicated components
          break;
        case "settings:updated":
          // Settings updated, could refresh settings if needed
          updateThemeBasedOnSettings();
          logger.info(`Settings updated`);
          break;
      }
    },
    [refreshConnection],
  );

  async function updateThemeBasedOnSettings() {
    const theme = await window.toolboxAPI.utils.getCurrentTheme();
    if (theme === "dark") {
      setTheme(teamsDarkTheme);
    } else {
      setTheme(teamsLightTheme);
    }
    logger.info(`Theme updated:${theme}`);
  }

  useToolboxEvents(handleEvent);

  // Add initial log (run only once on mount)
  useEffect(() => {
    const initialite = async () => {
      await updateThemeBasedOnSettings();
      logger.info(`Initialized`);
    };

    initialite();
  }, []);

  return (
    <FluentProvider theme={theme}>
      <div className={styles.container}>
        <div className={styles.header}>
          <img
            src={iconImage}
            alt="Ownership Mover"
            className={styles.headerIcon}
          />
          <div className={styles.headerTitle}>
            <Title3>Ownership Mover</Title3>
            <Text className={styles.subtitle}>
              A Power Platform Toolbox tool to move ownership of records easily.
            </Text>
          </div>
        </div>
        <div className={styles.content}>
          <Overview connection={connection} />
        </div>
      </div>
    </FluentProvider>
  );
}

export default App;
