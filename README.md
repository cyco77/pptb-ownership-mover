# Ownership Mover

![Ownership Mover](https://raw.githubusercontent.com/cyco77/pptb-ownership-mover/main/icon/ownership-mover_small.png)

A Power Platform Toolbox (PPTB) tool to analyze Dataverse record ownership and reassign owned records from users or teams to a new owner.

## Screenshots

### Dark Theme

![Ownership Mover - Dark Theme](https://raw.githubusercontent.com/cyco77/pptb-ownership-mover/main/screenshots/main_dark.png)

### Light Theme

![Ownership Mover - Light Theme](https://raw.githubusercontent.com/cyco77/pptb-ownership-mover/main/screenshots/main_light.png)

## Features

- Analyze ownership for selected system users or teams
- Scan user-owned and team-owned Dataverse entities using metadata
- Show live analysis progress (processed, analyzed, failed entities, current entity)
- Filter owners by:
  - Entity type (System Users / Teams)
  - User status (All / Enabled / Disabled)
  - User type (All / Users / Applications)
  - Business unit
  - Free-text search
- Review ownership per owner with:
  - Total owned records
  - Entities with owned records
  - Entity-level record counts
- Reassign selected records to a target system user or team
- Download CSV exports:
  - Per-owner analysis summary
  - Complete analysis summary
  - Assignment history summary

## License

MIT
