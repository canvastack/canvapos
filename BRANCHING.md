# Branching Strategy — CanvaPOS

## Environments

| Environment | Use | Confirmations | Debug Logging | Sheet Data |
|------------|-----|---------------|---------------|------------|
| **Production** | Live UMKM | ON | OFF | Real data |
| **Staging** | UAT / testing | ON | ON | Copy of production |
| **Development** | Local dev | OFF (skip) | ON | Dummy data |

## Setup

Set environment via:

```javascript
// In Script Editor → Execution Log
setupEnv("development"); // or "staging" or "production"
```

Or via menu: POS → Set Environment → pick from dialog.

## Git Branching

```
main        → Production (live GAS deployment)
├── staging → Staging (separate spreadsheet + GAS project)
└── dev     → Development (local editing, dummy data)
```

### Workflow

1. **dev** branch — all feature work, code changes, refactoring
2. Merge to **staging** — deploy to staging GAS project for UAT
3. Merge to **main** — deploy to production GAS project

## GAS Deployment

Each environment uses separate:
- Google Sheet (different spreadsheet ID)
- GAS project (different script ID) — or same project with different deployments

### Same GAS project approach

1. Create 3 copies of the spreadsheet
2. Bind the same GAS script to all 3
3. Set `CANVAPOS_ENV` via `PropertiesService.getDocumentProperties()` per sheet

### Separate GAS project approach (recommended)

1. Development: `clasp push` to dev project
2. Staging: `clasp push` to staging project
3. Production: `clasp push` to production project

## Configuration Per Environment

```javascript
// Production
HARGA_BASE = 5000
HARGA_TOPPING = 1000
Admin email: null (live)

// Staging
HARGA_BASE = 5000 (same as prod for accurate testing)
HARGA_TOPPING = 1000
Admin email: developer@example.com

// Development
HARGA_BASE = 5000
HARGA_TOPPING = 1000
Skip confirmations: true
Debug logging: true
```

## Security

- Never commit `PropertiesService` values to git
- Admin emails stored in PropertiesService, not in code
- Each environment's spreadsheet has separate sharing permissions
