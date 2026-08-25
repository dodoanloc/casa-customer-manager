# CASA rollback

## Before deploy

Record current commit/release and backup database:

```bash
cp data/app.db "data/app.db.$(date +%Y%m%d-%H%M%S).bak"
```

## Code rollback

Production deployment must use immutable release directories. Roll back to previous release SHA, restart service, then health-check:

```bash
sudo systemctl stop csv-report-webapp-8099.service
# point current release symlink to previous verified release
sudo systemctl start csv-report-webapp-8099.service
curl -fsS http://127.0.0.1:8099/
```

Production rollback requires admin approval and an audit record.

## Database

Do not restore database automatically. Restore only after confirming schema/data impact and preserving current backup.

## GitHub deploy rule

- Workflow runs only after merge to `main`.
- `production` Environment requires reviewer approval.
- Deploy records commit SHA, backup path, service and health result.
- Failed health check stops rollout and preserves previous release.

## Current limitation

Current service uses local working tree. Immutable release deployment must be enabled and tested in staging before production workflow restarts service automatically.

Service:

```text
csv-report-webapp-8099.service
```

Port:

```text
8099
``` 

Expected health status: HTTP `200`.
