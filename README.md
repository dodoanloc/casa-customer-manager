# CASA Customer Manager

CASA customer/account management and reporting webapp.

## Chức năng

- Nhập dữ liệu CSV/XLSX nghiệp vụ.
- Tra cứu và lọc dữ liệu khách hàng/tài khoản.
- Theo dõi dữ liệu DP01 và phạm vi nghiệp vụ liên quan.
- Xuất báo cáo XLSX/PDF.
- Quản lý file upload và export tại runtime.
- Backend API phục vụ frontend quản trị.

Dữ liệu thật không nằm trong GitHub. `data/`, `uploads/`, `exports/` là runtime storage trên máy triển khai.

## Cấu trúc

```text
server.py              FastAPI application, API, import, report export
public/index.html      Main UI
public/app.js          Frontend behavior/API client
public/styles.css      UI styling
requirements.txt       Python dependencies
.gitignore              Runtime/secret exclusions
.github/workflows/      CI and production deploy workflow
scripts/                Deployment and health-check scripts
```

## Runtime

```text
Service: csv-report-webapp-8099.service
Port: 8099
WorkingDirectory: /home/locdodoan/.openclaw/workspace/csv-report-webapp-6666
Health: http://127.0.0.1:8099/
```

## Local run

```bash
uv venv
source .venv/bin/activate
uv pip install -r requirements.txt
uvicorn server:app --host 127.0.0.1 --port 8099
```

## GitHub deploy flow

```text
feature/fix branch
  -> Pull Request to main
  -> CI: syntax + dependency/import checks
  -> CODEOWNERS review by @ktnqagribanktx-lang
  -> merge main
  -> production Environment approval
  -> SSH deploy release SHA
  -> backup DB
  -> restart csv-report-webapp-8099.service
  -> health check
```

Production deploy does not copy uncommitted local files and does not commit runtime data.

## Security

- Never commit DB, CSV upload, export, backup, credentials, token or PII.
- Backup `data/app.db` before deploy/migration.
- Preserve previous release SHA for rollback.
- Rotate any credential exposed outside secret storage.

See [`ROLLBACK.md`](ROLLBACK.md).
