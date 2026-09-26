# MyBook

MyBook is a local personal finance and asset tracking application. It imports bank statements, broker reports, email attachments, local files and selected web APIs, then stores transactions, balances, holdings, snapshots and fixed bootstrap data in MySQL.

The project was developed primarily through natural-language-driven development with OpenAI GPT-5 Codex, with a small amount of manual editing.

## Tech Stack

- .NET 8 WPF desktop application
- MySQL database
- SqlSugar ORM
- MailKit for mailbox access
- System.IO.Ports for USB SIM modem access
- PdfPig and HtmlAgilityPack for statement parsing
- Newtonsoft.Json for JSON and GraphQL data

## Repository Layout

- `MyBook/`: WPF application source.
- `Database/bootstrap.sql`: tracked schema for rebuilding an empty database.
- `Database/bootstrap.fixed-data.sql`: private local fixed-data export, used with the schema for a full rebuild.
- `MyBook/config.json.example`: tracked configuration template with blank or zero values.

Local statements, downloaded reports, `config.json`, database backups and other private or runtime files are excluded from version control.

## Local Initial Reports

Place initial IBKR CSV reports named `IBKR_INITIAL_*.csv` in the private, ignored `initialReports/` directory. Accounts without imported history read these local reports before continuing with mailbox imports.

## Data Sources

- **ICBC / BOC:** credit-card monthly statements from email and debit-card SMS through a SIM device; ICBC historical details from email attachments.
- **IBKR:** daily CSV reports from email, plus local initial CSV reports.
- **iFAST:** transaction emails and local monthly statements; interest rates from update emails and the official website.
- **ZA:** transaction notification emails.
- **Ant / Ele:** one configured account per bank, with transaction notifications fetched from Yahoo Mail.
- **FirstTrade:** balances, holdings and history through the account API, plus CSV/OFX downloads using the same login session and kept in memory. Automatic imports are temporarily paused; manual imports remain available.
- **Wise:** multi-currency balances, activities, transfer details and payment receipts through a read-only personal-token API.
- **Schwab:** investment account data through Plaid.
- **PayPal:** transaction information from each account's linked Plaid connection and mailbox.
- **Kraken:** balances and ledger entries through a read-only API; prices through public market-data endpoints.
- **Ethereum:** transactions and balances for configured addresses through blockchain query endpoints; prices through public market-data endpoints.
- **Nexus:** monthly Donation Points income through GraphQL.
- **Google Drive:** read-only report access restricted to the shared `Reports` folder.

Configure the corresponding accounts and fixed import starting points before importing.

## Configuration

Create a local configuration file from the example and fill in the integrations you use:

```powershell
Copy-Item MyBook\config.json.example MyBook\config.json
```

`config.json` contains private credentials and is excluded from Git. Main settings:

- `database_connection`: MySQL connection string; an empty value uses the built-in local default.
- `yahoo_user` / `yahoo_pass`, `gmail_user` / `gmail_app_pwd`: statement-mail credentials.
- `mail_proxy`: optional proxy for PayPal mail and FirstTrade. Other modules fetch mail and attachments directly, ignoring system proxies. `pubweb_proxy`: optional public market-data proxy. Empty proxy settings use direct connections.
- `alphavantage_key`: market-data key; `ib_gateway_port`: Interactive Brokers gateway port.
- `nexus_api_key`: Nexus personal API key used by current imports.
- `kraken_api_key` / `kraken_api_secret`, `etherscan_api_key`: credentials for read-only Kraken and Ethereum queries.
- `sim_imsi`: expected SIM IMSI; leave empty to disable polling. `sim_poll_interval_minutes` defaults to 5 when unset or below 1.
- `GoogleCloudServeAccountKey`: Google service-account JSON key. Share the `Reports` folder with its `client_email` as Viewer; no additional Google Cloud/IAM roles are required.

### Plaid / Schwab

Set `plaid_client_id` and `plaid_production_secret`, then authorize from the build-output directory:

```powershell
dotnet MyBook.dll --plaid-link --country US --product investments
```

Production is the default. Sandbox requires changing the compile-time environment and setting `plaid_sandbox_secret`.

Before importing Schwab data, bind the Plaid connection to its local account, with one investment account per connection.

### Wise

Set `wise_api_token` to a personal API token. Current imports use the API; the old Plaid Wise integration is no longer active.

### PayPal

Bind each PayPal account to its Plaid connection and configure its mailbox to fetch both sources. Mail access uses the `mail_proxy` setting.

### FirstTrade

Set `firsttrade_username`, `firsttrade_password` and `firsttrade_totp_secret` (the original Base32 authenticator key, not a six-digit code). The read-only integration references `MaxxRK/firstrade-api` and uses `mail_proxy` when configured.

### Nexus OAuth

Current imports use `nexus_api_key`; OAuth authorization is not currently used by scheduled imports. To authorize or refresh a local Nexus OAuth token:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --debug-authorize-nexus-oauth
```

This opens the Nexus authorization page in the browser and receives the callback at `http://127.0.0.1:4700/callback`.

## Database

The application validates its MySQL schema on startup. Accounts, registered account identifiers, Plaid connections, fixed import starting points and start snapshots are fixed data preserved during cleanup. Imported records, holdings, other snapshots and OAuth tokens are runtime data.

Rebuild an empty database using the tracked schema and local fixed-data file:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --rebuild-database-from-bootstrap-sql
```

`Database/bootstrap.fixed-data.sql` contains private account metadata and unencrypted Plaid access tokens. It and its backups are excluded from Git and should be stored securely.

Automatic backup and manual export are local debug extensions, not included in the repository. When installed, they save schema, fixed data and versioned backup pairs under `Database` in the application directory; normal startup also runs automatic backup.

Create a start snapshot:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --create-start-snapshot
```

## Build

```powershell
dotnet build MyBook\MyBook.csproj -v minimal /p:UseSharedCompilation=false
```

## Imports and Scheduling

Import progress and scheduling use dates in the runtime machine's local time zone. Source timestamps are converted before taking their dates. Statement periods and transaction dates retain their financial meaning; integrations handle UTC days, US Eastern dates and other source-specific query ranges. Changing the runtime time zone changes calendar-day scheduling boundaries.

Release builds run an import cycle on startup and daily afterward; Debug builds do not schedule imports. Each cycle runs configured, enabled integrations whose query intervals have elapsed. Intervals must be positive and become due on the specified day. Failed queries do not restart the interval. A zero missing-report deadline disables only overdue errors, not network, parsing or financial validation errors.

SMS polling uses its own configured interval. Mail imports share IMAP sessions and download matching attachments.

Each cycle also refreshes exchange rates and creates a snapshot. The UI shows the active task, last run and next run. Failures create `MyBook.import-failed.tmp` in the application directory; the clear-marker button removes the warning without retrying. Successful imports do not clear it automatically.
