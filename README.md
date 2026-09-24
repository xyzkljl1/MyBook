# MyBook

MyBook is a local personal finance and asset tracking application. It imports statements from banks, broker reports, mail attachments, local files, and selected web APIs, then stores records, balances, holdings, snapshots, and fixed bootstrap data in MySQL.

The project has been developed with extensive vibe coding using OpenAI GPT-5 Codex, plus a small amount of manual editing.

## Tech Stack

- .NET 8 WPF desktop application
- MySQL
- SqlSugar ORM
- MailKit for mailbox access
- System.IO.Ports for USB SIM modem access
- PdfPig and HtmlAgilityPack for statement parsing
- Newtonsoft.Json for JSON and GraphQL payloads

## Repository Layout

- `MyBook/` - WPF application source
- `Database/bootstrap.sql` - tracked database schema used to rebuild an empty database
- `Database/bootstrap.fixed-data.sql` - ignored local fixed-data export used with the schema for a full rebuild
- `MyBook/config.json.example` - tracked configuration template with blank or zero values
- `*.TODO.cs` modules - placeholders or not-yet-verified integrations
- `*.NotUsed.cs` modules - integrations kept for reference but not currently used

Local statements, downloaded reports, `config.json`, database backups, and other private/runtime files are intentionally ignored.

## Local Initial Reports

`initialReports/` is an ignored private directory for account history that predates normal recurring imports. When an IBKR account has no imported history, the importer first reads matching local files from this directory whose periods start after the provider checkpoint, validates that the initial statement starts from zero and that multi-part statements connect by balance, then continues with normal mailbox fetching. Reports ending at or before the checkpoint are skipped, while reports crossing it are rejected. Wise initial XML statements can still be parsed from this directory when Wise has no imported history.

Expected files:

- `IBKR_INITIAL_*.csv` - IBKR initial CSV reports.
- `statement_*.xml` - Wise initial XML statements.

## Implemented Account Sources

- **ICBC / BOC:** credit-card email statements and supported debit-card SMS. ICBC historical-detail imports remain available on demand; their scheduled task is temporarily disabled.
- **IBKR:** daily CSV email reports and local initial reports, including transactions, holdings, dividends, interest, fees and valuation changes.
- **iFAST:** transaction emails plus local monthly statements for GBP, USD, EUR, HKD, SGD and RMB. Missing monthly interest uses the latest published rate and a provisional daily-balance calculation; later statements must agree exactly.
- **ZA:** transaction emails on demand, not scheduled. Email notices do not provide a complete ledger or verified ending balance.
- **Wise:** local initial XML statements followed by daily Plaid transaction updates for all linked currencies. Statement-email downloads remain disabled.
- **FirstTrade:** read-only account balances, positions and transaction history through its web API.
- **Schwab:** direct Plaid investment imports. Former email and Google Drive importers are disabled.
- **Kraken / Ethereum:** completed-day transactions and asset valuations. Crypto quantities use `decimal(30,18)`; unsupported precision fails. Matching internal transfers requires the same chain event and opposite asset quantities.
- **Nexus:** monthly donation-point income through GraphQL.
- **Google Drive:** read-only report transport restricted to the shared `Reports` folder. No active Schwab importer uses it.

Imports require existing accounts and fixed starting checkpoints. They do not create accounts automatically.

## Configuration

### Plaid Imports

Configure `plaid_client_id` and `plaid_production_secret`, then authorize from the build-output directory:

```powershell
dotnet MyBook.dll --plaid-link --country US --product investments
```

Production is the default; Sandbox requires changing the compile-time environment switch. Each import must include every configured Schwab account or the entire batch fails. Imported history starts at the fixed checkpoint, then advances with a seven-day overlap. Transactions, holdings and values must reconcile exactly.

Wise uses an existing Transactions-enabled Item and maps its currency accounts to the single local Wise account. Initial XML statements establish the opening history; subsequent imports use an incremental cursor and validate every currency balance exactly. Pending transactions are not booked. Revised or removed posted transactions fail for manual review, without advancing the cursor. Transfers with an unknown fee split are recorded at their actual gross amount, marked pending split, and are not classified as internal transfers. Spending categories alone do not establish purchases or refunds, so those payments and receipts also remain pending split. Later XML enrichment is not implemented.

Local Wise XML parsing lives under `File/`; the former mail module is retained with a `.deprecated` suffix and is not compiled.

Plaid Items and their unencrypted access tokens are private fixed data. Changing them requires explicit approval; imports never repair authorization automatically. Quote dates are separate from the query date and do not indicate when the entire account last changed.

Plaid API requests reject HTTP redirects instead of forwarding credentials to another address.

Only if saving a newly authorized Item to the database fails, its token is written in plaintext to `plaid-token-recovery-*.local.json` beside the application, and the error reports the path. Use it for manual recovery, then delete it; never share or commit it. Successful saves create no recovery file.

### Deprecated Schwab Relay

The uncompiled `*.Schwab.cs.deprecated` files retain earlier implementations for reference. Their source was a ChatGPT financial plugin connected to Schwab through Plaid, with a scheduled task forwarding raw data through email and later Google Drive. They are not scheduled or used by the direct Plaid importer.

### FirstTrade Read-Only API Import

Configure `firsttrade_username`, `firsttrade_password` and `firsttrade_totp_secret` (the original Base32 authenticator key, not a six-digit code). The implementation references `MaxxRK/firstrade-api`, supports only login and read-only queries, and uses `mail_proxy` when configured.

Sessions and account captures are stored using Windows CurrentUser DPAPI under `%LOCALAPPDATA%/MyBook/FirstTrade/`. Keep these private and do not routinely delete session files: they also preserve login cooldowns. Credentials in local configuration are not encrypted. Unsupported authentication challenges fail rather than repeatedly logging in.

Imported values use cash plus individual positions. Only FirstTrade's equity-subtotal and account-total checks allow a difference strictly below USD 1; no adjustment Record is generated. Other checks remain exact.

Create a local configuration file from the example:

```powershell
Copy-Item MyBook\config.json.example MyBook\config.json
```

Fill only the values needed for the integrations you use. Do not commit `MyBook/config.json`.

Notable configuration keys:

- `database_connection` - MySQL connection string. If empty, the app falls back to the built-in local default.
- `yahoo_user` / `yahoo_pass` - mailbox credentials for statement mail imports.
- `gmail_user` / `gmail_app_pwd` - Gmail credentials used by supported mail fetches.
- `mail_proxy` - optional IMAP proxy for all mailbox fetches, for example `http://127.0.0.1:1196` or `socks5://127.0.0.1:1195`. Leave empty for direct connections.
- `pubweb_proxy` - optional HTTP proxy for public web market-data fetches, for example `http://127.0.0.1:8000`. Leave empty for direct connections; system proxy settings are not used by these fetches.
- `alphavantage_key` - exchange-rate or finance data key.
- `ib_gateway_port` - Interactive Brokers gateway port.
- `nexus_api_key` - legacy/personal Nexus API key fallback.
- `kraken_api_key` / `kraken_api_secret` - Kraken read-only API credentials for authenticated account queries.
- `nexus_oauth_client_id` - Nexus OAuth PKCE token refresh client id. `nexus_oauth_client_secret` is retained for local compatibility but is not sent by the PKCE refresh flow.
- `plaid_client_id` / `plaid_sandbox_secret` / `plaid_production_secret` - Plaid credentials used for account authorization and financial imports. The compile-time environment switch uses the Production secret by default; secrets must match the selected environment.
- `etherscan_api_key` - Etherscan API key for read-only Ethereum mainnet address balance and transaction queries.
- `sim_imsi` - expected IMSI for the local USB SIM modem. Leave empty to disable scheduled SMS polling.
- `sim_poll_interval_minutes` - optional SMS polling interval. Values less than 1 use the built-in default of 5 minutes.
- `GoogleCloudServeAccountKey` - Google service-account JSON key for read-only Drive reports. Share only the `Reports` folder with its `client_email` as Viewer; no Google Cloud/IAM roles are needed.

When adding, removing, or renaming configuration keys, update `MyBook/config.json.example` at the same time and keep all example values blank or zero.

## Database

The application validates its MySQL schema on startup. Accounts, registered account identifiers, Plaid Items, import checkpoints and Start snapshots are fixed data preserved by cleanup. Imported records, holdings, other snapshots and OAuth tokens are runtime data.

Rebuild an empty database from the tracked schema plus the local fixed-data file:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --rebuild-database-from-bootstrap-sql
```

Export the current schema to `Database/bootstrap.sql` and fixed data to the ignored `Database/bootstrap.fixed-data.sql`:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --export-bootstrap-sql
```

`Database/bootstrap.fixed-data.sql` contains private account metadata and unencrypted Plaid access tokens, so it is intentionally not tracked. Treat it and its backup copies as credential files: keep them private, never print their contents or commit them, and use protected storage for external backups.
Backup versions are kept as ignored `Database/bootstrap-*.schema.sql` and `Database/bootstrap-*.fixed-data.sql` file pairs.

Create a start snapshot:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --create-start-snapshot
```

## Build

```powershell
dotnet build MyBook\MyBook.csproj -v minimal /p:UseSharedCompilation=false
```

## Fetch Behavior

Release builds run an import cycle on startup and daily afterward; Debug builds do not schedule fetches. Configured API modules and daily reports run each cycle, while ICBC/BOC bills and Nexus monthly reports are checked when due. ICBC historical-detail scheduling is temporarily disabled.

SMS polling has its own configured interval. It verifies the SIM IMSI, combines complete long messages, and imports supported bank notifications. Unsupported bank formats fail visibly. Mail imports share IMAP sessions and download matching attachments.

Each cycle refreshes exchange rates and creates a snapshot. The UI shows the active task, last run and next run. Failures create `%TEMP%\MyBook.import-failed.tmp`; the clear-marker button removes the warning without retrying. Successful imports do not clear it automatically.

## Nexus OAuth

Nexus API requests include the application headers required by the Nexus API Acceptable Use Policy:

- `Application-Name: MyBook`
- `Application-Version: <assembly version>`

OAuth token storage uses the local database table `OAuthTokens`. Tokens are not stored in `config.json`.
Nexus OAuth uses the PKCE public-client flow; token refresh sends `client_id` and `refresh_token` without `client_secret`.
Nexus imports currently use `nexus_api_key`; stored OAuth tokens are not used by scheduled imports while OAuth is disabled.

Authorize or refresh the local Nexus OAuth token:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --debug-authorize-nexus-oauth
```

This opens the Nexus authorization page in the browser and listens for the callback on `http://127.0.0.1:4700/callback`. The command does not print the authorization URL or OAuth tokens.

## TODO Modules

The following modules are intentionally present as placeholders or not-yet-complete integrations:

- `FileUtil.WeChat.TODO.cs`
- `MailUtil.Steam.TODO.cs`
- `WebUtil.Bilibili.TODO.cs`
- `WebUtil.Meituan.TODO.cs`

These modules should fail loudly or remain unconnected until implemented and validated.

## NotUsed Modules

The following modules are kept in the codebase but are not part of the current active import workflow:

- `MailUtil.PayPal.NotUsed.cs`

## Accuracy Notes

- Records must be linked to a `StatementImport`.
- External imports should update records and holdings/balances as one atomic operation.
- Account balances are derived from holdings through the `AccountBalances` view.
- Snapshots represent database state at an import progress point, not natural-date account state.
- `Records.expenseAllocationDays` controls allocated-expense periods. When `expenseAllocationSkipDays` is empty, the original day-count behavior is used. When both values are present, `expenseAllocationSkipDays` may be 0 or must have the same sign as `expenseAllocationDays`, and `abs(expenseAllocationSkipDays) < abs(expenseAllocationDays)`; the two values then define the relative allocation boundaries around the record date.
