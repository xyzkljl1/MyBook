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

Local statements, downloaded reports, `config.json`, database backups, and other private/runtime files are intentionally ignored.

## Local Initial Reports

`initialReports/` is an ignored private directory for account history that predates normal recurring imports. When an IBKR account has no imported history, the importer first reads matching local files from this directory whose periods start after the provider checkpoint, validates that the initial statement starts from zero and that multi-part statements connect by balance, then continues with normal mailbox fetching. Reports ending at or before the checkpoint are skipped, while reports crossing it are rejected.

Expected files:

- `IBKR_INITIAL_*.csv` - IBKR initial CSV reports.

## Implemented Account Sources

- **ICBC / BOC:** credit-card email statements and supported debit-card SMS. ICBC historical-detail imports are available on demand and scheduled when due. Historical-detail email searches begin after the fixed starting checkpoint, or at the requested search date if later; earlier transactions inside those emails retain their existing treatment. Debit-card history supports only the registered demand-deposit account; deposit certificates and investment accounts under the same card are not supported. These accounts have separate balances, and closing a deposit can transfer funds into the demand-deposit account. Every transaction must belong to the registered demand-deposit account, otherwise the statement is rejected. Overlapping debit-card statements are checked against existing bank transactions throughout the covered period, with separate validation of SMS adjustments and opening balances to prevent duplicate entries. Ignored statements and unreadable attachments are marked as processed and skipped on later imports. This includes statements entirely before the opening balance or only partly covering an unresolved SMS balance adjustment; other statements continue processing.
- **IBKR:** daily CSV email reports and local initial reports, including transactions, holdings, dividends, interest, fees and valuation changes.
- **iFAST:** one account with separate GBP, USD, EUR, HKD, SGD and RMB cash holdings, imported from transaction emails and local monthly statements. Transfers with verified own-account counterparties are treated as internal. Interest-rate update emails provide effective dates; official Gross/AER observations and notices are retained in `StatementImports.sourceDataJson`. New interest calculations use the Gross rate applicable to each day, starting with the first saved observation; existing interest is not recalculated. Missing or conflicting rates fail rather than using AER or backfilling today's rate. Months with nonzero balances before the baseline require actual statements. Monthly rounding and exact statement validation remain unchanged.
- **ZA:** transaction emails on demand, not scheduled. Email notices do not provide a complete ledger or verified ending balance.
- **FirstTrade:** read-only account balances, positions and transaction history through its web API.
- **Wise:** daily read-only personal-token API imports, including multi-currency balances, activities, transfer details and payment receipts. Existing account identifiers are used to resolve counterparties; unavailable fee splits remain explicitly pending rather than estimated.
- **Schwab:** direct Plaid investment imports.
- **Kraken / Ethereum:** completed-day transactions and asset valuations. Kraken supports BTC, ETH, USDT and BABY, including BABY staking rewards valued using the daily BABY/USD close. Crypto quantities use `decimal(30,18)`; unsupported precision fails. Matching internal transfers requires the same chain event and opposite asset quantities.
- **Nexus:** monthly donation-point income through GraphQL.
- **Google Drive:** read-only report transport restricted to the shared `Reports` folder.

Imports require existing accounts and fixed starting checkpoints. They do not create accounts automatically.

For new transfers involving Wise, Schwab, FirstTrade, Kraken, Nexus, ZA or CICC, an explicit counterparty institution can resolve to its sole configured account when no account identifier matches. Conflicting or ambiguous evidence fails; fees and unsplit amounts retain their existing treatment. Existing records are not reclassified automatically.

## Configuration

Create a local configuration file from the example and fill the values for the integrations you use:

```powershell
Copy-Item MyBook\config.json.example MyBook\config.json
```

`config.json` contains private credentials and is excluded from Git. Main settings:

- `database_connection` - MySQL connection string. If empty, the app falls back to the built-in local default.
- `yahoo_user` / `yahoo_pass`, `gmail_user` / `gmail_app_pwd` - statement-mail credentials.
- `mail_proxy` - optional mailbox and FirstTrade proxy; `pubweb_proxy` - optional public market-data proxy. Leave empty for direct connections.
- `alphavantage_key` - market-data key; `ib_gateway_port` - Interactive Brokers gateway port.
- `nexus_api_key` - Nexus personal API key used by current imports.
- `kraken_api_key` / `kraken_api_secret`, `etherscan_api_key` - read-only Kraken and Ethereum queries.
- `sim_imsi` - expected SIM IMSI; leave empty to disable polling. `sim_poll_interval_minutes` defaults to 5 when unset or less than 1.
- `GoogleCloudServeAccountKey` - Google service-account JSON key. Share the `Reports` folder with its `client_email` as Viewer; no Google Cloud/IAM roles are needed.

### Plaid / Schwab

Set `plaid_client_id` and `plaid_production_secret`, then authorize from the build-output directory:

```powershell
dotnet MyBook.dll --plaid-link --country US --product investments
```

Production is the default. Sandbox requires changing the compile-time environment switch and using `plaid_sandbox_secret`. Authorizations are stored in the local database. If saving a new authorization fails, a plaintext `plaid-token-recovery-*.local.json` file is created beside the application for manual recovery; delete it after use.

Before importing Schwab data, bind the Plaid connection to the corresponding local account. Each connection supports only one investment account. Missing or invalid bindings stop the import.

### Wise

Set `wise_api_token` to a personal API token. The API importer replaces Plaid Wise; the old importer is no longer compiled. The retained Plaid importer does not use Plaid's inferred personal finance categories; descriptions remain generic payments or receipts pending further detail. Switching existing Wise data requires cleanup before the first API import. Some details, including conversion fees, are unavailable through the API and remain marked as pending.

### PayPal

PayPal combines each account's linked Plaid connection with its configured mailbox. Daily imports run after Wise and Nexus. Confirmed card-funded purchases are ignored; receipts, transfers, refunds and separately reported fees are reconciled without duplicating the same transaction from both sources. Identified withdrawals can be linked to existing bank records.

Imports retain source evidence and stop without changing financial records when history is incomplete, a match is ambiguous, or the detailed ledger disagrees with the reported balance. They do not infer missing funds or change the configured import starting point.

### FirstTrade

Set `firsttrade_username`, `firsttrade_password` and `firsttrade_totp_secret` (the original Base32 authenticator key, not a six-digit code). The read-only integration references `MaxxRK/firstrade-api` and uses `mail_proxy` when configured. Sessions are stored in the database; raw financial responses are saved only with successful imports. Login failures and HTTP 403/429 are reported immediately without retries, cooldowns or saved request pauses. Only a data-request HTTP 401 triggers one automatic session renewal and request retry per import. These sensitive database contents are not DPAPI-encrypted. No session or response files are written, and old file-based sessions are not loaded.

Firstrade uses different quote providers for balances and positions, so their valuations may differ. The equity subtotal and account total checks allow an absolute difference below USD 100; differences of USD 100 or more fail. Holdings and records use detail values without residual adjustments; all other exact validations remain unchanged.

## Database

The application validates its MySQL schema on startup. Accounts, registered account identifiers, Plaid Items, import checkpoints and Start snapshots are fixed data preserved by cleanup. Imported records, holdings, other snapshots and OAuth tokens are runtime data.

Rebuild an empty database from the tracked schema plus the local fixed-data file:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --rebuild-database-from-bootstrap-sql
```

`Database/bootstrap.fixed-data.sql` contains private account metadata and unencrypted Plaid access tokens. It and its backups are excluded from Git and require secure storage.
Automatic backup and manual export are local-debug extensions, not included in the repository. When present, they save schema/fixed-data files and versioned backup pairs in `Database` under the application directory; normal startup still runs automatic backup.

Create a start snapshot:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --create-start-snapshot
```

## Build

```powershell
dotnet build MyBook\MyBook.csproj -v minimal /p:UseSharedCompilation=false
```

## Fetch Behavior

Release builds run an import cycle on startup and daily afterward; Debug builds do not schedule fetches. Configured API modules and daily reports run each cycle, while ICBC/BOC bills and Nexus monthly reports are checked when due. ICBC historical details are fetched once more than 90 days have elapsed since the latest import, searching emails from the past five months subject to the fixed starting checkpoint.

SMS polling has its own configured interval. It verifies the SIM IMSI, combines complete long messages, and imports supported bank notifications. Unsupported bank formats fail visibly. Mail imports share IMAP sessions and download matching attachments.

Each cycle refreshes exchange rates and creates a snapshot. The UI shows the active task, last run and next run. Failures create `MyBook.import-failed.tmp` in the application directory; the clear-marker button removes the warning without retrying. Successful imports do not clear it automatically.

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
