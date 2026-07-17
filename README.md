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

`initialReports/` is an ignored private directory for account history that predates normal recurring imports. When an IBKR account has no imported history, the importer first reads matching local files from this directory, validates that the initial statement starts from zero and that multi-part statements connect by balance, then continues with normal mailbox fetching. Wise initial XML statements can still be parsed from this directory when Wise has no imported history.

Expected files:

- `IBKR_INITIAL_*.csv` - IBKR initial CSV reports.
- `statement_*.xml` - Wise initial XML statements.

## Implemented Account Sources

- `MailUtil.ICBC` fetches ICBC credit-card statements monthly from mailbox messages and imports card balances plus RMB and foreign-currency transaction details.
- `MailUtil.ICBCHistory` fetches ICBC historical-detail PDF mail attachments on demand, by date range, or by low-frequency scheduled scans, and imports debit-account history details and credit-card history supplements when the statement balance chain and overlap checks pass. Older debit history can be backfilled before newer existing records by validating the statement ending balance against the current balance rolled back through later records, without moving the current cash balance backward. For debit history that overlaps previously imported ICBC SIM SMS records, exact transaction matches are confirmed by source supplements instead of duplicated, more specific history-detail fields can fill the existing SMS record, and only explicitly marked small-transaction compensation records may be reversed and replaced by history-detail rows. If a SIM import created the account's initial cash-balance record, history-detail rows before that initialization point are used only to validate the initialization balance and are not imported or reversed.
- `MailUtil.IBKR` fetches Interactive Brokers `DailyMyBook` CSV reports daily from mailbox attachments and imports cash, NAV, positions, trades, commissions, transfers, cash/bond interest, accrued cash/bond interest, dividends, withholding tax, FX translation, and end-of-day holdings. It can also read local initial reports before daily reports exist.
- `KrakenUtil` imports every missing completed UTC day as a daily report. It pages through all balance-affecting Ledger entries, reconstructs and validates native asset quantities, and stores transactions, fees, holdings, and account balances directly in BTC, ETH, USDT, or USD. The current `BalanceEx` quantities must exactly equal the reconstructed Ledger chain before the missing daily reports are committed atomically.
- `MailUtil.Wise` imports local initial Wise XML statements from `initialReports` when the Wise account has no history, and keeps the XML parser for per-currency balances plus fees, conversions, card payments, direct debits, and sent/received transfers. Recurring mailbox downloads for monthly Wise statements are disabled because Wise no longer sends the statement files as mail attachments.
- `MailUtil.OCBC` fetches OCBC statement emails/PDFs monthly from the mailbox and imports configured OCBC account balances and transaction lines; if an old month is missing, it can also import a self-sent supplemental statement mail with the original subject and PDF attachment.
- `MailUtil.Steam.TODO` will fetch Steam account mail statements for Steam account transactions.
- `GraphQLUtil.Nexus` fetches Nexus Mods donation-point monthly summaries through the Nexus GraphQL API and imports monthly DP income for the configured Nexus account.
- `PlaidUtil.TODO` will evaluate Plaid-based account imports; it currently only verifies Plaid Sandbox credentials and searches supported institutions without creating Plaid Items or importing account data.
- `CryptoUtil` is the shared cryptocurrency entry point. Its `CryptoUtil.ETH` submodule provides read-only Ethereum mainnet queries through Etherscan V2 and imports every missing completed UTC day for `ETH_0x...` accounts. It validates exact ETH and official Ethereum USDT quantities against JSON-RPC balances, stores native-currency holdings and transaction/gas records, and ignores look-alike token symbols from other contracts.
- Crypto assets use the existing `Currency`, `Holding`, and `AccountBalance` models; `Finance` is reserved for their eventual valuation, but crypto prices are not fetched or updated yet. No price movement `Record` or provider-specific asset snapshot table is used. Ethereum and supported provider records retain chain, transaction hash, event index, and token contract. The internal-transfer matcher links records only when the same transaction contains one unique opposite amount in the same crypto currency; provider deposit and withdrawal addresses remain evidence for that event and are not stored as account addresses.
- `FileUtil.WeChat.TODO` will parse local WeChat bill files for WeChat account transactions.
- `SIMUtil` polls a USB SIM modem through a Windows COM port, verifies the SIM IMSI against local configuration, combines complete long SMS messages before dispatching by sender, and can import supported ICBC debit-card transaction SMS messages with balance validation. The first trusted ICBC SMS balance reuses the shared initial cash-balance record mechanism when no current cash balance exists. If an ICBC SMS balance proves that earlier unreported activity changed an already initialized debit-account balance, it writes an `ICBCSIMCompensation` record in the same atomic import so later history-detail rows can reverse and replace it.
- `WebUtil.Bilibili.TODO` will fetch Bilibili account balance information.
- `WebUtil.Meituan.TODO` will fetch Meituan account balance information.

## Configuration

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
- `ocbc_statement_passwords` - local list of OCBC statement passwords.
- `nexus_api_key` - legacy/personal Nexus API key fallback.
- `kraken_api_key` / `kraken_api_secret` - Kraken read-only API credentials for authenticated account queries.
- `nexus_oauth_client_id` - Nexus OAuth PKCE token refresh client id. `nexus_oauth_client_secret` is retained for local compatibility but is not sent by the PKCE refresh flow.
- `plaid_client_id` / `plaid_sandbox_secret` - Plaid Sandbox credentials for evaluating Plaid-based account imports.
- `etherscan_api_key` - Etherscan API key for read-only Ethereum mainnet address balance and transaction queries.
- `sim_imsi` - expected IMSI for the local USB SIM modem. Leave empty to disable scheduled SMS polling.
- `sim_poll_interval_minutes` - optional SMS polling interval. Values less than 1 use the built-in default of 5 minutes.

When adding, removing, or renaming configuration keys, update `MyBook/config.json.example` at the same time and keep all example values blank or zero.

## Database

The application validates the database schema on startup. Fixed data includes `Accounts`, fake/checkpoint `StatementImports`, and `Start` snapshots with their `SnapshotItems`; imported records, holdings, non-start snapshots, and OAuth tokens are runtime data.

Rebuild an empty database from the tracked schema plus the local fixed-data file:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --rebuild-database-from-bootstrap-sql
```

Export the current schema to `Database/bootstrap.sql` and fixed data to the ignored `Database/bootstrap.fixed-data.sql`:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --export-bootstrap-sql
```

`Database/bootstrap.fixed-data.sql` may contain private account metadata, so it is intentionally not tracked.
Backup versions are kept as ignored `Database/bootstrap-*.schema.sql` and `Database/bootstrap-*.fixed-data.sql` file pairs.

Create a start snapshot:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --create-start-snapshot
```

## Build

```powershell
dotnet build MyBook\MyBook.csproj -v minimal /p:UseSharedCompilation=false
```

The project currently builds with a known MailKit NU1902 advisory warning.

## Fetch Behavior

When the desktop app starts in a non-debug build, `Fetcher.RunSchedule()` starts scheduled background fetches. It runs one fetch cycle immediately, then schedules another cycle once per day.

During scheduled fetches:

- IBKR reports are checked every cycle.
- ICBC monthly bills, OCBC statements, and Nexus DP monthly reports are checked only when the latest import is more than 27 days old.
- ICBC historical-detail attachments are checked only when the latest history-detail import or scheduled empty-import checkpoint is more than 90 days old. Each scheduled scan searches the last 5 months and writes a `scheduled-empty-import-yyyyMMdd` checkpoint even if no statement is imported, so empty scans are not retried every day.
- Mail fetches share Yahoo/Gmail IMAP sessions within each fetch cycle or standalone mail import, use the configured `mail_proxy` when set, and reconnect once after connection-level failures.
- Public web market-data fetches use the configured `pubweb_proxy` when set; exchange-rate pages are requested concurrently and saved after all requests finish.
- Nexus DP imports use the monthly summary API first and fall back to per-month reports only for missing months or summary failures.
- SIM SMS polling runs on its own interval when `sim_imsi` is configured. The first poll must find exactly one responsive modem, later polls reuse the cached port and rescan only if it disappears. Each poll verifies IMSI, dispatches complete received SMS messages by sender, keeps incomplete long-SMS fragments until all parts arrive, currently routes ICBC messages to `SIMUtil.ICBC`, can initialize a debit account from the first trusted SMS balance, and logs every reason that no SMS was fetched, imported, retained, or deleted.
- IBKR and OCBC attachment imports search the missing date/month range in batches, then group downloaded attachments by statement date or month.
- Attachment-based mail imports first filter IMAP summaries and body structures, then download only matching attachment body parts instead of full messages.
- Imported records pass through one shared automatic expense-allocation judgment function after internal-transfer matching and before allocated-expense cache generation.
- Exchange rates are refreshed every cycle.
- A daily snapshot is created after the fetch cycle.
- If any scheduled import step fails, the app writes `%TEMP%\MyBook.import-failed.tmp`; while that file exists, the top summary area shows an import failure. Successful imports do not remove it, so it must be cleared manually after checking the failure.

The top summary area shows whether a scheduled import is currently running, the active step name when available, the last completed daily fetch time, and the next scheduled daily fetch time.

Debug builds do not run scheduled fetch tasks. The app prints `skip scheduled fetch in DEBUG`.

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
