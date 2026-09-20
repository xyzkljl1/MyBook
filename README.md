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

- `MailUtil.ZA` implements on-demand ZA Bank transaction email imports from `notification@service.bank.za.group`; it is not registered with the scheduled fetcher. The parser is based on actual simplified-Chinese receipt, completed transfer, and ZA Card purchase notices. It compares subject/body currency and amount exactly and uses the transaction timestamp in the body (Hong Kong time), rather than the later mail delivery time. Failed payments and unrelated notices create no records or processing markers; unsupported transaction formats fail. The shared Yahoo mailbox is accessed directly, with no new configuration keys. Imports require existing non-credit, relative-balance `ZA_<card-tail>` accounts and a fixed empty-key `ZAMail` checkpoint, neither of which is automatically created by the importer. Purchase notices select the account using their explicit ZA Card tail; transfer notices without a card identifier require exactly one ZA account and otherwise fail as ambiguous. Subsequent searches resume on the latest successful transaction mail's date, with same-day Message-ID deduplication and no indefinite historical rescan. Each email's record, cash balance change, transfer matching, and allocation processing use the existing atomic save path. Notifications do not include ending balances, original purchase amounts in another currency, or a complete interest/refund ledger: balances are inferred from imported transactions starting at zero, not asserted bank balances. No account/card identifiers are hardcoded or inferred from transaction counterparties; no new formal CLI entry point is added.
- `MailUtil.IFast` runs during the daily import task for the six relative-balance accounts `IFAST_GBP`, `IFAST_USD`, `IFAST_EUR`, `IFAST_HKD`, `IFAST_SGD`, and `IFAST_RMB`. GBP and EUR are supported throughout currency storage and exchange-rate lookup.
  - On first import, the consecutive iFAST `Monthly Statement*.pdf` files in `initialReports` define a fixed initialization period starting with zero balances. Only transaction rows create records; opening/closing balances, running balances, and debit/credit totals are checked exactly. Existing email transactions are matched without duplicating them. Keep the initial files until initialization completes. Subsequent files can import missing actual interest, but other transactions remain validation-only. Previously calculated interest must match exactly; a difference is an error, never an automatic correction.
  - Email imports use the shared Yahoo credentials and direct IMAP, starting on the latest successfully imported transaction email's persisted import date. Only mail/receipt/conversion statement keys determine this position; PDF and calculated-interest dates are excluded. The fixed `IFastMail` checkpoint is used only before any successful mail import. The shared database mechanism normalizes import times to dates, so the last imported day remains included and its successful mails are deduplicated. IMAP uses a small timezone margin for header-date searching; summaries older than the local start date are excluded. A newly arriving mail dated before that day requires a manual rescan. Imports process receipts, completed currency conversions (both legs atomically), and successful QR payments with the original purchase currency. Receipts/conversions use bank references for deduplication; QR payments use Message-ID. Emails inside the initialization period must match existing records. Non-transaction messages and statement-ready notices create no records or mail progress; monthly PDFs must be downloaded separately. Currency conversions use the explicit BST/GMT timestamp, receipts use the bank payment date, and QR payments use mail time.
  - When interest is not yet imported and its posting-month statement is unavailable, `PubWebUtil.IFast` reads the official website's public current-account Gross-rate JSON API using the existing `pubweb_proxy`. The latest rate for each currency is applied to the entire earning month, including historical months; no historical rate cache is required. Records retain the rate, retrieval timestamp, source URL, and an input hash. Already imported interest is not rewritten when rates change. A statement is preferred when available and does not require a rate lookup; for example, August interest is paid on September 1 and appears in the September statement.
  - After initialization, completed months generate `Reason=利息` records using the user-approved provisional convention: UK daily closing balance times the latest Gross annual rate divided by 365, no daily cent rounding, monthly rounding to two decimals away from zero, no fractional-cent carry, posting on the next month's first day. This convention is not a verified bank rounding specification. A new monthly statement must match every transaction and each currency's opening/closing balances exactly, including any existing calculated interest; differences fail import without residual or correction records.
  - All iFAST sources use `IFastMail`, separated by statement keys. `StatementImport.time` is the email time for mail, the UK period end for monthly PDFs, or the UK posting date for calculated monthly interest. Each source is written atomically with its balance changes; calculated interest is provisional until checked against a statement.
- `MailUtil.ICBC` fetches ICBC credit-card statements monthly from mailbox messages and imports card balances plus RMB and foreign-currency transaction details.
- ICBC SMS transactions for the same card and currency are ordered by an exact, unique balance chain within windows spanning at most two minutes. Windows are anchored to the earliest transaction and never expand transitively. Original transaction/receipt timestamps are preserved; timestamp inversions are logged, and broken or ambiguous chains still fail import.
- `MailUtil.BOC` fetches official Bank of China credit-card statement PDF attachments monthly and imports the account balances and transaction details after validating the statement totals. Wrapped transaction descriptions are collected within the PDF table row boundaries; missing descriptions or ambiguous row boundaries reject the statement without relaxing amount or balance validation.
- `MailUtil.ICBCHistory` fetches ICBC historical-detail PDF mail attachments on demand, by date range, or by low-frequency scheduled scans, and imports debit-account history details and credit-card history supplements when the statement balance chain and overlap checks pass. Older debit history can be backfilled before newer existing records by validating the statement ending balance against the current balance rolled back through later records, without moving the current cash balance backward. For debit history that overlaps previously imported ICBC SIM SMS records, exact transaction matches are confirmed by source supplements instead of duplicated, more specific history-detail fields can fill the existing SMS record, and only explicitly marked small-transaction compensation records may be reversed and replaced by history-detail rows. If a SIM import created the account's initial cash-balance record, history-detail rows before that initialization point are used only to validate the initialization balance and are not imported or reversed.
- `MailUtil.IBKR` fetches Interactive Brokers `DailyMyBook` CSV reports daily from mailbox attachments and imports cash, NAV, positions, trades, commissions, fees, transfers, cash/bond interest, accrued cash/bond interest, dividends, withholding tax, FX translation, and end-of-day holdings. It can also read local initial reports before daily reports exist.
- IBKR ACATS position transfers retain the reported signed quantity and market value as internal-transfer records, with the existing NAV and transfer-total checks unchanged. Counterparties are resolved by the full normalized account identifier across all existing accounts and registered identifiers, without a broker restriction or automatic account creation. An unknown counterparty remains unmatched and does not block import; ambiguous identifiers fail. ACATS records are paired only when both counterparties are known and reciprocal, with equal opposite amounts and quantities for the same security. No receiving-side record is synthesized.
- `KrakenUtil` imports every missing completed UTC day as a daily report. It pages through all balance-affecting Ledger entries, reconstructs and validates native asset quantities, stores USD cash plus BTC/ETH/USDT asset quantities, and values each completed day at the shared Kraken UTC daily close. `KrakenPubUtil` provides the credential-free Kraken OHLC prices shared by Kraken and Ethereum imports. The current `BalanceEx` quantities must exactly equal the reconstructed Ledger chain before the missing daily reports are committed atomically.
- `MailUtil.Wise` imports local initial Wise XML statements from `initialReports` when the Wise account has no history, and keeps the XML parser for per-currency balances plus fees, conversions, card payments, direct debits, and sent/received transfers. Recurring mailbox downloads for monthly Wise statements are disabled because Wise no longer sends the statement files as mail attachments.
- `MailUtil.Steam.TODO` will fetch Steam account mail statements for Steam account transactions.
- `GraphQLUtil.Nexus` fetches Nexus Mods donation-point monthly summaries through the Nexus GraphQL API and imports monthly DP income for the configured Nexus account.
- `PlaidUtil.TODO` will evaluate Plaid-based account imports; it currently only verifies Plaid Sandbox credentials and searches supported institutions without creating Plaid Items or importing account data.
- `CryptoUtil` is the shared cryptocurrency entry point. Its `CryptoUtil.ETH` submodule provides read-only Ethereum mainnet queries through Etherscan V2 and imports every missing completed UTC day for `ETH_0x...` accounts. It validates exact ETH and official Ethereum USDT quantities against JSON-RPC balances, stores asset quantities plus USD-valued transaction/gas records, and ignores look-alike token symbols from other contracts.
- Crypto quantities use `Holding.quantity` and `Record.HoldingQuantity` as `decimal(30,18)`; values outside that exact database range are rejected. `Currency` and `AccountBalance` remain fiat-only, while `Finance` caches the latest shared Kraken USD close. Each completed day produces explicit holding-price-change and transaction-price-impact records so the previous holdings plus records exactly equal the new USD account value. Ethereum and supported provider records retain chain, transaction hash, event index, and token contract. The internal-transfer matcher also requires an exact opposite asset quantity for the same chain event; provider deposit and withdrawal addresses remain evidence for that event and are not stored as account addresses.
- `FileUtil.WeChat.TODO` will parse local WeChat bill files for WeChat account transactions.
- `SIMUtil` polls a USB SIM modem through a Windows COM port, verifies the SIM IMSI against local configuration, combines complete long SMS messages before dispatching by sender, and can import supported ICBC and BOC debit-card transaction SMS messages with balance validation. The first trusted bank SMS balance reuses the shared initial cash-balance record mechanism when no current cash balance exists. ICBC can add an `ICBCSIMCompensation` record for known missing small transactions; BOC requires every later SMS beginning balance to match the current balance exactly and rejects differences. ICBC debit-card repayment SMS records are retained as the cash side of an internal transfer and can match the corresponding credit-card statement record. Recognized non-transaction credit-limit notifications are ignored and deleted. A message from a supported bank sender with an unsupported format fails the SIM import and writes the import-failure marker instead of being silently retained.
- `WebUtil.Bilibili.TODO` will fetch Bilibili account balance information.
- `WebUtil.Meituan.TODO` will fetch Meituan account balance information.

## Configuration

### FirstTrade Read-Only API Import

`WebUtil.FirstTrade` implements its own HTTP client based on the login and read-only protocol in `MaxxRK/firstrade-api`. It does not depend on Python. The only POST endpoints are password login and completing authenticator MFA. Account list, balances, positions and account history use a closed GET allowlist. No trading, cancellation, transfer, profile or watchlist mutation endpoints are implemented. HTTP redirects are disabled. All requests use `mail_proxy`, or direct connections when empty, never the system proxy. Normal TLS certificate validation remains enabled.

Configure `firsttrade_username`, `firsttrade_password` and `firsttrade_totp_secret`. The reference library's shared client token is a code constant, not a configuration value or personal session credential. `firsttrade_totp_secret` is the original Base32 key supplied when binding the authenticator, not a current six-digit code or a backup/recovery code. All example values are blank. Actual credentials, including the TOTP key, remain in the ignored local `config.json`; this configuration file is not encrypted by the module. Protect access to it and do not publish configuration, process dumps or HTTP protocol traces. Keys, generated codes and authentication responses are never logged or included in account captures. Decoded key bytes are cleared when the client is disposed. Email authentication and its configuration fields have been removed; this module no longer accesses the mailbox.

The non-debug daily task imports on startup and daily at 00:05 local time when configured, using the existing import status and failure marker. Debug builds do not run scheduled fetches. Login generates six-digit SHA-1 TOTP codes with a 30-second time step, matching the reference library. Keep the Windows system clock synchronized. Invalid Base32 keys are rejected before any network request; when a code is about to expire, the client waits for the next time step before sending it. It does not open a browser, wait for console input, send email/SMS codes, use recovery codes or change account security settings. HTTP requests have a 30-second timeout and network capture has a five-minute cancellation deadline. Saved session tokens and cookies are reused across imports/restarts without a password login. Only HTTP 401 on a private read invalidates the cached session and permits one retry after a cooldown-eligible login; a new login is forbidden within 15 minutes of the last successful login. Login failure delays grow from 1, 2, 4, 8, 16 to 24 hours and are reserved on disk before the first request, including when the process later crashes. HTTP 403 and 429 never trigger re-login and pause all requests for at least one hour, or longer according to `Retry-After`. Cooldown failures return immediately, not by sleeping or retrying in a loop. A non-waiting file lock also prevents overlapping clients from sharing a login budget. Other challenge types and rejected TOTP codes fail closed. These are local safety limits, not broker-published safe frequencies or a guarantee against risk controls.

Captures preserve original JSON numeric values, fetch all linked accounts, and keep balances plus paged holdings/history in Windows CurrentUser DPAPI-encrypted files under `%LOCALAPPDATA%/MyBook/FirstTrade/`, in a username-hashed subdirectory. No authentication responses or tokens are stored with the captures. A separate encrypted `session.firsttrade.dpapi` stores session tokens, cookies and persistent cooldown timestamps, but not the password or TOTP key. Do not publish or routinely delete that file: deletion resets the local login budget. Corrupt, unreadable or locked session state stops the client before network access instead of silently resetting it. DPAPI does not isolate secrets from other programs running as the same Windows user. Archives are saved before financial import and may include a capture that failed reconciliation. They are not the import checkpoint. History starts at the fixed `FirstTradeApi` database checkpoint; subsequent runs use the last successful import for each account with a seven-day overlap, bounded by that checkpoint. A failed database import never advances the scan cursor. The API's actual historical availability is not a guarantee of lifetime coverage; missing initial transactions cause reconciliation failure rather than an inferred opening balance.

Each explicitly configured initial account is named `FIRSTTRADE_<account>` and uses absolute balances with Investment usage; the importer only looks up accounts and never creates them. The account identifier comes from the account-list response. `FirstTradeApi` imports atomically save transaction Records, USD cash, equity Holdings and account values through the shared database mechanism. Deposits, withdrawals, interest (including securities-lending rebates), dividends and explicit fee/tax entries retain individual records. Equity trades create offsetting cash and asset legs, while paired cash/margin subaccount transfers retain both internal legs. Settlement cents must match rounded quantity times execution price; an unexplained fee difference fails and is never absorbed into valuation. Missing, unsupported or ambiguous product metadata fails: US equity exchange types are resolved from existing Holdings/Finance metadata, not hardcoded ticker lists. Options, bonds, short positions and unrecognized transaction types are rejected until their source-specific parsers are implemented.

Cash from transaction details, security quantities, quantity times current price, equity subtotal and total account value are checked exactly. Price changes are recorded per security from the previous imported value plus trade legs to the newly fetched value, dated on the current US/Eastern capture day. These are observed interval valuation changes, not reconstructed daily closing prices; the initial capture can include appreciation accumulated since purchase. History rows have no API transaction ID, so normalized content hashes plus occurrence counts preserve identical duplicate transactions. Replays do not create duplicate records, and removed or revised transactions within the overlap window fail instead of silently replacing old data. Failures use fixed safe diagnostics without response bodies, account identifiers, credentials, codes or request URLs in logs. The database and encrypted archives contain private financial data and must not be published.

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
- ICBC and BOC monthly bills and Nexus DP monthly reports are checked only when the latest import is more than 27 days old.
- ICBC historical-detail attachments are checked only when the latest history-detail import or scheduled empty-import checkpoint is more than 90 days old. Each scheduled scan searches the last 5 months and writes a `scheduled-empty-import-yyyyMMdd` checkpoint even if no statement is imported, so empty scans are not retried every day.
- Mail fetches share Yahoo/Gmail IMAP sessions within each fetch cycle or standalone mail import, use the configured `mail_proxy` when set, and reconnect once after connection-level failures.
- Public web market-data fetches use the configured `pubweb_proxy` when set; exchange-rate pages are requested concurrently and saved after all requests finish.
- Nexus DP imports use the monthly summary API first and fall back to per-month reports only for missing months or summary failures.
- SIM SMS polling runs on its own interval when `sim_imsi` is configured. The first poll must find exactly one responsive modem, later polls reuse the cached port and rescan only if it disappears. Each poll verifies IMSI, dispatches complete received SMS messages by sender, keeps incomplete long-SMS fragments until all parts arrive, routes supported bank messages to sender-specific `SIMUtil.ICBC` or `SIMUtil.BOC` modules, can initialize a debit account from the first trusted SMS balance, and logs every reason that no SMS was fetched, imported, retained, or deleted.
- IBKR attachment imports search the missing date range in batches, then group downloaded attachments by statement date.
- Attachment-based mail imports first filter IMAP summaries and body structures, then download only matching attachment body parts instead of full messages.
- Imported records pass through one shared automatic expense-allocation judgment function after internal-transfer matching and before allocated-expense cache generation.
- Exchange rates are refreshed every cycle.
- A daily snapshot is created after the fetch cycle.
- If any scheduled import step fails, the app writes `%TEMP%\MyBook.import-failed.tmp`; while that file exists, the top summary area shows an import failure with a clear-marker button. The button only removes the marker and refreshes the status; it does not retry imports. Successful imports do not remove it, so it must be cleared manually after checking the failure.

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
