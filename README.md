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
- **FirstTrade:** balances, holdings, transaction history and security information through the account API, plus CSV/OFX downloads using the same login session and kept in memory.
- **Wise:** multi-currency balances, activities, transfer details and payment receipts through a read-only personal-token API.
- **Schwab:** investment account data through Plaid.
- **PayPal:** transaction information from each account's linked Plaid connection and mailbox.
- **Stock and crypto quotes:** Google Finance for US stocks and ETFs, the existing public quote source for Shanghai stocks, and Kraken public Ticker for cryptocurrencies.
- **Kraken:** balances and ledger entries through a read-only API; prices through public market-data endpoints.
- **Ethereum:** transactions and balances for configured addresses through blockchain query endpoints; prices through public market-data endpoints.
- **Nexus:** monthly Donation Points income through GraphQL.
- **Bilibili:** shell wallet balance through the website API. Run `dotnet MyBook.dll --bilibili-login` from the build-output directory and scan the console QR code with the Bilibili app. This mode only authorizes and saves the program's independent session; it does not start the dashboard or background tasks. The session is renewed automatically when importing.
- **Steam:** account/wallet information through SteamKit2 and wallet transactions from authenticated purchase-history pages, both using `steam_proxy`. Run `dotnet MyBook.dll --steam-login` from the build-output directory to enter credentials and complete Steam Guard; `dotnet MyBook.dll --steam-login --saved` uses the stored session.
- **Exchange rates:** daily historical quotes against CNY from Google Finance and Kylc bank quote pages (CCB, ICBC, Industrial Bank and Hengfeng Bank; USD, HKD, GBP and EUR). Currency valuations use the latest available Google historical rates.

Configure the corresponding accounts and fixed import starting points before importing.

## Configuration

Create a local configuration file from the example and fill in the integrations you use:

```powershell
Copy-Item MyBook\config.json.example MyBook\config.json
```

`config.json` contains private credentials and is excluded from Git. Main settings:

- `database_connection`: MySQL connection string; an empty value uses the built-in local default.
- `yahoo_user` / `yahoo_pass`, `gmail_user` / `gmail_app_pwd`: statement-mail credentials.
- `mail_proxy`: optional proxy for PayPal mail and FirstTrade. Other modules fetch mail and attachments directly, ignoring system proxies. `pubweb_google_proxy`: optional HTTP proxy for Google Finance stock prices and exchange rates; `pubweb_proxy`: optional proxy for other public web requests. Empty proxy settings use direct connections.
- `alphavantage_key`: market-data key; `ib_gateway_port`: Interactive Brokers gateway port.
- `nexus_api_key`: Nexus personal API key used by current imports.
- `kraken_api_key` / `kraken_api_secret`, `etherscan_api_key`: credentials for read-only Kraken and Ethereum queries.
- `sim_imsi`: expected SIM IMSI; leave empty to disable polling. `sim_poll_interval_minutes` defaults to 5 when unset or below 1.

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

## Database

The application validates its MySQL schema on startup. Accounts, registered account identifiers, Plaid connections, fixed import starting points, initial rates for each source and currency, and start snapshots are fixed data preserved during cleanup. Imported records, holdings and other snapshots are runtime data.

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

Historical rate timestamps use local time, separately from retrieval time. Kylc supplies dates only; its dates use Beijing midnight converted to local time as date markers, not actual publication times.

Release builds check for daily imports immediately on startup and every 24 hours afterward, fetching exchange rates first. Automatic cycles are attempted at most once per local calendar day, including across restarts; a failed or interrupted cycle waits until a later day unless retried manually. Debug builds do not schedule imports. Each cycle runs configured, enabled integrations whose query intervals have elapsed. Intervals must be positive and become due on the specified day. Failed queries do not restart the interval. A zero missing-report deadline disables only overdue errors, not network, parsing or financial validation errors.

Both Debug and Release builds fetch current stock and cryptocurrency holding prices on startup and every 15 minutes. Live quotes stay in memory and are fetched again after a restart. Quotes may be delayed by their source or reflect the last trading session when markets are closed.

SMS polling uses its own configured interval. Mail imports share IMAP sessions and download matching attachments.

Each cycle also creates a snapshot. The UI shows the active task, last run and next run. Failures create `MyBook.import-failed.tmp` in the application directory; the clear-marker button removes the warning without retrying. Successful imports do not clear it automatically.
