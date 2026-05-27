# MyBook

MyBook is a local personal finance and asset tracking application. It imports statements from banks, broker reports, mail attachments, local files, and selected web APIs, then stores records, balances, holdings, snapshots, and fixed bootstrap data in MySQL.

## Tech Stack

- .NET 8 WPF desktop application
- MySQL
- SqlSugar ORM
- MailKit for mailbox access
- PdfPig and HtmlAgilityPack for statement parsing
- Newtonsoft.Json for JSON and GraphQL payloads

## Repository Layout

- `MyBook/` - WPF application source
- `Database/bootstrap.sql` - schema plus fixed data used to rebuild an empty database
- `MyBook/config.json.example` - tracked configuration template with blank or zero values
- `*.TODO.cs` modules - placeholders or not-yet-verified integrations

Local statements, downloaded reports, `config.json`, database backups, and other private/runtime files are intentionally ignored.

## Implemented Account Sources

- `MailUtil.ICBC` fetches ICBC credit-card statements monthly from mailbox messages and imports card balances plus RMB and foreign-currency transaction details.
- `MailUtil.ICBCHistory` fetches ICBC historical-detail PDF mail attachments on demand or by date range and imports debit-account history details and credit-card history supplements when the statement balance chain and overlap checks pass.
- `MailUtil.IBKR` fetches Interactive Brokers `DailyMyBook` CSV reports daily from mailbox attachments and imports cash, NAV, positions, trades, commissions, transfers, interest, dividends, withholding tax, FX translation, and end-of-day holdings. It can also read local initial reports before daily reports exist.
- `MailUtil.Wise` fetches Wise XML statements monthly from mailbox attachments and imports per-currency balances plus fees, conversions, card payments, direct debits, and sent/received transfers.
- `MailUtil.OCBC` fetches OCBC statement emails/PDFs monthly from the mailbox and imports configured OCBC account balances and transaction lines.
- `MailUtil.PayPal` fetches PayPal mail statements monthly from configured mailbox messages and imports supported PayPal payment events for configured PayPal accounts.
- `GraphQLUtil.Nexus` fetches Nexus Mods donation-point monthly reports through the Nexus GraphQL API and imports monthly DP income for the configured Nexus account.
- `LocalUtil.WeChat.TODO` will parse local WeChat bill files for WeChat account transactions.
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
- `alphavantage_key` - exchange-rate or finance data key.
- `ib_gateway_port` - Interactive Brokers gateway port.
- `ocbc_statement_passwords` - local list of OCBC statement passwords.
- `nexus_api_key` - legacy/personal Nexus API key fallback.
- `nexus_oauth_client_id`, `nexus_oauth_client_secret`, `nexus_oauth_scope` - Nexus OAuth app settings.

When adding, removing, or renaming configuration keys, update `MyBook/config.json.example` at the same time and keep all example values blank or zero.

## Database

The application validates the database schema on startup. Fixed data includes `Accounts` and fake/checkpoint `StatementImports`; imported records, holdings, snapshots, and OAuth tokens are runtime data.

Rebuild an empty database from tracked schema and fixed data:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --rebuild-database-from-bootstrap-sql
```

Export current schema plus fixed data to `Database/bootstrap.sql`:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --export-bootstrap-sql
```

Clean volatile imported data while preserving fixed data:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --clean-database
```

Create a start snapshot:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --create-start-snapshot
```

## Build

```powershell
dotnet build MyBook\MyBook.csproj -v minimal /p:UseSharedCompilation=false
```

The project currently builds with a known MailKit NU1902 advisory warning.

## Common Import Commands

```powershell
dotnet run --project MyBook\MyBook.csproj -- --fetch-icbc-bills
dotnet run --project MyBook\MyBook.csproj -- --fetch-icbc-history-details
dotnet run --project MyBook\MyBook.csproj -- --fetch-ibkr-reports
dotnet run --project MyBook\MyBook.csproj -- --fetch-wise-reports
dotnet run --project MyBook\MyBook.csproj -- --fetch-ocbc-reports
dotnet run --project MyBook\MyBook.csproj -- --fetch-paypal-reports
```

Useful local/debug commands:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --debug-fetch-local-ibkr-reports
dotnet run --project MyBook\MyBook.csproj -- --debug-fetch-local-wise-reports
dotnet run --project MyBook\MyBook.csproj -- --debug-fetch-local-icbc-history-details
dotnet run --project MyBook\MyBook.csproj -- --debug-sql "SHOW TABLES"
```

Debug builds do not run scheduled fetch tasks.

## Nexus OAuth

Nexus API requests include the application headers required by the Nexus API Acceptable Use Policy:

- `Application-Name: MyBook`
- `Application-Version: <assembly version>`

OAuth token storage uses the local database table `OAuthTokens`. Tokens are not stored in `config.json`.

Start the local OAuth authorization flow:

```powershell
dotnet run --project MyBook\MyBook.csproj -- --authorize-nexus-oauth-token
```

The callback URI is hard-coded and must match the redirect URI registered for the Nexus OAuth app:

```text
http://127.0.0.1:4700/callback
```

The current Nexus OAuth module is marked TODO because it has not been remotely verified with a valid client id.

## TODO Modules

The following modules are intentionally present as placeholders or not-yet-complete integrations:

- `GraphQLUtil.NexusOAuth.TODO.cs`
- `LocalUtil.WeChat.TODO.cs`
- `WebUtil.Bilibili.TODO.cs`
- `WebUtil.Meituan.TODO.cs`

These modules should fail loudly or remain unconnected until implemented and validated.

## Accuracy Notes

- Records must be linked to a `StatementImport`.
- External imports should update records and holdings/balances as one atomic operation.
- Account balances are derived from holdings through the `AccountBalances` view.
- Snapshots represent database state at an import progress point, not natural-date account state.
