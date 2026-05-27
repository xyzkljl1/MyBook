# MyBook

MyBook is a local personal finance and asset tracking application. It imports statements from banks, broker reports, mail attachments, local files, and selected web APIs, then stores records, balances, holdings, snapshots, and fixed bootstrap data in MySQL.

The project is intentionally strict about accounting accuracy: statement totals are used for validation, not as a substitute for detailed records, and imports are expected to be atomic.

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
- Do not fabricate residual records to force statement validation to pass.
