# GraphGetCalendar

A .NET 8.0 console application to log in to Microsoft Graph using username/password, retrieve events from a user or shared calendar, and save them to a SQL Server database. All configuration is managed via `appsettings.json` and optionally `appsecrets.json` for sensitive data.

## Features
- Authenticate to Microsoft Graph using username/password (MFA off, MSAL)
- Retrieve calendar events for a specified user or shared calendar (with correct permissions)
- Save events to SQL Server using parameterized queries (no Entity Framework)
- All configuration (credentials, SQL, calendar, debug) in `appsettings.json` or `appsecrets.json`
- Debug modes for login testing, displaying events, and listing visible calendars
- Dependency injection and configuration via Microsoft.Extensions

## Configuration

### appsettings.json
```json
{
  "Graph": {
    "UserId": "user@domain.com",
    "Password": "yourpassword", // recommend moving to appsecrets.json
    "ClientId": "your-client-id",
    "TenantId": "your-tenant-id"
  },
  "ConnectionStrings": {
    "SqlServer": "Server=YOUR_SERVER;Database=YOUR_DB;User Id=YOUR_USER;Password=YOUR_PASSWORD;Encrypt=True;TrustServerCertificate=True;"
  },
  "Calendar": {
    "SharedCalendarEmail": "user-or-shared@domain.com",
    "MonthsBefore": 1,
    "MonthsAfter": 1
  },
  "Debug": {
    "Login": false,
    "DisplayCalendar": false,
    "ListCalendars": false
  }
}
```

### appsecrets.json
- Optional. If present, overrides values in `appsettings.json`.
- Use for sensitive data (passwords, connection strings, etc.).
- **Do not commit to source control!**
- Example:
```json
{
  "Graph": {
    "UserId": "user@domain.com",
    "Password": "yourpassword",
    "ClientId": "your-client-id",
    "TenantId": "your-tenant-id"
  },
  "ConnectionStrings": {
    "SqlServer": "Server=YOUR_SERVER;Database=YOUR_DB;User Id=YOUR_USER;Password=YOUR_PASSWORD;Encrypt=True;TrustServerCertificate=True;"
  },
  "Calendar": {
    "SharedCalendarEmail": "user-or-shared@domain.com",
    "MonthsBefore": 1,
    "MonthsAfter": 1
  },
  "Debug": {
    "Login": false,
    "DisplayCalendar": false,
    "ListCalendars": false
  }
}
```

## Debug Options
- `Debug:Login`: If true, only tests login and prints user info.
- `Debug:DisplayCalendar`: If true, prints calendar events to console instead of saving to SQL.
- `Debug:ListCalendars`: If true, lists all calendars visible to the login user (for troubleshooting shared calendar access).

## Database Table
To save calendar events, create the following table in your SQL Server database:

```sql
CREATE TABLE CalendarEvents (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    Subject NVARCHAR(255),
    StartTime NVARCHAR(50),
    EndTime NVARCHAR(50),
    Organizer NVARCHAR(255),
    Location NVARCHAR(255),
    BodyPreview NVARCHAR(MAX)
);
```
- `StartTime` and `EndTime` are stored as strings in ISO 8601 format (as returned by Microsoft Graph).
- Adjust column sizes/types as needed for your requirements.

## Usage
1. Configure `appsettings.json` (and optionally `appsecrets.json`).
2. Build and run:
   ```powershell
   dotnet build
   dotnet run
   ```
3. Use debug options as needed for troubleshooting.

## Security
- Place sensitive data in `appsecrets.json` (which is in `.gitignore` by default).
- Never commit secrets to source control.

## Notes
- The app only supports user or shared calendars that are visible to the login user.
- For shared mailboxes, ensure the login user has full access and the mailbox is visible in their calendar list.
- If you encounter errors, use `Debug:ListCalendars` to verify calendar visibility.
