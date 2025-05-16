# GraphGetCalendar

This .NET 8.0 console application logs in to Microsoft Graph using username and password (MFA off), fetches events from a shared calendar, and saves them to a SQL Server database using `SqlConnection`. All configuration is in `appsettings.json`.

## appsettings.json Example
```json
{
  "Graph": {
    "UserId": "user@domain.com",
    "Password": "<your password>",
    "ClientId": "<your client id>",
    "TenantId": "<your tenant id>"
  },
  "ConnectionStrings": {
    "SqlServer": "Server=YOUR_SERVER;Database=YOUR_DB;User Id=YOUR_USER;Password=YOUR_PASSWORD;Encrypt=True;TrustServerCertificate=True;"
  },
  "Calendar": {
    "SharedCalendarEmail": "sharedmailboxname@domain.com",
    "MonthsBefore": 1,
    "MonthsAfter": 1
  },
  "Debug": {
    "Login": false,
    "DisplayCalendar": true
  }
}
```

## Settings
- **Graph.UserId**: The user email to authenticate with Microsoft Graph.
- **Graph.Password**: The password for the user (MFA must be off).
- **Graph.ClientId**: Azure AD Application (client) ID.
- **Graph.TenantId**: Azure AD Directory (tenant) ID.
- **ConnectionStrings.SqlServer**: SQL Server connection string.
- **Calendar.SharedCalendarEmail**: The email address of the calendar to fetch events from.
- **Calendar.MonthsBefore**: Number of months before today to include in the query.
- **Calendar.MonthsAfter**: Number of months after today to include in the query.
- **Debug.Login**: If true, only test login and display user info (no calendar or DB actions).
- **Debug.DisplayCalendar**: If true, print calendar events to the console instead of saving to the database.

## Usage
1. Update `appsettings.json` with your credentials, Azure AD app info, and SQL Server connection string.
2. Set the debug options as needed:
   - Set `Debug.Login` to `true` to test login only.
   - Set `Debug.DisplayCalendar` to `true` to print calendar events to the console.
3. Build and run the application:
   ```powershell
   dotnet build
   dotnet run
   ```

## Security Notes
- Do not hardcode credentials in code. Use `appsettings.json` or a secure store (e.g., Azure Key Vault) in production.
- Ensure MFA is disabled for the user account used for authentication.

## Disclaimer
- Username/password authentication is not recommended for production. Use modern auth flows (device code, interactive, managed identity) when possible.
