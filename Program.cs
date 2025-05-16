using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Collections.Generic;
using System.Net.Http;
using Microsoft.Kiota.Abstractions.Authentication;
using Microsoft.Data.SqlClient;

namespace GraphGetCalendar
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            var host = Host.CreateDefaultBuilder(args)
                .ConfigureAppConfiguration((context, config) =>
                {
                    config.SetBasePath(Directory.GetCurrentDirectory());
                    config.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
                    config.AddJsonFile("appsecrets.json", optional: true, reloadOnChange: true);
                })
                .ConfigureServices((context, services) =>
                {
                    services.AddSingleton<IConfiguration>(context.Configuration);
                    services.AddTransient<IGraphService, GraphService>();
                    services.AddTransient<ISqlService, SqlService>();
                })
                .Build();

            var graphService = host.Services.GetRequiredService<IGraphService>();
            var sqlService = host.Services.GetRequiredService<ISqlService>();
            var config = host.Services.GetRequiredService<IConfiguration>();

            var debugLogin = config.GetValue<bool>("Debug:Login", false);
            var debugDisplayCalendar = config.GetValue<bool>("Debug:DisplayCalendar", false);

            if (debugLogin)
            {
                var userId = config["Graph:UserId"];
                var password = config["Graph:Password"];
                var clientId = config["Graph:ClientId"];
                var tenantId = config["Graph:TenantId"];
                var publicClient = PublicClientApplicationBuilder
                    .Create(clientId)
                    .WithTenantId(tenantId)
                    .Build();
                var scopes = new[] { "https://graph.microsoft.com/.default" };
                var result = await publicClient.AcquireTokenByUsernamePassword(scopes, userId, password).ExecuteAsync();
                var authProvider = new BaseBearerTokenAuthenticationProvider(result.AccessToken);
                var graphClient = new GraphServiceClient(authProvider);
                var me = await graphClient.Me.GetAsync();
                Console.WriteLine($"Login successful. User: {me?.DisplayName} ({me?.UserPrincipalName})");
                return; // Bypass all calendar and database logic if debug login is true
            }

            var calendarEmail = config["Calendar:SharedCalendarEmail"] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(calendarEmail))
            {
                Console.WriteLine("Shared calendar email is not configured.");
                return;
            }

            var monthsBefore = config.GetValue<int>("Calendar:MonthsBefore", 1);
            var monthsAfter = config.GetValue<int>("Calendar:MonthsAfter", 1);
            var startDate = DateTime.UtcNow.AddMonths(-monthsBefore);
            var endDate = DateTime.UtcNow.AddMonths(monthsAfter);

            // Debug: List all calendars visible to the login user
            if (config.GetValue<bool>("Debug:ListCalendars", false))
            {
                await graphService.ListCalendarsForUserAsync();
                return;
            }

            var events = await graphService.GetCalendarEventsAsync(calendarEmail, startDate, endDate);

            if (debugDisplayCalendar)
            {
                foreach (var ev in events)
                {
                    Console.WriteLine($"Subject: {ev.Subject}\nStart: {ev.Start?.DateTime}\nEnd: {ev.End?.DateTime}\nOrganizer: {ev.Organizer?.EmailAddress?.Address}\nLocation: {ev.Location?.DisplayName}\nBodyPreview: {ev.BodyPreview}\n---");
                }
                return;
            }

            await sqlService.SaveEventsAsync(events);
        }
    }

    public interface IGraphService
    {
        Task<List<Microsoft.Graph.Models.Event>> GetCalendarEventsAsync(string calendarEmail, DateTime start, DateTime end);
        Task ListCalendarsForUserAsync();
    }

    public class GraphService : IGraphService
    {
        private readonly IConfiguration _config;
        public GraphService(IConfiguration config)
        {
            _config = config;
        }

        public async Task<List<Microsoft.Graph.Models.Event>> GetCalendarEventsAsync(string calendarEmail, DateTime start, DateTime end)
        {
            var userId = _config["Graph:UserId"];
            var password = _config["Graph:Password"];
            var clientId = _config["Graph:ClientId"];
            var tenantId = _config["Graph:TenantId"];

            var publicClient = PublicClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .Build();

            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var result = await publicClient.AcquireTokenByUsernamePassword(scopes, userId, password).ExecuteAsync();

            var authProvider = new BaseBearerTokenAuthenticationProvider(result.AccessToken);
            var graphClient = new GraphServiceClient(authProvider);

            // Only use user calendar endpoint (no group lookup)
            try
            {
                var userEvents = await graphClient.Users[calendarEmail].CalendarView.GetAsync(requestConfig =>
                {
                    requestConfig.QueryParameters.StartDateTime = start.ToString("o");
                    requestConfig.QueryParameters.EndDateTime = end.ToString("o");
                    requestConfig.QueryParameters.Select = new[] { "subject", "start", "end", "organizer", "location", "bodyPreview" };
                    requestConfig.QueryParameters.Top = 50;
                });
                return userEvents?.Value?.ToList() ?? new();
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError odataEx)
            {
                Console.WriteLine($"Graph API error: {odataEx.Error?.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error accessing user calendar: {ex.Message}");
                throw;
            }
        }

        public async Task ListCalendarsForUserAsync()
        {
            var userId = _config["Graph:UserId"];
            var password = _config["Graph:Password"];
            var clientId = _config["Graph:ClientId"];
            var tenantId = _config["Graph:TenantId"];

            var publicClient = PublicClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .Build();

            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var result = await publicClient.AcquireTokenByUsernamePassword(scopes, userId, password).ExecuteAsync();

            var authProvider = new BaseBearerTokenAuthenticationProvider(result.AccessToken);
            var graphClient = new GraphServiceClient(authProvider);

            var calendars = await graphClient.Me.Calendars.GetAsync();
            Console.WriteLine("Calendars visible to the login user:");
            if (calendars?.Value != null)
            {
                foreach (var cal in calendars.Value)
                {
                    Console.WriteLine($"Name: {cal.Name}, Id: {cal.Id}");
                }
            }
            else
            {
                Console.WriteLine("No calendars found or insufficient permissions.");
            }
        }
    }

    public interface ISqlService
    {
        Task SaveEventsAsync(List<Microsoft.Graph.Models.Event> events);
    }

    public class SqlService : ISqlService
    {
        private readonly IConfiguration _config;
        public SqlService(IConfiguration config)
        {
            _config = config;
        }

        public async Task SaveEventsAsync(List<Microsoft.Graph.Models.Event> events)
        {
            var connStr = _config.GetConnectionString("SqlServer");
            using var conn = new SqlConnection(connStr);
            await conn.OpenAsync();

            foreach (var ev in events)
            {
                using var cmd = new SqlCommand(@"INSERT INTO CalendarEvents (Subject, StartTime, EndTime, Organizer, Location, BodyPreview) VALUES (@Subject, @StartTime, @EndTime, @Organizer, @Location, @BodyPreview)", conn);
                cmd.Parameters.AddWithValue("@Subject", ev.Subject ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@StartTime", ev.Start?.DateTime ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@EndTime", ev.End?.DateTime ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@Organizer", ev.Organizer?.EmailAddress?.Address ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@Location", ev.Location?.DisplayName ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@BodyPreview", ev.BodyPreview ?? (object)DBNull.Value);
                await cmd.ExecuteNonQueryAsync();
            }
        }
    }

    // Helper for GraphServiceClient authentication
    public class BaseBearerTokenAuthenticationProvider : IAuthenticationProvider
    {
        private readonly string _accessToken;
        public BaseBearerTokenAuthenticationProvider(string accessToken)
        {
            _accessToken = accessToken;
        }
        public Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _accessToken);
            return Task.CompletedTask;
        }
        public Task AuthenticateRequestAsync(Microsoft.Kiota.Abstractions.RequestInformation request, System.Collections.Generic.Dictionary<string, object>? additionalAuthenticationContext = null, System.Threading.CancellationToken cancellationToken = default)
        {
            request.Headers["Authorization"] = new List<string> { $"Bearer {_accessToken}" };
            return Task.CompletedTask;
        }
    }
}
