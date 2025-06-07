using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ApiConsoleClient
{
    class Program
    {
        private static readonly HttpClient httpClient = new HttpClient();
        private static readonly string baseUrl = "http://localhost:8999/mcp/";

        private static readonly JsonSerializerOptions jsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        static async Task Main(string[] args)
        {
            Console.WriteLine("MCP API Console Client");
            Console.WriteLine("-----------------------");

            bool exit = false;
            while (!exit)
            {
                Console.WriteLine("\nChoose an action:");
                Console.WriteLine("1. Get Context");
                Console.WriteLine("2. Get Attendee Availability");
                Console.WriteLine("3. Create Meeting");
                Console.WriteLine("H. Health Check");
                Console.WriteLine("X. Exit");
                Console.Write("Enter choice: ");
                var choice = Console.ReadLine()?.Trim().ToUpperInvariant();

                try
                {
                    switch (choice)
                    {
                        case "1":
                            await GetContext();
                            break;
                        case "2":
                            await GetAvailability();
                            break;
                        case "3":
                            await CreateMeeting();
                            break;
                        case "H":
                            await HealthCheck();
                            break;
                        case "X":
                            exit = true;
                            break;
                        default:
                            Console.WriteLine("Invalid choice.");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Error: {ex.Message}");
                    Console.ResetColor();
                }
            }
        }

        static async Task HealthCheck()
        {
            var resp = await httpClient.GetAsync(baseUrl + "health");
            var body = await resp.Content.ReadAsStringAsync();
            Console.WriteLine($"Status: {resp.StatusCode}");
            Console.WriteLine(FormatJson(body));
        }

        static async Task GetContext()
        {
            Console.Write("Search query (blank for recent): ");
            string query = Console.ReadLine();

            var req = new McpContextRequest
            {
                Query = string.IsNullOrWhiteSpace(query) ? null : query,
                ContextTypesFilter = new List<string> { "email_document", "calendar_event" },
                MaxItemsPerType = 3
            };

            await PostAndPrint("getContext", req);
        }

        static async Task GetAvailability()
        {
            Console.Write("Attendee emails (comma separated): ");
            var input = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(input))
            {
                Console.WriteLine("No attendees provided.");
                return;
            }
            var attendees = input.Split(',').Select(e => e.Trim()).Where(e => !string.IsNullOrEmpty(e)).ToList();

            Console.Write("Meeting duration in minutes (default 60): ");
            if (!int.TryParse(Console.ReadLine(), out int duration))
                duration = 60;

            var req = new McpAvailabilityRequest
            {
                Attendees = attendees,
                StartDateUtc = DateTime.UtcNow.Date,
                EndDateUtc = DateTime.UtcNow.Date.AddDays(7),
                MeetingDurationMinutes = duration,
                WorkingHoursOnly = true,
                MinimumPercentageOfAttendeesFree = 100
            };

            await PostAndPrint("getAvailability", req);
        }

        static async Task CreateMeeting()
        {
            Console.Write("Required attendee emails (comma separated): ");
            var input = Console.ReadLine();
            var attendees = string.IsNullOrWhiteSpace(input)
                ? new List<string>()
                : input.Split(',').Select(e => e.Trim()).Where(e => !string.IsNullOrEmpty(e)).ToList();

            Console.Write("Meeting subject: ");
            var subject = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(subject))
                subject = "API Scheduled Meeting";

            Console.Write("Start time (yyyy-MM-dd HH:mm, local, blank = 1h from now): ");
            var startInput = Console.ReadLine();
            DateTime start;
            if (string.IsNullOrWhiteSpace(startInput) || !DateTime.TryParse(startInput, out start))
                start = DateTime.Now.AddHours(1);

            Console.Write("Duration minutes (default 30): ");
            if (!int.TryParse(Console.ReadLine(), out int duration))
                duration = 30;

            var details = new McpCalendarEventContext
            {
                Subject = subject,
                StartTimeUtc = start.ToUniversalTime(),
                EndTimeUtc = start.AddMinutes(duration).ToUniversalTime(),
                RequiredAttendees = attendees,
                Body = "Scheduled via console client",
                Location = "Console Location"
            };

            var req = new McpCreateMeetingRequest
            {
                MeetingDetails = details,
                SendInvitations = true
            };

            await PostAndPrint("createMeeting", req);
        }

        static async Task PostAndPrint(string endpoint, object payload)
        {
            var json = JsonSerializer.Serialize(payload, jsonOptions);
            Console.WriteLine("Request:");
            Console.WriteLine(FormatJson(json));

            var resp = await httpClient.PostAsync(baseUrl + endpoint,
                new StringContent(json, Encoding.UTF8, "application/json"));

            var body = await resp.Content.ReadAsStringAsync();
            Console.WriteLine($"Response (Status: {resp.StatusCode}):");
            Console.WriteLine(FormatJson(body));
        }

        static string FormatJson(string json)
        {
            try
            {
                using var doc = JsonDocument.Parse(json);
                return JsonSerializer.Serialize(doc.RootElement, new JsonSerializerOptions { WriteIndented = true });
            }
            catch
            {
                return json;
            }
        }
    }
}

