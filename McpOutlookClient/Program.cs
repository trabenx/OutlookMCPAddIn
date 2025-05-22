using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http; // For HttpClient
using System.Text;
using System.Text.Json; // For System.Text.Json
using System.Threading.Tasks;

// If McpDataModels.cs is in a different namespace in your client project,
// you might need a using directive for it, e.g.:
// using OutlookMcpAddIn; // Assuming models are in this namespace

namespace McpOutlookClient
{
    // --- Paste or reference your McpDataModels.cs content here or ensure it's accessible ---
    // For brevity, I'm assuming the McpContextRequest, McpAvailabilityRequest, McpCreateMeetingRequest,
    // McpContextResponse, McpAvailabilityResponse, McpCreateMeetingResponse, McpError,
    // McpContextItem (and derived types like McpEmailContext, McpCalendarEventContext),
    // MeetingSlot, AttendeeAvailabilityStatus classes are defined and accessible here.
    // If you copied McpDataModels.cs into this project, they should be.

    class Program
    {
        private static readonly HttpClient httpClient = new HttpClient();
        private static readonly string baseMcpUrl = "http://localhost:8999/mcp/"; // Use YOUR port

        private static readonly JsonSerializerOptions jsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true, // Good for deserializing
            WriteIndented = true,
            // Add JsonStringEnumConverter if you used it on the server for request enums
            // (though our current requests don't have enums directly in the request body)
            // Converters = { new JsonStringEnumConverter() }
        };


        static async Task Main(string[] args)
        {
            Console.WriteLine("Outlook MCP Client");
            Console.WriteLine("------------------");

            // Ensure your Outlook Add-in is running in Outlook!

            bool exit = false;
            while (!exit)
            {
                Console.WriteLine("\nChoose an action:");
                Console.WriteLine("1. Get Context (Emails/Calendar)");
                Console.WriteLine("2. Get Attendee Availability");
                Console.WriteLine("3. Create Meeting");
                Console.WriteLine("H. Health Check");
                Console.WriteLine("X. Exit");
                Console.Write("Enter choice: ");
                string choice = Console.ReadLine()?.ToUpper();

                try
                {
                    switch (choice)
                    {
                        case "1":
                            await TestGetContext();
                            break;
                        case "2":
                            await TestGetAvailability();
                            break;
                        case "3":
                            await TestCreateMeeting();
                            break;
                        case "H":
                            await TestHealthCheck();
                            break;
                        case "X":
                            exit = true;
                            break;
                        default:
                            Console.WriteLine("Invalid choice. Try again.");
                            break;
                    }
                }
                catch (HttpRequestException hre)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"HTTP Request Error: {hre.Message}");
                    if (hre.InnerException != null)
                        Console.WriteLine($"Inner Exception: {hre.InnerException.Message}");
                    Console.ResetColor();
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"An error occurred: {ex.Message}");
                    Console.ResetColor();
                }
            }
        }

        static async Task TestHealthCheck()
        {
            Console.WriteLine("\n--- Testing Health Check ---");
            HttpResponseMessage response = await httpClient.GetAsync(baseMcpUrl + "health");
            response.EnsureSuccessStatusCode(); // Throws if not 2xx

            string responseBody = await response.Content.ReadAsStringAsync();
            Console.WriteLine("Health Check Response:");
            Console.WriteLine(FormatJson(responseBody));
        }

        static async Task TestGetContext()
        {
            Console.WriteLine("\n--- Testing Get Context ---");
            Console.Write("Enter search query (e.g., 'meeting update', leave blank for recent): ");
            string query = Console.ReadLine();

            var request = new McpContextRequest
            {
                Query = string.IsNullOrWhiteSpace(query) ? null : query,
                ContextTypesFilter = new List<string> { "email_document", "calendar_event" },
                MaxItemsPerType = 3, // Get up to 3 of each
                //TimeRangeStart = DateTime.UtcNow.AddDays(-7), // Example time range
                //TimeRangeEnd = DateTime.UtcNow,
            };

            string jsonRequest = JsonSerializer.Serialize(request, jsonOptions);
            Console.WriteLine("Sending GetContext Request:");
            Console.WriteLine(FormatJson(jsonRequest));

            var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");
            HttpResponseMessage response = await httpClient.PostAsync(baseMcpUrl + "getContext", content);

            string responseBody = await response.Content.ReadAsStringAsync();
            Console.WriteLine($"GetContext Response (Status: {response.StatusCode}):");
            Console.WriteLine(FormatJson(responseBody));

            if (response.IsSuccessStatusCode)
            {
                // Optionally deserialize and process further if needed
                // var mcpResponse = JsonSerializer.Deserialize<McpContextResponse>(responseBody, jsonOptions);
                // foreach (var item in mcpResponse.ContextItems) { /* ... */ }
            }
        }

        static async Task TestGetAvailability()
        {
            Console.WriteLine("\n--- Testing Get Attendee Availability ---");
            Console.Write("Enter attendee emails (comma-separated, e.g., user1@example.com,user2@example.com): ");
            string emailsInput = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(emailsInput))
            {
                Console.WriteLine("No emails entered.");
                return;
            }
            List<string> attendees = new List<string>(emailsInput.Split(',').Select(e => e.Trim()).Where(e => !string.IsNullOrEmpty(e)));

            Console.Write("Enter meeting duration in minutes (default 60): ");
            if (!int.TryParse(Console.ReadLine(), out int duration) || duration <= 0)
            {
                duration = 60;
            }

            var request = new McpAvailabilityRequest
            {
                Attendees = attendees,
                StartDateUtc = DateTime.UtcNow.Date, // From today
                EndDateUtc = DateTime.UtcNow.Date.AddDays(7), // For the next 7 days
                MeetingDurationMinutes = duration,
                WorkingHoursOnly = true, // Try true and false
                MinimumPercentageOfAttendeesFree = 100 // Or try 50, etc.
            };

            string jsonRequest = JsonSerializer.Serialize(request, jsonOptions);
            Console.WriteLine("Sending GetAvailability Request:");
            Console.WriteLine(FormatJson(jsonRequest));

            var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");
            HttpResponseMessage response = await httpClient.PostAsync(baseMcpUrl + "getAvailability", content);

            string responseBody = await response.Content.ReadAsStringAsync();
            Console.WriteLine($"GetAvailability Response (Status: {response.StatusCode}):");
            Console.WriteLine(FormatJson(responseBody));
            if (response.IsSuccessStatusCode)
            {
                var mcpResponse = JsonSerializer.Deserialize<McpAvailabilityResponse>(responseBody, jsonOptions);
                if (mcpResponse.SuggestedSlots != null && mcpResponse.SuggestedSlots.Any())
                {
                    Console.WriteLine("Suggested Slots:");
                    foreach (var slot in mcpResponse.SuggestedSlots)
                    {
                        Console.WriteLine($"  - Start: {slot.StartTimeUtc.ToLocalTime()}, End: {slot.EndTimeUtc.ToLocalTime()}, AllRequiredFree: {slot.AllRequiredAttendeesFree}");
                    }
                }
                else
                {
                    Console.WriteLine("No suitable slots found based on criteria.");
                }
            }
        }

        static async Task TestCreateMeeting()
        {
            Console.WriteLine("\n--- Testing Create Meeting ---");
            Console.Write("Enter required attendee emails (comma-separated): ");
            string reqEmailsInput = Console.ReadLine();
            List<string> reqAttendees = string.IsNullOrWhiteSpace(reqEmailsInput) ?
                                        new List<string>() :
                                        new List<string>(reqEmailsInput.Split(',').Select(e => e.Trim()).Where(e => !string.IsNullOrEmpty(e)));

            Console.Write("Enter meeting subject: ");
            string subject = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(subject)) subject = "AI Scheduled Test Meeting";

            Console.Write("Enter start time (e.g., 'yyyy-MM-dd HH:mm' LOCAL TIME, or leave blank for 1 hour from now): ");
            string startTimeInput = Console.ReadLine();
            DateTime startTime;
            if (string.IsNullOrWhiteSpace(startTimeInput) || !DateTime.TryParse(startTimeInput, out startTime))
            {
                startTime = DateTime.Now.AddHours(1);
                Console.WriteLine($"Defaulting start time to: {startTime}");
            }

            Console.Write("Enter duration in minutes (default 30): ");
            if (!int.TryParse(Console.ReadLine(), out int duration) || duration <= 0)
            {
                duration = 30;
            }
            DateTime endTime = startTime.AddMinutes(duration);

            var meetingDetails = new McpCalendarEventContext // Re-using this for creating
            {
                Subject = subject,
                StartTimeUtc = startTime.ToUniversalTime(), // Send as UTC
                EndTimeUtc = endTime.ToUniversalTime(),   // Send as UTC
                RequiredAttendees = reqAttendees,
                Body = "This meeting was scheduled via the MCP Outlook Add-in test client.",
                Location = "AI Test Location"
            };

            var request = new McpCreateMeetingRequest
            {
                MeetingDetails = meetingDetails,
                SendInvitations = true // Or false to just save to calendar
            };

            string jsonRequest = JsonSerializer.Serialize(request, jsonOptions);
            Console.WriteLine("Sending CreateMeeting Request:");
            Console.WriteLine(FormatJson(jsonRequest));

            var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");
            HttpResponseMessage response = await httpClient.PostAsync(baseMcpUrl + "createMeeting", content);

            string responseBody = await response.Content.ReadAsStringAsync();
            Console.WriteLine($"CreateMeeting Response (Status: {response.StatusCode}):");
            Console.WriteLine(FormatJson(responseBody));
        }

        // Helper to pretty-print JSON
        static string FormatJson(string jsonString)
        {
            try
            {
                using (var jDoc = JsonDocument.Parse(jsonString))
                {
                    return JsonSerializer.Serialize(jDoc.RootElement, new JsonSerializerOptions { WriteIndented = true });
                }
            }
            catch
            {
                return jsonString; // Return original if parsing fails
            }
        }
    }
}