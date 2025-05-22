using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookMcpAddIn
{
    public static class McpHttpServer
    {
        private static HttpListener _listener;
        private static string _prefix = "http://localhost:8899/mcp/"; // Ensure this port is free
        private static bool _isRunning = false;
        private static IOutlookController _outlookController;
        private static SynchronizationContext _outlookSyncContext;
        private static readonly JsonSerializerOptions _jsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            WriteIndented = true,
            Converters = { new JsonStringEnumConverter() } // For serializing enums as strings
        };

        public static void Start(IOutlookController controller, SynchronizationContext syncContext) // syncContext can be null
        {
            if (_isRunning) return;

            _outlookController = controller ?? throw new ArgumentNullException(nameof(controller));
            _outlookSyncContext = syncContext; // Store it, even if it's null

            _listener = new HttpListener();
            _listener.Prefixes.Add(_prefix);
            try
            {
                _listener.Start();
                _isRunning = true;
                Task.Run(() => ListenLoop()); // Fire and forget the listen loop
                System.Diagnostics.Debug.WriteLine($"MCP HTTP Server listening on {_prefix}. SyncContext is {(_outlookSyncContext == null ? "NULL" : "Present")}");
            }
            catch (HttpListenerException hlex)
            {
                System.Diagnostics.Debug.WriteLine($"MCP HTTP Server start failed: {hlex.Message}. " +
                    "Ensure port is free and URL ACL is set (e.g., 'netsh http add urlacl url=http://localhost:8899/ user=EVERYONE')");
                _isRunning = false; // Ensure state is correct if start fails
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"MCP HTTP Server general error on start: {ex.Message}");
                _isRunning = false; // Ensure state is correct if start fails
            }
        }


        public static void Stop()
        {
            if (!_isRunning) return;
            _isRunning = false;
            try
            {
                _listener?.Stop(); // Stop accepting new requests
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error stopping HttpListener: {ex.Message}");
            }
            finally
            {
                _listener?.Close(); // Close and release resources
                _listener = null;
                System.Diagnostics.Debug.WriteLine("MCP HTTP Server stopped.");
            }
        }

        private static async Task ListenLoop()
        {
            while (_isRunning && _listener != null && _listener.IsListening)
            {
                try
                {
                    HttpListenerContext context = await _listener.GetContextAsync();
                    // Fire and forget processing to quickly return to listening
                    _ = Task.Run(() => ProcessRequestAsync(context));
                }
                catch (HttpListenerException ex) when (ex.ErrorCode == 995 || !_isRunning) // ERROR_OPERATION_ABORTED or server stopping
                {
                    System.Diagnostics.Debug.WriteLine("HttpListener operation aborted or server stopping.");
                    break;
                }
                catch (ObjectDisposedException)
                {
                    System.Diagnostics.Debug.WriteLine("HttpListener disposed, shutting down listen loop.");
                    break;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"MCP HTTP Server listen loop error: {ex.Message}");
                    if (!_isRunning || _listener == null || !_listener.IsListening) break;
                    await Task.Delay(1000); // Prevent tight loop on persistent error
                }
            }
            System.Diagnostics.Debug.WriteLine("MCP HTTP Server ListenLoop ended.");
        }

        private static async Task ProcessRequestAsync(HttpListenerContext context)
        {
            HttpListenerRequest request = context.Request;
            HttpListenerResponse response = context.Response;
            string responseString = "";
            int statusCode = 200;
            object mcpResponsePayload = null;
            string requestBody = null; // To store deserialized request for marshalling

            try
            {
                if (request.HttpMethod == "POST")
                {
                    using (var reader = new StreamReader(request.InputStream, request.ContentEncoding))
                    {
                        requestBody = await reader.ReadToEndAsync(); // Read the body once
                    }

                    Action<object> actionToMarshal = null; // Action to be sent to sync context
                    object stateForMarshal = null;     // State for that action

                    if (request.Url.AbsolutePath.EndsWith("/getContext"))
                    {
                        McpContextRequest mcpRequest = JsonSerializer.Deserialize<McpContextRequest>(requestBody, _jsonOptions);
                        stateForMarshal = mcpRequest;
                        actionToMarshal = (marshaledState) => {
                            mcpResponsePayload = _outlookController.GetMcpContext((McpContextRequest)marshaledState);
                        };
                    }
                    else if (request.Url.AbsolutePath.EndsWith("/getAvailability"))
                    {
                        McpAvailabilityRequest availabilityRequest = JsonSerializer.Deserialize<McpAvailabilityRequest>(requestBody, _jsonOptions);
                        stateForMarshal = availabilityRequest;
                        actionToMarshal = (marshaledState) => {
                            mcpResponsePayload = _outlookController.GetAttendeeAvailability((McpAvailabilityRequest)marshaledState);
                        };
                    }
                    else if (request.Url.AbsolutePath.EndsWith("/createMeeting"))
                    {
                        McpCreateMeetingRequest createRequest = JsonSerializer.Deserialize<McpCreateMeetingRequest>(requestBody, _jsonOptions);
                        stateForMarshal = createRequest;
                        actionToMarshal = (marshaledState) => {
                            mcpResponsePayload = _outlookController.CreateMeeting((McpCreateMeetingRequest)marshaledState);
                        };
                    }
                    else
                    {
                        statusCode = 404;
                        mcpResponsePayload = new { error = "Endpoint Not Found" };
                    }

                    if (actionToMarshal != null) // If a valid endpoint was hit
                    {
                        if (_outlookSyncContext != null)
                        {
                            _outlookSyncContext.Send(new SendOrPostCallback(actionToMarshal), stateForMarshal);
                        }
                        else
                        {
                            // CRITICAL: No SyncContext. Cannot reliably call Outlook.
                            System.Diagnostics.Debug.WriteLine($"CRITICAL: MCPHttpServer - No SynchronizationContext available for {request.Url.AbsolutePath}. Outlook calls will likely fail.");
                            statusCode = 503; // Service Unavailable (or 500 Internal Server Error)
                            mcpResponsePayload = new { error = "Service temporarily unavailable", details = "Cannot process Outlook request due to internal synchronization issue." };
                            // DO NOT attempt to call actionToMarshal directly here, as it will execute on the wrong thread.
                        }
                    }
                }
                else if (request.HttpMethod == "GET" && request.Url.AbsolutePath.EndsWith("/health"))
                {
                    mcpResponsePayload = new { status = "OK", timestamp = DateTime.UtcNow, syncContextAvailable = (_outlookSyncContext != null) };
                }
                else
                {
                    statusCode = 405;
                    mcpResponsePayload = new { error = "Method Not Allowed" };
                }
                responseString = JsonSerializer.Serialize(mcpResponsePayload, _jsonOptions);
            }
            catch (JsonException jex)
            {
                statusCode = 400;
                responseString = JsonSerializer.Serialize(new { error = "Invalid JSON request", details = jex.Message }, _jsonOptions);
            }
            catch (Exception ex)
            {
                statusCode = 500;
                responseString = JsonSerializer.Serialize(new { error = "Internal server error", details = ex.Message }, _jsonOptions);
                System.Diagnostics.Debug.WriteLine($"Error processing MCP request {request.Url}: {ex}");
            }
            finally
            {
                try
                {
                    response.ContentType = "application/json";
                    response.StatusCode = statusCode;
                    byte[] buffer = Encoding.UTF8.GetBytes(responseString);
                    response.ContentLength64 = buffer.Length;
                    await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error writing response: {ex.Message}");
                }
                finally
                {
                    response.OutputStream.Close();
                }
            }
        }
    }
}