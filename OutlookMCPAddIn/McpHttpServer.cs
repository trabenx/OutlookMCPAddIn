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

        public static void Start(IOutlookController controller, SynchronizationContext syncContext)
        {
            if (_isRunning) return;

            _outlookController = controller ?? throw new ArgumentNullException(nameof(controller));
            _outlookSyncContext = syncContext ?? throw new ArgumentNullException(nameof(syncContext));

            _listener = new HttpListener();
            _listener.Prefixes.Add(_prefix);
            try
            {
                _listener.Start();
                _isRunning = true;
                Task.Run(() => ListenLoop());
                System.Diagnostics.Debug.WriteLine($"MCP HTTP Server listening on {_prefix}");
            }
            catch (HttpListenerException hlex)
            {
                System.Diagnostics.Debug.WriteLine($"MCP HTTP Server start failed: {hlex.Message}. " +
                    "Ensure port is free and URL ACL is set (e.g., 'netsh http add urlacl url=http://localhost:8899/ user=EVERYONE')");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"MCP HTTP Server general error on start: {ex.Message}");
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

            try
            {
                if (request.HttpMethod == "POST")
                {
                    string requestBody;
                    using (var reader = new StreamReader(request.InputStream, request.ContentEncoding))
                    {
                        requestBody = await reader.ReadToEndAsync();
                    }

                    if (request.Url.AbsolutePath.EndsWith("/getContext"))
                    {
                        McpContextRequest mcpRequest = JsonSerializer.Deserialize<McpContextRequest>(requestBody, _jsonOptions);
                        _outlookSyncContext.Send(state =>
                        {
                            mcpResponsePayload = _outlookController.GetMcpContext((McpContextRequest)state);
                        }, mcpRequest);
                    }
                    else if (request.Url.AbsolutePath.EndsWith("/getAvailability"))
                    {
                        McpAvailabilityRequest availabilityRequest = JsonSerializer.Deserialize<McpAvailabilityRequest>(requestBody, _jsonOptions);
                        _outlookSyncContext.Send(state =>
                        {
                            mcpResponsePayload = _outlookController.GetAttendeeAvailability((McpAvailabilityRequest)state);
                        }, availabilityRequest);
                    }
                    else if (request.Url.AbsolutePath.EndsWith("/createMeeting"))
                    {
                        McpCreateMeetingRequest createRequest = JsonSerializer.Deserialize<McpCreateMeetingRequest>(requestBody, _jsonOptions);
                        _outlookSyncContext.Send(state =>
                        {
                            mcpResponsePayload = _outlookController.CreateMeeting((McpCreateMeetingRequest)state);
                        }, createRequest);
                    }
                    else
                    {
                        statusCode = 404;
                        mcpResponsePayload = new { error = "Endpoint Not Found" };
                    }
                    responseString = JsonSerializer.Serialize(mcpResponsePayload, _jsonOptions);
                }
                else if (request.HttpMethod == "GET" && request.Url.AbsolutePath.EndsWith("/health"))
                {
                    mcpResponsePayload = new { status = "OK", timestamp = DateTime.UtcNow };
                    responseString = JsonSerializer.Serialize(mcpResponsePayload, _jsonOptions);
                }
                else
                {
                    statusCode = 405; // Method Not Allowed
                    mcpResponsePayload = new { error = "Method Not Allowed" };
                    responseString = JsonSerializer.Serialize(mcpResponsePayload, _jsonOptions);
                }
            }
            catch (JsonException jex)
            {
                statusCode = 400; // Bad Request
                responseString = JsonSerializer.Serialize(new { error = "Invalid JSON request", details = jex.Message }, _jsonOptions);
            }
            catch (Exception ex)
            {
                statusCode = 500; // Internal Server Error
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