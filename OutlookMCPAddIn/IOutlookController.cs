using System;
using System.Runtime.InteropServices;

namespace OutlookMcpAddIn
{
    [ComVisible(true)]
    [Guid("BDA5CD2E-58BD-49C9-AAF0-4358C648B144")] // Generate a new GUID
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IOutlookController
    {
        McpContextResponse GetMcpContext(McpContextRequest mcpRequest);
        McpAvailabilityResponse GetAttendeeAvailability(McpAvailabilityRequest availabilityRequest);
        McpCreateMeetingResponse CreateMeeting(McpCreateMeetingRequest createMeetingRequest);

        // You can add back other non-MCP specific methods here if needed for direct COM access
        // e.g., string GetCurrentUserEmail();
    }
}