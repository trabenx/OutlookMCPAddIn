using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

// --- MCP Request from LLM App to Outlook Add-in ---
public class McpContextRequest
{
    [JsonPropertyName("request_id")]
    public string RequestId { get; set; } = Guid.NewGuid().ToString();

    [JsonPropertyName("query")]
    public string Query { get; set; }

    [JsonPropertyName("context_types_filter")]
    public List<string> ContextTypesFilter { get; set; }

    [JsonPropertyName("max_items_per_type")]
    public int? MaxItemsPerType { get; set; }

    [JsonPropertyName("time_range_start")]
    public DateTime? TimeRangeStart { get; set; }

    [JsonPropertyName("time_range_end")]
    public DateTime? TimeRangeEnd { get; set; }

    [JsonPropertyName("user_focus_identifier")]
    public string UserFocusIdentifier { get; set; }

    [JsonPropertyName("additional_params")]
    public Dictionary<string, string> AdditionalParams { get; set; }
}

// --- MCP Response from Outlook Add-in to LLM App ---
public class McpContextResponse
{
    [JsonPropertyName("request_id")]
    public string RequestId { get; set; }

    [JsonPropertyName("provider_id")]
    public string ProviderId { get; set; } = "OutlookAddInProvider/v1.0";

    [JsonPropertyName("context_items")]
    public List<McpContextItem> ContextItems { get; set; } = new List<McpContextItem>();

    [JsonPropertyName("summary")]
    public string Summary { get; set; }

    [JsonPropertyName("errors")]
    public List<McpError> Errors { get; set; } = new List<McpError>();
}

public class McpError
{
    [JsonPropertyName("code")]
    public string Code { get; set; }
    [JsonPropertyName("message")]
    public string Message { get; set; }
}

// --- Base Context Item ---
public abstract class McpContextItem
{
    [JsonPropertyName("id")]
    public string Id { get; set; }

    [JsonPropertyName("type")]
    public abstract string Type { get; }

    [JsonPropertyName("source_detail")]
    public string SourceDetail { get; set; }

    [JsonPropertyName("timestamp_retrieved_utc")]
    public DateTime TimestampRetrievedUtc { get; set; } = DateTime.UtcNow;

    [JsonPropertyName("relevance_score")]
    public double? RelevanceScore { get; set; }

    [JsonPropertyName("metadata")]
    public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();
}

// --- Specific Context Item Types ---
public class McpEmailContext : McpContextItem
{
    public override string Type => "email_document";
    [JsonPropertyName("subject")]
    public string Subject { get; set; }
    [JsonPropertyName("sender_name")]
    public string SenderName { get; set; }
    [JsonPropertyName("sender_email")]
    public string SenderEmail { get; set; }
    [JsonPropertyName("recipients_to")]
    public List<string> RecipientsTo { get; set; } = new List<string>();
    [JsonPropertyName("recipients_cc")]
    public List<string> RecipientsCc { get; set; } = new List<string>();
    [JsonPropertyName("date_received_utc")]
    public DateTime DateReceivedUtc { get; set; }
    [JsonPropertyName("date_sent_utc")]
    public DateTime? DateSentUtc { get; set; }
    [JsonPropertyName("body_plain_text")]
    public string BodyPlainText { get; set; }
    [JsonPropertyName("body_html")]
    public string BodyHtml { get; set; }
    [JsonPropertyName("importance")]
    public string Importance { get; set; }
    [JsonPropertyName("has_attachments")]
    public bool HasAttachments { get; set; }
}

public class McpCalendarEventContext : McpContextItem
{
    public override string Type => "calendar_event";
    [JsonPropertyName("subject")]
    public string Subject { get; set; }
    [JsonPropertyName("start_time_utc")]
    public DateTime StartTimeUtc { get; set; }
    [JsonPropertyName("end_time_utc")]
    public DateTime EndTimeUtc { get; set; }
    [JsonPropertyName("is_all_day")]
    public bool IsAllDay { get; set; }
    [JsonPropertyName("location")]
    public string Location { get; set; }
    [JsonPropertyName("organizer")]
    public string Organizer { get; set; }
    [JsonPropertyName("required_attendees")]
    public List<string> RequiredAttendees { get; set; } = new List<string>();
    [JsonPropertyName("optional_attendees")]
    public List<string> OptionalAttendees { get; set; } = new List<string>();
    [JsonPropertyName("body")]
    public string Body { get; set; }
    [JsonPropertyName("response_status")]
    public string ResponseStatus { get; set; }
}

public class McpContactContext : McpContextItem
{
    public override string Type => "contact_profile";
    [JsonPropertyName("full_name")]
    public string FullName { get; set; }
    [JsonPropertyName("first_name")]
    public string FirstName { get; set; }
    [JsonPropertyName("last_name")]
    public string LastName { get; set; }
    [JsonPropertyName("company_name")]
    public string CompanyName { get; set; }
    [JsonPropertyName("job_title")]
    public string JobTitle { get; set; }
    [JsonPropertyName("email_addresses")]
    public List<string> EmailAddresses { get; set; } = new List<string>();
    [JsonPropertyName("phone_numbers")]
    public Dictionary<string, string> PhoneNumbers { get; set; } = new Dictionary<string, string>();
    [JsonPropertyName("notes")]
    public string Notes { get; set; }
}


// --- MCP Request for Availability ---
public class McpAvailabilityRequest
{
    [JsonPropertyName("request_id")]
    public string RequestId { get; set; } = Guid.NewGuid().ToString();
    [JsonPropertyName("attendees")]
    public List<string> Attendees { get; set; }
    [JsonPropertyName("start_date_utc")]
    public DateTime StartDateUtc { get; set; }
    [JsonPropertyName("end_date_utc")]
    public DateTime EndDateUtc { get; set; }
    [JsonPropertyName("meeting_duration_minutes")]
    public int MeetingDurationMinutes { get; set; } = 60;
    [JsonPropertyName("working_hours_only")]
    public bool WorkingHoursOnly { get; set; } = true;
    [JsonPropertyName("minimum_percentage_of_attendees_free")]
    public int MinimumPercentageOfAttendeesFree { get; set; } = 100;
}

// --- MCP Response for Availability ---
public class McpAvailabilityResponse
{
    [JsonPropertyName("request_id")]
    public string RequestId { get; set; }
    [JsonPropertyName("provider_id")]
    public string ProviderId { get; set; } = "OutlookAddInProvider/v1.0";
    [JsonPropertyName("suggested_slots")]
    public List<MeetingSlot> SuggestedSlots { get; set; } = new List<MeetingSlot>();
    [JsonPropertyName("attendee_free_busy_details")]
    public Dictionary<string, string> AttendeeFreeBusyDetails { get; set; }
    [JsonPropertyName("errors")]
    public List<McpError> Errors { get; set; } = new List<McpError>();
}

public class MeetingSlot
{
    [JsonPropertyName("start_time_utc")]
    public DateTime StartTimeUtc { get; set; }
    [JsonPropertyName("end_time_utc")]
    public DateTime EndTimeUtc { get; set; }
    [JsonPropertyName("attendee_availability")]
    public Dictionary<string, AttendeeAvailabilityStatus> AttendeeAvailability { get; set; } = new Dictionary<string, AttendeeAvailabilityStatus>();
    [JsonPropertyName("all_required_attendees_free")]
    public bool AllRequiredAttendeesFree { get; set; }
}

[JsonConverter(typeof(JsonStringEnumConverter))] // Ensures enum is serialized as string
public enum AttendeeAvailabilityStatus
{
    Unknown,
    Free,
    Tentative,
    Busy,
    OutOfOffice
}

// --- MCP Request to Create Meeting ---
public class McpCreateMeetingRequest
{
    [JsonPropertyName("request_id")]
    public string RequestId { get; set; } = Guid.NewGuid().ToString();
    [JsonPropertyName("meeting_details")]
    public McpCalendarEventContext MeetingDetails { get; set; } // Re-using this for simplicity
    [JsonPropertyName("send_invitations")]
    public bool SendInvitations { get; set; } = true;
}

// --- MCP Response for Create Meeting ---
public class McpCreateMeetingResponse
{
    [JsonPropertyName("request_id")]
    public string RequestId { get; set; }
    [JsonPropertyName("provider_id")]
    public string ProviderId { get; set; } = "OutlookAddInProvider/v1.0";
    [JsonPropertyName("status")]
    public string Status { get; set; }
    [JsonPropertyName("meeting_id")]
    public string MeetingId { get; set; }
    [JsonPropertyName("message")]
    public string Message { get; set; }
    [JsonPropertyName("errors")]
    public List<McpError> Errors { get; set; } = new List<McpError>();
}