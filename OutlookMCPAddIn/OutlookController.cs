using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcpAddIn
{
    [ComVisible(true)]
    [Guid("F11CE6E8-D9D4-4904-9B97-259874922AD9")] // Generate a new GUID
    [ClassInterface(ClassInterfaceType.None)] // Explicitly implement IOutlookController
    public class OutlookController : IOutlookController
    {
        private readonly Outlook.Application _outlookApp;
        private readonly Outlook.NameSpace _mapiNamespace;

        public OutlookController(Outlook.Application outlookApp)
        {
            _outlookApp = outlookApp ?? throw new ArgumentNullException(nameof(outlookApp));
            _mapiNamespace = _outlookApp.GetNamespace("MAPI");
        }

        // --- Helper: Convert Outlook.MailItem to McpEmailContext ---
        private McpEmailContext ToMcpEmailContext(Outlook.MailItem mailItem, string sourceDetail)
        {
            if (mailItem == null) return null;
            var ctx = new McpEmailContext
            {
                Id = mailItem.EntryID,
                SourceDetail = sourceDetail,
                Subject = mailItem.Subject,
                SenderName = mailItem.SenderName,
                DateReceivedUtc = mailItem.ReceivedTime.ToUniversalTime(),
                BodyPlainText = mailItem.Body, // For LLMs, plain text is often preferred
                Importance = ((Outlook.OlImportance)mailItem.Importance).ToString().Replace("olImportance", "").ToLower(),
                HasAttachments = mailItem.Attachments.Count > 0,
                Metadata = new Dictionary<string, object>()
            };
            // Optionally add HTMLBody if truly needed, but be mindful of size
            // if (!string.IsNullOrEmpty(mailItem.HTMLBody))
            // {
            //      ctx.BodyHtml = mailItem.HTMLBody;
            // }

            try { ctx.SenderEmail = mailItem.SenderEmailAddress; }
            catch { /* Ignore if SenderEmailAddress is not available (e.g., for certain contact types) */ }


            if (mailItem.SentOn.Year > 1) // Check if SentOn is a valid date (avoids 01/01/4501 default)
            {
                ctx.DateSentUtc = mailItem.SentOn.ToUniversalTime();
            }

            Outlook.Recipients recipients = null;
            try
            {
                recipients = mailItem.Recipients;
                if (recipients != null && recipients.Count > 0)
                {
                    for (int i = 1; i <= recipients.Count; i++)
                    {
                        Outlook.Recipient recip = null;
                        try
                        {
                            recip = recipients[i];
                            // Ensure AddressEntry and Address are not null before accessing
                            string address = recip.AddressEntry?.Address ?? recip.Name;
                            if (recip.Type == (int)Outlook.OlMailRecipientType.olTo) ctx.RecipientsTo.Add(address);
                            else if (recip.Type == (int)Outlook.OlMailRecipientType.olCC) ctx.RecipientsCc.Add(address);
                        }
                        catch (COMException ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error accessing recipient property: {ex.Message}");
                        }
                        finally { if (recip != null) Marshal.ReleaseComObject(recip); }
                    }
                }
            }
            catch (COMException ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error accessing mailItem.Recipients: {ex.Message}");
            }
            finally { if (recipients != null) Marshal.ReleaseComObject(recipients); }

            return ctx;
        }

        // --- Helper: Convert Outlook.AppointmentItem to McpCalendarEventContext ---
        private McpCalendarEventContext ToMcpCalendarEventContext(Outlook.AppointmentItem apptItem, string sourceDetail)
        {
            if (apptItem == null) return null;
            var ctx = new McpCalendarEventContext
            {
                Id = apptItem.EntryID,
                SourceDetail = sourceDetail,
                Subject = apptItem.Subject,
                StartTimeUtc = apptItem.StartUTC,
                EndTimeUtc = apptItem.EndUTC,
                IsAllDay = apptItem.AllDayEvent,
                Location = apptItem.Location,
                Organizer = apptItem.Organizer,
                Body = apptItem.Body,
                // ***** CORRECTED PROPERTY *****
                ResponseStatus = ((Outlook.OlResponseStatus)apptItem.ResponseStatus).ToString().ToLower().Replace("olresponse", "") // Remove "olresponse" not "olresponsestatus"
            };

            Outlook.Recipients recipients = null;
            try
            {
                recipients = apptItem.Recipients;
                if (recipients != null && recipients.Count > 0)
                {
                    for (int i = 1; i <= recipients.Count; i++)
                    {
                        Outlook.Recipient recip = null;
                        try
                        {
                            recip = recipients[i];
                            string address = recip.AddressEntry?.Address ?? recip.Name;
                            // ***** Enum members olRequired and olOptional ARE CORRECT for OlMailRecipientType *****
                            if (recip.Type == (int)Outlook.OlMeetingRecipientType.olRequired) ctx.RequiredAttendees.Add(address);
                            else if (recip.Type == (int)Outlook.OlMeetingRecipientType.olOptional) ctx.OptionalAttendees.Add(address);
                        }
                        catch (COMException ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error accessing recipient property for appointment: {ex.Message}");
                        }
                        finally { if (recip != null) Marshal.ReleaseComObject(recip); }
                    }
                }
            }
            catch (COMException ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error accessing apptItem.Recipients: {ex.Message}");
            }
            finally { if (recipients != null) Marshal.ReleaseComObject(recipients); }
            return ctx;
        }

        // --- Helper: Convert Outlook.ContactItem to McpContactContext ---
        private McpContactContext ToMcpContactContext(Outlook.ContactItem contactItem, string sourceDetail)
        {
            if (contactItem == null) return null;
            var ctx = new McpContactContext
            {
                Id = contactItem.EntryID,
                SourceDetail = sourceDetail,
                FullName = contactItem.FullName,
                FirstName = contactItem.FirstName,
                LastName = contactItem.LastName,
                CompanyName = contactItem.CompanyName,
                JobTitle = contactItem.JobTitle,
                Notes = contactItem.Body
            };
            if (!string.IsNullOrEmpty(contactItem.Email1Address)) ctx.EmailAddresses.Add(contactItem.Email1Address);
            if (!string.IsNullOrEmpty(contactItem.Email2Address)) ctx.EmailAddresses.Add(contactItem.Email2Address);
            if (!string.IsNullOrEmpty(contactItem.Email3Address)) ctx.EmailAddresses.Add(contactItem.Email3Address);

            if (!string.IsNullOrEmpty(contactItem.BusinessTelephoneNumber)) ctx.PhoneNumbers["business"] = contactItem.BusinessTelephoneNumber;
            if (!string.IsNullOrEmpty(contactItem.HomeTelephoneNumber)) ctx.PhoneNumbers["home"] = contactItem.HomeTelephoneNumber;
            if (!string.IsNullOrEmpty(contactItem.MobileTelephoneNumber)) ctx.PhoneNumbers["mobile"] = contactItem.MobileTelephoneNumber;
            return ctx;
        }

        public McpContextResponse GetMcpContext(McpContextRequest mcpRequest)
        {
            // ... (No changes here based on the errors reported) ...
            // (Previous implementation for GetMcpContext)
            var response = new McpContextResponse { RequestId = mcpRequest.RequestId };
            int maxItems = mcpRequest.MaxItemsPerType ?? 10;
            string currentUserIdentity = _mapiNamespace.CurrentUser?.Name ?? "UnknownUser";

            try
            {
                if (mcpRequest.ContextTypesFilter == null || mcpRequest.ContextTypesFilter.Any(f => f.Equals("email_document", StringComparison.OrdinalIgnoreCase)))
                {
                    Outlook.MAPIFolder folder = null; Outlook.Items items = null;
                    try
                    {
                        folder = _mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                        string filter = BuildEmailFilter(mcpRequest.Query, mcpRequest.TimeRangeStart, mcpRequest.TimeRangeEnd);
                        items = string.IsNullOrEmpty(filter) ? folder.Items : folder.Items.Restrict(filter);
                        items.Sort("[ReceivedTime]", true);

                        for (int i = 1; i <= Math.Min(items.Count, maxItems); i++)
                        {
                            object itemObj = null; Outlook.MailItem mailItem = null;
                            try
                            {
                                itemObj = items[i]; // Get item by 1-based index
                                if (itemObj is Outlook.MailItem mi)
                                {
                                    mailItem = mi;
                                    var emailCtx = ToMcpEmailContext(mailItem, $"Outlook/Inbox/{currentUserIdentity}");
                                    if (emailCtx != null) response.ContextItems.Add(emailCtx);
                                }
                            }
                            catch (COMException ex) { System.Diagnostics.Debug.WriteLine($"Error processing email item {i}: {ex.Message}"); }
                            finally
                            {
                                if (mailItem != null) Marshal.ReleaseComObject(mailItem);
                                else if (itemObj != null) Marshal.ReleaseComObject(itemObj); // Release if not a MailItem but still an object
                            }
                        }
                    }
                    catch (COMException ex) { response.Errors.Add(new McpError { Code = "OutlookEmailError", Message = ex.Message }); }
                    finally
                    {
                        if (items != null) Marshal.ReleaseComObject(items);
                        if (folder != null) Marshal.ReleaseComObject(folder);
                    }
                }

                if (mcpRequest.ContextTypesFilter == null || mcpRequest.ContextTypesFilter.Any(f => f.Equals("calendar_event", StringComparison.OrdinalIgnoreCase)))
                {
                    Outlook.MAPIFolder folder = null; Outlook.Items items = null;
                    try
                    {
                        folder = _mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                        string filter = BuildCalendarFilter(mcpRequest.Query, mcpRequest.TimeRangeStart, mcpRequest.TimeRangeEnd);

                        items = string.IsNullOrEmpty(filter) ? folder.Items : folder.Items.Restrict(filter);
                        items.IncludeRecurrences = true;
                        items.Sort("[Start]", false);

                        int count = 0;
                        // Iterate using a for loop with 1-based index for Outlook.Items
                        for (int i = 1; i <= items.Count && count < maxItems; i++)
                        {
                            object itemObj = null; Outlook.AppointmentItem apptItem = null;
                            try
                            {
                                itemObj = items[i]; // Get item by 1-based index
                                if (itemObj is Outlook.AppointmentItem ai)
                                {
                                    apptItem = ai;
                                    DateTime itemStartTime = apptItem.StartUTC; // Use UTC for comparisons
                                    bool inRange = true; // Assume in range unless proven otherwise
                                    if (mcpRequest.TimeRangeStart.HasValue && itemStartTime < mcpRequest.TimeRangeStart.Value.ToUniversalTime())
                                        inRange = false;
                                    if (mcpRequest.TimeRangeEnd.HasValue && itemStartTime > mcpRequest.TimeRangeEnd.Value.ToUniversalTime()) // itemStartTime should not exceed end of range
                                        inRange = false;

                                    if (inRange)
                                    {
                                        var eventCtx = ToMcpCalendarEventContext(apptItem, $"Outlook/Calendar/{currentUserIdentity}");
                                        if (eventCtx != null) response.ContextItems.Add(eventCtx);
                                        count++;
                                    }
                                }
                            }
                            catch (COMException ex) { System.Diagnostics.Debug.WriteLine($"Error processing calendar item {i}: {ex.Message}"); }
                            finally
                            {
                                if (apptItem != null) Marshal.ReleaseComObject(apptItem);
                                else if (itemObj != null) Marshal.ReleaseComObject(itemObj);
                            }
                        }
                    }
                    catch (COMException ex) { response.Errors.Add(new McpError { Code = "OutlookCalendarError", Message = ex.Message }); }
                    finally
                    {
                        if (items != null) Marshal.ReleaseComObject(items);
                        if (folder != null) Marshal.ReleaseComObject(folder);
                    }
                }
                // Add Contact search logic here if needed
            }
            catch (COMException ex)
            {
                response.Errors.Add(new McpError { Code = "OutlookError", Message = $"COM Exception: {ex.Message} (ErrorCode: {ex.ErrorCode})" });
            }
            catch (Exception ex)
            {
                response.Errors.Add(new McpError { Code = "ProcessingError", Message = ex.Message });
            }
            return response;
        }

        public McpAvailabilityResponse GetAttendeeAvailability(McpAvailabilityRequest request)
        {
            // ... (No changes here based on the errors reported) ...
            // (Previous implementation for GetAttendeeAvailability)
            var response = new McpAvailabilityResponse { RequestId = request.RequestId };
            if (request.Attendees == null || !request.Attendees.Any())
            {
                response.Errors.Add(new McpError { Code = "InputError", Message = "No attendees specified." });
                return response;
            }

            var freeBusyDetails = new Dictionary<string, string>();
            var attendeeFreeBusyStatusLists = new Dictionary<string, List<AttendeeAvailabilityStatus>>();
            int freeBusyIntervalMinutes = 30;

            try
            {
                foreach (var attendeeEmail in request.Attendees)
                {
                    Outlook.Recipient recipient = null;
                    try
                    {
                        recipient = _mapiNamespace.CreateRecipient(attendeeEmail);
                        recipient.Resolve(); // Attempt to resolve the recipient
                        if (recipient.Resolved)
                        {
                            // Note: StartDateUtc must be a valid DateTime, MinPerChar an int (e.g., 15, 30, 60)
                            // CompleteFormat = true gives more details (0=free, 1=tentative, 2=busy, 3=OOF)
                            string fbData = recipient.FreeBusy(request.StartDateUtc, freeBusyIntervalMinutes, true);
                            freeBusyDetails[attendeeEmail] = fbData;

                            var statuses = fbData.Select(c =>
                            {
                                switch (c)
                                {
                                    case '0': return AttendeeAvailabilityStatus.Free;
                                    case '1': return AttendeeAvailabilityStatus.Tentative;
                                    case '2': return AttendeeAvailabilityStatus.Busy;
                                    case '3': return AttendeeAvailabilityStatus.OutOfOffice;
                                    default: return AttendeeAvailabilityStatus.Unknown;
                                }
                            }).ToList();
                            attendeeFreeBusyStatusLists[attendeeEmail] = statuses;
                        }
                        else
                        {
                            response.Errors.Add(new McpError { Code = "ResolutionFailed", Message = $"Could not resolve: {attendeeEmail}" });
                            attendeeFreeBusyStatusLists[attendeeEmail] = new List<AttendeeAvailabilityStatus>(); // Indicate data unavailable
                        }
                    }
                    catch (COMException ex)
                    {
                        response.Errors.Add(new McpError { Code = "FreeBusyError", Message = $"Error getting FreeBusy for {attendeeEmail}: {ex.Message}" });
                        attendeeFreeBusyStatusLists[attendeeEmail] = new List<AttendeeAvailabilityStatus>();
                    }
                    finally
                    {
                        if (recipient != null) Marshal.ReleaseComObject(recipient);
                    }
                }

                response.AttendeeFreeBusyDetails = freeBusyDetails;

                TimeSpan totalDuration = request.EndDateUtc - request.StartDateUtc;
                int numIntervalsTotal = (int)(totalDuration.TotalMinutes / freeBusyIntervalMinutes);
                int slotsNeededForMeeting = request.MeetingDurationMinutes / freeBusyIntervalMinutes;
                if (request.MeetingDurationMinutes % freeBusyIntervalMinutes != 0)
                {
                    // If meeting duration is not a multiple of interval, round up slots needed
                    slotsNeededForMeeting++;
                }


                for (int i = 0; i <= numIntervalsTotal - slotsNeededForMeeting; i++)
                {
                    DateTime slotStart = request.StartDateUtc.AddMinutes(i * freeBusyIntervalMinutes);
                    DateTime slotEnd = slotStart.AddMinutes(request.MeetingDurationMinutes);

                    if (slotEnd > request.EndDateUtc) break; // Ensure meeting ends within the requested EndDateUtc

                    if (request.WorkingHoursOnly)
                    {
                        DateTime localSlotStart = slotStart.ToLocalTime();
                        DateTime localSlotEnd = slotEnd.ToLocalTime();
                        if (localSlotStart.DayOfWeek == DayOfWeek.Saturday || localSlotStart.DayOfWeek == DayOfWeek.Sunday ||
                            localSlotStart.TimeOfDay < new TimeSpan(9, 0, 0) ||
                            localSlotEnd.TimeOfDay > new TimeSpan(17, 0, 0) || // Check end time as well
                            (localSlotEnd.Date > localSlotStart.Date && localSlotEnd.TimeOfDay > new TimeSpan(0, 0, 0))) // Spans past midnight into non-workday
                        {
                            continue;
                        }
                    }

                    var currentSlotAttendeeAvailability = new Dictionary<string, AttendeeAvailabilityStatus>();
                    int freeThisSlotCount = 0;

                    foreach (var attendeeEmail in request.Attendees)
                    {
                        AttendeeAvailabilityStatus worstStatusForAttendeeInSlot = AttendeeAvailabilityStatus.Free;
                        if (attendeeFreeBusyStatusLists.TryGetValue(attendeeEmail, out var statuses) && statuses.Any())
                        {
                            for (int k = 0; k < slotsNeededForMeeting; k++)
                            {
                                int intervalIndex = i + k;
                                if (intervalIndex < statuses.Count)
                                {
                                    // Higher enum value means less available
                                    if (statuses[intervalIndex] > worstStatusForAttendeeInSlot)
                                    {
                                        worstStatusForAttendeeInSlot = statuses[intervalIndex];
                                    }
                                }
                                else
                                {
                                    worstStatusForAttendeeInSlot = AttendeeAvailabilityStatus.Unknown; // Not enough data
                                    break;
                                }
                            }
                        }
                        else { worstStatusForAttendeeInSlot = AttendeeAvailabilityStatus.Unknown; }

                        currentSlotAttendeeAvailability[attendeeEmail] = worstStatusForAttendeeInSlot;
                        // Consider Tentative as "available enough" for counting purposes
                        if (worstStatusForAttendeeInSlot == AttendeeAvailabilityStatus.Free || worstStatusForAttendeeInSlot == AttendeeAvailabilityStatus.Tentative)
                        {
                            freeThisSlotCount++;
                        }
                    }

                    double percentageFree = request.Attendees.Any() ? ((double)freeThisSlotCount / request.Attendees.Count) * 100.0 : 100.0;

                    if (percentageFree >= request.MinimumPercentageOfAttendeesFree)
                    {
                        bool allRequiredTrulyFree = true; // Stricter check for the AllRequiredAttendeesFree flag
                        if (request.MinimumPercentageOfAttendeesFree == 100)
                        { // Only if 100% is required, check for strictly "Free"
                            allRequiredTrulyFree = currentSlotAttendeeAvailability.Values.All(s => s == AttendeeAvailabilityStatus.Free);
                        }
                        else
                        { // Otherwise, Tentative is okay for the flag if they contribute to percentage
                            allRequiredTrulyFree = currentSlotAttendeeAvailability.Values.All(s => s == AttendeeAvailabilityStatus.Free || s == AttendeeAvailabilityStatus.Tentative);
                        }


                        response.SuggestedSlots.Add(new MeetingSlot
                        {
                            StartTimeUtc = slotStart,
                            EndTimeUtc = slotEnd,
                            AttendeeAvailability = currentSlotAttendeeAvailability,
                            AllRequiredAttendeesFree = allRequiredTrulyFree
                        });
                    }
                }
            }
            catch (COMException ex)
            {
                response.Errors.Add(new McpError { Code = "OutlookAvailabilityError", Message = ex.Message });
            }
            catch (Exception ex)
            {
                response.Errors.Add(new McpError { Code = "GeneralAvailabilityError", Message = ex.Message });
            }
            return response;
        }

        public McpCreateMeetingResponse CreateMeeting(McpCreateMeetingRequest request)
        {
            System.Diagnostics.Debug.WriteLine($"[OutlookController.CreateMeeting] Entered. Thread ID: {Thread.CurrentThread.ManagedThreadId}");
            var response = new McpCreateMeetingResponse { RequestId = request.RequestId };
            if (request.MeetingDetails == null)
            {
                response.Errors.Add(new McpError { Code = "InputError", Message = "MeetingDetails not provided." });
                response.Status = "failure";
                return response;
            }

            Outlook.AppointmentItem appointment = null;
            bool allResolved = true; // Assume true initially

            try
            {
                System.Diagnostics.Debug.WriteLine("[CreateMeeting] Trying to create AppointmentItem...");
                appointment = _outlookApp.CreateItem(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;
                if (appointment == null)
                {
                    response.Errors.Add(new McpError { Code = "CreateItemFailed", Message = "OutlookApp.CreateItem returned null for AppointmentItem." });
                    response.Status = "failure";
                    System.Diagnostics.Debug.WriteLine("[CreateMeeting] CreateItem returned null!");
                    return response;
                }
                System.Diagnostics.Debug.WriteLine("[CreateMeeting] AppointmentItem created.");

                var details = request.MeetingDetails;
                appointment.Subject = details.Subject;
                System.Diagnostics.Debug.WriteLine($"[CreateMeeting] Subject set to: {details.Subject}");
                appointment.StartUTC = details.StartTimeUtc;
                System.Diagnostics.Debug.WriteLine($"[CreateMeeting] StartUTC set to: {details.StartTimeUtc}");
                appointment.EndUTC = details.EndTimeUtc;
                System.Diagnostics.Debug.WriteLine($"[CreateMeeting] EndUTC set to: {details.EndTimeUtc}");
                appointment.Location = details.Location;
                System.Diagnostics.Debug.WriteLine($"[CreateMeeting] Location set to: {details.Location}");
                appointment.Body = details.Body;
                System.Diagnostics.Debug.WriteLine($"[CreateMeeting] Body set.");
                appointment.AllDayEvent = details.IsAllDay;
                System.Diagnostics.Debug.WriteLine($"[CreateMeeting] AllDayEvent set to: {details.IsAllDay}");

                // Set meeting status to indicate it's a meeting, not just an appointment
                appointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                System.Diagnostics.Debug.WriteLine($"[CreateMeeting] MeetingStatus set to olMeeting.");


                Outlook.Recipients recipients = null;
                try
                {
                    recipients = appointment.Recipients;
                    if (details.RequiredAttendees != null)
                    {
                        foreach (var email in details.RequiredAttendees)
                        {
                            Outlook.Recipient r = null;
                            try { r = recipients.Add(email); r.Type = 1; /* olRequired */ } // Using integer literal
                            finally { if (r != null) Marshal.ReleaseComObject(r); }
                        }
                    }
                    if (details.OptionalAttendees != null)
                    {
                        foreach (var email in details.OptionalAttendees)
                        {
                            Outlook.Recipient r = null;
                            try { r = recipients.Add(email); r.Type = 2; /* olOptional */ } // Using integer literal
                            finally { if (r != null) Marshal.ReleaseComObject(r); }
                        }
                    }
                    System.Diagnostics.Debug.WriteLine("[CreateMeeting] Attempting to ResolveAll recipients...");
                    bool resolveAllResult = recipients.ResolveAll();
                    System.Diagnostics.Debug.WriteLine($"[CreateMeeting] ResolveAll result: {resolveAllResult}");
                }
                catch (COMException comExRecip)
                {
                    System.Diagnostics.Debug.WriteLine($"[CreateMeeting] COMException during recipient add/resolve: {comExRecip.ToString()}");
                    response.Errors.Add(new McpError { Code = "RecipientError", Message = $"Recipient processing error: {comExRecip.Message}" });
                    // Potentially set response.Status to failure here if this is critical
                }
                finally { if (recipients != null) Marshal.ReleaseComObject(recipients); }

                // Re-check resolution after ResolveAll
                Outlook.Recipients finalRecipients = null;
                try
                {
                    finalRecipients = appointment.Recipients;
                    if (finalRecipients != null && finalRecipients.Count > 0)
                    {
                        for (int i = 1; i <= finalRecipients.Count; i++)
                        {
                            Outlook.Recipient r = null;
                            try
                            {
                                r = finalRecipients[i];
                                if (!r.Resolved)
                                {
                                    allResolved = false;
                                    System.Diagnostics.Debug.WriteLine($"[CreateMeeting] Unresolved Recipient: {r.Name} (Type: {r.Type})");
                                    response.Errors.Add(new McpError { Code = "UnresolvedRecipient", Message = $"Could not resolve: {r.Name}" });
                                }
                            }
                            finally { if (r != null) Marshal.ReleaseComObject(r); }
                        }
                    }
                }
                catch (COMException comExResCheck)
                {
                    System.Diagnostics.Debug.WriteLine($"[CreateMeeting] COMException during final recipient check: {comExResCheck.ToString()}");
                    response.Errors.Add(new McpError { Code = "ResolutionCheckError", Message = $"Resolution check error: {comExResCheck.Message}" });
                }
                finally { if (finalRecipients != null) Marshal.ReleaseComObject(finalRecipients); }

                if (!allResolved && request.SendInvitations)
                {
                    System.Diagnostics.Debug.WriteLine("[CreateMeeting] Not attempting to Send due to unresolved recipients.");
                    response.Errors.Add(new McpError { Code = "SendBlocked", Message = "Cannot send meeting with unresolved recipients." });
                    response.Status = "failure";
                    // DO NOT PROCEED TO SEND/SAVE IF INTENDING TO SEND AND RECIPIENTS AREN'T RESOLVED
                }
                else // Proceed if all resolved OR if we are just saving (SendInvitations is false)
                {
                    if (request.SendInvitations)
                    {
                        System.Diagnostics.Debug.WriteLine("[CreateMeeting] Attempting to Send appointment...");
                        appointment.Send();
                        System.Diagnostics.Debug.WriteLine("[CreateMeeting] Appointment.Send() called.");
                        response.Message = "Meeting invitation sent.";
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("[CreateMeeting] Attempting to Save appointment...");
                        appointment.Save();
                        System.Diagnostics.Debug.WriteLine($"[CreateMeeting] Appointment.Save() called. Item IsSaved: {appointment.Saved}"); // CHECK THIS
                        response.Message = "Meeting saved to calendar.";
                    }

                    // Try to get EntryID. It might still be empty if Save/Send didn't truly persist it.
                    try
                    {
                        response.MeetingId = appointment.EntryID;
                        if (string.IsNullOrEmpty(response.MeetingId))
                        {
                            System.Diagnostics.Debug.WriteLine("[CreateMeeting] CRITICAL: EntryID is STILL EMPTY after Save/Send call.");
                            if (response.Status != "failure") // If not already marked as failure
                            {
                                response.Errors.Add(new McpError { Code = "PersistenceError", Message = "Meeting was processed but could not be persisted or EntryID not obtained." });
                                // Do not override status if it was already "failure" due to unresolved recipients
                                if (response.Status == "success" || response.Status == "partial_success") response.Status = "failure";
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[CreateMeeting] MeetingId obtained: {response.MeetingId}");
                        }
                    }
                    catch (COMException exEntryId)
                    {
                        System.Diagnostics.Debug.WriteLine($"[CreateMeeting] COMException when trying to get EntryID: {exEntryId.ToString()}");
                        response.Errors.Add(new McpError { Code = "EntryIDError", Message = $"Error getting EntryID: {exEntryId.Message}" });
                        if (response.Status != "failure") response.Status = "failure";
                    }

                    if (response.Status != "failure") // If not already set to failure
                    {
                        response.Status = allResolved ? "success" : "partial_success";
                    }
                }
            }
            catch (COMException comEx)
            {
                System.Diagnostics.Debug.WriteLine($"[CreateMeeting] Outer COMException: {comEx.ToString()}");
                response.Errors.Add(new McpError { Code = "OutlookComError", Message = $"COM Error: {comEx.Message} (ErrorCode: {comEx.ErrorCode})" });
                response.Status = "failure";
                response.Message = $"Failed to create meeting due to COM error: {comEx.Message}";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[CreateMeeting] Outer General Exception: {ex.ToString()}");
                response.Errors.Add(new McpError { Code = "CreateMeetingError", Message = ex.Message });
                response.Status = "failure";
                response.Message = $"Failed to create meeting: {ex.Message}";
            }
            finally
            {
                if (appointment != null)
                {
                    System.Diagnostics.Debug.WriteLine("[CreateMeeting] Releasing AppointmentItem COM object.");
                    Marshal.ReleaseComObject(appointment);
                }
            }
            System.Diagnostics.Debug.WriteLine($"[CreateMeeting] Exiting. Final Status: {response.Status}, Errors: {response.Errors.Count}, MeetingId: {response.MeetingId}");
            return response;
        }
        // --- Helper methods for building DASL filters ---
        private string BuildEmailFilter(string query, DateTime? start, DateTime? end)
        {
            var filters = new List<string>();
            if (!string.IsNullOrWhiteSpace(query))
            {
                // For DASL, use ci_phrasematch for case-insensitive phrase matching or ci_startswith / ci_contains
                // However, LIKE is often sufficient and simpler.
                filters.Add($"(\"urn:schemas:httpmail:subject\" LIKE '%{EscapeDASLString(query)}%' OR \"urn:schemas:httpmail:textdescription\" LIKE '%{EscapeDASLString(query)}%')");
            }
            if (start.HasValue) filters.Add($"\"urn:schemas:httpmail:datereceived\" >= '{start.Value.ToUniversalTime():yyyy/MM/dd HH:mm}'");
            if (end.HasValue) filters.Add($"\"urn:schemas:httpmail:datereceived\" <= '{end.Value.ToUniversalTime():yyyy/MM/dd HH:mm}'");
            return string.Join(" AND ", filters.Where(f => !string.IsNullOrEmpty(f)));
        }


        private string BuildCalendarFilter(string query, DateTime? start, DateTime? end)
        {
            var filters = new List<string>();
            // For calendar items, it's often better to filter broadly by date range and then refine in code if recurrences are complex.
            // Outlook's Restrict on calendar items can be tricky with recurrences spanning the filter boundary.
            // The filter below targets items that *overlap* the given range.
            DateTime effectiveStart = start ?? DateTime.UtcNow.Date.AddDays(-7); // Default to a week ago if no start
            DateTime effectiveEnd = end ?? DateTime.UtcNow.Date.AddDays(7);   // Default to a week from now if no end

            // This filter finds appointments that *overlap* the range [effectiveStart, effectiveEnd]
            // An appointment (ApptStart, ApptEnd) overlaps [RangeStart, RangeEnd] if:
            // ApptStart < RangeEnd AND ApptEnd > RangeStart
            filters.Add($"(\"[Start]\" < '{effectiveEnd.ToLocalTime().ToString("yyyy/MM/dd HH:mm")}' AND \"[End]\" > '{effectiveStart.ToLocalTime().ToString("yyyy/MM/dd HH:mm")}')");


            if (!string.IsNullOrWhiteSpace(query))
            {
                filters.Add($"(\"[Subject]\" LIKE '%{EscapeDASLString(query)}%' OR \"[Location]\" LIKE '%{EscapeDASLString(query)}%')");
            }
            return string.Join(" AND ", filters.Where(f => !string.IsNullOrEmpty(f)));
        }

        private string EscapeDASLString(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;
            // DASL uses single quotes for string literals. To include a single quote in the string, double it.
            // Percent signs are wildcards with LIKE, so if you want to search for a literal '%', it needs special handling
            // or use a different operator if not intending wildcard search. For LIKE, escaping '%' is often not needed
            // unless you want to match a literal percent. For simplicity, just escaping quotes.
            return value.Replace("'", "''");
        }
    }
}