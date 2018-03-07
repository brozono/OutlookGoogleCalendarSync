﻿using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.OutlookOgcs {
    /// <summary>
    /// Description of OutlookOgcs.Calendar.
    /// </summary>
    public class Calendar {
        private static Calendar instance;
        private static readonly ILog log = LogManager.GetLogger(typeof(Calendar));
        public Interface IOutlook;

        /// <summary>
        /// Whether instance of OutlookCalendar class should connect to Outlook application
        /// </summary>
        public static Boolean InstanceConnect = true;
        public static Boolean IsInstanceNull { get { return instance == null; } }
        public static Calendar Instance {
            get {
                try {
                    if (instance == null || instance.Folders == null) instance = new Calendar();
                } catch (System.ApplicationException ex) {
                    throw ex;
                } catch (System.Exception ex) {
                    OGCSexception.Analyse(ex);
                    log.Info("It appears Outlook has been restarted after OGCS was started. Reconnecting...");
                    instance = new Calendar();
                }
                return instance;
            }
        }
        public static Boolean OOMsecurityInfo = false;
        public const String GlobalIdPattern = "040000008200E00074C5B7101A82E008";
        public Sync.PushSyncTimer OgcsPushTimer;
        public MAPIFolder UseOutlookCalendar {
            get { return IOutlook.UseOutlookCalendar(); }
            set {
                IOutlook.UseOutlookCalendar(value);
                Settings.Instance.UseOutlookCalendar = new OutlookCalendarListEntry(value);
            }
        }
        public Folders Folders {
            get { return IOutlook.Folders(); }
        }
        public Dictionary<string, MAPIFolder> CalendarFolders {
            get { return IOutlook.CalendarFolders(); }
        }
        public enum Service {
            DefaultMailbox,
            AlternativeMailbox,
            SharedCalendar
        }

        public Calendar() {
            IOutlook = Factory.GetOutlookInterface();
            if (InstanceConnect) IOutlook.Connect();
        }

        public void Reset() {
            log.Info("Resetting connection to Outlook.");
            if (IOutlook != null) IOutlook.Disconnect();
            instance = new Calendar();
        }

        #region Push Sync
        //Multi-threaded, so need to protect against registering events more than once
        //Simply removing an event handler before adding isn't safe enough
        private int eventHandlerHooks = 0;

        public void RegisterForPushSync() {
            log.Info("Registering for Outlook appointment change events...");
            if (eventHandlerHooks != 0) purgeOutlookEventHandlers();

            if (Settings.Instance.SyncDirection != Sync.Direction.GoogleToOutlook) {
                log.Debug("Create the timer for the push synchronisation");
                if (OgcsPushTimer == null)
                    OgcsPushTimer = new Sync.PushSyncTimer();
                if (!OgcsPushTimer.Running())
                    OgcsPushTimer.Switch(true);

                UseOutlookCalendar.Items.ItemAdd += new ItemsEvents_ItemAddEventHandler(appointmentItem_Add);
                UseOutlookCalendar.Items.ItemChange += new ItemsEvents_ItemChangeEventHandler(appointmentItem_Change);
                UseOutlookCalendar.Items.ItemRemove += new ItemsEvents_ItemRemoveEventHandler(appointmentItem_Remove);
                eventHandlerHooks++;
            }
        }

        public void DeregisterForPushSync() {
            log.Info("Deregistering from Outlook appointment change events...");
            purgeOutlookEventHandlers();
            if (OgcsPushTimer != null && OgcsPushTimer.Running())
                OgcsPushTimer.Switch(false);
        }

        private void purgeOutlookEventHandlers() {
            log.Debug("Removing " + eventHandlerHooks + " Outlook event handler hooks.");
            while (eventHandlerHooks > 0) {
                try { UseOutlookCalendar.Items.ItemAdd -= new ItemsEvents_ItemAddEventHandler(appointmentItem_Add); } catch { }
                try { UseOutlookCalendar.Items.ItemChange -= new ItemsEvents_ItemChangeEventHandler(appointmentItem_Change); } catch { }
                try { UseOutlookCalendar.Items.ItemRemove -= new ItemsEvents_ItemRemoveEventHandler(appointmentItem_Remove); } catch { }
                eventHandlerHooks--;
            }
        }

        private void appointmentItem_Add(object Item) {
            if (Settings.Instance.SyncDirection == Sync.Direction.GoogleToOutlook) return;

            AppointmentItem ai = null;
            try {
                log.Debug("Detected Outlook item added.");
                ai = Item as AppointmentItem;

                DateTime syncMin = DateTime.Today.AddDays(-Settings.Instance.DaysInThePast);
                DateTime syncMax = DateTime.Today.AddDays(+Settings.Instance.DaysInTheFuture + 1);
                if (ai.Start < syncMax && ai.End >= syncMin) {
                    log.Debug(GetEventSummary(ai));
                    log.Debug("Item is in sync range, so push sync flagged for Go.");
                    OgcsPushTimer.ItemsQueued++;
                    log.Info(OgcsPushTimer.ItemsQueued + " items changed since last sync.");
                } else {
                    log.Fine("Item is outside of sync range.");
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            } finally {
                ai = (AppointmentItem)ReleaseObject(ai);
            }
        }
        private void appointmentItem_Change(object Item) {
            if (Settings.Instance.SyncDirection == Sync.Direction.GoogleToOutlook) return;

            AppointmentItem ai = null;
            try {
                log.Debug("Detected Outlook item changed.");
                ai = Item as AppointmentItem;

                DateTime syncMin = DateTime.Today.AddDays(-Settings.Instance.DaysInThePast);
                DateTime syncMax = DateTime.Today.AddDays(+Settings.Instance.DaysInTheFuture + 1);
                if (ai.Start < syncMax && ai.End >= syncMin) {
                    log.Debug(GetEventSummary(ai));
                    log.Debug("Item is in sync range, so push sync flagged for Go.");
                    OgcsPushTimer.ItemsQueued++;
                    log.Info(OgcsPushTimer.ItemsQueued + " items changed since last sync.");
                } else {
                    log.Fine("Item is outside of sync range.");
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            } finally {
                ai = (AppointmentItem)ReleaseObject(ai);
            }
        }
        private void appointmentItem_Remove() {
            if (Settings.Instance.SyncDirection == Sync.Direction.GoogleToOutlook) return;

            try {
                log.Debug("Detected Outlook item removed, so push sync flagged for Go.");
                OgcsPushTimer.ItemsQueued++;
                log.Info(OgcsPushTimer.ItemsQueued + " items changed since last sync.");
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }
        #endregion

        public List<AppointmentItem> GetCalendarEntriesInRange() {
            List<AppointmentItem> filtered = new List<AppointmentItem>();
            filtered = FilterCalendarEntries(UseOutlookCalendar.Items);
            return filtered;
        }

        public List<AppointmentItem> FilterCalendarEntries(Items OutlookItems, Boolean filterCategories = true, Boolean noDateFilter = false, String extraFilter = "") {
            //Filtering info @ https://msdn.microsoft.com/en-us/library/cc513841%28v=office.12%29.aspx

            List<AppointmentItem> result = new List<AppointmentItem>();
            if (OutlookItems != null) {
                log.Fine(OutlookItems.Count + " calendar items exist.");

                //OutlookItems.Sort("[Start]", Type.Missing);
                OutlookItems.IncludeRecurrences = false;

                if (!Settings.Instance.EnableUseRecurrence) {
                    OutlookItems.Sort("[Start]", Type.Missing);
                    OutlookItems.IncludeRecurrences = true;
                }

                DateTime min = DateTime.MinValue;
                DateTime max = DateTime.MaxValue;
                if (!noDateFilter) {
                    min = Settings.Instance.SyncStart;
                    max = Settings.Instance.SyncEnd;
                }

                string filter = "[End] >= '" + min.ToString(Settings.Instance.OutlookDateFormat) +
                    "' AND [Start] < '" + max.ToString(Settings.Instance.OutlookDateFormat) + "'" + extraFilter;
                log.Fine("Filter string: " + filter);
                Int32 categoryFiltered = 0;
                foreach (Object obj in OutlookItems.Restrict(filter)) {
                    AppointmentItem ai;
                    try {
                        ai = obj as AppointmentItem;
                    } catch {
                        log.Warn("Encountered a non-appointment item in the calendar.");
                        if (obj is MeetingItem) log.Debug("It is a meeting item.");
                        else if (obj is MailItem) log.Debug("It is a mail item.");
                        else if (obj is ContactItem) log.Debug("It is a contact item.");
                        else if (obj is TaskItem) log.Debug("It is a task item.");
                        else log.Debug("WTF is this item?!");
                        continue;
                    }
                    try {
                        if (ai.End == min) continue; //Required for midnight to midnight events 
                    } catch (System.Exception ex) {
                        OGCSexception.Analyse(ex, true);
                        try {
                            log.Debug("Unable to get End date for: " + OutlookOgcs.Calendar.GetEventSummary(ai));
                        } catch {
                            log.Error("Appointment item seems unusable!");
                        }
                        continue;
                    }
                    if (filterCategories) {
                        String categoryDelimiter = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator + " ";
                        if (Settings.Instance.CategoriesRestrictBy == Settings.RestrictBy.Include) {
                            if (Settings.Instance.Categories.Count() > 0 && ((ai.Categories == null && Settings.Instance.Categories.Contains("<No category assigned>")) ||
                                (ai.Categories != null && ai.Categories.Split(new[] { categoryDelimiter }, StringSplitOptions.None).Intersect(Settings.Instance.Categories).Count() > 0)))
                            {
                                result.Add(ai);
                            } else categoryFiltered++;

                        } else if (Settings.Instance.CategoriesRestrictBy == Settings.RestrictBy.Exclude) {
                            if (Settings.Instance.Categories.Count() == 0 || (ai.Categories == null && !Settings.Instance.Categories.Contains("<No category assigned>")) ||
                                (ai.Categories != null && ai.Categories.Split(new[] { categoryDelimiter }, StringSplitOptions.None).Intersect(Settings.Instance.Categories).Count() == 0))
                            {
                                result.Add(ai);
                            } else categoryFiltered++;
                        }
                    } else {
                        result.Add(ai);
                    }
                }
                if (categoryFiltered > 0) {
                    log.Info(categoryFiltered + " Outlook items excluded due to active category filter.");
                    if (result.Count == 0)
                        Forms.Main.Instance.Console.Update("Due to your category settings, all Outlook items have been filtered out!", Console.Markup.warning, notifyBubble: true);
                }
            }
            log.Fine("Filtered down to " + result.Count);
            return result;
        }

        #region Create
        public void CreateCalendarEntries(List<Event> events) {
            for (int g = 0; g < events.Count; g++) {
                if (Sync.Engine.Instance.CancellationPending) return;

                Event ev = events[g];
                AppointmentItem newAi = IOutlook.UseOutlookCalendar().Items.Add() as AppointmentItem;
                try {
                    try {
                        createCalendarEntry(ev, ref newAi);
                    } catch (System.Exception ex) {
                        String evSummary = "";
                        if (!Settings.Instance.VerboseOutput) evSummary = GoogleOgcs.Calendar.GetEventSummary(ev) + "<br/>";
                        if (ex.GetType() == typeof(ApplicationException)) {
                            Forms.Main.Instance.Console.Update(evSummary + "Appointment creation skipped.", Console.Markup.warning);
                            continue;
                        } else {
                            Forms.Main.Instance.Console.Update(evSummary + "Appointment creation failed.<br/>" + ex.Message, Console.Markup.error);
                            log.Error(ex.StackTrace);
                            if (CalMessageBox.Instance.ShowTrue("Outlook appointment creation failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question, DialogResult.Yes))
                                continue;
                            else
                                throw new UserCancelledSyncException("User chose not to continue sync.");
                        }
                    }

                    try {
                        createCalendarEntry_save(newAi, ref ev);
                        events[g] = ev;
                    } catch (System.Exception ex) {
                        String evSummary = "";
                        if (!Settings.Instance.VerboseOutput) evSummary = GoogleOgcs.Calendar.GetEventSummary(ev) + "<br/>";
                        Forms.Main.Instance.Console.Update(evSummary + "New appointment failed to save.<br/>" + ex.Message, Console.Markup.error);
                        log.Error(ex.StackTrace);
                        if (CalMessageBox.Instance.ShowTrue("New Outlook appointment failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question, DialogResult.Yes))
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }

                    if (ev.Recurrence != null && ev.RecurringEventId == null && Recurrence.Instance.HasExceptions(ev)) {
                        Forms.Main.Instance.Console.Update("This is a recurring item with some exceptions:-", verbose: true);
                        Recurrence.Instance.CreateOutlookExceptions(ref newAi, ev);
                        Forms.Main.Instance.Console.Update("Recurring exceptions completed.", verbose: true);
                    }
                } finally {
                    newAi = (AppointmentItem)ReleaseObject(newAi);
                }
            }
        }

        private void createCalendarEntry(Event ev, ref AppointmentItem ai) {
            string itemSummary = GoogleOgcs.Calendar.GetEventSummary(ev);
            log.Debug("Processing >> " + itemSummary);
            Forms.Main.Instance.Console.Update(itemSummary, Console.Markup.calendar, verbose: true);

            //Add the Google event IDs into Outlook appointment.
            AddGoogleIDs(ref ai, ev);

            ai.Start = new DateTime();
            ai.End = new DateTime();
            ai.AllDayEvent = (ev.Start.Date != null);
            ai = OutlookOgcs.Calendar.Instance.IOutlook.WindowsTimeZone_set(ai, ev);
            Recurrence.Instance.BuildOutlookPattern(ev, ai);

            ai.Subject = Obfuscate.ApplyRegex(ev.Summary, Sync.Direction.GoogleToOutlook);
            if (Settings.Instance.AddDescription && ev.Description != null) ai.Body = ev.Description;
            ai.Location = ev.Location;
            ai.Sensitivity = getPrivacy(ev.Visibility, null);
            ai.BusyStatus = getAvailability(ev.Transparency, null);

            if (Settings.Instance.AddAttendees && ev.Attendees != null) {
                foreach (EventAttendee ea in ev.Attendees) {
                    Recipients recipients = ai.Recipients;
                    createRecipient(ea, ref recipients);
                    recipients = (Recipients)ReleaseObject(recipients);
                }
            }

            //Reminder alert
            if (Settings.Instance.AddReminders && ev.Reminders != null && ev.Reminders.Overrides != null) {
                foreach (EventReminder reminder in ev.Reminders.Overrides) {
                    if (reminder.Method == "popup") {
                        ai.ReminderSet = true;
                        ai.ReminderMinutesBeforeStart = (int)reminder.Minutes;
                    }
                }
            }
        }

        private static void createCalendarEntry_save(AppointmentItem ai, ref Event ev) {
            if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional) {
                log.Debug("Saving timestamp when OGCS updated appointment.");
                setOGCSlastModified(ref ai);
            }

            ai.Save();

            if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional || GoogleOgcs.Calendar.HasOgcsProperty(ev)) {
                log.Debug("Storing the Outlook appointment IDs in Google event.");
                GoogleOgcs.Calendar.AddOutlookIDs(ref ev, ai);
                GoogleOgcs.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
            }
        }
        #endregion

        #region Update
        public void UpdateCalendarEntries(Dictionary<AppointmentItem, Event> entriesToBeCompared, ref int entriesUpdated) {
            entriesUpdated = 0;
            foreach (KeyValuePair<AppointmentItem, Event> compare in entriesToBeCompared) {
                if (Sync.Engine.Instance.CancellationPending) return;

                int itemModified = 0;
                AppointmentItem ai = compare.Key;
                try {
                    Boolean aiWasRecurring = ai.IsRecurring;
                    Boolean needsUpdating = false;
                    try {
                        needsUpdating = UpdateCalendarEntry(ref ai, compare.Value, ref itemModified);
                    } catch (System.Exception ex) {
                        String evSummary = "";
                        if (!Settings.Instance.VerboseOutput) evSummary = GoogleOgcs.Calendar.GetEventSummary(compare.Value) + "<br/>";
                        Forms.Main.Instance.Console.Update(evSummary + "Appointment update failed.<br/>" + ex.Message, Console.Markup.error);
                        log.Error(ex.StackTrace);
                        if (CalMessageBox.Instance.ShowTrue("Outlook appointment update failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question, DialogResult.Yes))
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }

                    if (itemModified > 0) {
                        try {
                            updateCalendarEntry_save(ref ai);
                            entriesUpdated++;
                        } catch (System.Exception ex) {
                            String evSummary = "";
                            if (!Settings.Instance.VerboseOutput) evSummary = GoogleOgcs.Calendar.GetEventSummary(compare.Value) + "<br/>";
                            Forms.Main.Instance.Console.Update(evSummary + "Updated appointment failed to save.<br/>" + ex.Message, Console.Markup.error);
                            log.Error(ex.StackTrace);
                            if (CalMessageBox.Instance.ShowTrue("Updated Outlook appointment failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question, DialogResult.Yes))
                                continue;
                            else
                                throw new UserCancelledSyncException("User chose not to continue sync.");
                        }
                        if (ai.IsRecurring) {
                            if (!aiWasRecurring) log.Debug("Appointment has changed from single instance to recurring.");
                            log.Debug("Recurring master appointment has been updated, so now checking if exceptions need reinstating.");
                            Recurrence.Instance.UpdateOutlookExceptions(ref ai, compare.Value, forceCompare: true);
                        }

                    } else {
                        if (ai.RecurrenceState == OlRecurrenceState.olApptMaster && compare.Value.Recurrence != null && compare.Value.RecurringEventId == null) {
                            log.Debug(GoogleOgcs.Calendar.GetEventSummary(compare.Value));
                            Recurrence.Instance.UpdateOutlookExceptions(ref ai, compare.Value, forceCompare: false);

                        } else if (needsUpdating || GetOGCSproperty(ai, MetadataId.forceSave)) {
                            log.Debug("Doing a dummy update in order to update the last modified date.");
                            setOGCSlastModified(ref ai);
                            updateCalendarEntry_save(ref ai);
                        }
                    }
                } finally {
                    ai = (AppointmentItem)ReleaseObject(ai);
                }
            }
        }

        public Boolean UpdateCalendarEntry(ref AppointmentItem ai, Event ev, ref int itemModified, Boolean forceCompare = false) {
            if (!(Sync.Engine.Instance.ManualForceCompare || forceCompare)) { //Needed if the exception has just been created, but now needs updating
                if (Settings.Instance.SyncDirection != Sync.Direction.Bidirectional) {
                    if (ai.LastModificationTime > ev.Updated)
                        return false;
                } else {
                    if (GoogleOgcs.Calendar.GetOGCSlastModified(ev).AddSeconds(5) >= ev.Updated)
                        //Google last modified by OGCS
                        return false;
                    if (ai.LastModificationTime > ev.Updated)
                        return false;
                }
            }

            if (ai.RecurrenceState == OlRecurrenceState.olApptMaster)
                log.Debug("Processing recurring master appointment.");
            
            String evSummary = GoogleOgcs.Calendar.GetEventSummary(ev);
            log.Debug("Processing >> " + evSummary);

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine(evSummary);

            if (ai.RecurrenceState != OlRecurrenceState.olApptMaster) {
                if (ai.AllDayEvent != (ev.Start.DateTime == null)) {
                    sb.AppendLine("All-Day: " + ai.AllDayEvent + " => " + (ev.Start.DateTime == null));
                    ai.AllDayEvent = (ev.Start.DateTime == null);
                    itemModified++;
                }
            }

            #region TimeZone
            String currentStartTZ = "UTC";
            String currentEndTZ = "UTC";
            String newStartTZ = "UTC";
            String newEndTZ = "UTC";
            IOutlook.WindowsTimeZone_get(ai, out currentStartTZ, out currentEndTZ);
            ai = OutlookOgcs.Calendar.Instance.IOutlook.WindowsTimeZone_set(ai, ev, onlyTZattribute: true);
            IOutlook.WindowsTimeZone_get(ai, out newStartTZ, out newEndTZ);
            Boolean startTzChange = Sync.Engine.CompareAttribute("Start Timezone", Sync.Direction.GoogleToOutlook, newStartTZ, currentStartTZ, sb, ref itemModified);
            Boolean endTzChange = Sync.Engine.CompareAttribute("End Timezone", Sync.Direction.GoogleToOutlook, newEndTZ, currentEndTZ, sb, ref itemModified);
            #endregion

            #region Start/End & Recurrence
            DateTime evStartParsedDate = ev.Start.DateTime ?? DateTime.Parse(ev.Start.Date);
            Boolean startChange = Sync.Engine.CompareAttribute("Start time", Sync.Direction.GoogleToOutlook, evStartParsedDate, ai.Start, sb, ref itemModified);

            DateTime evEndParsedDate = ev.End.DateTime ?? DateTime.Parse(ev.End.Date);
            Boolean endChange = Sync.Engine.CompareAttribute("End time", Sync.Direction.GoogleToOutlook, evEndParsedDate, ai.End, sb, ref itemModified);

            RecurrencePattern oPattern = null;
            try {
                if (startChange || endChange || startTzChange || endTzChange) {
                    if (ai.RecurrenceState == OlRecurrenceState.olApptMaster) {
                        if (startTzChange || endTzChange) {
                            oPattern = (RecurrencePattern)OutlookOgcs.Calendar.ReleaseObject(oPattern);
                            ai.ClearRecurrencePattern();
                            ai = OutlookOgcs.Calendar.Instance.IOutlook.WindowsTimeZone_set(ai, ev, onlyTZattribute: false);
                            ai.Save();
                            Recurrence.Instance.BuildOutlookPattern(ev, ai);
                            ai.Save(); //Explicit save required to make ai.IsRecurring true again
                        } else {
                            oPattern = (ai.RecurrenceState == OlRecurrenceState.olApptMaster) ? ai.GetRecurrencePattern() : null;
                            if (startChange) {
                                oPattern.PatternStartDate = evStartParsedDate;
                                oPattern.StartTime = TimeZoneInfo.ConvertTime(evStartParsedDate, TimeZoneInfo.FindSystemTimeZoneById(newStartTZ));
                            }
                            if (endChange) {
                                oPattern.PatternEndDate = evEndParsedDate;
                                oPattern.EndTime = TimeZoneInfo.ConvertTime(evEndParsedDate, TimeZoneInfo.FindSystemTimeZoneById(newEndTZ));
                            }
                        }
                    } else {
                        ai = OutlookOgcs.Calendar.Instance.IOutlook.WindowsTimeZone_set(ai, ev);
                    }
                }

                if (oPattern == null)
                    oPattern = (ai.RecurrenceState == OlRecurrenceState.olApptMaster) ? ai.GetRecurrencePattern() : null;
                if (oPattern != null) {
                    oPattern.Duration = Convert.ToInt32((evEndParsedDate - evStartParsedDate).TotalMinutes);
                    Recurrence.Instance.CompareOutlookPattern(ev, ref oPattern, Sync.Direction.GoogleToOutlook, sb, ref itemModified);
                }
            } finally {
                oPattern = (RecurrencePattern)ReleaseObject(oPattern);
            }

            if (ai.RecurrenceState == OlRecurrenceState.olApptMaster) {
                if (ev.Recurrence == null || ev.RecurringEventId != null) {
                    log.Debug("Converting to non-recurring events.");
                    ai.ClearRecurrencePattern();
                    itemModified++;
                }
            } else if (ai.RecurrenceState == OlRecurrenceState.olApptNotRecurring) {
                if (!ai.IsRecurring && ev.Recurrence != null && ev.RecurringEventId == null) {
                    log.Debug("Converting to recurring appointment.");
                    Recurrence.Instance.BuildOutlookPattern(ev, ai);
                    Recurrence.Instance.CreateOutlookExceptions(ref ai, ev);
                    itemModified++;
                }
            }
            #endregion

            String summaryObfuscated = Obfuscate.ApplyRegex(ev.Summary, Sync.Direction.GoogleToOutlook);
            if (Sync.Engine.CompareAttribute("Subject", Sync.Direction.GoogleToOutlook, summaryObfuscated, ai.Subject, sb, ref itemModified)) {
                ai.Subject = summaryObfuscated;
            }
            if (Settings.Instance.AddDescription) {
                if (Settings.Instance.SyncDirection == Sync.Direction.GoogleToOutlook || !Settings.Instance.AddDescription_OnlyToGoogle) {
                    if (Sync.Engine.CompareAttribute("Description", Sync.Direction.GoogleToOutlook, ev.Description, ai.Body, sb, ref itemModified))
                        ai.Body = ev.Description;
                }
            }

            if (Sync.Engine.CompareAttribute("Location", Sync.Direction.GoogleToOutlook, ev.Location, ai.Location, sb, ref itemModified))
                ai.Location = ev.Location;

            if (ai.RecurrenceState == OlRecurrenceState.olApptMaster ||
                ai.RecurrenceState == OlRecurrenceState.olApptNotRecurring) 
            {
                OlSensitivity gPrivacy = getPrivacy(ev.Visibility, ai.Sensitivity);
                if (Sync.Engine.CompareAttribute("Privacy", Sync.Direction.GoogleToOutlook, gPrivacy.ToString(), ai.Sensitivity.ToString(), sb, ref itemModified)) {
                    ai.Sensitivity = gPrivacy;
                }
            }
            OlBusyStatus gFreeBusy = getAvailability(ev.Transparency ?? "opaque", ai.BusyStatus);
            if (Sync.Engine.CompareAttribute("Free/Busy", Sync.Direction.GoogleToOutlook, gFreeBusy.ToString(), ai.BusyStatus.ToString(), sb, ref itemModified)) {
                ai.BusyStatus = gFreeBusy;
            }

            if (Settings.Instance.AddAttendees) {
                log.Fine("Comparing meeting attendees");
                Recipients recipients = ai.Recipients;
                List<EventAttendee> addAttendees = new List<EventAttendee>();
                try {
                    if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional &&
                        ev.Attendees != null && ev.Attendees.Count == 0 && recipients.Count > 150) {
                        log.Info("Attendees not being synced - there are too many (" + recipients.Count + ") for Google.");
                    } else {
                        //Build a list of Google attendees. Any remaining at the end of the diff must be added.
                        if (ev.Attendees != null) {
                            addAttendees = ev.Attendees.ToList();
                        }
                        for (int r = 1; r <= recipients.Count; r++) {
                            Recipient recipient = null;
                            Boolean foundAttendee = false;
                            try {
                                recipient = recipients[r];
                                if (recipient.Name == ai.Organizer) continue;

                                if (!recipient.Resolved) recipient.Resolve();
                                String recipientSMTP = IOutlook.GetRecipientEmail(recipient);

                                for (int g = (ev.Attendees == null ? -1 : ev.Attendees.Count - 1); g >= 0; g--) {
                                    GoogleOgcs.EventAttendee attendee = new GoogleOgcs.EventAttendee(ev.Attendees[g]);
                                    if (recipientSMTP.ToLower() == attendee.Email.ToLower()) {
                                        foundAttendee = true;

                                        //Optional attendee
                                        bool oOptional = (ai.OptionalAttendees != null && ai.OptionalAttendees.Contains(attendee.DisplayName ?? attendee.Email));
                                        bool gOptional = (attendee.Optional == null) ? false : (bool)attendee.Optional;
                                        if (Sync.Engine.CompareAttribute("Recipient " + recipient.Name + " - Optional Check",
                                            Sync.Direction.GoogleToOutlook, gOptional, oOptional, sb, ref itemModified)) {
                                            if (gOptional) {
                                                recipient.Type = (int)OlMeetingRecipientType.olOptional;
                                            } else {
                                                recipient.Type = (int)OlMeetingRecipientType.olRequired;
                                            }
                                        }
                                        //Response is readonly in Outlook :(
                                        addAttendees.Remove(ev.Attendees[g]);
                                        break;
                                    }
                                }
                                if (!foundAttendee) {
                                    sb.AppendLine("Recipient removed: " + recipient.Name);
                                    recipient.Delete();
                                    itemModified++;
                                }
                            } finally {
                                recipient = (Recipient)OutlookOgcs.Calendar.ReleaseObject(recipient);
                            }
                        }
                        foreach (EventAttendee gAttendee in addAttendees) {
                            GoogleOgcs.EventAttendee attendee = new GoogleOgcs.EventAttendee(gAttendee);
                            if (attendee.DisplayName == ai.Organizer) continue; //Attendee in Google is owner in Outlook, so can't also be added as a recipient)

                            sb.AppendLine("Recipient added: " + (attendee.DisplayName ?? attendee.Email));
                            createRecipient(attendee, ref recipients);
                            itemModified++;
                        }
                    }
                } finally {
                    recipients = (Recipients)OutlookOgcs.Calendar.ReleaseObject(recipients);
                }
            }

            //Reminders
            if (Settings.Instance.AddReminders) {
                if (ev.Reminders.Overrides != null) {
                    //Find the popup reminder in Google
                    for (int r = ev.Reminders.Overrides.Count - 1; r >= 0; r--) {
                        EventReminder reminder = ev.Reminders.Overrides[r];
                        if (reminder.Method == "popup") {
                            if (ai.ReminderSet) {
                                if (Sync.Engine.CompareAttribute("Reminder", Sync.Direction.GoogleToOutlook, reminder.Minutes.ToString(), ai.ReminderMinutesBeforeStart.ToString(), sb, ref itemModified)) {
                                    ai.ReminderMinutesBeforeStart = (int)reminder.Minutes;
                                }
                            } else {
                                sb.AppendLine("Reminder: nothing => " + reminder.Minutes);
                                ai.ReminderSet = true;
                                ai.ReminderMinutesBeforeStart = (int)reminder.Minutes;
                                itemModified++;
                            } //if Outlook reminders set
                        } //if google reminder found
                    } //foreach reminder

                } else { //no google reminders set
                    if (ai.ReminderSet && IsOKtoSyncReminder(ai)) {
                        sb.AppendLine("Reminder: " + ai.ReminderMinutesBeforeStart + " => removed");
                        ai.ReminderSet = false;
                        itemModified++;
                    }
                }
            }

            if (itemModified > 0) {
                Forms.Main.Instance.Console.FormatEventChanges(sb);
                Forms.Main.Instance.Console.Update(itemModified + " attributes updated.", Console.Markup.appointmentEnd, verbose: true, newLine: false);
                System.Windows.Forms.Application.DoEvents();
            }
            return true;
        }

        private void updateCalendarEntry_save(ref AppointmentItem ai) {
            if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional) {
                log.Debug("Saving timestamp when OGCS updated appointment.");
                setOGCSlastModified(ref ai);
            }
            removeOGCSproperty(ref ai, MetadataId.forceSave);
            ai.Save();
        }
        #endregion

        #region Delete
        public void DeleteCalendarEntries(List<AppointmentItem> oAppointments) {
            for (int o = oAppointments.Count - 1; o >= 0; o--) {
                if (Sync.Engine.Instance.CancellationPending) return;

                AppointmentItem ai = oAppointments[o];
                Boolean doDelete = false;
                try {
                    try {
                        doDelete = deleteCalendarEntry(ai);
                    } catch (System.Exception ex) {
                        String evSummary = "";
                        if (!Settings.Instance.VerboseOutput) evSummary = OutlookOgcs.Calendar.GetEventSummary(ai) + "<br/>";
                        Forms.Main.Instance.Console.Update(evSummary + "Appointment deletion failed.<br/>" + ex.Message, Console.Markup.error);
                        log.Error(ex.StackTrace);
                        if (MessageBox.Show("Outlook appointment deletion failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }

                    try {
                        if (doDelete) deleteCalendarEntry_save(ai);
                        else oAppointments.Remove(ai);
                    } catch (System.Exception ex) {
                        String evSummary = "";
                        if (!Settings.Instance.VerboseOutput) evSummary = OutlookOgcs.Calendar.GetEventSummary(ai) + "<br/>";
                        Forms.Main.Instance.Console.Update(evSummary + "Deleted appointment failed to remove.<br/>" + ex.Message, Console.Markup.error);
                        log.Error(ex.StackTrace);
                        if (CalMessageBox.Instance.ShowTrue("Deleted Outlook appointment failed to remove. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question, DialogResult.Yes))
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                } finally {
                    ai = (AppointmentItem)ReleaseObject(ai);
                }
            }
        }

        private Boolean deleteCalendarEntry(AppointmentItem ai) {
            String eventSummary = GetEventSummary(ai);
            Boolean doDelete = true;

            if (Settings.Instance.ConfirmOnDelete) {
                if (MessageBox.Show("Delete " + eventSummary + "?", "Deletion Confirmation",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) {
                    doDelete = false;
                    Forms.Main.Instance.Console.Update("Not deleted: " + eventSummary, Console.Markup.calendar);
                } else {
                    Forms.Main.Instance.Console.Update("Deleted: " + eventSummary, Console.Markup.calendar);
                }
            } else {
                Forms.Main.Instance.Console.Update(eventSummary, Console.Markup.calendar, verbose: true);
            }
            return doDelete;
        }

        private void deleteCalendarEntry_save(AppointmentItem ai) {
            ai.Delete();
        }
        #endregion

        public void ReclaimOrphanCalendarEntries(ref List<AppointmentItem> oAppointments, ref List<Event> gEvents) {
            log.Debug("Scanning " + oAppointments.Count + " Outlook appointments for orphans to reclaim...");

            //This is needed for people migrating from other tools, which do not have our GoogleID extendedProperty
            List<AppointmentItem> unclaimedAi = new List<AppointmentItem>();

            for (int o = oAppointments.Count - 1; o >= 0; o--) {
                AppointmentItem ai = oAppointments[o];
                String sigAi = signature(ai);

                //Find entries with no Google ID
                if (!GetOGCSproperty(ai, MetadataId.gEventID)) {
                    unclaimedAi.Add(ai);

                    for (int g = gEvents.Count - 1; g >= 0; g--) {
                        Event ev = gEvents[g];
                        String sigEv = GoogleOgcs.Calendar.signature(ev);
                        if (String.IsNullOrEmpty(sigEv)) {
                            gEvents.Remove(ev);
                            continue;
                        }

                        if (GoogleOgcs.Calendar.SignaturesMatch(sigEv, sigAi)) {
                            AddGoogleIDs(ref ai, ev);
                            updateCalendarEntry_save(ref ai);
                            unclaimedAi.Remove(ai);
                            Forms.Main.Instance.Console.Update("Reclaimed: " + GetEventSummary(ai), verbose: true);
                            break;
                        }
                    }
                }
            }
            log.Debug(unclaimedAi.Count + " unclaimed.");
            if (unclaimedAi.Count > 0 &&
                (Settings.Instance.SyncDirection == Sync.Direction.GoogleToOutlook ||
                 Settings.Instance.SyncDirection == Sync.Direction.Bidirectional))
            {
                log.Info(unclaimedAi.Count + " unclaimed orphan appointments found.");
                if (Settings.Instance.MergeItems || Settings.Instance.DisableDelete || Settings.Instance.ConfirmOnDelete) {
                    log.Info("These will be kept due to configuration settings.");
                } else if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional) {
                    log.Debug("These 'orphaned' items must not be deleted - they need syncing up.");
                } else {
                    if (MessageBox.Show(unclaimedAi.Count + " Outlook calendar items can't be matched to Google.\r\n" +
                        "Remember, it's recommended to have a dedicated Outlook calendar to sync with, " +
                        "or you may wish to merge with unmatched events. Continue with deletions?",
                        "Delete unmatched Outlook items?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) {

                        log.Info("User has requested to keep them.");
                        foreach (AppointmentItem ai in unclaimedAi) {
                            oAppointments.Remove(ai);
                        }
                    } else {
                        log.Info("User has opted to delete them.");
                    }
                }
            }
        }

        private void createRecipient(EventAttendee gea, ref Recipients recipients) {
            GoogleOgcs.EventAttendee ea = new GoogleOgcs.EventAttendee(gea);
            if (IOutlook.CurrentUserSMTP().ToLower() != ea.Email) {
                Recipient recipient = null;
                try {
                    recipient = recipients.Add(ea.DisplayName + "<" + ea.Email + ">");
                    recipient.Resolve();
                    //ReadOnly: recipient.Type = (int)((bool)ea.Organizer ? OlMeetingRecipientType.olOrganizer : OlMeetingRecipientType.olRequired);
                    recipient.Type = (int)(ea.Optional == null ? OlMeetingRecipientType.olRequired : ((bool)ea.Optional ? OlMeetingRecipientType.olOptional : OlMeetingRecipientType.olRequired));
                    //ReadOnly: ea.ResponseStatus
                } finally {
                    recipient = (Recipient)OutlookOgcs.Calendar.ReleaseObject(recipient);
                }
            }
        }

        /// <summary>
        /// Determine Appointment Item's privacy setting
        /// </summary>
        /// <param name="gVisibility">Google's current setting</param>
        /// <param name="oSensitivity">Outlook's current setting</param>
        private OlSensitivity getPrivacy(String gVisibility, OlSensitivity ?oSensitivity) {
            if (!Settings.Instance.SetEntriesPrivate)
                return (gVisibility == "private") ? OlSensitivity.olPrivate : OlSensitivity.olNormal;

            if (Settings.Instance.SyncDirection != Sync.Direction.Bidirectional) {
                return OlSensitivity.olPrivate;
            } else {
                if (Settings.Instance.TargetCalendar == Sync.Direction.OutlookToGoogle) { //Privacy enforcement is in other direction
                    if (oSensitivity == null)
                        return (gVisibility == "private") ? OlSensitivity.olPrivate : OlSensitivity.olNormal;
                    else if (oSensitivity == OlSensitivity.olPrivate && gVisibility != "private") {
                        log.Fine("Source of truth for enforced privacy is already set private and target is NOT - so syncing this back.");
                        return OlSensitivity.olNormal;
                    } else
                        return (OlSensitivity)oSensitivity;
                } else {
                    if (!Settings.Instance.CreatedItemsOnly || (Settings.Instance.CreatedItemsOnly && oSensitivity == null))
                        return OlSensitivity.olPrivate;
                    else
                        return (gVisibility == "private") ? OlSensitivity.olPrivate : OlSensitivity.olNormal;
                }
            }
        }

        /// <summary>
        /// Determine Appointment's availability setting
        /// </summary>
        /// <param name="gTransparency">Google's current setting</param>
        /// <param name="oBusyStatus">Outlook's current setting</param>
        private OlBusyStatus getAvailability(String gTransparency, OlBusyStatus ?oBusyStatus) {
            if (!Settings.Instance.SetEntriesAvailable)
                return (gTransparency == "transparent") ? OlBusyStatus.olFree : OlBusyStatus.olBusy;

            if (Settings.Instance.SyncDirection != Sync.Direction.Bidirectional) {
                return OlBusyStatus.olFree;
            } else {
                if (Settings.Instance.TargetCalendar == Sync.Direction.OutlookToGoogle) { //Availability enforcement is in other direction
                    if (oBusyStatus == null)
                        return (gTransparency == "transparent") ? OlBusyStatus.olFree : OlBusyStatus.olBusy;
                    else if (oBusyStatus == OlBusyStatus.olFree && gTransparency != "transparent") {
                        log.Fine("Source of truth for Availability is already set available and target is NOT - so syncing this back.");
                        return OlBusyStatus.olBusy;
                    } else
                        return (OlBusyStatus)oBusyStatus;
                } else {
                    if (!Settings.Instance.CreatedItemsOnly || (Settings.Instance.CreatedItemsOnly && oBusyStatus == null))
                        return OlBusyStatus.olFree;
                    else
                        return (gTransparency == "transparent") ? OlBusyStatus.olFree : OlBusyStatus.olBusy;
                }
            }
        }

        #region STATIC functions
        public static void AttachToOutlook(ref Microsoft.Office.Interop.Outlook.Application oApp, Boolean openOutlookOnFail = true, Boolean withSystemCall = false) {
            if (System.Diagnostics.Process.GetProcessesByName("OUTLOOK").Count() > 0) {
                log.Info("Attaching to the already running Outlook process.");
                try {
                    oApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
                } catch (System.Exception ex) {
                    if (OGCSexception.GetErrorCode(ex) == "0x800401E3") { //MK_E_UNAVAILABLE
                        log.Warn("Attachment failed - Outlook is running without GUI for programmatic access.");
                    } else {
                        log.Warn("Attachment failed.");
                        OGCSexception.Analyse(ex);
                    }
                    if (openOutlookOnFail) openOutlookHandler(ref oApp, withSystemCall);
                }
            } else {
                log.Warn("No Outlook process available to attach to.");
                if (openOutlookOnFail) openOutlookHandler(ref oApp, withSystemCall);
            }
        }

        private static void openOutlookHandler(ref Microsoft.Office.Interop.Outlook.Application oApp, Boolean withSystemCall = false) {
            int openAttempts = 1;
            int maxAttempts = 3;
            while (openAttempts <= maxAttempts) {
                try {
                    openOutlook(ref oApp, withSystemCall);
                    openAttempts = maxAttempts;
                } catch (ApplicationException aex) {
                    if (aex.Message == "Outlook is busy.") {
                        log.Warn(aex.Message + " Attempt " + openAttempts + "/" + maxAttempts);
                        if (openAttempts == maxAttempts) {
                            String message = "Outlook has been unresponsive for " + maxAttempts * 10 + " seconds.\n" +
                                "Please try running OGCS again later" +
                                (Settings.Instance.StartOnStartup ? " or " + ((Settings.Instance.StartupDelay == 0) ? "set a" : "increase the") + " delay on startup." : ".");

                            if (aex.InnerException.Message.Contains("CO_E_SERVER_EXEC_FAILURE"))
                                message += "\nAlso check that one of OGCS and Outlook are not running 'as Administrator'.";
                            
                            throw new ApplicationException(message);                            
                        }
                        System.Threading.Thread.Sleep(10000);
                    } else {
                        log.Error("openOutlookHandler: " + aex.Message);
                        throw aex;
                    }
                }
                openAttempts++;
            }
        }
        private static void openOutlook(ref Microsoft.Office.Interop.Outlook.Application oApp, Boolean withSystemCall = false) {
            log.Info("Starting a new instance of Outlook.");
            try {
                if (!withSystemCall)
                    oApp = new Microsoft.Office.Interop.Outlook.Application();
                else {
                    System.Diagnostics.Process oProcess = new System.Diagnostics.Process();
                    oProcess.StartInfo.FileName = "outlook";
                    oProcess.StartInfo.Arguments = "/recycle";
                    oProcess.Start();

                    int maxWaits = 8;
                    while (maxWaits > 0 && oApp == null) {
                        if (maxWaits % 2 == 0) log.Info("Waiting for Outlook to start...");
                        oProcess.WaitForInputIdle(15000);
                        OutlookOgcs.Calendar.AttachToOutlook(ref oApp, openOutlookOnFail: false);
                        if (oApp == null) {
                            log.Debug("Reattempting starting Outlook without system call.");
                            try { oApp = new Microsoft.Office.Interop.Outlook.Application(); } catch (System.Exception ex) { log.Debug("Errored with: " + ex.Message); }
                        }
                        maxWaits--;
                    }
                    if (oApp == null) {
                        log.Error("Giving up waiting for Outlook to open!");
                        throw new System.ApplicationException("Could not establish a connection with Outlook.");
                    }
                }
            } catch (System.Runtime.InteropServices.COMException ex) {
                oApp = null;
                String hResult = OGCSexception.GetErrorCode(ex);

                if (ex.ErrorCode == -2147221164) {
                    OGCSexception.Analyse(ex);
                    throw new ApplicationException("Outlook does not appear to be installed!\nThis is a pre-requisite for this software.");

                } else if (hResult == "0x80010001" && ex.Message.Contains("RPC_E_CALL_REJECTED") ||
                    (hResult == "0x80080005" && ex.Message.Contains("CO_E_SERVER_EXEC_FAILURE")) ||
                    (hResult == "0x800706BA" || hResult == "0x800706BE") ) //Remote Procedure Call failed.
                {
                    log.Warn(ex.Message);
                    throw new ApplicationException("Outlook is busy.", ex);

                } else if (OGCSexception.GetErrorCode(ex, 0x000FFFFF) == "0x00040115") {
                    log.Warn(ex.Message);
                    log.Debug("OGCS is not able to run as Outlook is not properly connected to the Exchange server?");
                    throw new ApplicationException("Outlook is busy.", ex);

                } else if (OGCSexception.GetErrorCode(ex, 0x000FFFFF) == "0x000702E4") {
                    log.Error(ex.Message);
                    throw new ApplicationException("Outlook and OGCS are running in different security elevations.\n" +
                        "Both must be running in Standard or Administrator mode.");

                } else {
                    log.Error("COM Exception encountered.");
                    OGCSexception.Analyse(ex);
                    System.Diagnostics.Process.Start("https://github.com/phw198/OutlookGoogleCalendarSync/wiki/FAQs---COM-Errors");
                    throw new ApplicationException("COM error " + OGCSexception.GetErrorCode(ex) + " encountered.\r\n" +
                        "Please check if there is a published solution on the OGCS wiki.");
                }

            } catch (System.InvalidCastException ex) {
                if (ex.Message.Contains("0x80004002 (E_NOINTERFACE)")) {
                    log.Warn(ex.Message);
                    throw new ApplicationException("A problem was encountered with your Office install.\r\n" +
                        "Please perform an Office Repair and then try running OGCS again.");
                } else if (ex.Message.Contains("0x80040155")) {
                    log.Warn(ex.Message);
                    System.Diagnostics.Process.Start("https://github.com/phw198/OutlookGoogleCalendarSync/wiki/FAQs---COM-Errors#0x80040155---interface-not-registered");
                    throw new ApplicationException("A problem was encountered with your Office install.\r\n" +
                        "Please see the wiki for a solution.");
                } else
                    throw ex;

            } catch (System.Exception ex) {
                log.Warn("Early binding to Outlook appears to have failed.");
                OGCSexception.Analyse(ex, true);
                log.Debug("Could try late binding??");
                //System.Type oAppType = System.Type.GetTypeFromProgID("Outlook.Application");
                //ApplicationClass oAppClass = System.Activator.CreateInstance(oAppType) as ApplicationClass;
                //oApp = oAppClass.CreateObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
                throw ex;
            }
        }

        public static string signature(AppointmentItem ai) {
            return (ai.Subject + ";" + ai.Start.ToPreciseString() + ";" + ai.End.ToPreciseString()).Trim();
        }

        public static void ExportToCSV(String action, String filename, List<AppointmentItem> ais) {
            if (!Settings.Instance.CreateCSVFiles) return;

            log.Debug(action);

            TextWriter tw;
            try {
                tw = new StreamWriter(Path.Combine(Program.UserFilePath, filename));
            } catch (System.Exception ex) {
                Forms.Main.Instance.Console.Update("Failed to create CSV file '" + filename + "'.", Console.Markup.error);
                log.Error("Error opening file '" + filename + "' for writing.");
                OGCSexception.Analyse(ex);
                return;
            }
            try {
                String CSVheader = "Start Time,Finish Time,Subject,Location,Description,Privacy,FreeBusy,";
                CSVheader += "Required Attendees,Optional Attendees,Reminder Set,Reminder Minutes,";
                CSVheader += "Outlook GlobalID,Outlook EntryID,Outlook CalendarID,";
                CSVheader += "Google EventID,Google CalendarID";
                tw.WriteLine(CSVheader);
                foreach (AppointmentItem ai in ais) {
                    try {
                        tw.WriteLine(exportToCSV(ai));
                    } catch (System.Exception ex) {
                        Forms.Main.Instance.Console.Update("Failed to output following Outlook appointment to CSV:-<br/>" + GetEventSummary(ai), Console.Markup.warning);
                        OGCSexception.Analyse(ex);
                    }
                }
            } catch (System.Exception ex) {
                Forms.Main.Instance.Console.Update("Failed to output Outlook events to CSV.", Console.Markup.error);
                OGCSexception.Analyse(ex);
            } finally {
                if (tw != null) tw.Close();
            }
            log.Debug("Done.");
        }
        private static string exportToCSV(AppointmentItem ai) {
            System.Text.StringBuilder csv = new System.Text.StringBuilder();

            csv.Append(ai.Start.ToPreciseString() + ",");
            csv.Append(ai.End.ToPreciseString() + ",");
            csv.Append("\"" + ai.Subject + "\",");

            if (ai.Location == null) csv.Append(",");
            else csv.Append("\"" + ai.Location + "\",");

            if (ai.Body == null) csv.Append(",");
            else {
                String csvBody = ai.Body.Replace("\"", "");
                csvBody = csvBody.Replace("\r\n", " ");
                csv.Append("\"" + csvBody.Substring(0, System.Math.Min(csvBody.Length, 100)) + "\",");
            }

            csv.Append("\"" + ai.Sensitivity.ToString() + "\",");
            csv.Append("\"" + ai.BusyStatus.ToString() + "\",");
            csv.Append("\"" + (ai.RequiredAttendees == null ? "" : ai.RequiredAttendees) + "\",");
            csv.Append("\"" + (ai.OptionalAttendees == null ? "" : ai.OptionalAttendees) + "\",");
            csv.Append(ai.ReminderSet + ",");
            csv.Append(ai.ReminderMinutesBeforeStart.ToString() + ",");
            csv.Append(OutlookOgcs.Calendar.Instance.IOutlook.GetGlobalApptID(ai) + ",");
            csv.Append(ai.EntryID + "," + OutlookOgcs.Calendar.instance.UseOutlookCalendar.EntryID + ",");
            String googleIdValue;
            GetOGCSproperty(ai, MetadataId.gEventID, out googleIdValue); csv.Append((googleIdValue ?? "") + ",");
            GetOGCSproperty(ai, MetadataId.gCalendarId, out googleIdValue); csv.Append(googleIdValue ?? "");

            return csv.ToString();
        }

        public static string GetEventSummary(AppointmentItem ai) {
            String eventSummary = "";
            try {
                if (ai.AllDayEvent) {
                    log.Fine("GetSummary - all day event");
                    eventSummary += ai.Start.Date.ToShortDateString();
                } else {
                    log.Fine("GetSummary - not all day event");
                    eventSummary += ai.Start.ToShortDateString() + " " + ai.Start.ToShortTimeString();
                }
                eventSummary += " " + (ai.IsRecurring ? "(R) " : "") + "=> ";
                eventSummary += '"' + ai.Subject + '"';

            } catch (System.Exception ex) {
                log.Warn("Failed to get appointment summary: " + eventSummary);
                OGCSexception.Analyse(ex, true);
            }
            return eventSummary;
        }

        public static void IdentifyEventDifferences(
            ref List<Event> google,             //need creating
            ref List<AppointmentItem> outlook,  //need deleting
            Dictionary<AppointmentItem, Event> compare) {
            log.Debug("Comparing Google events to Outlook items...");

            // Count backwards so that we can remove found items without affecting the order of remaining items
            String compare_oEventID;
            int metadataEnhanced = 0;
            for (int o = outlook.Count - 1; o >= 0; o--) {
                log.Fine("Checking " + GetEventSummary(outlook[o]));

                if (GetOGCSproperty(outlook[o], MetadataId.gEventID, out compare_oEventID)) {
                    Boolean googleIDmissing = GoogleIdMissing(outlook[o]);

                    for (int g = google.Count - 1; g >= 0; g--) {
                        log.UltraFine("Checking " + GoogleOgcs.Calendar.GetEventSummary(google[g]));

                        if (compare_oEventID == google[g].Id.ToString()) {
                            if (googleIDmissing) {
                                log.Info("Enhancing appointment's metadata...");
                                AppointmentItem ai = outlook[o];
                                AddGoogleIDs(ref ai, google[g]);
                                AddOGCSproperty(ref ai, MetadataId.forceSave, "True");
                                outlook[o] = ai;
                                metadataEnhanced++;
                            }
                            if (ItemIDsMatch(outlook[o], google[g])) {
                                compare.Add(outlook[o], google[g]);
                                outlook.Remove(outlook[o]);
                                google.Remove(google[g]);
                                break;
                            }
                        }
                    }
                } else if (Settings.Instance.MergeItems) {
                    //Remove the non-Google item so it doesn't get deleted
                    outlook.Remove(outlook[o]);
                }
            }
            if (metadataEnhanced > 0) log.Info(metadataEnhanced + " item's metadata enhanced.");

            if (Settings.Instance.DisableDelete) {
                if (outlook.Count > 0)
                    Forms.Main.Instance.Console.Update(outlook.Count + " Outlook items would have been deleted, but you have deletions disabled.", Console.Markup.warning);
                outlook = new List<AppointmentItem>();
            }
            if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional) {
                //Don't recreate any items that have been deleted in Outlook
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (GoogleOgcs.Calendar.GetOGCSproperty(google[g], GoogleOgcs.Calendar.MetadataId.oEntryId))
                        google.Remove(google[g]);
                }
                //Don't delete any items that aren't yet in Google or just created in Google during this sync
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (!GetOGCSproperty(outlook[o], MetadataId.gEventID) ||
                        outlook[o].LastModificationTime > Settings.Instance.LastSyncDate)
                        outlook.Remove(outlook[o]);
                }
            }
            if (Settings.Instance.CreateCSVFiles) {
                ExportToCSV("Appointments for deletion in Outlook", "outlook_delete.csv", outlook);
                GoogleOgcs.Calendar.ExportToCSV("Events for creation in Outlook", "outlook_create.csv", google);
            }
        }

        public static Boolean ItemIDsMatch(AppointmentItem ai, Event ev) {
            //For format of Entry ID : https://msdn.microsoft.com/en-us/library/ee201952(v=exchg.80).aspx
            //For format of Global ID: https://msdn.microsoft.com/en-us/library/ee157690%28v=exchg.80%29.aspx

            String oCompareID;
            log.Fine("Comparing Google Event ID");
            if (GetOGCSproperty(ai, MetadataId.gEventID, out oCompareID) && oCompareID == ev.Id) {
                log.Fine("Comparing Google Calendar ID");
                if (GetOGCSproperty(ai, MetadataId.gCalendarId, out oCompareID) &&
                    oCompareID == Settings.Instance.UseGoogleCalendar.Id) return true;
                else {
                    log.Warn("Could not find Google calendar ID against Outlook appointment item.");
                    return true;
                }
            } else {
                log.Warn("Could not find Google event ID against Outlook appointment item.");
            }
            return false;
        }

        public static object ReleaseObject(object obj) {
            try {
                if (obj != null && System.Runtime.InteropServices.Marshal.IsComObject(obj)) {
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(obj) > 0)
                        System.Windows.Forms.Application.DoEvents();
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex, true);
            }
            GC.Collect();
            return null;
        }

        public Boolean IsOKtoSyncReminder(AppointmentItem ai) {
            if (Settings.Instance.ReminderDND) {
                DateTime alarm;
                if (ai.ReminderSet)
                    alarm = ai.Start.AddMinutes(-ai.ReminderMinutesBeforeStart);
                else {
                    if (Settings.Instance.UseGoogleDefaultReminder && GoogleOgcs.Calendar.Instance.MinDefaultReminder != long.MinValue) {
                        log.Fine("Using default Google reminder value: " + GoogleOgcs.Calendar.Instance.MinDefaultReminder);
                        alarm = ai.Start.AddMinutes(-GoogleOgcs.Calendar.Instance.MinDefaultReminder);
                    } else
                        return false;
                }
                return isOKtoSyncReminder(alarm);
            }
            return true;
        }
        private Boolean isOKtoSyncReminder(DateTime alarm) {
            if (Settings.Instance.ReminderDNDstart.TimeOfDay > Settings.Instance.ReminderDNDend.TimeOfDay) {
                //eg 22:00 to 06:00
                //Make sure end time is the day following the start time
                Settings.Instance.ReminderDNDstart = alarm.Date.AddDays(-1).Add(Settings.Instance.ReminderDNDstart.TimeOfDay);
                Settings.Instance.ReminderDNDend = alarm.Date.Add(Settings.Instance.ReminderDNDend.TimeOfDay);

                if (alarm > Settings.Instance.ReminderDNDstart && alarm < Settings.Instance.ReminderDNDend) {
                    log.Debug("Reminder (@" + alarm.ToString("HH:mm") + ") falls in DND range - not synced.");
                    return false;
                } else
                    return true;

            } else {
                //eg 01:00 to 06:00
                if (alarm.TimeOfDay < Settings.Instance.ReminderDNDstart.TimeOfDay ||
                    alarm.TimeOfDay > Settings.Instance.ReminderDNDend.TimeOfDay) {
                    return true;
                } else {
                    log.Debug("Reminder (@" + alarm.ToString("HH:mm") + ") falls in DND range - not synced.");
                    return false;
                }
            }
        }

        #region OGCS Outlook properties
        public enum MetadataId {
            gEventID,
            gCalendarId,
            ogcsModified,
            forceSave,
            locallyCopied
        }
        public static String MetadataIdKeyName(MetadataId Id) {
            switch (Id) {
                case MetadataId.gEventID: return "googleEventID";
                case MetadataId.gCalendarId: return "googleCalendarID";
                case MetadataId.ogcsModified: return "OGCSmodified";
                case MetadataId.forceSave: return "forceSave";
                default: return Id.ToString();
            }
        }

        public static Boolean GoogleIdMissing(AppointmentItem ai) {
            //Make sure Outlook appointment has all Google IDs stored
            String missingIds = "";
            if (!GetOGCSproperty(ai, MetadataId.gEventID)) missingIds += MetadataIdKeyName(MetadataId.gEventID) + "|";
            if (!GetOGCSproperty(ai, MetadataId.gCalendarId)) missingIds += MetadataIdKeyName(MetadataId.gCalendarId) + "|";
            if (!string.IsNullOrEmpty(missingIds))
                log.Warn("Found Outlook item missing Google IDs (" + missingIds.TrimEnd('|') + "). " + GetEventSummary(ai));
            return !string.IsNullOrEmpty(missingIds);
        }

        public static Boolean HasOgcsProperty(AppointmentItem ai) {
            if (GetOGCSproperty(ai, MetadataId.gEventID)) return true;
            if (GetOGCSproperty(ai, MetadataId.gCalendarId)) return true;
            return false;
        }

        public static void AddGoogleIDs(ref AppointmentItem ai, Event ev) {
            //Add the Google event IDs into Outlook appointment.
            AddOGCSproperty(ref ai, MetadataId.gEventID, ev.Id);
            AddOGCSproperty(ref ai, MetadataId.gCalendarId, Settings.Instance.UseGoogleCalendar.Id);
        }

        public static void AddOGCSproperty(ref AppointmentItem ai, MetadataId key, String value) {
            if (!GetOGCSproperty(ai, key)) {
                try {
                    ai.UserProperties.Add(MetadataIdKeyName(key), OlUserPropertyType.olText);
                } catch (System.Exception ex) {
                    OGCSexception.Analyse(ex);
                    ai.UserProperties.Add(MetadataIdKeyName(key), OlUserPropertyType.olText, false);
                }
            }
            ai.UserProperties[MetadataIdKeyName(key)].Value = value;
        }
        private static void addOGCSproperty(ref AppointmentItem ai, MetadataId key, DateTime value) {
            if (!GetOGCSproperty(ai, key)) {
                try {
                    ai.UserProperties.Add(MetadataIdKeyName(key), OlUserPropertyType.olDateTime);
                } catch (System.Exception ex) {
                    OGCSexception.Analyse(ex);
                    ai.UserProperties.Add(MetadataIdKeyName(key), OlUserPropertyType.olDateTime, false);
                }
            }
            ai.UserProperties[MetadataIdKeyName(key)].Value = value;
        }

        public static Boolean GetOGCSproperty(AppointmentItem ai, MetadataId key) {
            String throwAway;
            return GetOGCSproperty(ai, key, out throwAway);
        }
        public static Boolean GetOGCSproperty(AppointmentItem ai, MetadataId key, out String value) {
            UserProperty prop = ai.UserProperties.Find(MetadataIdKeyName(key));
            if (prop == null) {
                value = null;
                return false;
            } else {
                value = prop.Value.ToString();
                return true;
            }
        }
        private static Boolean getOGCSproperty(AppointmentItem ai, MetadataId key, out DateTime value) {
            UserProperty prop = ai.UserProperties.Find(MetadataIdKeyName(key));
            if (prop == null) {
                value = new DateTime();
                return false;
            } else {
                value = (DateTime)prop.Value;
                return true;
            }
        }

        public static void RemoveOGCSproperties(ref AppointmentItem ai) {
            removeOGCSproperty(ref ai, MetadataId.gEventID);
            removeOGCSproperty(ref ai, MetadataId.gCalendarId);
        }
        private static void removeOGCSproperty(ref AppointmentItem ai, MetadataId key) {
            if (GetOGCSproperty(ai, key)) {
                UserProperty prop = ai.UserProperties.Find(MetadataIdKeyName(key));
                prop.Delete();
                log.Debug("Removed " + MetadataIdKeyName(key) + " property.");
            }
        }

        public static DateTime GetOGCSlastModified(AppointmentItem ai) {
            DateTime lastModded;
            getOGCSproperty(ai, MetadataId.ogcsModified, out lastModded);
            return lastModded;
        }
        private static void setOGCSlastModified(ref AppointmentItem ai) {
            addOGCSproperty(ref ai, MetadataId.ogcsModified, DateTime.Now);
        }

        public static String GetOGCSEntryID(AppointmentItem ai) {
            String entryID = ai.EntryID;

            if (!Settings.Instance.EnableUseRecurrence) {
                entryID += " " + ai.Start.ToShortDateString() + " " + ai.Start.ToShortTimeString();
                entryID += " " + ai.End.ToShortDateString() + " " + ai.End.ToShortTimeString();
            }

            return entryID;
                
        }
        #endregion
        #endregion

    }
}
