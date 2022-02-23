using Microsoft.Office.Interop.Outlook;
using OutlookAppointment.AppointmentMonitor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAppointment.Helper
{
    public class CalendarMonitor
    {
        private Explorer _explorer;
        private List<string> _folderPaths;
        private List<MAPIFolder> _calendarFolders;
        private List<Items> _calendarItems;
        private MAPIFolder _deletedItemsFolder;

        public event EventHandler<EventArgs<AppointmentItem>> AppointmentAdded;
        public event EventHandler<EventArgs<AppointmentItem>> AppointmentModified;
        public event EventHandler<CancelEventArgs<AppointmentItem>> AppointmentDeleting;

        public CalendarMonitor(Explorer anExplorer)
        {
            _folderPaths = new List<string>();
            _calendarFolders = new List<MAPIFolder>();
            _calendarItems = new List<Items>();

            _explorer = anExplorer;
            _explorer.BeforeFolderSwitch +=
              new ExplorerEvents_10_BeforeFolderSwitchEventHandler(Explorer_BeforeFolderSwitch);

            NameSpace session = _explorer.Session;
            try
            {
                _deletedItemsFolder = session.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
                HookupDefaultCalendarEvents(session);
            }
            finally
            {
                Marshal.ReleaseComObject(session);
                session = null;
            }
        }

        private void HookupDefaultCalendarEvents(NameSpace aSession)
        {
            MAPIFolder folder = aSession.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            if (folder != null)
            {
                try
                {
                    HookupCalendarEvents(folder);
                }
                finally
                {
                    Marshal.ReleaseComObject(folder);
                    folder = null;
                }
            }
        }

        private void Explorer_BeforeFolderSwitch(object aNewFolder, ref bool Cancel)
        {
            MAPIFolder folder = (aNewFolder as MAPIFolder);
            //
            // Hookup events to any other Calendar folder opened.
            //
            if (folder != null)
            {
                try
                {
                    if (folder.DefaultItemType == OlItemType.olAppointmentItem)
                    {
                        HookupCalendarEvents(folder);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(folder);
                    folder = null;
                }
            }
        }

        private void HookupCalendarEvents(MAPIFolder aCalendarFolder)
        {
            if (aCalendarFolder.DefaultItemType != OlItemType.olAppointmentItem)
            {
                throw new ArgumentException("The MAPIFolder must use " +
                  "AppointmentItems as the default type.");
            }
            //
            // Ignore other user's calendars.
            //
            if ((_folderPaths.Contains(aCalendarFolder.FolderPath) == false)
              && (IsUsersCalendar(aCalendarFolder)))
            {
                Items items = aCalendarFolder.Items;
                //
                // Store folder path to prevent double ups on our listeners.
                //
                _folderPaths.Add(aCalendarFolder.FolderPath);
                //
                // Store a reference to the folder and to the items collection so that it remains alive for
                // as long as we want. This keeps the ref count up on the underlying COM object and prevents
                // it from being intermittently released (then the events don't get fired).
                //
                _calendarFolders.Add(aCalendarFolder);
                _calendarItems.Add(items);
                //
                // Add listeners for the events we need.
                //
                ((MAPIFolderEvents_12_Event)aCalendarFolder).BeforeItemMove +=
                  new MAPIFolderEvents_12_BeforeItemMoveEventHandler(Calendar_BeforeItemMove);
                items.ItemChange += new ItemsEvents_ItemChangeEventHandler(CalendarItems_ItemChange);
                items.ItemAdd += new ItemsEvents_ItemAddEventHandler(CalendarItems_ItemAdd);
            }
        }

        private void CalendarItems_ItemAdd(object anItem)
        {
            AppointmentItem appointment = (anItem as AppointmentItem);
            if (appointment != null)
            {
                try
                {
                    if (this.AppointmentAdded != null)
                    {
                        this.AppointmentAdded(this, new EventArgs<AppointmentItem>(appointment));
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(appointment);
                    appointment = null;
                }
            }
        }

        private void CalendarItems_ItemChange(object anItem)
        {
            AppointmentItem appointment = (anItem as AppointmentItem);
            if (appointment != null)
            {
                try
                {
                    if (this.AppointmentModified != null)
                    {
                        this.AppointmentModified(this, new EventArgs<AppointmentItem>(appointment));
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(appointment);
                    appointment = null;
                }
            }
        }

        private void Calendar_BeforeItemMove(object anItem, MAPIFolder aMoveToFolder, ref bool Cancel)
        {
            if ((aMoveToFolder == null) || (IsDeletedItemsFolder(aMoveToFolder)))
            {
                AppointmentItem appointment = (anItem as AppointmentItem);
                if (appointment != null)
                {
                    try
                    {
                        if (this.AppointmentDeleting != null)
                        {
                            //
                            // Listeners to the AppointmentDeleting event can cancel the move operation if moving
                            // to the deleted items folder.
                            //
                            CancelEventArgs<AppointmentItem> args = new CancelEventArgs<AppointmentItem>(appointment);
                            this.AppointmentDeleting(this, args);
                            Cancel = args.Cancel;
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(appointment);
                        appointment = null;
                    }
                }
            }
        }

        private bool IsUsersCalendar(MAPIFolder aFolder)
        {
            //
            // This is based purely on my observations so far - a better way?
            //
            return (aFolder.Store != null);
        }

        private bool IsDeletedItemsFolder(MAPIFolder aFolder)
        {
            return (aFolder.EntryID == _deletedItemsFolder.EntryID);
        }
    }
}
