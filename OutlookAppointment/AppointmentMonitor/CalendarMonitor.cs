using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAppointment.AppointmentMonitor
{
    public class CalendarMonitor
    {
        private NameSpace _session;
        private List<string> _folderPaths;
        private List<MAPIFolder> _calendarFolders;
        private List<Items> _calendarItems;
        private MAPIFolder _deletedItemsFolder;

        public event EventHandler<EventArgs<AppointmentItem>> AppointmentAdded;
        public event EventHandler<EventArgs<AppointmentItem>> AppointmentModified;
        public event EventHandler<CancelEventArgs<AppointmentItem>> AppointmentDeleting;

        public CalendarMonitor(NameSpace aSession)
        {
            _folderPaths = new List<string>();
            _calendarFolders = new List<MAPIFolder>();
            _calendarItems = new List<Items>();

            _session = aSession;
            _deletedItemsFolder = aSession.GetDefaultFolder(
                OlDefaultFolders.olFolderDeletedItems);

            HookupDefaultCalendarEvents();
        }

        private void HookupDefaultCalendarEvents()
        {
            MAPIFolder folder = _session.GetDefaultFolder(
                OlDefaultFolders.olFolderCalendar);
            if (folder != null)
            {
                HookupCalendarEvents(folder);
            }
        }

        private void HookupCalendarEvents(MAPIFolder aCalendarFolder)
        {
            if (aCalendarFolder.DefaultItemType != OlItemType.olAppointmentItem)
            {
                throw new ArgumentException("The MAPIFolder must use " +
                    "AppointmentItems as the default type.");
            }

            if (_folderPaths.Contains(aCalendarFolder.FolderPath) == false)
            {
                Items items = aCalendarFolder.Items;
                //
                // Store folder path to prevent double ups on our listeners.
                //
                _folderPaths.Add(aCalendarFolder.FolderPath);
                //
                // Store a reference to the folder and to the items collection
                // so that it remains alive for as long as we want. This keeps
                // the ref count up on the underlying COM object and prevents
                // it from being intermittently released (then the events don't
                // get fired).
                //
                _calendarFolders.Add(aCalendarFolder);
                _calendarItems.Add(items);
                //
                // Add listeners for the events we need.
                //
                ((MAPIFolderEvents_12_Event)aCalendarFolder).BeforeItemMove +=
                    new MAPIFolderEvents_12_BeforeItemMoveEventHandler(Calendar_BeforeItemMove);
                items.ItemChange +=
                    new ItemsEvents_ItemChangeEventHandler(CalendarItems_ItemChange);
                items.ItemAdd +=
                    new ItemsEvents_ItemAddEventHandler(CalendarItems_ItemAdd);
            }
        }

        private void CalendarItems_ItemAdd(object anItem)
        {
            if (anItem is AppointmentItem)
            {
                OutlookAppointment.Global.OutlookState.appointmentState = (anItem as AppointmentItem);

                if (this.AppointmentAdded != null)
                {
                    this.AppointmentAdded(this,
                        new EventArgs<AppointmentItem>((AppointmentItem)anItem));
                }
            }


        }

        private void CalendarItems_ItemChange(object anItem)
        {
            if (anItem is AppointmentItem)
            {
                if (this.AppointmentModified != null)
                {
                    this.AppointmentModified(this,
                        new EventArgs<AppointmentItem>((AppointmentItem)anItem));
                }
            }
        }

        private void Calendar_BeforeItemMove(object anItem, MAPIFolder aMoveToFolder,
            ref bool Cancel)
        {
            if ((aMoveToFolder == null) || (IsDeletedItemsFolder(aMoveToFolder)))
            {
                if (anItem is AppointmentItem)
                {
                    if (this.AppointmentDeleting != null)
                    {
                        //
                        // Listeners to the AppointmentDeleting event can cancel
                        // the move operation if moving to the deleted items folder.
                        //
                        CancelEventArgs<AppointmentItem> args =
                            new CancelEventArgs<AppointmentItem>((AppointmentItem)anItem);
                        this.AppointmentDeleting(this, args);
                        Cancel = args.Cancel;
                    }
                }
            }
        }

        private bool IsDeletedItemsFolder(MAPIFolder aFolder)
        {
            return (aFolder.EntryID == _deletedItemsFolder.EntryID);
        }
    }



 
}
