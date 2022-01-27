using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Newtonsoft.Json;
using System.Net.Http;
using System.Diagnostics;
using System.IO;
using Microsoft.Office.Interop.Outlook;

namespace OutlookAppointment
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;
        Outlook.CalendarModule Calendar;
        static HttpClient client = new HttpClient();
        string workingDirectory = Environment.CurrentDirectory;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
              

            //using (StreamReader r = new StreamReader($"../../appsettings.json"))
            //{
            //    string json = r.ReadToEnd();
            //    AppConfig items = JsonConvert.DeserializeObject<AppConfig>(json);
            //}

            Outlook.Folder calFolder =
            Application.Session.GetDefaultFolder(
            Outlook.OlDefaultFolders.olFolderCalendar)
            as Outlook.Folder;
            DateTime start = DateTime.Now.AddDays(-30);
            DateTime end = start.AddDays(60);
            Outlook.Items rangeAppts = GetAppointmentsInRange(calFolder, start, end);
            if (rangeAppts != null)
            {
                Debug.WriteLine("All appointment are here");
                foreach (Outlook.AppointmentItem appt in rangeAppts)
                {


                    List<string> userList = appt.RequiredAttendees.Split(';').ToList();
                    string CompanyName = "";

                    for (int i = 1; i < appt.Recipients.Count; i++)
                    {
                        string email = GetEmailAddressOfAttendee(appt.Recipients[i]);
                        string splittedValue = email.Split('@').ToList()[1];
                        CompanyName = splittedValue.Split('.').ToList()[0];
                    }

                    foreach (var trackUser in userList.Skip(1))
                    {

                        var row = new Appointment()
                        {
                            CompanyName = CompanyName,
                            FullName = trackUser,
                            MeetingPurpose = 1,
                            VisitingEmployee = appt.Organizer,
                            CheckIn = appt.Start,
                            MeetingDescription = appt.Subject
                        };

                        AddAppointment(row);
                    }

 

                   

                    Debug.WriteLine("Subject: " + appt.Companies
                        + " Start: " + appt.Start.ToString("g"));
                }
            }
            else
            {
                Debug.WriteLine("No appointmnet  are here");
            }

        }

        string GetEmailAddressOfAttendee(Recipient TheRecipient)
        {

            // See http://msdn.microsoft.com/en-us/library/cc513843%28v=office.12%29.aspx#AddressBooksAndRecipients_TheRecipientsCollection
            // for more info

            string PROPERTY_TAG_SMTP_ADDRESS = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

            if (TheRecipient.Type == (int)Outlook.OlMailRecipientType.olTo)
            {
                PropertyAccessor pa = TheRecipient.PropertyAccessor;
                return pa.GetProperty(PROPERTY_TAG_SMTP_ADDRESS);
            }
            return null;
        }

        public static object GetProperty(object target, string name)
        {
            var site = System.Runtime.CompilerServices.CallSite<Func<System.Runtime.CompilerServices.CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.GetMember(0, name, target.GetType(), new[] { Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo.Create(0, null) }));
            return site.Target(site, target);
        }

        private Outlook.Items GetAppointmentsInRange(Outlook.Folder folder, DateTime startTime, DateTime endTime)
        {
            string filter = "[Start] >= '"
                + startTime.ToString("g")
                + "' AND [End] <= '"
                + endTime.ToString("g") + "'";
            Debug.WriteLine(filter);
            try
            {
                Outlook.Items calItems = folder.Items;
                calItems.IncludeRecurrences = true;
                calItems.Sort("[Start]", Type.Missing);
                Outlook.Items restrictItems = calItems.Restrict(filter);
                if (restrictItems.Count > 0)
                {
                    return restrictItems;
                }
                else
                {
                    return null;
                }
            }
            catch { return null; }
        }


        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "This text was added by using code";
                    mailItem.Body = "This text was added by using code";


                }

            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        void AddAppointment(Appointment param)
        {
            string output = JsonConvert.SerializeObject(param);


            client.PostAsync("https://localhost:44308/api/visitor/add-appointment", new StringContent(output, Encoding.UTF8, "application/json"));
        }
    }

    public class AppConfig
    { 
        public string Tenant { get; set; }
    }

    public class Appointment
    {
        public string FullName { get; set; }
        public string CompanyName { get; set; }
        public int MeetingPurpose { get; set; }
        public string VisitingEmployee { get; set; }
        public DateTime CheckIn { get; set; }
        public string MeetingDescription { get; set; }

    }
}
