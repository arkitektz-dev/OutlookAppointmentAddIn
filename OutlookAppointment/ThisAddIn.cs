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
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Net.Http.Headers;
using OutlookAppointment.Model;
using System.Drawing;
using System.Drawing.Imaging;
using OutlookAppointment.AppointmentMonitor;

namespace OutlookAppointment
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;
        Outlook.CalendarModule Calendar;
        static HttpClient client = new HttpClient();
        string workingDirectory = Environment.CurrentDirectory;
        static int TenenatId = 1;
        static string baseUrl = "https://localhost:44308/";
        static string uploadButton = "Start";
       

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            CalendarMonitor monitor = new CalendarMonitor(this.Application.Session);
            monitor.AppointmentAdded +=
                new EventHandler<EventArgs<AppointmentItem>>(monitor_AppointmentAdded);
            monitor.AppointmentModified +=
                new EventHandler<EventArgs<AppointmentItem>>(monitor_AppointmentModified);
            monitor.AppointmentDeleting +=
                new EventHandler<CancelEventArgs<AppointmentItem>>(monitor_AppointmentDeleting);
        }


        private void monitor_AppointmentAdded(object sender, EventArgs<Outlook.AppointmentItem> e)
        {
            Debug.Print("Appointment Added: {0}", e.Value.GlobalAppointmentID);
            Outlook.AppointmentItem appt = (AppointmentItem)e;
            UploadAppointment();
        }

        private void monitor_AppointmentModified(object sender, EventArgs<AppointmentItem> e)
        {
            Debug.Print("Appointment Modified: {0}", e.Value.GlobalAppointmentID);
        }

        private void monitor_AppointmentDeleting(object sender, CancelEventArgs<AppointmentItem> e)
        {
            Debug.Print("Appointment Deleting: {0}", e.Value.GlobalAppointmentID);
            DialogResult dr = MessageBox.Show("Delete appointment?", "Confirm",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.No)
            {
                e.Cancel = true;
            }
        }


        public  void UploadAppointment() {
             
                var task = Task.Run(() => GetStartTimeAsync(TenenatId));
                task.Wait();
                DateTime startTime = task.Result;
             
                Outlook.Folder calFolder =
                Application.Session.GetDefaultFolder(
                Outlook.OlDefaultFolders.olFolderCalendar)
                as Outlook.Folder;
                DateTime start = startTime;
                DateTime end = start.AddDays(12);
                Outlook.Items rangeAppts = GetAppointmentsInRange(calFolder, start, end);
                if (rangeAppts != null)
                {
                    Debug.WriteLine("All appointment are here");
                    foreach (Outlook.AppointmentItem appt in rangeAppts)
                    {

                       
                        List<string> userList = appt.RequiredAttendees.Split(';').ToList();
                        List<AttendeeDetail> attendee = new List<AttendeeDetail>();
                        
                        string CompanyName = "";

                        for (int i = 1; i < appt.Recipients.Count; i++)
                        {
                            string email = GetEmailAddressOfAttendee(appt.Recipients[i]);
                            string splittedValue = email.Split('@').ToList()[1];
                            CompanyName = splittedValue.Split('.').ToList()[0];
                            
                           

                            attendee.Add(new AttendeeDetail()
                            {
                                Email = email,
                                Name = userList[i - 1], 
                            });

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
                                MeetingDescription = appt.Subject,
                                GlobalAppointmentId = appt.GlobalAppointmentID
                            };

                            var task2 = Task.Run(() => AddAppointment(row));
                            task2.Wait();


                               AppointmentSaveDto result = task2.Result;

                                if (result.Id != null && result.Id != 0) {
                                    var getAttendeeEmail = attendee.Where(x => x.Name.Trim().ToLower() == trackUser.Trim().ToLower()).Select(x => x.Email).FirstOrDefault();
                                    if (getAttendeeEmail != null)
                                    {
                                         CreateEmailItem("Appointment", getAttendeeEmail, $"Please use the following barcode to check in if then enter this {result.Id}", result.Id);
                                    }
                                }

                   
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

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new MyRibbon();
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

        async Task<AppointmentSaveDto> AddAppointment(Appointment param)
        {
            string output = JsonConvert.SerializeObject(param);


            var result = await client.PostAsync($"{baseUrl}api/visitor/add-appointment", new StringContent(output, Encoding.UTF8, "application/json"));
            var response = await result.Content.ReadAsStringAsync();
            AppointmentSaveDto myDeserializedClass = JsonConvert.DeserializeObject<AppointmentSaveDto>(response);
            return myDeserializedClass;
        }

        async Task<DateTime> GetStartTimeAsync(int TenantId) {
             
            HttpResponseMessage response = await client.GetAsync($"{baseUrl}api/visitor/outlook-tenant-intitaltime?tenantId={TenantId}");
            if (response.IsSuccessStatusCode)
            {
                DateTime intitalTime = await response.Content.ReadAsAsync<DateTime>();
                return intitalTime;
            }

            return DateTime.Now.Date;  
        }

        private void CreateEmailItem(string subjectEmail,
       string toEmail, string bodyEmail,int? number)
        {
            BarcodeLib.Barcode b = new BarcodeLib.Barcode();
            Image img = b.Encode(BarcodeLib.TYPE.CODE39, number.ToString(), Color.Black, Color.White, 290, 120);
            Bitmap bImage = (Bitmap)img;  // Your Bitmap Image
            System.IO.MemoryStream ms = new MemoryStream();
            bImage.Save(ms, ImageFormat.Jpeg);
            byte[] byteImage = ms.ToArray();
            var SigBase64 = Convert.ToBase64String(byteImage);

            Outlook.MailItem eMail = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);
            eMail.Subject = subjectEmail;
            eMail.To = toEmail;
            //eMail.Body = bodyEmail; 
            eMail.HTMLBody = $"<p>{bodyEmail}</p></br></br></br><img src='data:image/png;base64, {SigBase64}' />";
            eMail.Importance = Outlook.OlImportance.olImportanceLow;
            ((Outlook._MailItem)eMail).Send();
        }

    }

    public class AttendeeDetail
    { 
        public string Email { get; set; }
        public string Name { get; set; }
        public string GlobalAppointmentId { get; set; }
    }

    public class AppConfig
    { 
        public string Tenant { get; set; }
    }

    public class IntitalTime
    { 
        public DateTime StartTime { get; set; }
    }

    public class Appointment
    {
        public string GlobalAppointmentId { get; set; }
        public string FullName { get; set; }
        public string CompanyName { get; set; }
        public int MeetingPurpose { get; set; }
        public string VisitingEmployee { get; set; }
        public DateTime CheckIn { get; set; }
        public string MeetingDescription { get; set; }

    }
}
