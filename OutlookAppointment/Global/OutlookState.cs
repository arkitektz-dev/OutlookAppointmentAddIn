using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAppointment.Global
{
    public static class OutlookState
    {
        public static bool isPressed = false;
        public  static AppointmentItem appointmentState;
    }
}
