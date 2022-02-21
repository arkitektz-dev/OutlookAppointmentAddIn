using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAppointment.AppointmentMonitor
{
    public class CancelEventArgs<T> : EventArgs<T>
    {
        private bool _cancel;

        public CancelEventArgs(T aValue)
            : base(aValue)
        {
        }

        public bool Cancel
        {
            get { return _cancel; }
            set { _cancel = value; }
        }
    }
     
}
