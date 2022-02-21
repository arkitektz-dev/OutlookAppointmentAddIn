using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAppointment.AppointmentMonitor
{
    public class EventArgs<T> : EventArgs
    {
        private T _value;

        public EventArgs(T aValue)
        {
            _value = aValue;
        }

        public T Value
        {
            get { return _value; }
            set { _value = value; }
        }
    }
     
}
