using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LotusLyncer
{
    class CalendarEvent
    {
        public DateTime Starts { get; set; }
        public DateTime Ends { get; set; }
        public string Title { get; set; }
        public string Location { get; set; }
    }
}
