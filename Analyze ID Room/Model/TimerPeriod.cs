using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Analyze_ID_Room.Model
{
    public class TimerPeriod
    {
        public string StartTime { get; set; }
        public string EndTime { get; set; }

        public TimerPeriod(string start, string end)
        {
            StartTime = start;
            EndTime = end;
        }
    }
}
