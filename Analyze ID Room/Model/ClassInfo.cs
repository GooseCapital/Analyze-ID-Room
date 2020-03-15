using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Analyze_ID_Room.Model
{
    public class ClassInfo
    {
        public long OnlineId { get; set; }
        public int Day { get; set; }
        public string LectureCode { get; set; }
        public string CourseName { get; set; }
        public string CourseCode { get; set; }
        public string Seat { get; set; }
        public string TeacherName { get; set; }
        public string Lession { get; set; }
        public string RoomName { get; set; }
        public string RawDateTime { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public string Details { get; set; }
    }
}
