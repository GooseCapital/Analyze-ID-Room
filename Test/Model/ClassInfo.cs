using System;
using System.Collections.Generic;
using System.Data;
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
        public string DurationTime { get; private set; }

        public ClassInfo(DataRow dataTableRow, List<OnlineClassInfo> onlineClass)
        {
            Day = Convert.ToInt32(dataTableRow.ItemArray[0].ToString());
            CourseCode = dataTableRow.ItemArray[1].ToString();
            CourseName = dataTableRow.ItemArray[3].ToString();
            LectureCode = dataTableRow.ItemArray[4].ToString();
            Seat = dataTableRow.ItemArray[6].ToString();
            TeacherName = dataTableRow.ItemArray[7].ToString();
            Lession = dataTableRow.ItemArray[8].ToString();
            RoomName = dataTableRow.ItemArray[9].ToString();
            RawDateTime = dataTableRow.ItemArray[10].ToString();
            StartTime = DateTime.ParseExact(dataTableRow.ItemArray[10].ToString().Split('-')[0],
                "dd/MM/yyyy", CultureInfo.InvariantCulture);
            EndTime = DateTime.ParseExact(dataTableRow.ItemArray[10].ToString().Split('-')[1],
                "dd/MM/yyyy", CultureInfo.InvariantCulture);
            try
            {
                OnlineId = onlineClass.FirstOrDefault(n => n.RoomName == dataTableRow.ItemArray[9].ToString()).RoomId;
            }
            catch
            {
                OnlineId = -1;
            }
            Details =
                $@"Giảng viên: {dataTableRow.ItemArray[7].ToString()}. Lớp học phần: {dataTableRow.ItemArray[4].ToString()}";
        }

        public void EncodeDurationTime(List<TimerPeriod> timerList)
        {
            string[] timeArray = Lession.Split(',');
            DurationTime =
                $@"{timerList[Convert.ToInt32(timeArray[0]) - 1].StartTime} - {timerList[Convert.ToInt32(timeArray[timeArray.Length - 1]) - 1].EndTime}";
        }
    }
}
