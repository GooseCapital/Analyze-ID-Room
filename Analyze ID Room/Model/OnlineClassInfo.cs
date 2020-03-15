using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Analyze_ID_Room.Model
{
    public class OnlineClassInfo
    {
        private string roomName = "";
        public string RoomName
        {
            get => roomName;
            set
            {
                roomName = value.Replace("-", " ");
            }
        }
        public long RoomId { get; set; }
        public string Building { get; set; }
        public string Location { get; set; }
    }
}
