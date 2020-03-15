using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security;
using System.Windows;
using Analyze_ID_Room.Model;
using ExcelDataReader;

namespace Analyze_ID_Room
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public List<ClassInfo> ListClassInfo = new List<ClassInfo>();

        private static readonly string[] ExampleObjects = new string[11]
        {
            "Thứ", "Mã học phần", "", "Tên học phần", "Lớp học phần", "", "Ghế", "CBGD", "Tiết học", "Phòng học",
            "Thời gian học"
        };

        private void BtnInputFile_Click(object sender, RoutedEventArgs e)
        {
            if (!File.Exists("ID.xlsx"))
            {
                MessageBox.Show("Thiếu file ID.xlsx", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            List<OnlineClassInfo> onlineClassData = ParseOnlineClass("ID.xlsx");
            if (GetScheduleData(@"C:\Users\GooseCapital\Downloads\ThoiKhoaBieuSinhVien (4).xls", onlineClassData))
            {
                ListClassInfo.Sort((a, b) => a.StartTime.CompareTo(b.StartTime));
                DgClassInfo.ItemsSource = ListClassInfo;
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private DataTable ReadExcelFile(string fileName)
        {
            try
            {
                using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
                {
                    IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                    var conf = new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true
                        }
                    };

                    var dataSet = reader.AsDataSet(conf);

                    // Now you can get data from each sheet by its index or its "name"
                    return dataSet.Tables[0];
                }
            }
            catch (Exception e)
            {
                MessageBox.Show($@"{e.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        private List<OnlineClassInfo> ParseOnlineClass(string fileName)
        {
            List<OnlineClassInfo> onlineClass = new List<OnlineClassInfo>();
            DataTable dataTable = ReadExcelFile("ID.xlsx");
            foreach (DataRow dataTableRow in dataTable.Rows)
            {
                try
                {
                    if (int.TryParse(dataTableRow.ItemArray[0].ToString(), out _))
                    {
                        onlineClass.Add(new OnlineClassInfo()
                        {
                            RoomName = dataTableRow.ItemArray[1].ToString(),
                            Building = dataTableRow.ItemArray[2].ToString(),
                            RoomId = Convert.ToInt64(dataTableRow.ItemArray[3].ToString()),
                            Location = dataTableRow.ItemArray[4].ToString()
                        });
                    }
                }
                catch (Exception e)
                {
                    File.AppendAllText("Exception.txt", $@"{e.Message} ----- ParseOnlineClass{Environment.NewLine}");
                    continue;
                }
            }

            return onlineClass;
        }

        private List<ClassInfo> AddClassInfoToList(DataTable dataTable, List<OnlineClassInfo> onlineClass)
        {
            List<ClassInfo> classList = new List<ClassInfo>();
            foreach (DataRow dataTableRow in dataTable.Rows)
            {
                try
                {
                    if (int.TryParse(dataTableRow.ItemArray[0].ToString(), out _))
                    {
                        classList.Add(new ClassInfo()
                        {
                            Day = Convert.ToInt32(dataTableRow.ItemArray[0].ToString()),
                            CourseCode = dataTableRow.ItemArray[1].ToString(),
                            CourseName = dataTableRow.ItemArray[3].ToString(),
                            LectureCode = dataTableRow.ItemArray[4].ToString(),
                            Seat = dataTableRow.ItemArray[6].ToString(),
                            TeacherName = dataTableRow.ItemArray[7].ToString(),
                            Lession = dataTableRow.ItemArray[8].ToString(),
                            RoomName = dataTableRow.ItemArray[9].ToString(),
                            RawDateTime = dataTableRow.ItemArray[10].ToString(),
                            StartTime = DateTime.ParseExact(dataTableRow.ItemArray[10].ToString().Split('-')[0],
                                "dd/MM/yyyy", CultureInfo.InvariantCulture),
                            EndTime = DateTime.ParseExact(dataTableRow.ItemArray[10].ToString().Split('-')[1],
                                "dd/MM/yyyy", CultureInfo.InvariantCulture),
                            OnlineId = onlineClass.FirstOrDefault(n => n.RoomName == dataTableRow.ItemArray[9].ToString()).RoomId,
                            Details = $@"Giảng viên: {dataTableRow.ItemArray[7].ToString()}. Lớp học phần: {dataTableRow.ItemArray[4].ToString()}"
                        });
                    }
                }
                catch (Exception e)
                {
                    File.AppendAllText("Exception.txt", $@"{e.Message} ----- ReadExcelFile{Environment.NewLine}");
                    continue;
                }
            }

            return classList;
        }

        private bool GetScheduleData(string fileName, List<OnlineClassInfo> onlineClass)
        {
            DataTable dataTable = ReadExcelFile(@"C:\Users\GooseCapital\Downloads\ThoiKhoaBieuSinhVien (4).xls");
            string[] data = dataTable.Rows[8].ItemArray.Select(x => x.ToString()).ToArray();

            if (data.SequenceEqual(ExampleObjects))
            {
                ListClassInfo = AddClassInfoToList(dataTable, onlineClass);
                return true;
            }

            return false;
        }

        private void BtnLinkOnline_OnClick(object sender, RoutedEventArgs e)
        {
            ClassInfo obj = ((FrameworkElement)sender).DataContext as ClassInfo;
            Process.Start($@"https://zoom.us/j/{obj.OnlineId}?status=success");
        }
    }
}
