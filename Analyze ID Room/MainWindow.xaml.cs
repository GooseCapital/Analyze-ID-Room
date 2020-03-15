using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using Analyze_ID_Room.Model;
using ExcelDataReader;
using Microsoft.Win32;
using xNet;

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
        private HttpRequest _httpRequest = new HttpRequest();

        private static readonly string[] ExampleObjects = new string[11]
        {
            "Thứ", "Mã học phần", "", "Tên học phần", "Lớp học phần", "", "Ghế", "CBGD", "Tiết học", "Phòng học",
            "Thời gian học"
        };

        public static string GetMD5Hash(string str)
        {
            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
            byte[] bytes = ASCIIEncoding.Default.GetBytes(str);
            byte[] encoded = md5.ComputeHash(bytes);

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < encoded.Length; i++)
                sb.Append(encoded[i].ToString("x2"));

            return sb.ToString();
        }

        private bool Login(string username, string password)
        {
            _httpRequest = new HttpRequest();
            CookieDictionary cookieDictionary = new CookieDictionary();
            // {
            //     {"__dtsu", "6BB6E72D5DF8DE9E18CD3F45102ACD02"},
            //     {"_ga", "GA1.3.1169250952.1583144203" },
            //     {"ASP.NET_SessionId", "45c5rck10lqiq2xejdinyqyn" },
            //     {"SL_GWPT_Show_Hide_tmp", "1" },
            //     {"SL_wptGlobTipTmp", "1"}
            // };
            _httpRequest.Cookies = cookieDictionary;
            _httpRequest.Get("http://dkh.tlu.edu.vn/CMCSoft.IU.Web.Info/Login.aspx");
            var reqPrams = new RequestParams();
            reqPrams["__EVENTTARGET"] = "";
            // reqPrams["__EVENTARGUMENT"] = "";
            // reqPrams["__LASTFOCUS"] = "";
            // reqPrams["__VIEWSTATEGENERATOR"] = "D620498B";
            // reqPrams["PageHeader1$drpNgonNgu"] = "2D8E12D246D64A7BB3B89CC6484D2776";
            // reqPrams["PageHeader1$hidisNotify"] = "0";
            // reqPrams["PageHeader1$hidValueNotify"] = ".";
            reqPrams["btnSubmit"] = "Đăng+nhập";
            // reqPrams["hidUserId"] = "";
            // reqPrams["hidUserFullName"] = "";
            // reqPrams["hidTrainingSystemId"] = "";
            reqPrams["txtUserName"] = username;
            reqPrams["txtPassword"] = GetMD5Hash(password);
            // reqPrams["__EVENTVALIDATION"] = "/wEdAA90tNej9blFYHJDs+U5oBbCb8csnTIorMPSfpUKU79Fa8zr1tijm/dVbgMI0MJ/5MhJxi/pyypRM6i+Ay0tyJXpHW/isUyw6w8trNAGHDe5T/w2lIs9E7eeV2CwsZKam8yG9tEt/TDyJa1fzAdIcnRuY3plgk0YBAefRz3MyBlTcHY2+Mc6SrnAqio3oCKbxYY85pbWlDO2hADfoPXD/5tdAxTm4XTnH1XBeB1RAJ3owlx3skko0mmpwNmsvoT+s7J0y/1mTDOpNgKEQo+otMEzMS21+fhYdbX7HjGORawQMqpdNpKktwtkFUYS71DzGv66E0F2ty4ucCjkSoWNZgQd08YQgWdRT/H0yF4ypQtQjQ==";
            // reqPrams["__VIEWSTATE"] = "/wEPDwUKMTkwNDg4MTQ5MQ9kFgICAQ9kFgpmD2QWCgIBDw8WAh4EVGV4dAU5SOG7hiBUSOG7kE5HIMSQxIJORyBLw50gSOG7jEMgLSDEkOG6oEkgSOG7jEMgVEjhu6ZZIEzhu6JJZGQCAg9kFgJmDw8WBB8ABQ3EkMSDbmcgbmjhuq1wHhBDYXVzZXNWYWxpZGF0aW9uaGRkAgMPEA8WBh4NRGF0YVRleHRGaWVsZAUGa3loaWV1Hg5EYXRhVmFsdWVGaWVsZAUCSUQeC18hRGF0YUJvdW5kZ2QQFQECVk4VASAyRDhFMTJEMjQ2RDY0QTdCQjNCODlDQzY0ODREMjc3NhQrAwFnFgFmZAIEDw8WAh4ISW1hZ2VVcmwFKC9DTUNTb2Z0LklVLldlYi5JbmZvL0ltYWdlcy9Vc2VySW5mby5naWZkZAIFD2QWBgIBDw8WAh8ABQZLaMOhY2hkZAIDDw8WAh8AZWRkAgcPDxYCHgdWaXNpYmxlaGRkAgIPZBYEAgMPD2QWAh4Gb25ibHVyBQptZDUodGhpcyk7ZAIHDw8WAh8AZWRkAgQPDxYCHwZoZGQCBg8PFgIfBmhkFgYCAQ8PZBYCHwcFCm1kNSh0aGlzKTtkAgUPD2QWAh8HBQptZDUodGhpcyk7ZAIJDw9kFgIfBwUKbWQ1KHRoaXMpO2QCCw9kFghmDw8WAh8ABQwwMjQuMzg1MjE0NDFkZAIBD2QWAmYPDxYCHwFoZGQCAg9kFgJmDw8WBB8ABQ3EkMSDbmcgbmjhuq1wHwFoZGQCAw8PFgIfAAW0BTxhIGhyZWY9IiMiIG9uY2xpY2s9ImphdmFzY3JpcHQ6d2luZG93LnByaW50KCkiPjxkaXYgc3R5bGU9IkZMT0FUOmxlZnQiPgk8aW1nIHNyYz0iL0NNQ1NvZnQuSVUuV2ViLkluZm8vaW1hZ2VzL3ByaW50LnBuZyIgYm9yZGVyPSIwIj48L2Rpdj48ZGl2IHN0eWxlPSJGTE9BVDpsZWZ0O1BBRERJTkctVE9QOjZweCI+SW4gdHJhbmcgbsOgeTwvZGl2PjwvYT48YSBocmVmPSJtYWlsdG86P3N1YmplY3Q9SGUgdGhvbmcgdGhvbmcgdGluIElVJmFtcDtib2R5PWh0dHA6Ly9ka2gudGx1LmVkdS52bi9DTUNTb2Z0LklVLldlYi5JbmZvL0xvZ2luLmFzcHgiPjxkaXYgc3R5bGU9IkZMT0FUOmxlZnQiPjxpbWcgc3JjPSIvQ01DU29mdC5JVS5XZWIuSW5mby9pbWFnZXMvc2VuZGVtYWlsLnBuZyIgIGJvcmRlcj0iMCI+PC9kaXY+PGRpdiBzdHlsZT0iRkxPQVQ6bGVmdDtQQURESU5HLVRPUDo2cHgiPkfhu61pIGVtYWlsIHRyYW5nIG7DoHk8L2Rpdj48L2E+PGEgaHJlZj0iIyIgb25jbGljaz0iamF2YXNjcmlwdDphZGRmYXYoKSI+PGRpdiBzdHlsZT0iRkxPQVQ6bGVmdCI+PGltZyBzcmM9Ii9DTUNTb2Z0LklVLldlYi5JbmZvL2ltYWdlcy9hZGR0b2Zhdm9yaXRlcy5wbmciICBib3JkZXI9IjAiPjwvZGl2PjxkaXYgc3R5bGU9IkZMT0FUOmxlZnQ7UEFERElORy1UT1A6NnB4Ij5UaMOqbSB2w6BvIMawYSB0aMOtY2g8L2Rpdj48L2E+ZGRkbiplXuTlfiOl6mrLaLXQzmg+SN2Z85mFV4xgTKvmLQs=";
            // _webClient.Headers["Connection"] = "keep-alive";
            _httpRequest.AddHeader("Cache-Control", "max-age=0");
            _httpRequest.AddHeader("Origin", "http://dkh.tlu.edu.vn");
            _httpRequest.AddHeader("Upgrade-Insecure-Requests", "1");
            // _httpRequest.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            _httpRequest.AllowAutoRedirect = true;
            _httpRequest.KeepAlive = true;
            _httpRequest.Referer = "http://dkh.tlu.edu.vn/CMCSoft.IU.Web.Info/Login.aspx";
            _httpRequest.UserAgent =
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36";
            _httpRequest.AddHeader("Accept-Language", "en-US,en;q=0.9,vi;q=0.8");
            _httpRequest.AddHeader("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9");

            string content = _httpRequest.Post("http://dkh.tlu.edu.vn/CMCSoft.IU.Web.Info/Login.aspx", reqPrams)
                .ToString();
            return content.Contains(username);
        }

        private void DownloadScheduleFile()
        {
            var reqPrams = new RequestParams();
            reqPrams["__EVENTTARGET"] = "";
            reqPrams["__EVENTARGUMENT"] = "";
            reqPrams["__LASTFOCUS"] = "";
            reqPrams["__VIEWSTATEGENERATOR"] = "A0622FE5";
            reqPrams["PageHeader1$drpNgonNgu"] = "2D8E12D246D64A7BB3B89CC6484D2776";
            reqPrams["PageHeader1$hidisNotify"] = "0";
            reqPrams["PageHeader1$hidValueNotify"] = ".";
            reqPrams["drpSemester"] = "5ad062230c5c4973a1bd241ed876d6fb";
            reqPrams["drpTerm"] = "1";
            reqPrams["hidDiscountFactor"] = "";
            reqPrams["hidReduceTuitionType"] = "";
            reqPrams["hidXetHeSoHocPhiTheoDoiTuong"] = "N1";
            reqPrams["hidTuitionFactorMode"] = "2";
            reqPrams["hidLoaiUuTienHeSoHocPhi"] = "1";
            reqPrams["hidSecondFieldId"] = "";
            reqPrams["hidSecondAyId"] = "";
            reqPrams["hidSecondFacultyId"] = "";
            reqPrams["hidSecondAdminClassId"] = "";
            reqPrams["hidSecondFieldLevel"] = "";
            reqPrams["hidXetHeSoHocPhiDoiTuongTheoNganh"] = "";
            reqPrams["hidUnitPriceDetail"] = "1";
            reqPrams["hidFacultyId"] = "702AF27DB0D248EF835549965FD42EA9";
            reqPrams["hidFieldLevel"] = "MF";
            reqPrams["hidPrintLocationByCode"] = "0";
            reqPrams["hidUnitPriceKP"] = "0";
            reqPrams["drpType"] = "B";
            reqPrams["btnView"] = "Xuất+file+Excel";
            reqPrams["hidStudentId"] = "519a4a0f54bc407ba97cd6f1822d14f0";
            reqPrams["hidAcademicYearId"] = "F732C737EB6E49A8824BC5725BC035E6";
            reqPrams["hidFieldId"] = "F9FB5998C9AB4E47B9CA44617E4A8C3F";
            reqPrams["hidSemester"] = "";
            reqPrams["hidTerm"] = "";
            reqPrams["hidShowTeacher"] = "";
            reqPrams["hidUnitPrice"] = "";
            reqPrams["hidTuitionCalculating"] = "1";
            reqPrams["hidTrainingSystemId"] = "dda1424e45dd4676ad5fb25d33a31b43";
            reqPrams["hidAdminClassId"] = "7273A45232744142A556E7DA35D1B11E";
            reqPrams["hidTuitionStudentType"] = "1";
            reqPrams["hidStudentTypeId"] = "1";
            reqPrams["hidUntuitionStudentTypeId"] = "1";
            reqPrams["hidUniversityCode"] = "THUYLOI";
            reqPrams["hidTuanBatDauHK2"] = "";
            reqPrams["hidSoTietSang"] = "";
            reqPrams["hidSoTietChieu"] = "";
            reqPrams["hidSoTietToi"] = "";
            reqPrams["__EVENTVALIDATION"] = "/wEdAFZB/z3iSzo44vytjx+zJvMcb8csnTIorMPSfpUKU79Fa8zr1tijm/dVbgMI0MJ/5MhJxi/pyypRM6i+Ay0tyJXpHW/isUyw6w8trNAGHDe5T/w2lIs9E7eeV2CwsZKam8yG9tEt/TDyJa1fzAdIcnRuq940A0sVd2nflhG7GplI5+8XeUh3gRTV1fmhPau35QRJEm+/71JNhmPUTYDle8ZO7TFUfuayrl2bG1x4vP8V54TyUfJH0FlNJdsjU+SfhUKFrhSym0kpHdm3aCioRhHRrFz7X5A42bQsMkJUVcrLSXUals9EQIVT4Vt1R2RjkROSGh/ELDAQiaS6T50iYENaAzk9/Geq/sZ1KX6XFCkRS44y60/oE5uSLvWsRgLhnjNhQvKmsy75GEzFEYmNq2/iQc35Hltk4PvMuMIHHlQ0tJbosCYu6/9rByw74uMzxNrrfXccQwymcDOo6uDm0AChYF/s32XUp730vb66Z3UPK09QEm7Br10Kdpp9zVr2HD/fcNpATQjkTQGC6O3p7uQAnnOB8+1TN12OqHVW4+d21vYLQldpo0WlnpvFOln321M18UkHKEulXwtAzgDJ7P8wRsL+UkwV30r4Ps+OUUVf63EpHOPNdnsbTTPvtG2LUSevhB3CTxvO1QKtGNjN1Ei/4X8C8PuLuSE41Ubm5XoyWP3ZXtFnpNG5xXzmu9m2/8sDzpMzsnfHoMKXR+quG2FTKpIfjgv9t9HrKfZMt+m0MAZUgidnen0rtIRR2U+izgjPX8NVc8o5AzUkW6rx9C2qkNR5I3cGvZSFIHpMebZ75gMEYFjANUUE0FzGlBwEbwLdY2RVuKY+RsZTEykD4prdWcm1AXsQKwELQQHR+HzShR00HoZWsYIOJdy0kIb/ha9OoFYGBLuwYLUC7fRFpDkMRfRkxiAdhHjSpoo9EiNqGRzqAy9mzO4eRO9tmdX+AOmhXdwedMRtjEkwMRERRfLa4O6Dcre4aeWV+mEzVuXJcuF7gKJV61Lfrm36r1Hvy1kqnlZ4lHGzXC/AkArxwcc4RWqMDlYgtCimcY6dcwMnTxb1gr3l/Js8zwiiia0zg8NyPyBhbHismGnStK0wEuMVhKeIa91xirYo4z0DsK4P0qAp69YzfQfcO/Oudp8GofsK+a0kU/Aut3Ztpw5yu5NL8q17YyrNeSSBueLyFv6EHiEViRAZVjwEnURcEuZXhJpt2fskQ8Pr8dJAusLWaNv5TTSU2gFdhUDvE0V7MYqBZFbffkBC6cdPR1wIypQV9bVDdnqFTWd9VddHmeUMv5uo/vhCyiaQJyvmpq37xyq0L/donmu4HLHhM6vfj2yKHfEdRv0W6jo/2pr3hHTpucxjp68JyG36llXJXQhzTDs0AH4vupOzk0lkmj6auLolpNJwzGL/84LP5nmY4Whj6i0bd35OcDyJUMOWUqtE15SJOtj/k5FcHkMmGlxA52K+Sba3T2JsLCkAg/HQm1Xzn5c96j1DFJwuPBalDwDViVsHMfiMTZJQPiz+KQtXXcwptCRKXm3XhFQLxFSz0cGbIcskIhzAoGZPTQXL+UawAdRqOXTL/WZMM6k2AoRCj6i0wTPS4i0yr4VuPXQcphLUg/BObWY0FgCqMPJvilc8zUnQhacX23vL9Y+10bhu3VvMRCizUlxfEqJUrU03FTtSqh/lh1I303i5Iru9PEXv2mO93/NkNDUKFaaRFASh0Ssc7a/k9C+S22gH6cBFgiPuyGy3I7/ZboXgCQ8Gz+Eb7vUtYJ7QClToIvGluZt4u5AeblsxLbX5+Fh1tfseMY5FrBAyql02kqS3C2QVRhLvUPMa/j4KgbflRoYpNrZLbEzgtIihhPu7PUz0UmXkRmgDoR8o";
            reqPrams["__VIEWSTATE"] = "/wEPDwUKMTUwOTAzNzUwNA9kFgICAQ9kFiQCAQ9kFgwCAQ8PFgIeBFRleHQFOUjhu4YgVEjhu5BORyDEkMSCTkcgS8OdIEjhu4xDIC0gxJDhuqBJIEjhu4xDIFRI4bumWSBM4buiSWRkAgIPZBYCZg8PFgQfAAUGVGhvw6F0HhBDYXVzZXNWYWxpZGF0aW9uaGRkAgMPEA8WBh4NRGF0YVRleHRGaWVsZAUGa3loaWV1Hg5EYXRhVmFsdWVGaWVsZAUCSUQeC18hRGF0YUJvdW5kZ2QQFQECVk4VASAyRDhFMTJEMjQ2RDY0QTdCQjNCODlDQzY0ODREMjc3NhQrAwFnFgFmZAIEDw8WAh4ISW1hZ2VVcmwFKC9DTUNTb2Z0LklVLldlYi5JbmZvL0ltYWdlcy9Vc2VySW5mby5naWZkZAIFD2QWCAIBDw8WAh8ABRzDgnUgVGnhur9uIETFqW5nKDE2NTExMjI1NjgpZGQCBQ8PFgIfAAUKU2luaCB2acOqbmRkAgcPDxYCHgdWaXNpYmxlaGRkAgkPDxYCHwAFEEjhu5lwIHRpbiBuaOG6r25kZAIHDw8WAh8ABZkBxJDEg25nIGvDvSBo4buNYyA8c3BhbiBzdHlsZT0iZm9udC1zaXplOjEwcHgiPj48L3NwYW4+IDxhIGhyZWY9Ii9DTUNTb2Z0LklVLldlYi5JbmZvL1JlcG9ydHMvRm9ybS9TdHVkZW50VGltZVRhYmxlLmFzcHgiPkvhur90IHF14bqjIMSRxINuZyBrw70gaOG7jWM8L2E+ZGQCBA9kFgJmD2QWAgIBDw8WAh8ABRRLP3QgcXU/IGRhbmcga8O9IGg/Y2RkAgYPZBYCZg8PFgIfAWhkZAIID2QWAmYPDxYEHwAFBlRob8OhdB8BaGRkAgoPZBYCZg8WAh4JaW5uZXJodG1sBesNPHVsIGNsYXNzPSdzaWRlYmFyLW1lbnUnPjxsaSBjbGFzcz0naGVhZGVyJz5EQU5IIE3hu6RDIENIw41OSDwvbGk+PGxpIGNsYXNzPSd0cmVldmlldyc+IDxhIGhyZWY9Ii9DTUNTb2Z0LklVLldlYi5JbmZvL1Byb2ZpbGUvTGlua1VuaS5hc3B4Ij48c3Bhbj5D4buVbmcgS+G6v3QgYuG6oW4gU2luaCB2acOqbiwgU2nDqnUgdGnhu4duIMOtY2ghPC9zcGFuPjwvYT48L2xpPjxsaT4gPGE+PGkgY2xhc3M9J2ZhIGZhLWxhcHRvcCc+PC9pPjxzcGFuPsSQxINuZyBrw70gaOG7jWM8L3NwYW4+PGkgY2xhc3M9J2ZhIGZhLWFuZ2xlLWxlZnQgcHVsbC1yaWdodCc+PC9pPjwvYT48dWwgY2xhc3M9J3RyZWV2aWV3LW1lbnUnPjxsaSBjbGFzcz0ndHJlZXZpZXcnPiA8YSBocmVmPSIvQ01DU29mdC5JVS5XZWIuaW5mby9TdHVkeVJlZ2lzdGVyL1N0dWR5UmVnaXN0ZXIuYXNweCI+U2luaCB2acOqbiDEkcSDbmcga8O9IGjhu41jPC9hPjwvbGk+PGxpIGNsYXNzPSd0cmVldmlldyc+IDxhIGhyZWY9Ii9DTUNTb2Z0LklVLldlYi5JbmZvL1JlcG9ydHMvRm9ybS9TdHVkZW50VGltZVRhYmxlLmFzcHgiPkvhur90IHF14bqjIMSRxINuZyBrw70gaOG7jWM8L2E+PC9saT48bGkgY2xhc3M9J3RyZWV2aWV3Jz4gPGEgaHJlZj0iL0NNQ1NvZnQuSVUuV2ViLkluZm8vU3R1ZHlSZWdpc3Rlci9SZWdpc3RyYXRpb25IaXN0b3J5LmFzcHgiPlF1w6EgdHLDrG5oIMSRxINuZyBrw70gaOG7jWM8L2E+PC9saT48bGkgY2xhc3M9J3RyZWV2aWV3Jz4gPGEgaHJlZj0iL0NNQ1NvZnQuSVUuV2ViLkluZm8vU3R1ZHlSZWdpc3Rlci9EZW1hbmRSZWdpc3Rlci5hc3B4Ij7EkMSDbmcga8O9IG5ndXnhu4duIHbhu41uZzwvYT48L2xpPjwvdWw+PC9saT48bGkgY2xhc3M9J3RyZWV2aWV3Jz4gPGEgaHJlZj0iL0NNQ1NPRlQuSVUuV0VCLklORk8vTWFya0FuZFZpZXcuYXNweCI+PHNwYW4+VHJhIGPhu6l1IMSRaeG7g20gdOG7lW5nIGjhu6NwPC9zcGFuPjwvYT48L2xpPjxsaSBjbGFzcz0ndHJlZXZpZXcnPiA8YSBocmVmPSIvQ01DU29mdC5JVS5XZWIuaW5mby9TdHVkZW50TWFyay5hc3B4Ij48c3Bhbj5UcmEgY+G7qXUgxJFp4buDbSBo4buNYyB04bqtcDwvc3Bhbj48L2E+PC9saT48bGkgY2xhc3M9J3RyZWV2aWV3Jz4gPGEgaHJlZj0iL0NNQ1NvZnQuSVUuV2ViLkluZm8vU3R1ZGVudFNlcnZpY2UvUHJhY3Rpc2VNYXJrQW5kU3R1ZHlXYXJuaW5nLmFzcHgiPjxzcGFuPlLDqG4gbHV54buHbiwgaOG7jWMgduG7pSB2w6AgdOG7kXQgbmdoaeG7h3A8L3NwYW4+PC9hPjwvbGk+PGxpIGNsYXNzPSd0cmVldmlldyc+IDxhIGhyZWY9Ii9DTUNTb2Z0LklVLldlYi5JbmZvL1N0dWRlbnRTZXJ2aWNlL1N0dWRlbnRUdWl0aW9uLmFzcHgiPjxzcGFuPlRyYSBj4bupdSBo4buNYyBwaMOtPC9zcGFuPjwvYT48L2xpPjxsaSBjbGFzcz0ndHJlZXZpZXcnPiA8YSBocmVmPSIvQ01DU29mdC5JVS5XZWIuSW5mby9Db3Vyc2VCeUZpZWxkVHJlZS5hc3B4Ij48c3Bhbj5DaMawxqFuZyB0csOsbmggxJHDoG8gdOG6oW88L3NwYW4+PC9hPjwvbGk+PGxpIGNsYXNzPSd0cmVldmlldyc+IDxhIGhyZWY9Ii9DTUNTb2Z0LklVLldlYi5pbmZvL0NoYW5nZVBhc3NXb3JkU3R1ZGVudC5hc3B4Ij48c3Bhbj7EkOG7lWkgbeG6rXQga2jhuql1PC9zcGFuPjwvYT48L2xpPjxsaSBjbGFzcz0ndHJlZXZpZXcnPiA8YSBocmVmPSIvY21jc29mdC5pdS53ZWIuaW5mby9TdHVkZW50Vmlld0V4YW1MaXN0LmFzcHgiPjxzcGFuPlhlbSBs4buLY2ggdGhpIGPDoSBuaMOibjwvc3Bhbj48L2E+PC9saT48L3VsPmQCDA8QDxYGHwIFCFNlbWVzdGVyHwMFAklkHwRnZBAVGwsxXzIwMjBfMjAyMQsyXzIwMTlfMjAyMAsxXzIwMTlfMjAyMAsyXzIwMThfMjAxOQsxXzIwMThfMjAxOQsyXzIwMTdfMjAxOAsxXzIwMTdfMjAxOAsyXzIwMTZfMjAxNwsxXzIwMTZfMjAxNwsyXzIwMTVfMjAxNgsxXzIwMTVfMjAxNgsyXzIwMTRfMjAxNQsxXzIwMTRfMjAxNQsyXzIwMTNfMjAxNAsxXzIwMTNfMjAxNAsyXzIwMTJfMjAxMwsxXzIwMTJfMjAxMwsyXzIwMTFfMjAxMgsxXzIwMTFfMjAxMgsyXzIwMTBfMjAxMQsxXzIwMTBfMjAxMQsyXzIwMDlfMjAxMAsxXzIwMDlfMjAxMAsyXzIwMDhfMjAwOQsxXzIwMDhfMjAwOQsyXzIwMDdfMjAwOAsxXzIwMDdfMjAwOBUbIDIyMGI4OTdmM2JiODQ0ZWI4NzZjZTc2ZGEwYWRhMDY5IDVhZDA2MjIzMGM1YzQ5NzNhMWJkMjQxZWQ4NzZkNmZiIGJmMjk4YjQ3MjJjODQxMzhiNGRlYTBmNDk4ZThiYzU5IDM0MDU0NTFmZDQ4MjQ0NmE5NmJhYWFlNDIwNjBhNjg5IDVmZTcxOWJlNzAyZTQ3MmI4YjAxOTdjMTdkZDRkNzcwIDY4NDgyRUM3NUZCOTQ2RDA4ODM3QTI4OThDNUM1ODg0IDRhYTFiYjY5MzQwMjQ4OWJhOTdhMTJmZWJmNmFjNDhlIGUwYWIyNjJmMWMwYzQ4Y2Q5ZDg1YzJjODMzYWFhNTVjIDk1ZGNiYTQwYjkwYjQzZDI4YWVjYzY3M2VjZDQ3NTJiIDNCOERGQ0RFMzZEODRFNkQ4RjE5RkI4RjJGREVFMTRDIDMyODQ5MDVmZmQyYTQyMjdiMzA3OGM5Zjc0YTRkNGMxIDJFNDc5OTRCNzdFMzQ5MENBQ0I5MDAwRUQxNThBOTRGIGU4ZjgyMzAzNGUxZTQyZTg4MzBjNDBiNjYxYTk4NmQ5IEFCMjBEMUE5NzVFODQ5NDA4ODA5QUE3NzJEN0UzRDBDIGE1MzlhOGI2OWJiOTRjNTc4ZWMxNmFlNDU2NjgxMzczIEE4Nzc4M0YzNzMzODQwNENBMUFFRjY1NkIwNTE5MTgyIDdiNDJkN2Q0MjEzNTQzNTQ5NGI4NzhkMWI1Mjg5OWFkIDgwMzdmMzUwNzU1MTQ5YjJhZmQzZjBiYTE1NjYxYzA3IGNkOWY3MjUwNzVmMzQ2ZDViNzJhODljYjZmMzRkZjRmIEJDQ0VFRTE3NDUwRDQ2NTFBMTZFMUY0QUM4Q0E1QzEyIDVhYjRlMjQzOWY4OTRmNzY4YTNmYTAxZTBhZTU4YzgzIGVjYjYxYjQ5NzE4MjRmZGY5YmZkM2U3M2MwMjFmMWFhIGI0OGY2MjhiYjdmZTQ5MDhiZGE3MWRjMmJmZTM3Mzg4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aG/DoXQfAWhkZAIDDw8WAh8ABcwFPGEgaHJlZj0iIyIgb25jbGljaz0iamF2YXNjcmlwdDp3aW5kb3cucHJpbnQoKSI+PGRpdiBzdHlsZT0iRkxPQVQ6bGVmdCI+CTxpbWcgc3JjPSIvQ01DU29mdC5JVS5XZWIuSW5mby9pbWFnZXMvcHJpbnQucG5nIiBib3JkZXI9IjAiPjwvZGl2PjxkaXYgc3R5bGU9IkZMT0FUOmxlZnQ7UEFERElORy1UT1A6NnB4Ij5JbiB0cmFuZyBuw6B5PC9kaXY+PC9hPjxhIGhyZWY9Im1haWx0bzo/c3ViamVjdD1IZSB0aG9uZyB0aG9uZyB0aW4gSVUmYW1wO2JvZHk9aHR0cDovL2RraC50bHUuZWR1LnZuL0NNQ1NvZnQuSVUuV2ViLkluZm8vUmVwb3J0cy9Gb3JtL1N0dWRlbnRUaW1lVGFibGUuYXNweCI+PGRpdiBzdHlsZT0iRkxPQVQ6bGVmdCI+PGltZyBzcmM9Ii9DTUNTb2Z0LklVLldlYi5JbmZvL2ltYWdlcy9zZW5kZW1haWwucG5nIiAgYm9yZGVyPSIwIj48L2Rpdj48ZGl2IHN0eWxlPSJGTE9BVDpsZWZ0O1BBRERJTkctVE9QOjZweCI+R+G7rWkgZW1haWwgdHJhbmcgbsOgeTwvZGl2PjwvYT48YSBocmVmPSIjIiBvbmNsaWNrPSJqYXZhc2NyaXB0OmFkZGZhdigpIj48ZGl2IHN0eWxlPSJGTE9BVDpsZWZ0Ij48aW1nIHNyYz0iL0NNQ1NvZnQuSVUuV2ViLkluZm8vaW1hZ2VzL2FkZHRvZmF2b3JpdGVzLnBuZyIgIGJvcmRlcj0iMCI+PC9kaXY+PGRpdiBzdHlsZT0iRkxPQVQ6bGVmdDtQQURESU5HLVRPUDo2cHgiPlRow6ptIHbDoG8gxrBhIHRow61jaDwvZGl2PjwvYT5kZGSN0WY8wEKGA9EvyCfc26Lc7ZbvpWRH3+O6b/NO5xYQLg==";

            _httpRequest.Post("http://dkh.tlu.edu.vn/CMCSoft.IU.Web.Info/Reports/Form/StudentTimeTable.aspx", reqPrams).ToFile("1.xls");
        }

        private void BtnInputFile_Click(object sender, RoutedEventArgs e)
        {
            // bool isLoggedIn = Login("1651122568", "03/08/1998");
            // DownloadScheduleFile();
            if (!File.Exists("ID.xlsx"))
            {
                MessageBox.Show("Thiếu file ID.xlsx", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            MessageBox.Show("Chọn file excel sắp xếp theo thứ học", "Thông báo", MessageBoxButton.OK,
                MessageBoxImage.Asterisk);
            OpenFileDialog open = new OpenFileDialog();
            open.Multiselect = false;
            open.Title = "Chọn file Excel";
            open.DefaultExt = "*.xls";
            if (open.ShowDialog() == true)
            {
                string fileName = open.FileName;
                List<OnlineClassInfo> onlineClassData = ParseOnlineClass("ID.xlsx");
                if (GetScheduleData(fileName, onlineClassData))
                {
                    ListClassInfo.Sort((a, b) => a.StartTime.CompareTo(b.StartTime));
                    DgClassInfo.ItemsSource = ListClassInfo;
                }
                else
                {
                    MessageBox.Show("Có lỗi xảy ra", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
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
            DataTable dataTable = ReadExcelFile(fileName);
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
