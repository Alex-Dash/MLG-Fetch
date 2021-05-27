using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace MLG_Fetch
{

    class Responder
    {
        public bool success { get; set; }
    }

    class Reporter
    {
        //Internal vars block
        private string excel_ver, word_ver;

        private bool WordInstalled()
        {
            Type Excel_check = Type.GetTypeFromProgID("Word.Application");
            if (Excel_check == null)
            {
                return false;
            }
            else
            {
                try
                {
                    Microsoft.Office.Interop.Word._Application base1 = new Microsoft.Office.Interop.Word.Application();
                    word_ver = base1.Version + " | " + base1.System.OperatingSystem;
                    base1.Quit();
                }
                catch
                {
                    word_ver = "Detected. Failed to open app";
                }
                
                return true;
            }
        }

        private bool ExcelInstalled()
        {
            Type Word_check = Type.GetTypeFromProgID("Excel.Application");
            if (Word_check == null)
            {
                return false;
            }
            else
            {
                try
                {
                    Microsoft.Office.Interop.Excel._Application base1 = new Microsoft.Office.Interop.Excel.Application();
                    excel_ver = base1.Version + " | " + base1.OperatingSystem;
                    base1.Quit();
                }
                catch
                {
                    excel_ver = "Detected. Failed to open app";
                }
                return true;
            }
        }


        //System info block
        public string Version {
            get { return Properties.Settings.Default.build_type+Properties.Settings.Default.version; }
        }

        public string Name
        {
            get { return System.Windows.Forms.SystemInformation.UserName+ " on "+ System.Windows.Forms.SystemInformation.ComputerName; }
        }

        public string SysLocale
        {
            get { return Convert.ToString(CultureInfo.InstalledUICulture.TextInfo); }
        }

        public string ConnectionPresent
        {
            get { return System.Windows.Forms.SystemInformation.Network.ToString(); }
        }

        public string Date
        {
            get { return DateTime.UtcNow.Day.ToString()+"/"+DateTime.UtcNow.Month.ToString()+ "/" + DateTime.UtcNow.Year.ToString()+" | "+DateTime.UtcNow.Hour.ToString()+":"+ DateTime.UtcNow.Minute.ToString()+":"+ DateTime.UtcNow.Second.ToString() + ":" + DateTime.UtcNow.Millisecond.ToString(); }
        }

        public string ExcelStatus
        {
            get
            {
                if (ExcelInstalled())
                {
                    return excel_ver;
                }
                else
                {
                    return "Not detected";
                }
            }
        }

        public string WordStatus
        {
            get
            {
                if (WordInstalled())
                {
                    return word_ver;
                }
                else
                {
                    return "Not detected";
                }
            }
        }


        //Crash site block
        public string EventType { get; set; }

        public string ReportType { get; set; }

        public string Stage { get; set; }

        public string ExceptionDescription { get; set; }


    }

    class ReportSender
    {

        //Crash action block
        public bool SendReport(Reporter data)
        {
            //stop if concent is not given
            if (!Properties.Settings.Default.report_concent)
            {
                return true;
            }
            Responder res = new Responder();
            try
            {
                //get reporting Url
                string url = Properties.Settings.Default.report_url;
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(data);

                //create request
                var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Method = "POST";

                //Write stream
                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                {
                    streamWriter.Write(json);
                }

                //Get and validate response
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                    res = Newtonsoft.Json.JsonConvert.DeserializeObject<Responder>(result);
                }
            }
            catch
            {
                return false;
            }

            //Check if the data is recieved properly
            if (res.success)
            {
                return true;
            }

            //fail otherwise
            return false;
        }

    }
}
